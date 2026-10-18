#!/usr/bin/env bash
# ==============================================================================
#  jboss_pool_collector.sh
#    JBoss EAP 8.1 データソース (IronJacamar) コネクションプール統計 収集ツール
#
#  概要
#    ECS タスク上の UBI 9.x (9.8) コンテナで稼働する JBoss EAP 8.1 から、
#    jboss-cli (既定) または JMX でデータソースのプール統計を定期取得し、
#    レポート用の CSV を出力します。
#      - 使用件数の時系列推移 / 時間別推移 / 最大・最小・平均・95%値
#      - コネクションプール枯渇 (取得タイムアウト・待機) の検出とイベント化
#      - 接続待ちスレッド (dump-queued-threads-in-pool) の集計
#      - プール設定値 / JDBC(PreparedStatement キャッシュ) / XA 統計
#      - (任意) 枯渇検出時のスレッドダンプ取得
#
#  動作要件
#    bash 4.4 以降, coreutils, awk(gawk), sed, grep, java(EAP 同梱)
#    ※ jq / python / procps 等の追加パッケージは不要です (ubi9-minimal でも動作)
#
#  使い方
#    bash jboss_pool_collector.sh -h
# ==============================================================================

set -u
set -o pipefail
export LC_ALL=C

readonly TOOL_VERSION="1.0.0"
readonly SCRIPT_PATH="$0"
ORIG_ARGS=("$@")

# ------------------------------------------------------------------------------
# 既定値 (環境変数でも指定可能)
# ------------------------------------------------------------------------------
MODE="${POOL_MODE:-cli}"                                      # cli | jmx
JBOSS_HOME="${JBOSS_HOME:-}"
CONTROLLER="${JBOSS_CONTROLLER:-}"                            # 空 = 稼働中プロセスから自動検出
CLI_USER="${JBOSS_CLI_USER:-}"
CLI_PASSWORD="${JBOSS_CLI_PASSWORD:-}"
INTERVAL=10                                                   # 取得間隔(秒)
DURATION=3600                                                 # 収集時間(秒) 0=停止まで
COUNT=""                                                      # 取得回数 (指定時は DURATION より優先)
OUT_DIR=""
DS_FILTER=""                                                  # 対象DS (カンマ区切り, 空=全て)
BUCKET_MIN=60                                                 # 時間別集計の単位(分)
WARN_PCT=80                                                   # 高負荷判定の使用率(%)
REPORT_TZ="${REPORT_TZ:-Asia/Tokyo}"
ENABLE_STATS=1                                                # 統計が無効なら実行時に有効化
RESTORE_STATS=1                                               # 終了時に有効化した統計を元に戻す
QUEUED_DUMP=1                                                 # 接続待ちスレッドを取得
THREAD_DUMP=0                                                 # 枯渇/待機検出時にスレッドダンプ取得
THREAD_DUMP_GAP=300                                           # スレッドダンプの最小間隔(秒)
REDISCOVER_SEC=600                                            # データソース再検出間隔(秒)
SUMMARY_EVERY=30                                              # サマリ再生成間隔(サンプル数)
WRITE_BOM=1                                                   # CSV に UTF-8 BOM を付与 (Excel 直接表示用)
BACKGROUND=0
SUMMARIZE_DIR=""
HOST_LABEL=""
CLI_TIMEOUT=60                                                # jboss-cli 1回あたりのタイムアウト(秒)
CLI_JAVA_OPTS="${POOL_CLI_JAVA_OPTS:--Xms16m -Xmx128m -XX:+UseSerialGC -XX:TieredStopAtLevel=1 -XX:-UsePerfData -Djava.awt.headless=true}"

# 出力ファイル名
readonly F_TS="pool_timeseries.csv"
readonly F_HOURLY="pool_hourly.csv"
readonly F_SUMMARY="pool_summary.csv"
readonly F_EVENTS="pool_events.csv"
readonly F_CONFIG="pool_config.csv"
readonly F_QUEUED="pool_queued_threads.csv"
readonly F_QSTACK="pool_queued_stacks.csv"
readonly F_TDTHREADS="pool_threaddump_threads.csv"
readonly F_INFO="collection_info.csv"
readonly F_COLUMNS="columns.csv"

# 時系列 CSV の列 (順序固定)
TS_COLS="timestamp,epoch,elapsed_sec,sample_no,host,datasource,ds_type,jndi_name,stats_enabled,status,usage_pct"
TS_COLS+=",InUseCount,ActiveCount,IdleCount,AvailableCount,MaxUsedCount,min_pool_size,max_pool_size,peak_usage_pct,queued_threads"
TS_COLS+=",d_WaitCount,d_BlockingFailureCount,d_TimedOut,d_CreatedCount,d_DestroyedCount,d_TotalBlockingTime,d_TotalGetTime,d_TotalUsageTime"
TS_COLS+=",WaitCount,MaxWaitCount,BlockingFailureCount,TimedOut,CreatedCount,DestroyedCount"
TS_COLS+=",AverageBlockingTime,MaxWaitTime,TotalBlockingTime,AverageGetTime,MaxGetTime,TotalGetTime"
TS_COLS+=",AverageUsageTime,MaxUsageTime,TotalUsageTime,AverageCreationTime,MaxCreationTime,TotalCreationTime"
TS_COLS+=",AveragePoolTime,MaxPoolTime,TotalPoolTime"
TS_COLS+=",PreparedStatementCacheAccessCount,PreparedStatementCacheHitCount,PreparedStatementCacheMissCount"
TS_COLS+=",PreparedStatementCacheAddCount,PreparedStatementCacheDeleteCount,PreparedStatementCacheCurrentSize,ps_cache_hit_pct"
TS_COLS+=",XACommitCount,XACommitAverageTime,XACommitMaxTime,XAPrepareCount,XAPrepareAverageTime,XAPrepareMaxTime"
TS_COLS+=",XARollbackCount,XARollbackAverageTime,XARollbackMaxTime,XAStartCount,XAStartAverageTime,XAStartMaxTime"
TS_COLS+=",XAEndCount,XAEndAverageTime,XAEndMaxTime,XARecoverCount,XAForgetCount,note"
readonly TS_COLS
readonly Q_COLS="timestamp,epoch,sample_no,host,datasource,signature_id,waiting_threads,caller_frame,app_frame,thread_names"
readonly QS_COLS="signature_id,datasource,first_seen,frame_count,caller_frame,app_frame,stack"
readonly TD_COLS="dump_time,dump_file,trigger,thread_name,thread_state,category,jdbc_frame,app_frame,lock_name"
readonly CFG_COLS="datasource,ds_type,category,attribute,label,value,default_value,note"

# ------------------------------------------------------------------------------
# 共通関数
# ------------------------------------------------------------------------------
LOG_FILE=""
STOP=0
SLEEP_PID=""
JAVA_PID=""
JMX_STOP_FILE=""
TD_TRIGGER=""

now_ms()  { date +%s%3N; }
now_str() { date '+%Y-%m-%d %H:%M:%S'; }

log() {
    local level=$1; shift
    local msg
    msg="$(now_str) [$level] $*"
    printf '%s\n' "$msg" >&2
    if [ -n "$LOG_FILE" ]; then printf '%s\n' "$msg" >> "$LOG_FILE"; fi
}
die() { log ERROR "$*"; exit 1; }

usage() {
    cat <<'USAGE'
使い方: bash jboss_pool_collector.sh [オプション]

  JBoss EAP 8.1 のデータソース・コネクションプール統計を定期取得し CSV を出力します。

接続オプション
  -m, --mode MODE          取得方式: cli (jboss-cli, 既定) | jmx (JMX remote+http)
  -H, --jboss-home DIR     JBOSS_HOME (省略時は自動検出: /opt/server, /opt/eap 等)
  -c, --controller URL     管理インタフェース (既定: 稼働中の EAP の -bmanagement / ポートオフセット
                           から自動検出。検出できなければ remote+http://127.0.0.1:9990)
  -u, --user USER          管理ユーザ (省略時はローカル認証。環境変数 JBOSS_CLI_USER)
  -p, --password PASS      管理パスワード (環境変数 JBOSS_CLI_PASSWORD 推奨)

収集オプション
  -i, --interval SEC       取得間隔(秒) 既定: 10
  -d, --duration SEC       収集時間(秒) 既定: 3600 / 0 = Ctrl+C か kill で停止するまで
  -n, --count N            取得回数 (指定時は --duration より優先)
  -s, --datasource LIST    対象データソース名 (カンマ区切り, 既定: 全て)
  -o, --output-dir DIR     出力先 (既定: /tmp/pool_report_<ホスト>_<日時>)
  -w, --warn-pct PCT       高負荷と判定する使用率(%) 既定: 80
  -b, --bucket-min MIN     時間別集計の単位(分): 1,2,3,4,5,6,10,12,15,20,30,60 既定: 60
  -z, --tz TZ              タイムスタンプのタイムゾーン 既定: Asia/Tokyo
      --host-label NAME    CSV の host 列に出力する名前 (既定: ECS タスクID または ホスト名)

統計・調査オプション
      --no-enable-stats    統計(statistics-enabled)が無効でも有効化しない
      --keep-stats         ツールが有効化した統計を終了時に元に戻さない
      --no-queued-dump     接続待ちスレッド(dump-queued-threads-in-pool)を取得しない
  -t, --thread-dump        枯渇/待機を検出したらスレッドダンプを取得 (間隔は --thread-dump-gap)
      --thread-dump-gap S  スレッドダンプの最小取得間隔(秒) 既定: 300
      --no-bom             CSV に UTF-8 BOM を付けない

その他
  -D, --background         nohup でバックグラウンド実行 (ECS Exec セッション切断対策)
      --summarize DIR      既存の出力ディレクトリから集計 CSV のみ再生成
  -h, --help               このヘルプを表示

例
  # 10秒間隔で1時間収集 (ローカル認証)
  bash jboss_pool_collector.sh
  # 5秒間隔で24時間、バックグラウンド実行
  bash jboss_pool_collector.sh -i 5 -d 86400 -D
  # JMX で取得、特定データソースのみ、枯渇時にスレッドダンプ
  bash jboss_pool_collector.sh -m jmx -s MyDS,MyXADS -t

注意
  - 統計が無効なデータソースは実行時のみ有効化します (設定ファイルは変更しません)。
    終了時に元へ戻します (--keep-stats で維持)。
  - jboss-cli 方式は 1 回の取得ごとに JVM を起動します (約1～2秒, ヒープ最大128MB)。
    5 秒未満の間隔で取得したい場合は -m jmx を推奨します。
USAGE
}

need_arg() { [ $# -ge 2 ] && [ -n "$2" ] || die "オプション $1 には値が必要です"; }
is_uint()  { case "$1" in ''|*[!0-9]*) return 1 ;; *) return 0 ;; esac; }

csvq() { local s=${1//\"/\"\"}; printf '"%s"' "$s"; }

write_header() {   # $1=ファイル $2=カンマ区切り列名
    {
        if [ "$WRITE_BOM" = 1 ]; then printf '\357\273\277'; fi
        printf '%s\n' "$2"
    } > "$1"
}

# ------------------------------------------------------------------------------
# 引数解析
# ------------------------------------------------------------------------------
parse_args() {
    while [ $# -gt 0 ]; do
        case "$1" in
            -m|--mode)            need_arg "$@"; MODE=$2; shift 2 ;;
            -H|--jboss-home)      need_arg "$@"; JBOSS_HOME=$2; shift 2 ;;
            -c|--controller)      need_arg "$@"; CONTROLLER=$2; shift 2 ;;
            -u|--user)            need_arg "$@"; CLI_USER=$2; shift 2 ;;
            -p|--password)        need_arg "$@"; CLI_PASSWORD=$2; shift 2 ;;
            -i|--interval)        need_arg "$@"; INTERVAL=$2; shift 2 ;;
            -d|--duration)        need_arg "$@"; DURATION=$2; shift 2 ;;
            -n|--count)           need_arg "$@"; COUNT=$2; shift 2 ;;
            -s|--datasource)      need_arg "$@"; DS_FILTER=$2; shift 2 ;;
            -o|--output-dir)      need_arg "$@"; OUT_DIR=$2; shift 2 ;;
            -w|--warn-pct)        need_arg "$@"; WARN_PCT=$2; shift 2 ;;
            -b|--bucket-min)      need_arg "$@"; BUCKET_MIN=$2; shift 2 ;;
            -z|--tz)              need_arg "$@"; REPORT_TZ=$2; shift 2 ;;
            --host-label)         need_arg "$@"; HOST_LABEL=$2; shift 2 ;;
            --no-enable-stats)    ENABLE_STATS=0; shift ;;
            --keep-stats)         RESTORE_STATS=0; shift ;;
            --no-queued-dump)     QUEUED_DUMP=0; shift ;;
            -t|--thread-dump)     THREAD_DUMP=1; shift ;;
            --thread-dump-gap)    need_arg "$@"; THREAD_DUMP_GAP=$2; shift 2 ;;
            --no-bom)             WRITE_BOM=0; shift ;;
            -D|--background)      BACKGROUND=1; shift ;;
            --summarize)          need_arg "$@"; SUMMARIZE_DIR=$2; shift 2 ;;
            -h|--help)            usage; exit 0 ;;
            *) die "不明なオプションです: $1 (-h でヘルプを表示)" ;;
        esac
    done

    case "$MODE" in cli|jmx) ;; *) die "--mode は cli か jmx を指定してください: $MODE" ;; esac
    is_uint "$INTERVAL" && [ "$INTERVAL" -ge 1 ] || die "--interval は 1 以上の整数で指定してください"
    is_uint "$DURATION" || die "--duration は 0 以上の整数で指定してください"
    if [ -n "$COUNT" ]; then is_uint "$COUNT" && [ "$COUNT" -ge 1 ] || die "--count は 1 以上の整数で指定してください"; fi
    is_uint "$WARN_PCT" && [ "$WARN_PCT" -ge 1 ] && [ "$WARN_PCT" -le 100 ] || die "--warn-pct は 1～100 で指定してください"
    case "$BUCKET_MIN" in 1|2|3|4|5|6|10|12|15|20|30|60) ;; *) die "--bucket-min は 60 の約数(1,2,3,4,5,6,10,12,15,20,30,60)で指定してください" ;; esac
    is_uint "$THREAD_DUMP_GAP" || die "--thread-dump-gap は整数で指定してください"
}

setup_tz() {
    # tzdata が無いコンテナでも日本時間で出力できるよう POSIX 形式へフォールバック
    if [ -n "$REPORT_TZ" ] && [ ! -e "/usr/share/zoneinfo/$REPORT_TZ" ]; then
        case "$REPORT_TZ" in
            Asia/Tokyo|Japan|JST) REPORT_TZ="JST-9" ;;
            UTC|Etc/UTC)          REPORT_TZ="UTC0" ;;
            *[!A-Za-z0-9+-]*)     log WARN "タイムゾーン $REPORT_TZ が見つかりません (UTC として扱われる可能性があります)" ;;
        esac
    fi
    export TZ="$REPORT_TZ"
}

# ------------------------------------------------------------------------------
# 環境検出
# ------------------------------------------------------------------------------
detect_jboss_home() {
    local d f
    if [ -n "$JBOSS_HOME" ]; then
        [ -x "$JBOSS_HOME/bin/jboss-cli.sh" ] || die "JBOSS_HOME に jboss-cli.sh がありません: $JBOSS_HOME"
        return
    fi
    for d in /opt/server /opt/eap /opt/jboss-eap /opt/jboss/eap /opt/jboss-eap-8.1 /opt/jboss-eap-8.0 /opt/rh/eap8/root/usr/share/wildfly /opt/wildfly; do
        if [ -x "$d/bin/jboss-cli.sh" ]; then JBOSS_HOME=$d; return; fi
    done
    # 稼働中プロセスの -Djboss.home.dir から検出 (procps 不要)
    for f in /proc/[0-9]*/cmdline; do
        d=$(tr '\0' '\n' < "$f" 2>/dev/null | sed -n 's/^-Djboss\.home\.dir=//p' | head -n 1)
        if [ -n "$d" ] && [ -x "$d/bin/jboss-cli.sh" ]; then JBOSS_HOME=$d; return; fi
    done
    die "JBOSS_HOME を検出できませんでした。-H オプションで指定してください"
}

# 稼働中サーバの起動引数から管理インタフェースのアドレス/ポートを決める
detect_controller() {
    local f args a prev="" host="" offset=0 port=9990
    [ -z "$CONTROLLER" ] || return 0
    for f in /proc/[0-9]*/cmdline; do
        args=$(tr '\0' '\n' < "$f" 2>/dev/null) || continue
        printf '%s\n' "$args" | grep -qx 'org.jboss.as.standalone' || continue
        while IFS= read -r a; do
            case "$a" in
                -bmanagement=*)                      host=${a#-bmanagement=} ;;
                -Djboss.bind.address.management=*)   host=${a#*=} ;;
                -Djboss.socket.binding.port-offset=*) offset=${a#*=} ;;
                -Djboss.management.http.port=*)      port=${a#*=} ;;
            esac
            if [ "$prev" = "-bmanagement" ]; then host=$a; fi
            prev=$a
        done <<< "$args"
        break
    done
    case "$host" in ""|0.0.0.0|::|"[::]"|*'${'*) host=127.0.0.1 ;; esac
    is_uint "$offset" || offset=0
    is_uint "$port" || port=9990
    CONTROLLER="remote+http://$host:$((port + offset))"
}

ECS_TASK_ARN=""; ECS_CLUSTER=""; ECS_FAMILY=""; ECS_AZ=""; ECS_CONTAINER=""
detect_ecs() {
    local uri="${ECS_CONTAINER_METADATA_URI_V4:-${ECS_CONTAINER_METADATA_URI:-}}" task cont
    [ -n "$uri" ] || return 0
    command -v curl >/dev/null 2>&1 || return 0
    task=$(curl -s --noproxy '*' --max-time 3 "$uri/task" 2>/dev/null) || task=""
    cont=$(curl -s --noproxy '*' --max-time 3 "$uri" 2>/dev/null) || cont=""
    jget() { printf '%s' "$1" | grep -o "\"$2\" *: *\"[^\"]*\"" | head -n 1 | sed 's/.*: *"//; s/"$//'; }
    ECS_TASK_ARN=$(jget "$task" TaskARN)
    ECS_CLUSTER=$(jget "$task" Cluster)
    ECS_FAMILY="$(jget "$task" Family):$(jget "$task" Revision)"
    [ "$ECS_FAMILY" = ":" ] && ECS_FAMILY=""
    ECS_AZ=$(jget "$task" AvailabilityZone)
    ECS_CONTAINER=$(jget "$cont" Name)
}

host_label() {
    if [ -n "$HOST_LABEL" ]; then printf '%s' "$HOST_LABEL"; return; fi
    if [ -n "$ECS_TASK_ARN" ]; then printf 'task-%s' "${ECS_TASK_ARN##*/}" | cut -c1-17; return; fi
    printf '%s' "${HOSTNAME:-$(uname -n)}"
}

# 子プロセスの JVM にサーバ用設定 (javaagent / 大きなヒープ等) を引き継がない
run_java_env() {
    env -u JAVA_TOOL_OPTIONS -u _JAVA_OPTIONS -u JDK_JAVA_OPTIONS -u JAVA_OPTS_APPEND \
        JAVA_OPTS="$CLI_JAVA_OPTS" "$@"
}

# ------------------------------------------------------------------------------
# awk / Java プログラムの展開
# ------------------------------------------------------------------------------
AWK_BIN="${AWK:-}"
setup_awk() {
    if [ -z "$AWK_BIN" ]; then
        if command -v gawk >/dev/null 2>&1; then AWK_BIN="gawk"; else AWK_BIN="awk"; fi
    fi
    command -v "$AWK_BIN" >/dev/null 2>&1 || die "awk が見つかりません (dnf/microdnf install gawk)"
}

write_programs() {
# --- jboss-cli の JSON 出力を行形式に平坦化 ------------------------------------
cat > "$WORK_DIR/flatten.awk" <<'AWK'
# 入力 : jboss-cli --output-json の出力 (複数の JSON と @@ マーカー行)
# 出力 : @@... (マーカーをそのまま) / O<TAB>outcome / F<TAB>failure-description
#        L<TAB>path<TAB>S|N<TAB>value  (path は result 配下の階層を \034 で連結)
BEGIN { SEP = "\034"; depth = 0; pkey = ""; expect_key = 0 }
function jname(   n) {
    if (ktype[depth] == "a") { n = aidx[depth]; aidx[depth]++; return n }
    return pkey
}
function jpush(kind, name) {
    depth++; ktype[depth] = kind; kname[depth] = name; aidx[depth] = 0
    expect_key = (kind == "o")
}
function jpop() { if (depth > 0) depth--; expect_key = 0 }
function jstrend(s,   i, j, k, n, rest) {
    i = 2
    while (1) {
        rest = substr(s, i); j = index(rest, "\"")
        if (j == 0) return 0
        j = i + j - 1; n = 0; k = j - 1
        while (k > 1 && substr(s, k, 1) == "\\") { n++; k-- }
        if (n % 2 == 0) return j
        i = j + 1
    }
}
function emit(name, t, v,   i, p) {
    if (depth == 1) {
        if (name == "outcome") print "O\t" v
        else if (name == "failure-description") print "F\t" v
        else if (name == "result") print "L\t\t" t "\t" v
        return
    }
    if (kname[2] == "failure-description") { print "F\t" v; return }
    if (kname[2] != "result") return
    p = ""
    for (i = 3; i <= depth; i++) p = p kname[i] SEP
    print "L\t" p name "\t" t "\t" v
}
{
    line = $0
    sub(/\r$/, "", line)
    if (depth == 0) {
        if (line ~ /^@@/) { print line; next }
        sub(/^[ \t]+/, "", line)
        c = substr(line, 1, 1)
        if (c != "{" && c != "[") next
    }
    while (1) {
        sub(/^[ \t]+/, "", line)
        if (line == "") break
        c = substr(line, 1, 1)
        if (c == "{" || c == "[") {
            nm = ""
            if (depth > 0) nm = jname()
            jpush((c == "{") ? "o" : "a", nm)
            line = substr(line, 2)
        } else if (c == "}" || c == "]") {
            jpop(); line = substr(line, 2)
        } else if (c == ",") {
            if (depth > 0 && ktype[depth] == "o") expect_key = 1
            line = substr(line, 2)
        } else if (c == ":") {
            line = substr(line, 2)
        } else if (c == "\"") {
            e = jstrend(line)
            if (e == 0) { tok = substr(line, 2); line = "" }
            else { tok = substr(line, 2, e - 2); line = substr(line, e + 1) }
            if (depth > 0 && ktype[depth] == "o" && expect_key) { pkey = tok; expect_key = 0 }
            else if (depth > 0) emit(jname(), "S", tok)
        } else {
            if (match(line, /^[^],:{}[ \t]+/)) { tok = substr(line, 1, RLENGTH); line = substr(line, RLENGTH + 1) }
            else { tok = c; line = substr(line, 2) }
            if (depth > 0) emit(jname(), "N", tok)
        }
    }
}
AWK

# --- 検出結果 → 設定 CSV / 状態ファイル ----------------------------------------
cat > "$WORK_DIR/discover.awk" <<'AWK'
# 入力 : 平坦化した検出結果 (@@INFO / @@DS セクション)
# 出力 : cfgcsv (プール設定 CSV 行を追記), statef (DS 一覧), infof (サーバ情報)
BEGIN {
    FS = "\t"; SEP = "\034"
    nf = split(filter, FL, ",")
    for (i = 1; i <= nf; i++) { x = FL[i]; gsub(/^ +| +$/, "", x); if (x != "") { WANT[x] = 1; nwant++ } }
    nimp = 0
    imp("jndi-name", "JNDI名", "")
    imp("enabled", "有効", "true")
    imp("driver-name", "JDBCドライバ名", "")
    imp("connection-url", "接続URL", "")
    imp("xa-datasource-class", "XAデータソースクラス", "")
    imp("user-name", "接続ユーザ", "")
    imp("min-pool-size", "最小プールサイズ", "0")
    imp("initial-pool-size", "初期プールサイズ", "0")
    imp("max-pool-size", "最大プールサイズ", "20")
    imp("pool-prefill", "起動時に最小数まで接続作成", "false")
    imp("pool-use-strict-min", "最小数を厳守", "false")
    imp("pool-fair", "公平モード(FIFO)", "true")
    imp("blocking-timeout-wait-millis", "接続取得待機タイムアウト(ms)", "30000")
    imp("idle-timeout-minutes", "アイドルタイムアウト(分)", "30")
    imp("background-validation", "バックグラウンド検証", "false")
    imp("background-validation-millis", "バックグラウンド検証間隔(ms)", "0")
    imp("validate-on-match", "取得時検証", "false")
    imp("check-valid-connection-sql", "検証SQL", "")
    imp("valid-connection-checker-class-name", "接続検証クラス", "")
    imp("exception-sorter-class-name", "例外ソータークラス", "")
    imp("stale-connection-checker-class-name", "ステイル接続チェッカークラス", "")
    imp("use-fast-fail", "検証失敗時に即時失敗", "false")
    imp("flush-strategy", "フラッシュ戦略", "FailingConnectionOnly")
    imp("allocation-retry", "取得リトライ回数", "0")
    imp("allocation-retry-wait-millis", "取得リトライ間隔(ms)", "5000")
    imp("query-timeout", "クエリタイムアウト(秒)", "")
    imp("set-tx-query-timeout", "Txタイムアウトをクエリに適用", "false")
    imp("use-try-lock", "tryLockタイムアウト(秒)", "")
    imp("xa-resource-timeout", "XAリソースタイムアウト(秒)", "")
    imp("same-rm-override", "同一RM判定の上書き", "")
    imp("interleaving", "インターリーブ", "false")
    imp("no-tx-separate-pool", "非Tx用の別プール", "false")
    imp("pad-xid", "XIDパディング", "false")
    imp("wrap-xa-resource", "XAResourceラップ", "true")
    imp("prepared-statements-cache-size", "PSキャッシュサイズ", "0")
    imp("share-prepared-statements", "PS共有", "false")
    imp("track-statements", "ステートメント追跡", "nowarn")
    imp("transaction-isolation", "トランザクション分離レベル", "")
    imp("new-connection-sql", "接続時実行SQL", "")
    imp("jta", "JTA連携", "true")
    imp("use-ccm", "CCM(キャッシュ接続マネージャ)使用", "true")
    imp("mcp", "プール実装(MCP)", "")
    imp("statistics-enabled", "統計有効(設定値)", "false")
    imp("capacity-incrementer-class", "容量増加ポリシー", "")
    imp("capacity-decrementer-class", "容量減少ポリシー", "")
    imp("connectable", "Connectable(CMR)", "false")
    imp("tracking", "接続追跡", "false")
    imp("enlistment-trace", "エンリストトレース", "false")
}
function imp(a, l, d) { IMP[++nimp] = a; LBL[a] = l; DEF[a] = d; ISIMP[a] = 1 }
function q(s) { gsub(/"/, "\"\"", s); return "\"" s "\"" }
function unesc(s) { gsub(/\\"/, "\"", s); gsub(/\\n/, " ", s); gsub(/\\t/, " ", s); return s }
function mask_url(v,   out, lv) {
    out = ""
    while (1) {
        lv = tolower(v)
        if (!match(lv, /(password|passwd|pwd)=/)) break
        out = out substr(v, 1, RSTART + RLENGTH - 1) "****"
        v = substr(v, RSTART + RLENGTH)
        if (match(v, /[;&,)]/)) v = substr(v, RSTART); else v = ""
    }
    v = out v
    if (match(v, /\/\/[^\/:@]+:[^\/@]+@/)) {
        seg = substr(v, RSTART, RLENGTH); i = index(substr(seg, 3), ":") + 2
        v = substr(v, 1, RSTART - 1) substr(seg, 1, i) "****@" substr(v, RSTART + RLENGTH)
    }
    return v
}
function masked(attr, v,   la) {
    if (v == "") return v
    la = tolower(attr)
    if (la ~ /password|passwd|secret|credential|token/) return "********"
    if (la ~ /url/) return mask_url(v)
    return v
}
function gv(k, a) { return ((k, a) in VAL) ? VAL[k, a] : "" }   # 参照で要素を作らない
function effnum(v, d) {           # 式 ${env.X:20} の既定値も解釈
    if (v ~ /^-?[0-9]+$/) return v
    if (match(v, /:-?[0-9]+}$/)) return substr(v, RSTART + 1, RLENGTH - 2)
    return d
}
/^@@INFO/ { sec = "INFO"; next }
/^@@DS/   { sec = "DS"; next }
/^@@/     { sec = ""; next }
$1 == "L" && sec == "INFO" {
    if ($2 != "" && index($2, SEP) == 0 && $4 != "null") INFO[$2] = unesc($4)
    next
}
$1 == "L" && sec == "DS" {
    n = split($2, p, SEP); v = ($4 == "null") ? "" : unesc($4)
    if (p[1] == "data-source" || p[1] == "xa-data-source") {
        k = p[1] SUBSEP p[2]
        if (!(k in SEEN)) { SEEN[k] = 1; ORDER[++nds] = k; DT[k] = p[1]; DN[k] = p[2] }
        if (n == 3) { a = p[3] }
        else if (n == 4 && p[4] == "EXPRESSION_VALUE") { a = p[3] }
        else if (n == 5 && (p[3] == "connection-properties" || p[3] == "xa-datasource-properties") && p[5] == "value") { a = p[3] "/" p[4] }
        else next
        if (!((k, a) in VAL)) ALIST[k] = ALIST[k] SEP a
        VAL[k, a] = v
    } else if (p[1] == "jdbc-driver" && n == 3) {
        k = p[2]
        if (!(k in DSEEN)) { DSEEN[k] = 1; DORDER[++ndrv] = k }
        if (!((k, p[3]) in DVAL)) DLIST[k] = DLIST[k] SEP p[3]
        DVAL[k, p[3]] = v
    }
    next
}
END {
    printf "" > statef
    for (i = 1; i <= nds; i++) {
        k = ORDER[i]; t = DT[k]; nm = DN[k]
        en = tolower(gv(k, "enabled")); if (en == "") en = "true"
        mx = effnum(gv(k, "max-pool-size"), "20"); mxdef = (gv(k, "max-pool-size") == "") ? 1 : 0
        mn = effnum(gv(k, "min-pool-size"), "0")
        bt = effnum(gv(k, "blocking-timeout-wait-millis"), "30000")
        tgt = (nwant == 0 || (nm in WANT)) ? 1 : 0
        print t "\t" nm "\t" en "\t" gv(k, "jndi-name") "\t" mn "\t" mx "\t" mxdef "\t" bt "\t" tgt > statef
        # 主要属性 → その他属性 → 接続/XA プロパティ の順に出力
        for (j = 1; j <= nimp; j++) {
            a = IMP[j]
            if (!((k, a) in VAL)) continue
            v = VAL[k, a]; note = ""
            if (v == "" && DEF[a] != "") note = "未設定 (既定値 " DEF[a] " が適用)"
            print q(nm) "," q(t) "," q("主要") "," q(a) "," q(LBL[a]) "," q(masked(a, v)) "," q(DEF[a]) "," q(note) >> cfgcsv
        }
        na = split(ALIST[k], AL, SEP)
        for (j = 2; j <= na; j++) {
            a = AL[j]
            if (a in ISIMP) continue
            cat = "その他"
            if (index(a, "connection-properties/") == 1) cat = "接続プロパティ"
            if (index(a, "xa-datasource-properties/") == 1) cat = "XAプロパティ"
            print q(nm) "," q(t) "," q(cat) "," q(a) "," q("") "," q(masked(a, VAL[k, a])) "," q("") "," q("") >> cfgcsv
        }
    }
    for (i = 1; i <= ndrv; i++) {
        k = DORDER[i]; na = split(DLIST[k], AL, SEP)
        for (j = 2; j <= na; j++) {
            if (DVAL[k, AL[j]] == "") continue
            print q(k) "," q("jdbc-driver") "," q("JDBCドライバ") "," q(AL[j]) "," q("") "," q(DVAL[k, AL[j]]) "," q("") "," q("") >> cfgcsv
        }
    }
    printf "" > infof
    split("product-name product-version release-version release-codename name server-state running-mode launch-type suspend-state", IK, " ")
    for (i = 1; i <= 9; i++) if (IK[i] in INFO) print IK[i] "\t" INFO[IK[i]] > infof
}
AWK

# --- 1 サンプル分 → 時系列 CSV 行 / 接続待ちスレッド CSV --------------------------
cat > "$WORK_DIR/rows.awk" <<'AWK'
# 入力 : 平坦化したサンプル (@@DS|type|name, @@P/@@J/@@Q セクション, @@ENDDS)
# 変数 : cols ts epoch elapsed sno host warn cfgf prevf newprevf statusf qsigf qcsvf qstackf
BEGIN {
    FS = "\t"; SEP = "\034"
    ncol = split(cols, COL, ",")
    nd = split("WaitCount,BlockingFailureCount,TimedOut,CreatedCount,DestroyedCount,TotalBlockingTime,TotalGetTime,TotalUsageTime", DC, ",")
    for (i = 1; i < 256; i++) ORD[sprintf("%c", i)] = i
    while ((getline ln < cfgf) > 0) {
        split(ln, a, "\t"); k = a[1] SUBSEP a[2]
        CJNDI[k] = a[4]; CMIN[k] = a[5]; CMAX[k] = a[6]; CMAXDEF[k] = a[7]
    }
    close(cfgf)
    while ((getline ln < prevf) > 0) {
        n = split(ln, a, "\t"); k = a[1] SUBSEP a[2]; HASPREV[k] = 1
        for (i = 3; i <= n; i++) { p = index(a[i], "="); PREV[k, substr(a[i], 1, p - 1)] = substr(a[i], p + 1) }
    }
    close(prevf)
    maxqid = 0
    while ((getline ln < qsigf) > 0) { split(ln, a, "\t"); QID[a[1]] = a[2]; if (a[2] + 0 > maxqid) maxqid = a[2] + 0 }
    close(qsigf)
    printf "" > newprevf
    printf "" > statusf
}
function num(x) { return (x == "" || x == "null") ? 0 : x + 0 }
function clean(s) { gsub(/[\r\n]+/, " ", s); gsub(/,/, ";", s); gsub(/"/, "'", s); return s }
function q(s) { gsub(/"/, "\"\"", s); return "\"" s "\"" }
function djb2(s,   i, n, h) { h = 5381; n = length(s); for (i = 1; i <= n; i++) h = (h * 33 + ORD[substr(s, i, 1)]) % 2147483647; return h }
function is_fw(f) {
    return f ~ /^(java\.|javax\.|jakarta\.|jdk\.|sun\.|com\.sun\.|org\.jboss\.|org\.wildfly\.|io\.undertow\.|org\.hibernate\.|org\.apache\.jasper\.|org\.apache\.catalina\.|org\.apache\.cxf\.|org\.apache\.commons\.|org\.springframework\.|io\.smallrye\.|org\.eclipse\.|com\.zaxxer\.|io\.opentelemetry\.|io\.micrometer\.|net\.bytebuddy\.|com\.fasterxml\.|org\.glassfish\.|io\.netty\.|org\.xnio\.|org\.infinispan\.|com\.arjuna\.|io\.agroal\.|io\.quarkus\.|datadog\.|com\.datadog\.|co\.elastic\.|com\.newrelic\.|software\.amazon\.)/
}
function frames_info(frames,   n, f, j, lastgc) {
    n = split(frames, f, "\n"); lastgc = 0
    for (j = 1; j <= n; j++) if (f[j] ~ /^(org\.jboss\.jca\.|org\.jboss\.as\.connector\.)/ && f[j] ~ /getConnection/) lastgc = j
    CALLER = (lastgc > 0 && lastgc < n) ? f[lastgc + 1] : ""
    APPF = ""
    for (j = lastgc + 1; j <= n; j++) if (!is_fw(f[j])) { APPF = f[j]; break }
    NFR = n
}
function reset_ds() { delete P; delete J; delete Q; delete ERRS; nq = 0; fmsg = ""; sec = ""; qseen = 0 }
/^@@DS\|/ {
    reset_ds()
    rest = substr($0, 6); p = index(rest, "|")
    dtype = substr(rest, 1, p - 1); dname = substr(rest, p + 1); key = dtype SUBSEP dname
    next
}
/^@@P$/ { sec = "P"; next }
/^@@J$/ { sec = "J"; next }
/^@@Q$/ { sec = "Q"; qseen = 1; next }
/^@@ERR$/ { ERRS[sec] = 1; next }
/^@@ENDDS$/ { finish(); next }
/^@@/ { next }
$1 == "O" { if ($2 != "success") ERRS[sec] = 1; next }
$1 == "F" { if (fmsg == "") fmsg = $2; next }
$1 == "L" {
    if (sec == "P") P[$2] = $4
    else if (sec == "J") J[$2] = $4
    else if (sec == "Q" && $2 ~ /^[0-9]+$/) Q[++nq] = $4
    next
}
function emit_row(   i, c, line) {
    line = ""
    for (i = 1; i <= ncol; i++) { c = COL[i]; line = line ((i > 1) ? "," : "") ((c in R) ? R[c] : "") }
    print line
}
function carry_prev(   i, s) {
    s = ""
    for (i = 1; i <= nd; i++) if ((key, DC[i]) in PREV) s = s "\t" DC[i] "=" PREV[key, DC[i]]
    if (s != "") print dtype "\t" dname s > newprevf
}
function finish(   i, c, cur, d, st, mx, inuse, reset, note, sen, np, acc, hit) {
    delete R
    mx = CMAX[key] + 0
    R["timestamp"] = ts; R["epoch"] = epoch; R["elapsed_sec"] = elapsed; R["sample_no"] = sno
    R["host"] = clean(host); R["datasource"] = clean(dname); R["ds_type"] = dtype; R["jndi_name"] = clean(CJNDI[key])
    R["min_pool_size"] = CMIN[key]; R["max_pool_size"] = CMAX[key]
    note = ""
    if (CMAXDEF[key] == "1") note = note "max-pool-size未設定(既定値20で計算) "
    if ("P" in ERRS) {
        R["status"] = "ERROR"; R["note"] = clean(note "取得失敗: " fmsg)
        emit_row(); carry_prev(); print "ERROR\t" dtype "\t" dname > statusf
        return
    }
    sen = (P["statistics-enabled"] == "true")
    R["stats_enabled"] = sen ? "true" : "false"
    for (i = 1; i <= ncol; i++) {
        c = COL[i]
        if (c in P) R[c] = P[c]
        else if (c in J) R[c] = J[c]
    }
    inuse = num(P["InUseCount"])
    if (mx > 0) {
        R["usage_pct"] = sprintf("%.1f", inuse * 100 / mx)
        R["peak_usage_pct"] = sprintf("%.1f", num(P["MaxUsedCount"]) * 100 / mx)
    }
    if (!sen) {
        R["status"] = "NOSTATS"; R["note"] = clean(note "statistics-enabled=false")
        emit_row(); carry_prev(); print "NOSTATS\t" dtype "\t" dname > statusf
        return
    }
    reset = 0; np = ""
    for (i = 1; i <= nd; i++) {
        c = DC[i]
        if (!(c in P)) continue
        cur = num(P[c])
        if (HASPREV[key] && ((key, c) in PREV)) {
            d = cur - PREV[key, c]
            if (d < 0) { d = cur; reset = 1 }
            R["d_" c] = d
        }
        np = np "\t" c "=" cur
    }
    print dtype "\t" dname np > newprevf
    if (reset) note = note "カウンタリセット検出(サーバ再起動/統計再有効化) "
    if (qseen) R["queued_threads"] = nq
    acc = num(J["PreparedStatementCacheAccessCount"]); hit = num(J["PreparedStatementCacheHitCount"])
    if (acc > 0) R["ps_cache_hit_pct"] = sprintf("%.1f", hit * 100 / acc)
    if ((mx > 0 && inuse >= mx) || num(R["d_BlockingFailureCount"]) > 0 || (mx > 0 && ("AvailableCount" in P) && num(P["AvailableCount"]) <= 0))
        st = "EXHAUSTED"
    else if (num(R["d_WaitCount"]) > 0 || nq > 0)
        st = "WAIT"
    else if (mx > 0 && inuse * 100 >= warn * mx)
        st = "HIGH"
    else
        st = "OK"
    R["status"] = st
    if (note != "") R["note"] = clean(note)
    emit_row()
    print st "\t" dtype "\t" dname > statusf
    if (nq > 0) proc_queued()
}
function proc_queued(   i, j, n, fr, t, tname, frames, sk, h, no) {
    delete SIG; delete CNT; delete NAMES; delete SO; no = 0
    for (i = 1; i <= nq; i++) {
        n = split(Q[i], fr, /\\n/)
        tname = fr[1]; sub(/^Queued thread: */, "", tname); gsub(/\\"/, "\"", tname)
        frames = ""
        for (j = 2; j <= n; j++) { t = fr[j]; sub(/^[ \t]+/, "", t); if (t != "") frames = frames ((frames == "") ? "" : "\n") t }
        sk = dname SEP frames
        if (!(sk in SIG)) { SIG[sk] = djb2(sk); SO[++no] = sk }
        CNT[sk]++
        if (CNT[sk] <= 3) NAMES[sk] = NAMES[sk] ((NAMES[sk] == "") ? "" : " / ") tname
    }
    for (j = 1; j <= no; j++) {
        sk = SO[j]; h = SIG[sk]
        frames = substr(sk, length(dname) + 2)
        frames_info(frames)
        if (!(h in QID)) {
            QID[h] = ++maxqid
            print h "\t" maxqid >> qsigf
            print maxqid "," q(dname) "," q(ts) "," NFR "," q(CALLER) "," q(APPF) "," q(substr(frames, 1, 30000)) >> qstackf
        }
        print q(ts) "," epoch "," sno "," q(host) "," q(dname) "," QID[h] "," CNT[sk] "," q(CALLER) "," q(APPF) "," q(NAMES[sk] ((CNT[sk] > 3) ? " ..." : "")) >> qcsvf
    }
}
AWK

# --- 時系列 CSV → サマリ / 時間別 / イベント CSV ----------------------------------
cat > "$WORK_DIR/summary.awk" <<'AWK'
# 変数 : interval warn bucket bom sumf hourf evf
BEGIN {
    FS = ","; SEP = SUBSEP
    RANK["OK"] = 0; RANK["HIGH"] = 1; RANK["WAIT"] = 2; RANK["EXHAUSTED"] = 3
    LBL["OK"] = "正常"; LBL["HIGH"] = "高負荷"; LBL["WAIT"] = "待機発生"; LBL["EXHAUSTED"] = "枯渇"
    LBL["NOSTATS"] = "統計無効"; LBL["ERROR"] = "取得失敗"; LBL[""] = ""
    nds = 0; nev = 0
    split("d_WaitCount d_BlockingFailureCount d_TimedOut d_CreatedCount d_DestroyedCount d_TotalBlockingTime", DSUM, " ")
    split("MaxWaitCount MaxWaitTime AverageBlockingTime AverageGetTime MaxGetTime AverageUsageTime MaxUsageTime AverageCreationTime MaxCreationTime MaxUsedCount PreparedStatementCacheHitCount PreparedStatementCacheAccessCount", LASTS, " ")
}
function v(name) { return (name in H) ? $(H[name]) : "" }
function q(s) { gsub(/"/, "\"\"", s); return "\"" s "\"" }
function f1(x) { return sprintf("%.1f", x) }
function acc(k, m, x) {
    if (!((k, m, "n") in A)) { A[k, m, "n"] = 0; A[k, m, "min"] = x; A[k, m, "max"] = x; A[k, m, "sum"] = 0 }
    A[k, m, "n"]++; A[k, m, "sum"] += x
    if (x < A[k, m, "min"]) A[k, m, "min"] = x
    if (x > A[k, m, "max"]) A[k, m, "max"] = x
}
function amin(k, m) { return ((k, m, "n") in A) ? A[k, m, "min"] : "" }
function amax(k, m) { return ((k, m, "n") in A) ? A[k, m, "max"] : "" }
function aavg(k, m) { return ((k, m, "n") in A && A[k, m, "n"] > 0) ? f1(A[k, m, "sum"] / A[k, m, "n"]) : "" }
function close_exh(k,   len) {
    if (!(k in EXS)) return
    len = EXL[k] - EXS[k] + interval
    EXT[k] += len
    if (len > LONG[k]) LONG[k] = len
    delete EXS[k]
}
function close_ev(k,   e) {
    if (!(k in EVS)) return
    e = ++nev
    E_K[e] = k; E_SEV[e] = EVSEV[k]; E_STS[e] = EVSTS[k]; E_ETS[e] = EVETS[k]; E_SEP[e] = EVS[k]
    E_DUR[e] = EVE[k] - EVS[k] + interval; E_N[e] = EVN[k]; E_IN[e] = EVIN[k]; E_MAX[e] = EVMAXP[k]
    E_USE[e] = EVUSE[k]; E_Q[e] = EVQ[k]; E_W[e] = EVW[k]; E_B[e] = EVB[k]; E_MWC[e] = EVMWC[k]; E_MWT[e] = EVMWT[k]
    delete EVS[k]
}
function bucket_of(t,   mm) { mm = int(substr(t, 15, 2) / bucket) * bucket; return substr(t, 1, 14) sprintf("%02d", mm) }
function bucket_end(b,   hh, mm, t) {
    hh = substr(b, 12, 2) + 0; mm = substr(b, 15, 2) + bucket
    if (mm >= 60) { hh++; mm -= 60 }
    return sprintf("%02d:%02d", hh, mm)
}
NR == 1 { sub(/^\357\273\277/, "", $1); for (i = 1; i <= NF; i++) H[$i] = i; next }
{
    ds = v("datasource"); ty = v("ds_type"); k = ds SEP ty
    if (!(k in N)) { DSO[++nds] = k; DSN[k] = ds; DST[k] = ty; LONG[k] = 0 }
    N[k]++
    t = v("timestamp"); ep = v("epoch") + 0; st = v("status")
    if (N[k] == 1) { FTS[k] = t; FEP[k] = ep }
    LTS[k] = t; LEP[k] = ep
    if (v("jndi_name") != "") JNDI[k] = v("jndi_name")
    if (v("max_pool_size") != "") MAXP[k] = v("max_pool_size")
    if (v("min_pool_size") != "") MINP[k] = v("min_pool_size")
    if (index(v("note"), "max-pool-size") > 0) MAXDEF[k] = 1
    if (index(v("note"), "カウンタリセット") > 0) RST[k]++

    b = bucket_of(t); bk = k SEP b
    if (!(bk in BN)) { NB[k]++; BO[k, NB[k]] = b }
    BN[bk]++
    if (v("max_pool_size") != "") BMAXP[bk] = v("max_pool_size")

    if (st == "ERROR" || st == "NOSTATS") {
        if (st == "ERROR") { ERRN[k]++; BERR[bk]++ } else NOST[k]++
        close_exh(k); close_ev(k)
        next
    }
    VN[k]++
    inuse = v("InUseCount") + 0; usage = v("usage_pct"); qn = v("queued_threads") + 0
    acc(k, "inuse", inuse); acc(bk, "inuse", inuse)
    acc(k, "active", v("ActiveCount") + 0); acc(bk, "active", v("ActiveCount") + 0)
    acc(k, "idle", v("IdleCount") + 0); acc(bk, "idle", v("IdleCount") + 0)
    acc(k, "avail", v("AvailableCount") + 0); acc(bk, "avail", v("AvailableCount") + 0)
    acc(k, "queued", qn); acc(bk, "queued", qn)
    if (usage != "") { acc(k, "usage", usage + 0); acc(bk, "usage", usage + 0) }
    HIST[k, inuse]++; if (inuse > HMAX[k]) HMAX[k] = inuse
    XS[k, VN[k]] = ep; YS[k, VN[k]] = inuse
    for (i = 1; i <= 6; i++) { x = v(DSUM[i]); if (x != "") { S[k, DSUM[i]] += x; S[bk, DSUM[i]] += x } }
    for (i = 1; i <= 12; i++) { x = v(LASTS[i]); if (x != "") { L[k, LASTS[i]] = x; if (!((bk, LASTS[i]) in L) || x + 0 > L[bk, LASTS[i]] + 0) L[bk, LASTS[i]] = x } }
    CNT[k, st]++; CNT[bk, st]++
    if (!(bk in BW) || RANK[st] > RANK[BW[bk]]) BW[bk] = st

    if (st == "EXHAUSTED") {
        if (!(k in EXS)) EXS[k] = ep
        EXL[k] = ep
        if (!(k in FEXH)) FEXH[k] = t
        LEXH[k] = t
    } else close_exh(k)

    if (st == "HIGH" || st == "WAIT" || st == "EXHAUSTED") {
        if (!(k in EVS)) {
            EVS[k] = ep; EVSTS[k] = t; EVSEV[k] = st; EVN[k] = 0; EVIN[k] = 0; EVUSE[k] = 0; EVQ[k] = 0; EVW[k] = 0; EVB[k] = 0
        }
        EVE[k] = ep; EVETS[k] = t; EVN[k]++
        if (RANK[st] > RANK[EVSEV[k]]) EVSEV[k] = st
        if (inuse > EVIN[k]) EVIN[k] = inuse
        if (usage + 0 > EVUSE[k]) EVUSE[k] = usage + 0
        if (qn > EVQ[k]) EVQ[k] = qn
        EVW[k] += v("d_WaitCount") + 0; EVB[k] += v("d_BlockingFailureCount") + 0
        EVMWC[k] = v("MaxWaitCount"); EVMWT[k] = v("MaxWaitTime"); EVMAXP[k] = v("max_pool_size")
    } else close_ev(k)
}
function p95(k,   tot, target, c, x) {
    tot = VN[k]; if (tot <= 0) return ""
    target = int(tot * 0.95); if (target < tot * 0.95) target++
    c = 0
    for (x = 0; x <= HMAX[k]; x++) { c += HIST[k, x]; if (c >= target) return x }
    return HMAX[k]
}
function trend(k,   n, i, x, y, sx, sy, sxx, sxy, den) {
    n = VN[k]; if (n < 3) return ""
    sx = sy = sxx = sxy = 0
    for (i = 1; i <= n; i++) { x = (XS[k, i] - XS[k, 1]) / 3600; y = YS[k, i]; sx += x; sy += y; sxx += x * x; sxy += x * y }
    den = n * sxx - sx * sx
    if (den == 0) return ""
    return sprintf("%.2f", (n * sxy - sx * sy) / den)
}
function quarter_min(k, from, to,   i, m) {
    m = ""
    for (i = from; i <= to; i++) if (m == "" || YS[k, i] < m) m = YS[k, i]
    return m
}
END {
    for (i = 1; i <= nds; i++) { k = DSO[i]; close_exh(k); close_ev(k) }
    if (bom) { printf "\357\273\277" > sumf; printf "\357\273\277" > hourf; printf "\357\273\277" > evf }
    print "datasource,ds_type,jndi_name,verdict_code,verdict,min_pool_size,max_pool_size,first_timestamp,last_timestamp,duration_sec,samples,valid_samples,error_samples,nostats_samples,inuse_min,inuse_max,inuse_avg,inuse_p95,usage_pct_max,usage_pct_avg,usage_pct_p95,active_min,active_max,active_avg,idle_min,idle_max,idle_avg,available_min,MaxUsedCount,peak_usage_pct,high_samples,wait_samples,exhausted_samples,exhausted_ratio_pct,exhausted_time_sec,longest_exhausted_sec,first_exhausted_at,last_exhausted_at,queued_threads_max,d_WaitCount,d_BlockingFailureCount,d_TimedOut,d_CreatedCount,d_DestroyedCount,d_TotalBlockingTime,MaxWaitCount,MaxWaitTime,AverageBlockingTime,AverageGetTime,MaxGetTime,AverageUsageTime,MaxUsageTime,AverageCreationTime,MaxCreationTime,ps_cache_hit_pct,inuse_trend_per_hour,leak_suspect,counter_resets,comment" > sumf
    for (i = 1; i <= nds; i++) {
        k = DSO[i]; mx = MAXP[k] + 0; vn = VN[k] + 0
        dur = LEP[k] - FEP[k] + interval
        exn = CNT[k, "EXHAUSTED"] + 0; wn = CNT[k, "WAIT"] + 0; hn = CNT[k, "HIGH"] + 0
        dW = S[k, "d_WaitCount"] + 0; dB = S[k, "d_BlockingFailureCount"] + 0
        dT = S[k, "d_TimedOut"] + 0; dD = S[k, "d_DestroyedCount"] + 0
        pk = p95(k)
        upct95 = (pk != "" && mx > 0) ? f1(pk * 100 / mx) : ""
        peak = (mx > 0 && ((k, "MaxUsedCount") in L)) ? f1(L[k, "MaxUsedCount"] * 100 / mx) : ""
        hitp = (L[k, "PreparedStatementCacheAccessCount"] + 0 > 0) ? f1(L[k, "PreparedStatementCacheHitCount"] * 100 / L[k, "PreparedStatementCacheAccessCount"]) : ""
        # リーク疑い: 30 分以上の収集で、増加傾向(期間中に +2 本以上)かつ下限(最小値)も上昇している
        tr = (dur >= 1800) ? trend(k) : ""; leak = "no"
        if (vn >= 8 && tr != "" && tr + 0 > 0 && (tr + 0) * dur / 3600 >= 2) {
            qn4 = int(vn / 4)
            if (quarter_min(k, vn - qn4 + 1, vn) > quarter_min(k, 1, qn4)) leak = "yes"
        }
        if (vn == 0) { vc = (NOST[k] > 0) ? "NOSTATS" : "ERROR"; vd = (vc == "NOSTATS") ? "統計無効" : "取得失敗" }
        else if (exn > 0 || dB > 0) { vc = "EXHAUSTED"; vd = "枯渇あり" }
        else if (wn > 0 || dW > 0) { vc = "WAIT"; vd = "待機発生あり" }
        else if (hn > 0) { vc = "HIGH"; vd = "高負荷あり" }
        else if (leak == "yes") { vc = "HIGH"; vd = "リーク疑い" }
        else { vc = "OK"; vd = "正常" }
        c = ""
        if (vn == 0 && vc == "NOSTATS") c = c "statistics-enabled=false のため統計値を取得できません。--no-enable-stats を外して実行するか設定で有効化してください。"
        if (vn == 0 && vc == "ERROR") c = c "統計の取得に失敗しました。collector.log を確認してください。"
        if (dB > 0) c = c "接続取得タイムアウト(BlockingFailure)が" dB "件発生。"
        if (exn > 0) c = c "枯渇状態を" exn "サンプル検出(約" EXT[k] + 0 "秒、最長連続" LONG[k] "秒、初回 " FEXH[k] ")。"
        if (dW > 0) c = c "接続待機が" dW "件発生(最大同時待機" (L[k, "MaxWaitCount"] + 0) "スレッド)。"
        if (vn > 0 && amax(k, "usage") != "" && amax(k, "usage") + 0 >= warn) c = c "最大使用率" amax(k, "usage") "%(閾値" warn "%)。"
        if (exn > 0 || dB > 0) c = c "max-pool-size(現在" mx ")の拡張、長時間保持処理やコネクションリークの調査を推奨。"
        else if (wn > 0 || dW > 0) c = c "ピーク時に接続が不足しています。max-pool-size の見直しを推奨。"
        if (leak == "yes") c = c "使用中接続数が増加傾向(+" tr "本/時)で下限も上昇しており、コネクションリークの可能性があります。"
        if (L[k, "MaxUsageTime"] + 0 >= 60000) c = c "最大使用時間" L[k, "MaxUsageTime"] "msの長時間保持あり。"
        if (L[k, "MaxCreationTime"] + 0 >= 3000) c = c "接続作成に最大" L[k, "MaxCreationTime"] "ms要しています(DB/ネットワーク確認推奨)。"
        if (dD > dT) c = c "アイドルタイムアウト以外の接続破棄が" (dD - dT) "件(検証失敗・例外・フラッシュ等)。"
        if (RST[k] > 0) c = c "期間中にカウンタのリセットを" RST[k] "回検出(再起動等)。"
        if (ERRN[k] > 0 && vn > 0) c = c "取得失敗" ERRN[k] "回。"
        if (MAXDEF[k]) c = c "max-pool-size未設定のため既定値20で評価。"
        if (vc == "OK") c = c "期間中に枯渇・待機は検出されませんでした。"
        line = q(DSN[k]) "," DST[k] "," q(JNDI[k]) "," vc "," q(vd) "," MINP[k] "," MAXP[k] "," FTS[k] "," LTS[k] "," dur "," N[k] "," vn "," ERRN[k] + 0 "," NOST[k] + 0
        line = line "," amin(k, "inuse") "," amax(k, "inuse") "," aavg(k, "inuse") "," pk "," amax(k, "usage") "," aavg(k, "usage") "," upct95
        line = line "," amin(k, "active") "," amax(k, "active") "," aavg(k, "active") "," amin(k, "idle") "," amax(k, "idle") "," aavg(k, "idle") "," amin(k, "avail")
        line = line "," L[k, "MaxUsedCount"] "," peak "," hn "," wn "," exn "," ((vn > 0) ? f1(exn * 100 / vn) : "") "," EXT[k] + 0 "," LONG[k] + 0 "," FEXH[k] "," LEXH[k]
        line = line "," amax(k, "queued") "," dW "," dB "," dT "," S[k, "d_CreatedCount"] + 0 "," dD "," S[k, "d_TotalBlockingTime"] + 0
        line = line "," L[k, "MaxWaitCount"] "," L[k, "MaxWaitTime"] "," L[k, "AverageBlockingTime"] "," L[k, "AverageGetTime"] "," L[k, "MaxGetTime"]
        line = line "," L[k, "AverageUsageTime"] "," L[k, "MaxUsageTime"] "," L[k, "AverageCreationTime"] "," L[k, "MaxCreationTime"]
        line = line "," hitp "," tr "," leak "," RST[k] + 0 "," q(c)
        print line > sumf
    }
    print "bucket_start,bucket_end,datasource,ds_type,samples,max_pool_size,inuse_min,inuse_max,inuse_avg,usage_pct_max,usage_pct_avg,active_min,active_max,active_avg,idle_min,idle_max,idle_avg,available_min,queued_threads_max,high_samples,wait_samples,exhausted_samples,error_samples,d_WaitCount,d_BlockingFailureCount,d_TimedOut,d_CreatedCount,d_DestroyedCount,d_TotalBlockingTime,MaxWaitCount,MaxWaitTime,MaxGetTime,MaxUsageTime,MaxCreationTime,status_worst" > hourf
    for (i = 1; i <= nds; i++) {
        k = DSO[i]
        for (j = 1; j <= NB[k]; j++) {
            b = BO[k, j]; bk = k SEP b
            line = b "," bucket_end(b) "," q(DSN[k]) "," DST[k] "," BN[bk] "," BMAXP[bk]
            line = line "," amin(bk, "inuse") "," amax(bk, "inuse") "," aavg(bk, "inuse") "," amax(bk, "usage") "," aavg(bk, "usage")
            line = line "," amin(bk, "active") "," amax(bk, "active") "," aavg(bk, "active") "," amin(bk, "idle") "," amax(bk, "idle") "," aavg(bk, "idle")
            line = line "," amin(bk, "avail") "," amax(bk, "queued") "," CNT[bk, "HIGH"] + 0 "," CNT[bk, "WAIT"] + 0 "," CNT[bk, "EXHAUSTED"] + 0 "," BERR[bk] + 0
            line = line "," S[bk, "d_WaitCount"] + 0 "," S[bk, "d_BlockingFailureCount"] + 0 "," S[bk, "d_TimedOut"] + 0 "," S[bk, "d_CreatedCount"] + 0 "," S[bk, "d_DestroyedCount"] + 0 "," S[bk, "d_TotalBlockingTime"] + 0
            line = line "," L[bk, "MaxWaitCount"] "," L[bk, "MaxWaitTime"] "," L[bk, "MaxGetTime"] "," L[bk, "MaxUsageTime"] "," L[bk, "MaxCreationTime"]
            line = line "," ((bk in BW) ? BW[bk] : ((BERR[bk] > 0) ? "ERROR" : "NOSTATS"))
            print line > hourf
        }
    }
    # 開始時刻順に並べ替え (挿入ソート)
    for (i = 1; i <= nev; i++) IDX[i] = i
    for (i = 2; i <= nev; i++) {
        x = IDX[i]; j = i - 1
        while (j >= 1 && E_SEP[IDX[j]] > E_SEP[x]) { IDX[j + 1] = IDX[j]; j-- }
        IDX[j + 1] = x
    }
    print "event_no,datasource,ds_type,severity,severity_label,start_time,end_time,duration_sec,samples,inuse_max,max_pool_size,usage_pct_max,queued_threads_max,d_WaitCount,d_BlockingFailureCount,MaxWaitCount,MaxWaitTime" > evf
    for (i = 1; i <= nev; i++) {
        e = IDX[i]; k = E_K[e]
        print i "," q(DSN[k]) "," DST[k] "," E_SEV[e] "," q(LBL[E_SEV[e]]) "," E_STS[e] "," E_ETS[e] "," E_DUR[e] "," E_N[e] "," E_IN[e] "," E_MAX[e] "," f1(E_USE[e]) "," E_Q[e] "," E_W[e] "," E_B[e] "," E_MWC[e] "," E_MWT[e] > evf
    }
}
AWK

# --- スレッドダンプ (平坦化) → テキスト / 分類 CSV --------------------------------
cat > "$WORK_DIR/threaddump.awk" <<'AWK'
# 変数 : txtf csvf dtime dfile trigger
BEGIN { FS = "\t"; SEP = "\034"; cur = ""; nthr = 0; nj = 0; nw = 0; nc = 0 }
function unesc(s) { gsub(/\\"/, "\"", s); gsub(/\\n/, " ", s); gsub(/\\t/, " ", s); return s }
function q(s) { gsub(/"/, "\"\"", s); return "\"" s "\"" }
function nn(s) { return (s == "null") ? "" : s }
function is_fw(f) {
    return f ~ /^(java\.|javax\.|jakarta\.|jdk\.|sun\.|com\.sun\.|org\.jboss\.|org\.wildfly\.|io\.undertow\.|org\.hibernate\.|org\.apache\.jasper\.|org\.apache\.catalina\.|org\.apache\.cxf\.|org\.apache\.commons\.|org\.springframework\.|io\.smallrye\.|org\.eclipse\.|com\.zaxxer\.|io\.opentelemetry\.|io\.micrometer\.|net\.bytebuddy\.|com\.fasterxml\.|org\.glassfish\.|io\.netty\.|org\.xnio\.|org\.infinispan\.|com\.arjuna\.|io\.agroal\.|io\.quarkus\.|datadog\.|com\.datadog\.|co\.elastic\.|com\.newrelic\.|software\.amazon\.|oracle\.jdbc\.|com\.mysql\.|org\.postgresql\.|com\.microsoft\.sqlserver\.|org\.mariadb\.|com\.ibm\.db2\.|org\.h2\.)/
}
function flush(   j, cls, fr, cat, jf, af, wait, create, jdbcidx, ln) {
    if (cur == "") return
    nthr++
    printf "\"%s\" id=%s state=%s", unesc(T["thread-name"]), T["thread-id"], T["thread-state"] >> txtf
    if (nn(T["lock-name"]) != "") printf " waiting-on=%s", unesc(T["lock-name"]) >> txtf
    if (nn(T["lock-owner-name"]) != "") printf " owner=\"%s\"", unesc(T["lock-owner-name"]) >> txtf
    printf "\n" >> txtf
    wait = 0; create = 0; jdbcidx = 0
    for (j = 0; j < nfr; j++) {
        cls = F[j, "class-name"]
        fr = cls "." F[j, "method-name"] "(" nn(F[j, "file-name"]) ":" F[j, "line-number"] ")"
        FR[j] = fr
        print "    at " fr >> txtf
        if (cls ~ /^org\.jboss\.jca\.core\.connectionmanager\.pool\./ && F[j, "method-name"] == "tryAcquire") wait = 1
        if (F[j, "method-name"] == "createManagedConnection") create = 1
        if (jdbcidx == 0 && cls ~ /^org\.jboss\.jca\.adapters\.jdbc\./) jdbcidx = j + 1
    }
    printf "\n" >> txtf
    cat = ""
    if (wait) { cat = "接続待ち"; nw++ }
    else if (create) { cat = "接続作成中"; nc++ }
    else if (jdbcidx > 0) { cat = "JDBC実行中"; nj++ }
    if (cat != "") {
        jf = (jdbcidx > 0) ? FR[jdbcidx - 1] : ""
        af = ""
        for (j = (jdbcidx > 0 ? jdbcidx : 0); j < nfr; j++) if (!is_fw(F[j, "class-name"])) { af = FR[j]; break }
        print q(dtime) "," q(dfile) "," q(trigger) "," q(unesc(T["thread-name"])) "," q(T["thread-state"]) "," q(cat) "," q(jf) "," q(af) "," q(unesc(nn(T["lock-name"]))) >> csvf
    }
    delete T; delete F; delete FR; nfr = 0
}
/^@@/ { next }
$1 == "L" {
    n = split($2, p, SEP)
    if (p[1] != cur) { flush(); cur = p[1] }
    if (n == 2) T[p[2]] = $4
    else if (n == 4 && p[2] == "stack-trace") { j = p[3] + 0; F[j, p[4]] = $4; if (j + 1 > nfr) nfr = j + 1 }
}
END { flush(); print nthr "\t" nj "\t" nw "\t" nc }
AWK
}

write_java() {
cat > "$WORK_DIR/PoolJmx.java" <<'JAVA'
import java.io.File;
import java.io.FileDescriptor;
import java.io.FileOutputStream;
import java.io.PrintStream;
import java.lang.management.ThreadInfo;
import java.util.*;
import javax.management.*;
import javax.management.openmbean.CompositeData;
import javax.management.remote.*;

/** JMX (remote+http) で IronJacamar プール統計を取得し、jboss-cli 平坦化形式で出力する */
public class PoolJmx {
    static final String SEP = "\u001c";
    static PrintStream out;
    static MBeanServerConnection mbs;
    static final List<String[]> enabledByUs = new ArrayList<>();   // {objectName}

    public static void main(String[] a) throws Exception {
        out = new PrintStream(new FileOutputStream(FileDescriptor.out), false, "UTF-8");
        String url = a[0];
        long intervalMs = Long.parseLong(a[1]) * 1000L;
        long count = Long.parseLong(a[2]);
        boolean enable = "1".equals(a[3]);
        final boolean restore = "1".equals(a[4]);
        boolean queued = "1".equals(a[5]);
        long rediscoverMs = Long.parseLong(a[6]) * 1000L;
        Set<String> filter = new HashSet<>();
        for (String s : a[7].split(",")) if (!s.trim().isEmpty()) filter.add(s.trim());
        File tdRequest = new File(a[8]);
        File stopRequest = new File(a[9]);

        Map<String, Object> env = new HashMap<>();
        String user = System.getenv("JBOSS_CLI_USER");
        if (user != null && !user.isEmpty()) {
            String pw = System.getenv("JBOSS_CLI_PASSWORD");
            env.put(JMXConnector.CREDENTIALS, new String[] {user, pw == null ? "" : pw});
        }
        final JMXConnector conn;
        try {
            conn = JMXConnectorFactory.connect(new JMXServiceURL(url), env);
            mbs = conn.getMBeanServerConnection();
        } catch (Exception e) {
            out.println("@@FATAL|" + oneLine(e));
            out.flush();
            System.exit(2);
            return;
        }
        Runtime.getRuntime().addShutdownHook(new Thread(() -> {
            if (restore) restoreStats();
            try { conn.close(); } catch (Exception ignore) { }
            out.flush();
        }));
        out.println("@@READY");
        List<ObjectName> targets = discover(filter, enable);
        long start = System.currentTimeMillis();
        long due = start;
        long lastDisc = start;
        for (long i = 0; (count == 0 || i < count) && !stopRequest.exists(); i++) {
            while (System.currentTimeMillis() < due && !stopRequest.exists()) {
                Thread.sleep(Math.min(200L, due - System.currentTimeMillis()));
            }
            if (stopRequest.exists()) break;
            if (System.currentTimeMillis() - lastDisc >= rediscoverMs) {
                targets = discover(filter, enable);
                lastDisc = System.currentTimeMillis();
            }
            sample(targets, queued);
            if (tdRequest.exists()) {
                tdRequest.delete();
                threadDump();
            }
            due += intervalMs;
            long now = System.currentTimeMillis();
            if (due < now) due = now;
        }
        if (restore) restoreStats();
        out.println("@@EOF");
        out.flush();
        System.exit(0);
    }

    static String oneLine(Throwable e) {
        String m = e.getClass().getName() + ": " + e.getMessage();
        if (e.getCause() != null) m += " / " + e.getCause();
        return m.replace('\n', ' ').replace('\r', ' ').replace('\t', ' ');
    }

    static String esc(String s) {
        StringBuilder b = new StringBuilder(s.length() + 16);
        for (int i = 0; i < s.length(); i++) {
            char c = s.charAt(i);
            switch (c) {
                case '\\': b.append("\\\\"); break;
                case '"': b.append("\\\""); break;
                case '\n': b.append("\\n"); break;
                case '\r': b.append("\\r"); break;
                case '\t': b.append("\\t"); break;
                default: b.append(c);
            }
        }
        return b.toString();
    }

    static void leaf(String path, Object v) {
        if (v == null) out.println("L\t" + path + "\tN\tnull");
        else if (v instanceof Number || v instanceof Boolean) out.println("L\t" + path + "\tN\t" + v);
        else if (v instanceof String[]) {
            String[] arr = (String[]) v;
            for (int i = 0; i < arr.length; i++) leaf(path + (path.isEmpty() ? "" : SEP) + i, arr[i]);
        } else out.println("L\t" + path + "\tS\t" + esc(String.valueOf(v)));
    }

    /** camelCase → hyphen-case (先頭が大文字の統計属性名はそのまま) */
    static String hyphen(String n) {
        if (n.isEmpty() || Character.isUpperCase(n.charAt(0))) return n;
        StringBuilder b = new StringBuilder();
        for (char c : n.toCharArray()) {
            if (Character.isUpperCase(c)) b.append('-').append(Character.toLowerCase(c));
            else b.append(c);
        }
        return b.toString();
    }

    static Map<String, Object> attrs(ObjectName on) throws Exception {
        MBeanAttributeInfo[] infos = mbs.getMBeanInfo(on).getAttributes();
        String[] names = new String[infos.length];
        for (int i = 0; i < infos.length; i++) names[i] = infos[i].getName();
        Map<String, Object> m = new TreeMap<>();
        for (Attribute at : mbs.getAttributes(on, names).asList()) m.put(hyphen(at.getName()), at.getValue());
        return m;
    }

    static List<ObjectName> discover(Set<String> filter, boolean enable) {
        List<ObjectName> targets = new ArrayList<>();
        out.println("@@DISCOVERY");
        out.println("@@INFO");
        try {
            for (Map.Entry<String, Object> e : attrs(new ObjectName("jboss.as:management-root=server")).entrySet()) {
                Object v = e.getValue();
                if (v == null || v instanceof String || v instanceof Number || v instanceof Boolean) leaf(e.getKey(), v);
            }
            out.println("O\tsuccess");
        } catch (Exception e) {
            out.println("O\tfailed");
            out.println("F\t" + oneLine(e));
        }
        out.println("@@DS");
        try {
            for (String type : new String[] {"data-source", "xa-data-source"}) {
                Set<ObjectName> set = new TreeSet<>(mbs.queryNames(new ObjectName("jboss.as:subsystem=datasources," + type + "=*"), null));
                for (ObjectName on : set) {
                    String name = on.getKeyProperty(type);
                    String base = type + SEP + name + SEP;
                    Map<String, Object> m = attrs(on);
                    for (Map.Entry<String, Object> e : m.entrySet()) {
                        Object v = e.getValue();
                        if (v instanceof CompositeData || v instanceof Object[] && !(v instanceof String[])) v = String.valueOf(v);
                        leaf(base + e.getKey(), v);
                    }
                    for (String child : new String[] {"connection-properties", "xa-datasource-properties"}) {
                        ObjectName pat = new ObjectName(on.getDomain() + ":" + on.getKeyPropertyListString() + "," + child + "=*");
                        for (ObjectName c : mbs.queryNames(pat, null)) {
                            leaf(base + child + SEP + c.getKeyProperty(child) + SEP + "value", mbs.getAttribute(c, "value"));
                        }
                    }
                    boolean en = !Boolean.FALSE.equals(m.get("enabled"));
                    if (en && (filter.isEmpty() || filter.contains(name))) {
                        targets.add(on);
                        if (enable) {
                            enableStats(on, "pool", type, name);
                            enableStats(on, "jdbc", type, name);
                        }
                    }
                }
            }
            for (ObjectName on : mbs.queryNames(new ObjectName("jboss.as:subsystem=datasources,jdbc-driver=*"), null)) {
                String base = "jdbc-driver" + SEP + on.getKeyProperty("jdbc-driver") + SEP;
                for (Map.Entry<String, Object> e : attrs(on).entrySet()) {
                    Object v = e.getValue();
                    if (v == null || v instanceof String || v instanceof Number || v instanceof Boolean) leaf(base + e.getKey(), v);
                }
            }
            out.println("O\tsuccess");
        } catch (Exception e) {
            out.println("O\tfailed");
            out.println("F\t" + oneLine(e));
        }
        out.println("@@ENDDISCOVERY");
        out.flush();
        return targets;
    }

    static ObjectName stats(ObjectName ds, String kind) throws MalformedObjectNameException {
        return new ObjectName(ds.getDomain() + ":" + ds.getKeyPropertyListString() + ",statistics=" + kind);
    }

    static void enableStats(ObjectName ds, String kind, String type, String name) {
        try {
            ObjectName s = stats(ds, kind);
            if (Boolean.FALSE.equals(mbs.getAttribute(s, "statisticsEnabled"))) {
                mbs.setAttribute(s, new Attribute("statisticsEnabled", Boolean.TRUE));
                synchronized (enabledByUs) { enabledByUs.add(new String[] {s.getCanonicalName()}); }
                out.println("@@ENABLED|" + kind + "|" + type + "|" + name);
            }
        } catch (Exception e) {
            out.println("@@WARN|statistics " + kind + " enable failed: " + name + ": " + oneLine(e));
        }
    }

    static void restoreStats() {
        synchronized (enabledByUs) {
            for (String[] s : enabledByUs) {
                try {
                    mbs.setAttribute(new ObjectName(s[0]), new Attribute("statisticsEnabled", Boolean.FALSE));
                    out.println("@@RESTORED|" + s[0]);
                } catch (Exception e) {
                    out.println("@@WARN|restore failed: " + s[0] + ": " + oneLine(e));
                }
            }
            enabledByUs.clear();
        }
    }

    static void section(String marker, ObjectName on) {
        out.println(marker);
        try {
            for (Map.Entry<String, Object> e : attrs(on).entrySet()) leaf(e.getKey(), e.getValue());
            out.println("O\tsuccess");
        } catch (Exception e) {
            out.println("@@ERR");
            out.println("F\t" + oneLine(e));
        }
    }

    static void sample(List<ObjectName> targets, boolean queued) {
        out.println("@@SAMPLE|" + System.currentTimeMillis());
        for (ObjectName ds : targets) {
            String type = ds.getKeyProperty("data-source") != null ? "data-source" : "xa-data-source";
            out.println("@@DS|" + type + "|" + ds.getKeyProperty(type));
            try {
                section("@@P", stats(ds, "pool"));
                section("@@J", stats(ds, "jdbc"));
            } catch (MalformedObjectNameException e) {
                out.println("@@ERR");
            }
            if (queued) {
                out.println("@@Q");
                try {
                    Object r = mbs.invoke(ds, "dumpQueuedThreadsInPool", new Object[0], new String[0]);
                    if (r instanceof String) {
                        // JMX ファサードはリストを JSON 文字列で返す: エスケープ済みのまま出力
                        List<String> items = jsonStrings((String) r);
                        for (int i = 0; i < items.size(); i++) out.println("L\t" + i + "\tS\t" + items.get(i));
                    } else if (r != null) leaf("", r);
                    out.println("O\tsuccess");
                } catch (Exception e) {
                    out.println("@@ERR");
                }
            }
            out.println("@@ENDDS");
        }
        out.println("@@ENDSAMPLE");
        out.flush();
    }

    /** JSON 配列文字列から文字列要素をエスケープされた状態のまま取り出す */
    static List<String> jsonStrings(String s) {
        List<String> list = new ArrayList<>();
        int i = 0, n = s.length();
        while (i < n) {
            if (s.charAt(i) != '"') { i++; continue; }
            int j = i + 1;
            while (j < n && s.charAt(j) != '"') j += (s.charAt(j) == '\\') ? 2 : 1;
            list.add(s.substring(i + 1, Math.min(j, n)));
            i = j + 1;
        }
        return list;
    }

    static void threadDump() {
        out.println("@@TD|" + System.currentTimeMillis());
        try {
            Object r = mbs.invoke(new ObjectName("java.lang:type=Threading"), "dumpAllThreads",
                    new Object[] {Boolean.FALSE, Boolean.FALSE}, new String[] {"boolean", "boolean"});
            CompositeData[] arr = (CompositeData[]) r;
            for (int i = 0; i < arr.length; i++) {
                ThreadInfo ti = ThreadInfo.from(arr[i]);
                String p = i + SEP;
                leaf(p + "thread-id", ti.getThreadId());
                leaf(p + "thread-name", ti.getThreadName());
                leaf(p + "thread-state", String.valueOf(ti.getThreadState()));
                leaf(p + "lock-name", ti.getLockName());
                leaf(p + "lock-owner-name", ti.getLockOwnerName());
                StackTraceElement[] st = ti.getStackTrace();
                for (int j = 0; j < st.length; j++) {
                    String f = p + "stack-trace" + SEP + j + SEP;
                    leaf(f + "file-name", st[j].getFileName());
                    leaf(f + "line-number", st[j].getLineNumber());
                    leaf(f + "class-name", st[j].getClassName());
                    leaf(f + "method-name", st[j].getMethodName());
                }
            }
            out.println("O\tsuccess");
        } catch (Exception e) {
            out.println("O\tfailed");
            out.println("F\t" + oneLine(e));
        }
        out.println("@@ENDTD");
        out.flush();
    }
}
JAVA
}

# ------------------------------------------------------------------------------
# 項目定義 CSV (Excel 変換スクリプトが見出し・書式・説明に使用)
#   file,key,label,unit,format,description
#   format: datetime / int / dec1 / pct / text / wrap / status
# ------------------------------------------------------------------------------
write_columns_csv() {
    {
        if [ "$WRITE_BOM" = 1 ]; then printf '\357\273\277'; fi
        cat <<'CSV'
file,key,label,unit,format,description
timeseries,timestamp,取得日時,,datetime,統計を取得した日時 (--tz のタイムゾーン)
timeseries,epoch,エポック秒,秒,int,取得日時の UNIX 時刻
timeseries,elapsed_sec,経過秒,秒,int,収集開始からの経過秒数
timeseries,sample_no,サンプル番号,,int,取得の通し番号
timeseries,host,ホスト,,text,ECS タスクID またはホスト名
timeseries,datasource,データソース,,text,データソース名
timeseries,ds_type,種別,,text,data-source (非XA) / xa-data-source (XA)
timeseries,jndi_name,JNDI名,,text,データソースの JNDI 名
timeseries,stats_enabled,統計有効,,text,statistics=pool の statistics-enabled (実行時の値)
timeseries,status,状態,,status,OK=正常 / HIGH=高負荷(使用率が閾値以上) / WAIT=待機発生 / EXHAUSTED=枯渇 / NOSTATS=統計無効 / ERROR=取得失敗
timeseries,usage_pct,使用率,%,pct,使用中接続数 ÷ 最大プールサイズ × 100
timeseries,InUseCount,使用中接続数,本,int,アプリケーションが現在使用中の接続数 (InUseCount)
timeseries,ActiveCount,作成済接続数,本,int,プール内に作成済みの物理接続数 (使用中 + アイドル) (ActiveCount)
timeseries,IdleCount,アイドル接続数,本,int,プール内で未使用の接続数 (IdleCount) ※瞬間的に負値になる場合あり
timeseries,AvailableCount,取得可能数,本,int,最大プールサイズまでに追加で取得できる接続数 (AvailableCount) 0 は枯渇
timeseries,MaxUsedCount,最大使用数(累積),本,int,統計開始以降の使用中接続数の最大値 (MaxUsedCount)
timeseries,min_pool_size,最小プールサイズ,本,int,min-pool-size (未設定時は 0)
timeseries,max_pool_size,最大プールサイズ,本,int,max-pool-size (未設定時は既定値 20)
timeseries,peak_usage_pct,ピーク使用率(累積),%,pct,MaxUsedCount ÷ 最大プールサイズ × 100
timeseries,queued_threads,接続待ちスレッド数,件,int,取得時点で接続の空きを待っているスレッド数 (dump-queued-threads-in-pool)
timeseries,d_WaitCount,待機発生数(区間),件,int,前回取得からの WaitCount 増分 (接続取得で待たされた要求数)
timeseries,d_BlockingFailureCount,取得タイムアウト数(区間),件,int,前回取得からの BlockingFailureCount 増分 (blocking-timeout 超過で取得失敗 = 枯渇)
timeseries,d_TimedOut,アイドルタイムアウト数(区間),件,int,前回取得からの TimedOut 増分 (アイドルタイムアウトで閉じた接続)
timeseries,d_CreatedCount,接続作成数(区間),件,int,前回取得からの CreatedCount 増分
timeseries,d_DestroyedCount,接続破棄数(区間),件,int,前回取得からの DestroyedCount 増分
timeseries,d_TotalBlockingTime,待機時間合計(区間),ms,int,前回取得からの TotalBlockingTime 増分
timeseries,d_TotalGetTime,取得時間合計(区間),ms,int,前回取得からの TotalGetTime 増分
timeseries,d_TotalUsageTime,使用時間合計(区間),ms,int,前回取得からの TotalUsageTime 増分 (÷区間ミリ秒 ≒ 平均同時使用数)
timeseries,WaitCount,待機発生数(累積),件,int,接続取得で待たされた要求数の累計 (WaitCount)
timeseries,MaxWaitCount,最大同時待機数,件,int,同時に接続を待っていたスレッド数の最大値 (MaxWaitCount)
timeseries,BlockingFailureCount,取得タイムアウト数(累積),件,int,接続取得タイムアウトの累計 (BlockingFailureCount)
timeseries,TimedOut,アイドルタイムアウト数(累積),件,int,アイドルタイムアウトで閉じた接続の累計 (TimedOut)
timeseries,CreatedCount,接続作成数(累積),件,int,物理接続の作成数の累計 (CreatedCount)
timeseries,DestroyedCount,接続破棄数(累積),件,int,物理接続の破棄数の累計 (DestroyedCount)
timeseries,AverageBlockingTime,平均待機時間,ms,int,接続取得時の平均待機時間 (AverageBlockingTime)
timeseries,MaxWaitTime,最大待機時間,ms,int,接続取得時の最大待機時間 (MaxWaitTime)
timeseries,TotalBlockingTime,待機時間合計,ms,int,接続取得時の待機時間の合計 (TotalBlockingTime)
timeseries,AverageGetTime,平均取得時間,ms,int,接続取得に要した平均時間 (AverageGetTime)
timeseries,MaxGetTime,最大取得時間,ms,int,接続取得に要した最大時間 (MaxGetTime)
timeseries,TotalGetTime,取得時間合計,ms,int,接続取得に要した時間の合計 (TotalGetTime)
timeseries,AverageUsageTime,平均使用時間,ms,int,アプリが接続を保持した平均時間 (AverageUsageTime)
timeseries,MaxUsageTime,最大使用時間,ms,int,アプリが接続を保持した最大時間 (MaxUsageTime)
timeseries,TotalUsageTime,使用時間合計,ms,int,アプリが接続を保持した時間の合計 (TotalUsageTime)
timeseries,AverageCreationTime,平均接続作成時間,ms,int,物理接続の作成に要した平均時間 (AverageCreationTime)
timeseries,MaxCreationTime,最大接続作成時間,ms,int,物理接続の作成に要した最大時間 (MaxCreationTime)
timeseries,TotalCreationTime,接続作成時間合計,ms,int,物理接続の作成に要した時間の合計 (TotalCreationTime)
timeseries,AveragePoolTime,平均プール滞留時間,ms,int,接続がプール内でアイドルだった平均時間 (AveragePoolTime)
timeseries,MaxPoolTime,最大プール滞留時間,ms,int,接続がプール内でアイドルだった最大時間 (MaxPoolTime)
timeseries,TotalPoolTime,プール滞留時間合計,ms,int,接続がプール内でアイドルだった時間の合計 (TotalPoolTime)
timeseries,PreparedStatementCacheAccessCount,PSキャッシュ参照数,件,int,PreparedStatement キャッシュの参照数
timeseries,PreparedStatementCacheHitCount,PSキャッシュヒット数,件,int,PreparedStatement キャッシュのヒット数
timeseries,PreparedStatementCacheMissCount,PSキャッシュミス数,件,int,PreparedStatement キャッシュのミス数
timeseries,PreparedStatementCacheAddCount,PSキャッシュ追加数,件,int,PreparedStatement キャッシュへの追加数
timeseries,PreparedStatementCacheDeleteCount,PSキャッシュ削除数,件,int,PreparedStatement キャッシュからの削除数
timeseries,PreparedStatementCacheCurrentSize,PSキャッシュ現在サイズ,件,int,PreparedStatement キャッシュの現在のエントリ数
timeseries,ps_cache_hit_pct,PSキャッシュヒット率,%,pct,ヒット数 ÷ 参照数 × 100
timeseries,XACommitCount,XAコミット数,件,int,XA commit の回数
timeseries,XACommitAverageTime,XAコミット平均時間,ms,int,XA commit の平均時間
timeseries,XACommitMaxTime,XAコミット最大時間,ms,int,XA commit の最大時間
timeseries,XAPrepareCount,XAプリペア数,件,int,XA prepare の回数
timeseries,XAPrepareAverageTime,XAプリペア平均時間,ms,int,XA prepare の平均時間
timeseries,XAPrepareMaxTime,XAプリペア最大時間,ms,int,XA prepare の最大時間
timeseries,XARollbackCount,XAロールバック数,件,int,XA rollback の回数
timeseries,XARollbackAverageTime,XAロールバック平均時間,ms,int,XA rollback の平均時間
timeseries,XARollbackMaxTime,XAロールバック最大時間,ms,int,XA rollback の最大時間
timeseries,XAStartCount,XAスタート数,件,int,XA start の回数
timeseries,XAStartAverageTime,XAスタート平均時間,ms,int,XA start の平均時間
timeseries,XAStartMaxTime,XAスタート最大時間,ms,int,XA start の最大時間
timeseries,XAEndCount,XAエンド数,件,int,XA end の回数
timeseries,XAEndAverageTime,XAエンド平均時間,ms,int,XA end の平均時間
timeseries,XAEndMaxTime,XAエンド最大時間,ms,int,XA end の最大時間
timeseries,XARecoverCount,XAリカバリ数,件,int,XA recover の回数
timeseries,XAForgetCount,XAフォーゲット数,件,int,XA forget の回数
timeseries,note,備考,,text,カウンタリセット・取得失敗などの補足
hourly,bucket_start,時間帯(開始),,datetime,集計単位(--bucket-min)の開始日時
hourly,bucket_end,時間帯(終了),,text,集計単位の終了時刻
hourly,datasource,データソース,,text,データソース名
hourly,ds_type,種別,,text,data-source / xa-data-source
hourly,samples,サンプル数,件,int,時間帯内の取得回数
hourly,max_pool_size,最大プールサイズ,本,int,max-pool-size
hourly,inuse_min,使用中 最小,本,int,使用中接続数の最小値
hourly,inuse_max,使用中 最大,本,int,使用中接続数の最大値
hourly,inuse_avg,使用中 平均,本,dec1,使用中接続数の平均値
hourly,usage_pct_max,使用率 最大,%,pct,使用率の最大値
hourly,usage_pct_avg,使用率 平均,%,pct,使用率の平均値
hourly,active_min,作成済 最小,本,int,作成済接続数の最小値
hourly,active_max,作成済 最大,本,int,作成済接続数の最大値
hourly,active_avg,作成済 平均,本,dec1,作成済接続数の平均値
hourly,idle_min,アイドル 最小,本,int,アイドル接続数の最小値
hourly,idle_max,アイドル 最大,本,int,アイドル接続数の最大値
hourly,idle_avg,アイドル 平均,本,dec1,アイドル接続数の平均値
hourly,available_min,取得可能数 最小,本,int,取得可能数の最小値 (0 は枯渇)
hourly,queued_threads_max,接続待ちスレッド 最大,件,int,接続待ちスレッド数の最大値
hourly,high_samples,高負荷サンプル数,件,int,状態が HIGH のサンプル数
hourly,wait_samples,待機サンプル数,件,int,状態が WAIT のサンプル数
hourly,exhausted_samples,枯渇サンプル数,件,int,状態が EXHAUSTED のサンプル数
hourly,error_samples,取得失敗数,件,int,状態が ERROR のサンプル数
hourly,d_WaitCount,待機発生数,件,int,時間帯内の WaitCount 増分の合計
hourly,d_BlockingFailureCount,取得タイムアウト数,件,int,時間帯内の BlockingFailureCount 増分の合計
hourly,d_TimedOut,アイドルタイムアウト数,件,int,時間帯内の TimedOut 増分の合計
hourly,d_CreatedCount,接続作成数,件,int,時間帯内の CreatedCount 増分の合計
hourly,d_DestroyedCount,接続破棄数,件,int,時間帯内の DestroyedCount 増分の合計
hourly,d_TotalBlockingTime,待機時間合計,ms,int,時間帯内の TotalBlockingTime 増分の合計
hourly,MaxWaitCount,最大同時待機数(累積),件,int,時間帯末時点の MaxWaitCount
hourly,MaxWaitTime,最大待機時間(累積),ms,int,時間帯末時点の MaxWaitTime
hourly,MaxGetTime,最大取得時間(累積),ms,int,時間帯末時点の MaxGetTime
hourly,MaxUsageTime,最大使用時間(累積),ms,int,時間帯末時点の MaxUsageTime
hourly,MaxCreationTime,最大接続作成時間(累積),ms,int,時間帯末時点の MaxCreationTime
hourly,status_worst,最悪状態,,status,時間帯内で最も深刻な状態
summary,datasource,データソース,,text,データソース名
summary,ds_type,種別,,text,data-source / xa-data-source
summary,jndi_name,JNDI名,,text,JNDI 名
summary,verdict_code,判定コード,,status,OK / HIGH / WAIT / EXHAUSTED / NOSTATS / ERROR
summary,verdict,判定,,text,正常 / 高負荷あり / リーク疑い / 待機発生あり / 枯渇あり / 統計無効 / 取得失敗
summary,min_pool_size,最小プールサイズ,本,int,min-pool-size
summary,max_pool_size,最大プールサイズ,本,int,max-pool-size
summary,first_timestamp,収集開始,,datetime,最初のサンプルの取得日時
summary,last_timestamp,収集終了,,datetime,最後のサンプルの取得日時
summary,duration_sec,収集時間,秒,int,収集期間 (最終 - 最初 + 間隔)
summary,samples,サンプル数,件,int,取得回数
summary,valid_samples,有効サンプル数,件,int,統計値を取得できた回数
summary,error_samples,取得失敗数,件,int,取得に失敗した回数
summary,nostats_samples,統計無効数,件,int,統計が無効だった回数
summary,inuse_min,使用中 最小,本,int,使用中接続数の最小値
summary,inuse_max,使用中 最大,本,int,使用中接続数の最大値
summary,inuse_avg,使用中 平均,本,dec1,使用中接続数の平均値
summary,inuse_p95,使用中 95%値,本,int,使用中接続数の 95 パーセンタイル
summary,usage_pct_max,使用率 最大,%,pct,使用率の最大値
summary,usage_pct_avg,使用率 平均,%,pct,使用率の平均値
summary,usage_pct_p95,使用率 95%値,%,pct,使用中 95%値 ÷ 最大プールサイズ × 100
summary,active_min,作成済 最小,本,int,作成済接続数の最小値
summary,active_max,作成済 最大,本,int,作成済接続数の最大値
summary,active_avg,作成済 平均,本,dec1,作成済接続数の平均値
summary,idle_min,アイドル 最小,本,int,アイドル接続数の最小値
summary,idle_max,アイドル 最大,本,int,アイドル接続数の最大値
summary,idle_avg,アイドル 平均,本,dec1,アイドル接続数の平均値
summary,available_min,取得可能数 最小,本,int,取得可能数の最小値 (0 は枯渇)
summary,MaxUsedCount,最大使用数(累積),本,int,統計開始以降の使用中接続数の最大値
summary,peak_usage_pct,ピーク使用率(累積),%,pct,MaxUsedCount ÷ 最大プールサイズ × 100
summary,high_samples,高負荷サンプル数,件,int,状態が HIGH のサンプル数
summary,wait_samples,待機サンプル数,件,int,状態が WAIT のサンプル数
summary,exhausted_samples,枯渇サンプル数,件,int,状態が EXHAUSTED のサンプル数
summary,exhausted_ratio_pct,枯渇サンプル率,%,pct,枯渇サンプル数 ÷ 有効サンプル数 × 100
summary,exhausted_time_sec,枯渇時間(概算),秒,int,連続した枯渇期間(最終 - 最初 + 取得間隔)の合計
summary,longest_exhausted_sec,最長連続枯渇時間,秒,int,枯渇状態が連続した最長時間
summary,first_exhausted_at,初回枯渇日時,,datetime,最初に枯渇を検出した日時
summary,last_exhausted_at,最終枯渇日時,,datetime,最後に枯渇を検出した日時
summary,queued_threads_max,接続待ちスレッド 最大,件,int,接続待ちスレッド数の最大値
summary,d_WaitCount,待機発生数,件,int,期間中の WaitCount 増分
summary,d_BlockingFailureCount,取得タイムアウト数,件,int,期間中の BlockingFailureCount 増分 (> 0 は枯渇によるエラー発生)
summary,d_TimedOut,アイドルタイムアウト数,件,int,期間中の TimedOut 増分
summary,d_CreatedCount,接続作成数,件,int,期間中の CreatedCount 増分
summary,d_DestroyedCount,接続破棄数,件,int,期間中の DestroyedCount 増分
summary,d_TotalBlockingTime,待機時間合計,ms,int,期間中の TotalBlockingTime 増分
summary,MaxWaitCount,最大同時待機数,件,int,MaxWaitCount (最終値)
summary,MaxWaitTime,最大待機時間,ms,int,MaxWaitTime (最終値)
summary,AverageBlockingTime,平均待機時間,ms,int,AverageBlockingTime (最終値)
summary,AverageGetTime,平均取得時間,ms,int,AverageGetTime (最終値)
summary,MaxGetTime,最大取得時間,ms,int,MaxGetTime (最終値)
summary,AverageUsageTime,平均使用時間,ms,int,AverageUsageTime (最終値)
summary,MaxUsageTime,最大使用時間,ms,int,MaxUsageTime (最終値)
summary,AverageCreationTime,平均接続作成時間,ms,int,AverageCreationTime (最終値)
summary,MaxCreationTime,最大接続作成時間,ms,int,MaxCreationTime (最終値)
summary,ps_cache_hit_pct,PSキャッシュヒット率,%,pct,PreparedStatement キャッシュのヒット率 (最終値)
summary,inuse_trend_per_hour,使用中 増加傾向,本/時,dec1,使用中接続数の回帰直線の傾き (1 時間あたり, 収集 30 分以上で算出)
summary,leak_suspect,リーク疑い,,text,使用中接続数が期間中に 2 本以上増加傾向で、最小値(下限)も上昇している場合 yes
summary,counter_resets,カウンタリセット回数,回,int,累積カウンタの減少を検出した回数 (再起動等)
summary,comment,所見,,wrap,自動判定のコメント
events,event_no,No,,int,イベント番号 (開始日時順)
events,datasource,データソース,,text,データソース名
events,ds_type,種別,,text,data-source / xa-data-source
events,severity,重大度,,status,HIGH / WAIT / EXHAUSTED (期間中で最も深刻な状態)
events,severity_label,重大度(表示),,text,高負荷 / 待機発生 / 枯渇
events,start_time,開始日時,,datetime,連続した異常サンプルの最初の取得日時
events,end_time,終了日時,,datetime,連続した異常サンプルの最後の取得日時
events,duration_sec,継続時間,秒,int,終了 - 開始 + 取得間隔
events,samples,サンプル数,件,int,イベント中のサンプル数
events,inuse_max,使用中 最大,本,int,イベント中の使用中接続数の最大値
events,max_pool_size,最大プールサイズ,本,int,max-pool-size
events,usage_pct_max,使用率 最大,%,pct,イベント中の使用率の最大値
events,queued_threads_max,接続待ちスレッド 最大,件,int,イベント中の接続待ちスレッド数の最大値
events,d_WaitCount,待機発生数,件,int,イベント中の WaitCount 増分
events,d_BlockingFailureCount,取得タイムアウト数,件,int,イベント中の BlockingFailureCount 増分
events,MaxWaitCount,最大同時待機数(累積),件,int,イベント終了時点の MaxWaitCount
events,MaxWaitTime,最大待機時間(累積),ms,int,イベント終了時点の MaxWaitTime
config,datasource,データソース,,text,データソース名 (jdbc-driver はドライバ名)
config,ds_type,種別,,text,data-source / xa-data-source / jdbc-driver
config,category,分類,,text,主要 / その他 / 接続プロパティ / XAプロパティ / JDBCドライバ
config,attribute,属性,,text,管理モデルの属性名
config,label,説明,,text,主要属性の日本語名
config,value,設定値,,text,設定値 (パスワード等はマスク)
config,default_value,既定値,,text,未設定時に適用される既定値
config,note,備考,,text,未設定時の補足
queued,timestamp,取得日時,,datetime,接続待ちスレッドを検出した日時
queued,epoch,エポック秒,秒,int,UNIX 時刻
queued,sample_no,サンプル番号,,int,時系列のサンプル番号
queued,host,ホスト,,text,ECS タスクID またはホスト名
queued,datasource,データソース,,text,データソース名
queued,signature_id,スタックID,,int,待機スタックの識別番号 (待機スタックシート参照)
queued,waiting_threads,待機スレッド数,件,int,同じスタックで接続を待っていたスレッド数
queued,caller_frame,getConnection 呼出元,,text,getConnection を呼び出したフレーム
queued,app_frame,アプリケーションフレーム,,text,フレームワーク以外で最初に現れるフレーム
queued,thread_names,スレッド名(例),,text,待機していたスレッド名 (最大3件)
qstack,signature_id,スタックID,,int,待機スタックの識別番号
qstack,datasource,データソース,,text,データソース名
qstack,first_seen,初回検出日時,,datetime,このスタックを最初に検出した日時
qstack,frame_count,フレーム数,,int,スタックフレーム数
qstack,caller_frame,getConnection 呼出元,,text,getConnection を呼び出したフレーム
qstack,app_frame,アプリケーションフレーム,,text,フレームワーク以外で最初に現れるフレーム
qstack,stack,スタックトレース,,wrap,接続待ちスレッドのスタックトレース
threaddump,dump_time,ダンプ日時,,datetime,スレッドダンプの取得日時
threaddump,dump_file,ファイル,,text,スレッドダンプのテキストファイル名
threaddump,trigger,契機,,text,取得契機 (状態とデータソース)
threaddump,thread_name,スレッド名,,text,スレッド名
threaddump,thread_state,スレッド状態,,text,RUNNABLE / WAITING / TIMED_WAITING / BLOCKED
threaddump,category,分類,,text,接続待ち / 接続作成中 / JDBC実行中 (接続を保持して SQL 実行中)
threaddump,jdbc_frame,JDBCフレーム,,text,IronJacamar JDBC ラッパーの最上位フレーム
threaddump,app_frame,アプリケーションフレーム,,text,フレームワーク以外で最初に現れるフレーム
threaddump,lock_name,待機ロック,,text,待機しているロック
info,key,項目キー,,text,項目キー
info,label,項目,,text,収集条件・環境の項目名
info,value,値,,text,値
status,OK,正常,,status,使用率が閾値未満で待機・枯渇なし
status,HIGH,高負荷,,status,使用率が --warn-pct 以上
status,WAIT,待機発生,,status,接続取得の待機が発生 (WaitCount 増加 または 待機スレッドあり)
status,EXHAUSTED,枯渇,,status,使用中 = 最大プールサイズ / 取得可能数 0 / 取得タイムアウト発生
status,NOSTATS,統計無効,,status,statistics-enabled=false のため値なし
status,ERROR,取得失敗,,status,統計の取得に失敗
ui,report_title,JBoss EAP コネクションプール 調査レポート,,text,レポートタイトル
ui,sheet_summary,サマリ,,text,シート名
ui,sheet_charts,グラフ,,text,シート名
ui,sheet_hourly,時間別推移,,text,シート名
ui,sheet_events,枯渇イベント,,text,シート名
ui,sheet_timeseries,時系列,,text,シート名
ui,sheet_queued,待機スレッド,,text,シート名
ui,sheet_qstack,待機スタック,,text,シート名
ui,sheet_threaddump,スレッドダンプ,,text,シート名
ui,sheet_config,プール設定,,text,シート名
ui,sheet_columns,項目説明,,text,シート名
ui,sec_info,収集条件・環境,,text,見出し
ui,sec_summary,データソース別サマリ,,text,見出し
ui,sec_legend,状態の凡例,,text,見出し
ui,col_item,項目,,text,見出し
ui,col_unit,単位,,text,見出し
ui,col_sheet,シート,,text,見出し
ui,col_key,キー,,text,見出し
ui,col_desc,説明,,text,見出し
ui,col_format,書式,,text,見出し
ui,generated,作成日時,,text,ラベル
ui,source_dir,入力フォルダ,,text,ラベル
ui,chart_ts,接続数の推移,,text,グラフタイトル
ui,chart_hourly,時間別 使用中接続数,,text,グラフタイトル
ui,chart_hourly_err,時間別 待機・取得タイムアウト発生数,,text,グラフタイトル
ui,axis_conn,接続数(本),,text,軸タイトル
ui,axis_count,件数,,text,軸タイトル
ui,ser_inuse,使用中,,text,系列名
ui,ser_active,作成済,,text,系列名
ui,ser_idle,アイドル,,text,系列名
ui,ser_maxpool,最大プールサイズ,,text,系列名
ui,ser_inuse_max,使用中 最大,,text,系列名
ui,ser_inuse_avg,使用中 平均,,text,系列名
ui,ser_inuse_min,使用中 最小,,text,系列名
ui,ser_wait,待機発生数,,text,系列名
ui,ser_bf,取得タイムアウト数,,text,系列名
ui,no_data,データがありません,,text,メッセージ
ui,truncated,行数が Excel の上限を超えたため一部のみ出力しています,,text,メッセージ
ui,msg_created,Excel を作成しました,,text,メッセージ
ui,msg_sheets,シート,,text,メッセージ
CSV
    } > "$OUT_DIR/$F_COLUMNS"
}

# ------------------------------------------------------------------------------
# 収集条件・環境 CSV
# ------------------------------------------------------------------------------
INFO_PRODUCT_NAME=""; INFO_PRODUCT_VERSION=""; INFO_RELEASE_VERSION=""; INFO_SERVER_NAME=""; INFO_SERVER_STATE=""
START_STR=""; END_STR=""; SAMPLE_NO=0; PLANNED=""

load_info_state() {
    [ -f "$INFO_STATE" ] || return 0
    local k v
    while IFS=$'\t' read -r k v; do
        case "$k" in
            product-name)    INFO_PRODUCT_NAME=$v ;;
            product-version) INFO_PRODUCT_VERSION=$v ;;
            release-version) INFO_RELEASE_VERSION=$v ;;
            name)            INFO_SERVER_NAME=$v ;;
            server-state)    INFO_SERVER_STATE=$v ;;
        esac
    done < "$INFO_STATE"
}

write_info_csv() {
    local f="$OUT_DIR/$F_INFO" os="" targets="" enabled_list=""
    [ -r /etc/os-release ] && os=$(sed -n 's/^PRETTY_NAME=//p' /etc/os-release | tr -d '"')
    if [ -f "$CONFIG_STATE" ]; then
        targets=$("$AWK_BIN" -F'\t' '$9 == 1 && $3 == "true" { s = s (s == "" ? "" : " ") $2 } END { print s }' "$CONFIG_STATE")
    fi
    if [ -s "$ENABLED_STATE" ]; then
        enabled_list=$("$AWK_BIN" -F'|' '{ s = s (s == "" ? "" : " ") $3 "(" $1 ")" } END { print s }' "$ENABLED_STATE")
    fi
    local mode_label="jboss-cli"
    [ "$MODE" = jmx ] && mode_label="JMX (remote+http)"
    {
        if [ "$WRITE_BOM" = 1 ]; then printf '\357\273\277'; fi
        echo "key,label,value"
        row() { printf '%s,%s,%s\n' "$1" "$(csvq "$2")" "$(csvq "$3")"; }
        row tool_version     "ツールバージョン"               "jboss_pool_collector.sh $TOOL_VERSION"
        row mode             "取得方式"                       "$mode_label"
        row host             "ホスト(CSV の host 列)"         "$HOST"
        row hostname         "ホスト名"                       "${HOSTNAME:-$(uname -n)}"
        row ecs_cluster      "ECS クラスタ"                   "$ECS_CLUSTER"
        row ecs_task_arn     "ECS タスク ARN"                 "$ECS_TASK_ARN"
        row ecs_task_family  "ECS タスク定義"                 "$ECS_FAMILY"
        row ecs_container    "ECS コンテナ名"                 "$ECS_CONTAINER"
        row ecs_az           "アベイラビリティゾーン"         "$ECS_AZ"
        row os               "OS"                             "$os"
        row controller       "管理インタフェース"             "$CONTROLLER"
        row jboss_home       "JBOSS_HOME"                     "$JBOSS_HOME"
        row product_name     "製品名"                         "$INFO_PRODUCT_NAME"
        row product_version  "製品バージョン"                 "$INFO_PRODUCT_VERSION"
        row release_version  "コアバージョン"                 "$INFO_RELEASE_VERSION"
        row server_name      "サーバ名"                       "$INFO_SERVER_NAME"
        row server_state     "サーバ状態(開始時)"             "$INFO_SERVER_STATE"
        row start_time       "収集開始"                       "$START_STR"
        row end_time         "収集終了"                       "$END_STR"
        row interval_sec     "取得間隔(秒)"                   "$INTERVAL"
        row planned_samples  "予定取得回数"                   "${PLANNED:-停止まで}"
        row samples          "取得回数"                       "$SAMPLE_NO"
        row datasources      "対象データソース"               "$targets"
        row warn_pct         "高負荷判定の使用率(%)"          "$WARN_PCT"
        row bucket_min       "時間別集計の単位(分)"           "$BUCKET_MIN"
        row timezone         "タイムゾーン"                   "$REPORT_TZ"
        row stats_enabled_by_tool "ツールが有効化した統計"    "${enabled_list:-なし}"
        row stats_restore    "終了時に統計を元に戻す"         "$([ "$RESTORE_STATS" = 1 ] && echo はい || echo いいえ)"
        row queued_dump      "接続待ちスレッド取得"           "$([ "$QUEUED_DUMP" = 1 ] && echo はい || echo いいえ)"
        row thread_dump      "枯渇時スレッドダンプ"           "$([ "$THREAD_DUMP" = 1 ] && echo "はい(最小間隔 ${THREAD_DUMP_GAP}秒)" || echo いいえ)"
    } > "$f"
}

# ------------------------------------------------------------------------------
# jboss-cli 実行
# ------------------------------------------------------------------------------
run_cli() {   # $1=CLI スクリプト $2=出力ファイル
    local -a args pre=()
    args=("--controller=$CONTROLLER" --connect --output-json "--file=$1" "--command-timeout=$CLI_TIMEOUT")
    if [ -n "$CLI_USER" ]; then args+=("--user=$CLI_USER" "--password=$CLI_PASSWORD"); fi
    if command -v timeout >/dev/null 2>&1; then pre=(timeout -k 10 "$((CLI_TIMEOUT + 30))"); fi
    "${pre[@]}" env -u JAVA_TOOL_OPTIONS -u _JAVA_OPTIONS -u JDK_JAVA_OPTIONS -u JAVA_OPTS_APPEND \
        JAVA_OPTS="$CLI_JAVA_OPTS" "$JBOSS_HOME/bin/jboss-cli.sh" "${args[@]}" > "$2" 2>> "$CLI_ERR_LOG"
}

flatten() { "$AWK_BIN" -f "$WORK_DIR/flatten.awk" "$1"; }

cli_addr() {   # $1=type $2=name
    local n=$2
    case "$n" in
        *[!A-Za-z0-9_.-]*) n="\"${n//\"/\\\"}\"" ;;
    esac
    printf '/subsystem=datasources/%s=%s' "$1" "$n"
}

discover_cli() {
    local cli="$WORK_DIR/discover.cli" raw="$WORK_DIR/discover.json" flat="$WORK_DIR/discover.flat"
    cat > "$cli" <<'CLI'
echo @@INFO
:read-resource(include-runtime=true)
echo @@DS
try
/subsystem=datasources:read-resource(recursive=true,resolve-expressions=true)
catch
echo @@DS_UNRESOLVED
/subsystem=datasources:read-resource(recursive=true)
end-try
CLI
    if ! run_cli "$cli" "$raw"; then
        log ERROR "jboss-cli で管理インタフェースに接続できません ($CONTROLLER)。詳細: $CLI_ERR_LOG"
        tail -n 5 "$CLI_ERR_LOG" >&2 2>/dev/null
        return 1
    fi
    flatten "$raw" > "$flat"
    apply_discovery "$flat"
}

apply_discovery() {   # $1=平坦化した検出結果
    write_header "$OUT_DIR/$F_CONFIG.tmp" "$CFG_COLS"
    "$AWK_BIN" -v filter="$DS_FILTER" -v cfgcsv="$OUT_DIR/$F_CONFIG.tmp" -v statef="$CONFIG_STATE" -v infof="$INFO_STATE" \
        -f "$WORK_DIR/discover.awk" "$1" || return 1
    mv -f "$OUT_DIR/$F_CONFIG.tmp" "$OUT_DIR/$F_CONFIG"
    load_info_state
    local n_all n_tgt
    n_all=$(wc -l < "$CONFIG_STATE" | tr -d ' ')
    n_tgt=$("$AWK_BIN" -F'\t' '$9 == 1 && $3 == "true"' "$CONFIG_STATE" | wc -l | tr -d ' ')
    if [ "$n_all" -eq 0 ]; then log WARN "データソースが見つかりません"; fi
    if [ -n "$DS_FILTER" ] && [ "$n_tgt" -eq 0 ]; then log WARN "指定されたデータソースが見つからないか無効です: $DS_FILTER"; fi
    "$AWK_BIN" -F'\t' '$9 != 1 || $3 != "true" { print $2 (($3 != "true") ? " (無効)" : " (対象外)") }' "$CONFIG_STATE" |
        while IFS= read -r l; do log INFO "収集対象外: $l"; done
    return 0
}

build_sample_cli() {
    local type name enabled target addr
    : > "$SAMPLE_CLI"
    while IFS=$'\t' read -r type name enabled _ _ _ _ _ target; do
        [ "$target" = 1 ] && [ "$enabled" = "true" ] || continue
        addr=$(cli_addr "$type" "$name")
        {
            printf 'echo @@DS|%s|%s\n' "$type" "$name"
            printf 'echo @@P\ntry\n%s/statistics=pool:read-resource(include-runtime=true)\ncatch\necho @@ERR\nend-try\n' "$addr"
            printf 'echo @@J\ntry\n%s/statistics=jdbc:read-resource(include-runtime=true)\ncatch\necho @@ERR\nend-try\n' "$addr"
            if [ "$QUEUED_DUMP" = 1 ]; then
                printf 'echo @@Q\ntry\n%s:dump-queued-threads-in-pool\ncatch\necho @@ERR\nend-try\n' "$addr"
            fi
            printf 'echo @@ENDDS\n'
        } >> "$SAMPLE_CLI"
    done < "$CONFIG_STATE"
}

enable_stats_cli() {
    [ "$ENABLE_STATS" = 1 ] || return 0
    local cli="$WORK_DIR/enable.cli" raw="$WORK_DIR/enable.json" type name enabled target addr kind
    : > "$cli"
    while IFS=$'\t' read -r type name enabled _ _ _ _ _ target; do
        [ "$target" = 1 ] && [ "$enabled" = "true" ] || continue
        addr=$(cli_addr "$type" "$name")
        for kind in pool jdbc; do
            printf 'if (result == false) of %s/statistics=%s:read-attribute(name=statistics-enabled)\necho @@ENABLED|%s|%s|%s\n%s/statistics=%s:write-attribute(name=statistics-enabled,value=true)\nend-if\n' \
                "$addr" "$kind" "$kind" "$type" "$name" "$addr" "$kind" >> "$cli"
        done
    done < "$CONFIG_STATE"
    [ -s "$cli" ] || return 0
    if ! run_cli "$cli" "$raw"; then
        log WARN "統計の有効化に失敗しました (権限不足の可能性)。詳細: $CLI_ERR_LOG"
    fi
    grep '^@@ENABLED|' "$raw" 2>/dev/null | sed 's/^@@ENABLED|//' | while IFS= read -r l; do
        printf '%s\n' "$l" >> "$ENABLED_STATE"
        log INFO "統計を実行時に有効化しました: $l (設定ファイルは変更しません)"
    done
}

restore_stats_cli() {
    [ "$RESTORE_STATS" = 1 ] && [ -s "$ENABLED_STATE" ] || return 0
    local cli="$WORK_DIR/restore.cli" raw="$WORK_DIR/restore.json" kind type name
    : > "$cli"
    while IFS='|' read -r kind type name; do
        printf 'try\n%s/statistics=%s:write-attribute(name=statistics-enabled,value=false)\ncatch\necho @@ERR\nend-try\n' \
            "$(cli_addr "$type" "$name")" "$kind" >> "$cli"
    done < "$ENABLED_STATE"
    if run_cli "$cli" "$raw"; then
        log INFO "ツールが有効化した統計を元に戻しました"
    else
        log WARN "統計を元に戻せませんでした。必要に応じて手動で statistics-enabled=false にしてください"
    fi
}

# ------------------------------------------------------------------------------
# サンプル処理 (cli / jmx 共通)
# ------------------------------------------------------------------------------
LAST_STATUS=""; LAST_TRIGGER=""
process_sample() {   # $1=平坦化サンプル $2=epoch
    local flat=$1 ep=$2 ts elapsed
    SAMPLE_NO=$((SAMPLE_NO + 1))
    ts=$(date -d "@$ep" '+%Y-%m-%d %H:%M:%S')
    elapsed=$((ep - START_EPOCH))
    "$AWK_BIN" -v cols="$TS_COLS" -v ts="$ts" -v epoch="$ep" -v elapsed="$elapsed" -v sno="$SAMPLE_NO" \
        -v host="$HOST" -v warn="$WARN_PCT" -v cfgf="$CONFIG_STATE" -v prevf="$PREV_STATE" \
        -v newprevf="$PREV_STATE.new" -v statusf="$STATUS_FILE" -v qsigf="$QSIG_STATE" \
        -v qcsvf="$OUT_DIR/$F_QUEUED" -v qstackf="$OUT_DIR/$F_QSTACK" \
        -f "$WORK_DIR/rows.awk" "$flat" >> "$OUT_DIR/$F_TS"
    mv -f "$PREV_STATE.new" "$PREV_STATE"
    # 最も深刻な状態を判定 (ログ / スレッドダンプ契機)
    LAST_STATUS=$("$AWK_BIN" -F'\t' 'BEGIN { r["OK"] = 1; r["HIGH"] = 2; r["WAIT"] = 3; r["EXHAUSTED"] = 4; w = ""; best = 0 }
        { s = ($1 in r) ? r[$1] : 0; if (s > best) { best = s; w = $1; d = $3 } else if (w == "" && $1 != "") { w = $1; d = $3 } }
        END { print w "\t" d }' "$STATUS_FILE")
    LAST_TRIGGER="${LAST_STATUS%%$'\t'*}:${LAST_STATUS#*$'\t'}"
    LAST_STATUS="${LAST_STATUS%%$'\t'*}"
    case "$LAST_STATUS" in
        EXHAUSTED) log WARN "#$SAMPLE_NO コネクションプール枯渇を検出: ${LAST_TRIGGER#*:}" ;;
        WAIT)      log WARN "#$SAMPLE_NO 接続待機を検出: ${LAST_TRIGGER#*:}" ;;
        ERROR)     log WARN "#$SAMPLE_NO 統計の取得に失敗しました: ${LAST_TRIGGER#*:}" ;;
    esac
    if [ $((SAMPLE_NO % 30)) -eq 0 ] || [ "$SAMPLE_NO" -eq 1 ]; then
        log INFO "#$SAMPLE_NO 取得完了 (状態: ${LAST_STATUS:-なし})"
    fi
    if [ $((SAMPLE_NO % SUMMARY_EVERY)) -eq 0 ]; then summarize "$OUT_DIR" quiet; fi
}

error_sample_flat() {   # 取得自体に失敗した場合に ERROR 行を作るための疑似サンプル
    local msg=$1 type name enabled target
    while IFS=$'\t' read -r type name enabled _ _ _ _ _ target; do
        [ "$target" = 1 ] && [ "$enabled" = "true" ] || continue
        printf '@@DS|%s|%s\n@@P\n@@ERR\nF\t%s\n@@ENDDS\n' "$type" "$name" "$msg"
    done < "$CONFIG_STATE"
}

LAST_TD=0
should_thread_dump() {
    [ "$THREAD_DUMP" = 1 ] || return 1
    case "$LAST_STATUS" in EXHAUSTED|WAIT) ;; *) return 1 ;; esac
    [ $(( $(date +%s) - LAST_TD )) -ge "$THREAD_DUMP_GAP" ] || return 1
    return 0
}

process_threaddump() {   # $1=平坦化ダンプ $2=epoch $3=契機
    local flat=$1 ep=$2 trig=$3 name txt res
    mkdir -p "$OUT_DIR/threaddumps"
    name="threaddump_$(date -d "@$ep" '+%Y%m%d_%H%M%S').txt"
    txt="$OUT_DIR/threaddumps/$name"
    {
        printf 'Thread dump: %s  trigger: %s  host: %s\n\n' "$(date -d "@$ep" '+%Y-%m-%d %H:%M:%S')" "$trig" "$HOST"
    } > "$txt"
    [ -f "$OUT_DIR/$F_TDTHREADS" ] || write_header "$OUT_DIR/$F_TDTHREADS" "$TD_COLS"
    res=$("$AWK_BIN" -v txtf="$txt" -v csvf="$OUT_DIR/$F_TDTHREADS" -v dtime="$(date -d "@$ep" '+%Y-%m-%d %H:%M:%S')" \
        -v dfile="$name" -v trigger="$trig" -f "$WORK_DIR/threaddump.awk" "$flat")
    LAST_TD=$ep
    IFS=$'\t' read -r t_all t_jdbc t_wait t_create <<< "$res"
    log INFO "スレッドダンプ取得: threaddumps/$name (全${t_all:-0} / JDBC実行中 ${t_jdbc:-0} / 接続待ち ${t_wait:-0} / 接続作成中 ${t_create:-0})"
}

thread_dump_cli() {
    local cli="$WORK_DIR/td.cli" raw="$WORK_DIR/td.json" flat="$WORK_DIR/td.flat" ep
    ep=$(date +%s)
    echo '/core-service=platform-mbean/type=threading:dump-all-threads(locked-monitors=false,locked-synchronizers=false)' > "$cli"
    if run_cli "$cli" "$raw"; then
        flatten "$raw" > "$flat"
        process_threaddump "$flat" "$ep" "$LAST_TRIGGER"
    else
        log WARN "スレッドダンプの取得に失敗しました"
        LAST_TD=$ep
    fi
}

sleep_ms() {   # 中断可能な sleep
    local ms=$1
    [ "$ms" -gt 0 ] || return 0
    sleep "$((ms / 1000)).$(printf '%03d' $((ms % 1000)))" &
    SLEEP_PID=$!
    wait "$SLEEP_PID" 2>/dev/null
    SLEEP_PID=""
}

# ------------------------------------------------------------------------------
# 収集ループ
# ------------------------------------------------------------------------------
collect_cli() {
    local due now ep raw="$WORK_DIR/sample.json" flat="$WORK_DIR/sample.flat" last_disc rc need_disc=0
    discover_cli || die "初回のデータソース検出に失敗しました"
    enable_stats_cli
    build_sample_cli
    [ -s "$SAMPLE_CLI" ] || die "収集対象のデータソースがありません"
    write_info_csv
    last_disc=$(date +%s)
    due=$(now_ms)
    while [ "$STOP" = 0 ]; do
        if [ -n "$PLANNED" ] && [ "$SAMPLE_NO" -ge "$PLANNED" ]; then break; fi
        now=$(now_ms)
        if [ "$due" -gt "$now" ]; then sleep_ms $((due - now)); fi
        [ "$STOP" = 0 ] || break
        if [ "$need_disc" = 1 ] || [ $(( $(date +%s) - last_disc )) -ge "$REDISCOVER_SEC" ]; then
            if discover_cli; then enable_stats_cli; build_sample_cli; need_disc=0; fi
            last_disc=$(date +%s)
        fi
        ep=$(date +%s)
        run_cli "$SAMPLE_CLI" "$raw"; rc=$?
        if [ "$rc" -eq 0 ] && grep -q '^@@ENDDS' "$raw"; then
            flatten "$raw" > "$flat"
        else
            log WARN "jboss-cli の実行に失敗しました (rc=$rc)。詳細: $CLI_ERR_LOG"
            error_sample_flat "jboss-cli rc=$rc" > "$flat"
            need_disc=1
        fi
        process_sample "$flat" "$ep"
        if should_thread_dump; then thread_dump_cli; fi
        due=$((due + INTERVAL * 1000))
        now=$(now_ms)
        if [ "$due" -lt "$now" ]; then due=$now; fi
    done
}

jmx_url() {
    case "$CONTROLLER" in
        service:jmx:*) printf '%s' "$CONTROLLER" ;;
        *://*)         printf 'service:jmx:%s' "$CONTROLLER" ;;
        *)             printf 'service:jmx:remote+http://%s' "$CONTROLLER" ;;
    esac
}

collect_jmx() {
    local java="${JAVA_HOME:+$JAVA_HOME/bin/}java" client="$JBOSS_HOME/bin/client/jboss-cli-client.jar"
    local line mode="" disc="$WORK_DIR/discover.flat" flat="$WORK_DIR/sample.flat" td="$WORK_DIR/td.flat"
    local td_req="$WORK_DIR/jmx_td.request" ep td_ep=0 ready=0 count=0 waited=0
    command -v "$java" >/dev/null 2>&1 || die "java が見つかりません (JAVA_HOME を設定してください)"
    [ -f "$client" ] || die "JMX クライアントライブラリがありません: $client"
    if ! run_java_env "$java" --list-modules 2>/dev/null | grep -q '^jdk\.compiler'; then
        die "この Java には jdk.compiler モジュールが無いため JMX 方式を使用できません。-m cli を使用してください"
    fi
    write_java
    [ -n "$PLANNED" ] && count=$PLANNED
    JMX_STOP_FILE="$WORK_DIR/jmx_stop.request"
    rm -f "$td_req" "$JMX_STOP_FILE"
    exec 3< <(exec env -u JAVA_TOOL_OPTIONS -u _JAVA_OPTIONS -u JDK_JAVA_OPTIONS -u JAVA_OPTS_APPEND \
        JBOSS_CLI_USER="$CLI_USER" JBOSS_CLI_PASSWORD="$CLI_PASSWORD" \
        "$java" $CLI_JAVA_OPTS -Djava.util.logging.config.file=/dev/null -cp "$client" "$WORK_DIR/PoolJmx.java" \
        "$(jmx_url)" "$INTERVAL" "$count" "$ENABLE_STATS" "$RESTORE_STATS" "$QUEUED_DUMP" "$REDISCOVER_SEC" "$DS_FILTER" "$td_req" "$JMX_STOP_FILE" \
        2>> "$CLI_ERR_LOG")
    JAVA_PID=$!
    log INFO "JMX クライアントを起動しました (pid=$JAVA_PID, $(jmx_url))"
    while [ "$STOP" = 0 ] && IFS= read -r -u 3 line; do
        case "$line" in
            @@READY)          ready=1; log INFO "JMX 接続完了" ;;
            @@FATAL\|*)       log ERROR "JMX 接続に失敗しました: ${line#@@FATAL|}"; break ;;
            @@DISCOVERY)      mode=disc; : > "$disc" ;;
            @@ENDDISCOVERY)   mode=""; apply_discovery "$disc"; write_info_csv ;;
            @@ENABLED\|*)     printf '%s\n' "${line#@@ENABLED|}" >> "$ENABLED_STATE"
                              log INFO "統計を実行時に有効化しました: ${line#@@ENABLED|} (設定ファイルは変更しません)" ;;
            @@RESTORED\|*)    log INFO "統計を元に戻しました: ${line#@@RESTORED|}" ;;
            @@WARN\|*)        log WARN "${line#@@WARN|}" ;;
            @@SAMPLE\|*)      mode=sample; ep=${line#@@SAMPLE|}; ep=$((ep / 1000)); exec 4> "$flat" ;;
            @@ENDSAMPLE)      mode=""; exec 4>&-
                              process_sample "$flat" "$ep"
                              if should_thread_dump; then : > "$td_req"; LAST_TD=$(date +%s); TD_TRIGGER=$LAST_TRIGGER; fi ;;
            @@TD\|*)          mode=td; td_ep=${line#@@TD|}; td_ep=$((td_ep / 1000)); exec 5> "$td" ;;
            @@ENDTD)          mode=""; exec 5>&-; process_threaddump "$td" "$td_ep" "$TD_TRIGGER" ;;
            @@EOF)            break ;;
            *)
                case "$mode" in
                    disc)   printf '%s\n' "$line" >> "$disc" ;;
                    sample) printf '%s\n' "$line" >&4 ;;
                    td)     printf '%s\n' "$line" >&5 ;;
                esac ;;
        esac
    done
    if [ "$mode" = sample ]; then exec 4>&-; fi
    if [ "$mode" = td ]; then exec 5>&-; fi
    # 停止要求 → Java 側で統計を元に戻して終了するまで待つ (最大 60 秒)
    : > "$JMX_STOP_FILE"
    while [ "$waited" -lt 60 ]; do
        if ! IFS= read -r -t 5 -u 3 line; then
            [ $? -gt 128 ] || break
            waited=$((waited + 5)); continue
        fi
        case "$line" in
            @@RESTORED\|*) log INFO "統計を元に戻しました: ${line#@@RESTORED|}" ;;
            @@WARN\|*)     log WARN "${line#@@WARN|}" ;;
            @@EOF)         break ;;
        esac
    done
    if [ -n "$JAVA_PID" ] && kill -0 "$JAVA_PID" 2>/dev/null; then kill -TERM "$JAVA_PID" 2>/dev/null; fi
    wait "$JAVA_PID" 2>/dev/null
    JAVA_PID=""
    exec 3<&-
    [ "$ready" = 1 ] || die "JMX で接続できませんでした。詳細: $CLI_ERR_LOG"
}

# ------------------------------------------------------------------------------
# 集計
# ------------------------------------------------------------------------------
summarize() {   # $1=出力ディレクトリ [$2=quiet]
    local dir=$1 interval=$INTERVAL warn=$WARN_PCT bucket=$BUCKET_MIN v
    [ -s "$dir/$F_TS" ] || { log WARN "時系列 CSV がありません: $dir/$F_TS"; return 1; }
    if [ -n "$SUMMARIZE_DIR" ] && [ -f "$dir/$F_INFO" ]; then
        v=$(grep '^interval_sec,' "$dir/$F_INFO" | cut -d, -f3 | tr -d '"'); is_uint "$v" && interval=$v
        v=$(grep '^warn_pct,' "$dir/$F_INFO" | cut -d, -f3 | tr -d '"'); is_uint "$v" && warn=$v
        v=$(grep '^bucket_min,' "$dir/$F_INFO" | cut -d, -f3 | tr -d '"'); is_uint "$v" && bucket=$v
    fi
    "$AWK_BIN" -v interval="$interval" -v warn="$warn" -v bucket="$bucket" -v bom="$WRITE_BOM" \
        -v sumf="$dir/$F_SUMMARY.tmp" -v hourf="$dir/$F_HOURLY.tmp" -v evf="$dir/$F_EVENTS.tmp" \
        -f "$WORK_DIR/summary.awk" "$dir/$F_TS" || { log WARN "集計に失敗しました"; return 1; }
    mv -f "$dir/$F_SUMMARY.tmp" "$dir/$F_SUMMARY"
    mv -f "$dir/$F_HOURLY.tmp" "$dir/$F_HOURLY"
    mv -f "$dir/$F_EVENTS.tmp" "$dir/$F_EVENTS"
    [ "${2:-}" = quiet ] || log INFO "集計 CSV を出力しました: $F_SUMMARY / $F_HOURLY / $F_EVENTS"
}

print_summary_console() {
    [ -s "$OUT_DIR/$F_SUMMARY" ] || return 0
    echo "------------------------------------------------------------------"
    echo " データソース別 判定"
    echo "------------------------------------------------------------------"
    # 先頭列(データソース) と 判定(5列目) / 使用中 最大 / 最大プールサイズ を表示
    "$AWK_BIN" 'NR == 1 { sub(/^\357\273\277/, ""); next }
        {
            line = $0; n = 0; f = ""; inq = 0
            for (i = 1; i <= length(line); i++) {
                c = substr(line, i, 1)
                if (c == "\"") { inq = !inq; continue }
                if (c == "," && !inq) { F[++n] = f; f = ""; continue }
                f = f c
            }
            F[++n] = f
            printf "  %-24s %-14s 使用中 最大 %s / 最大プールサイズ %s\n", F[1], F[5], F[16], F[7]
        }' "$OUT_DIR/$F_SUMMARY"
    echo "------------------------------------------------------------------"
}

# ------------------------------------------------------------------------------
# 終了処理
# ------------------------------------------------------------------------------
FINALIZED=0
finalize() {
    [ "$FINALIZED" = 0 ] || return 0
    FINALIZED=1
    [ -n "${WORK_DIR:-}" ] || return 0
    if [ "$MODE" = cli ]; then restore_stats_cli; fi
    END_STR=$(now_str)
    write_info_csv
    if [ "$SAMPLE_NO" -gt 0 ]; then
        summarize "$OUT_DIR"
        print_summary_console
    fi
    log INFO "収集を終了しました (取得 $SAMPLE_NO 回)。出力先: $OUT_DIR"
    log INFO "Excel 変換: pool_csv_to_xlsx.sh (RHEL) / Convert-PoolCsvToXlsx.ps1 (Windows) に上記フォルダを指定してください"
}

on_signal() {
    STOP=1
    log INFO "停止要求を受け付けました。終了処理を行います..."
    if [ -n "$SLEEP_PID" ]; then kill "$SLEEP_PID" 2>/dev/null; fi
    if [ -n "$JMX_STOP_FILE" ]; then : > "$JMX_STOP_FILE"; fi
}

# ------------------------------------------------------------------------------
# メイン
# ------------------------------------------------------------------------------
main() {
    parse_args "$@"
    setup_tz
    setup_awk

    if [ -n "$SUMMARIZE_DIR" ]; then
        [ -d "$SUMMARIZE_DIR" ] || die "ディレクトリがありません: $SUMMARIZE_DIR"
        OUT_DIR=$SUMMARIZE_DIR
        WORK_DIR=$(mktemp -d "${TMPDIR:-/tmp}/pool_summarize.XXXXXX") || die "作業ディレクトリを作成できません"
        trap 'rm -rf "$WORK_DIR"' EXIT
        write_programs
        summarize "$OUT_DIR" && print_summary_console
        [ -f "$OUT_DIR/$F_COLUMNS" ] || write_columns_csv
        exit 0
    fi

    detect_jboss_home
    detect_controller
    detect_ecs
    HOST=$(host_label)

    if [ -z "$OUT_DIR" ]; then
        OUT_DIR="${TMPDIR:-/tmp}/pool_report_$(printf '%s' "$HOST" | tr -c 'A-Za-z0-9._-' '_')_$(date '+%Y%m%d_%H%M%S')"
    fi
    mkdir -p "$OUT_DIR" || die "出力ディレクトリを作成できません: $OUT_DIR"
    OUT_DIR=$(cd "$OUT_DIR" && pwd)

    # バックグラウンド実行: 自分自身を nohup で起動し直す
    if [ "$BACKGROUND" = 1 ] && [ -z "${_POOL_COLLECTOR_CHILD:-}" ]; then
        local -a args=()
        local a
        for a in "${ORIG_ARGS[@]}"; do
            case "$a" in -D|--background) ;; *) args+=("$a") ;; esac
        done
        _POOL_COLLECTOR_CHILD=1 nohup bash "$SCRIPT_PATH" "${args[@]}" -o "$OUT_DIR" \
            > "$OUT_DIR/collector.log" 2>&1 < /dev/null &
        local pid=$!
        echo "$pid" > "$OUT_DIR/collector.pid"
        echo "バックグラウンドで収集を開始しました (PID: $pid)"
        echo "  出力先 : $OUT_DIR"
        echo "  ログ   : tail -f $OUT_DIR/collector.log"
        echo "  停止   : kill $pid   (集計 CSV を出力して終了します)"
        exit 0
    fi
    if [ -z "${_POOL_COLLECTOR_CHILD:-}" ]; then LOG_FILE="$OUT_DIR/collector.log"; fi

    WORK_DIR="$OUT_DIR/.work"
    mkdir -p "$WORK_DIR" || die "作業ディレクトリを作成できません"
    CONFIG_STATE="$WORK_DIR/config.state"
    INFO_STATE="$WORK_DIR/info.state"
    PREV_STATE="$WORK_DIR/prev.state"
    QSIG_STATE="$WORK_DIR/qsig.state"
    STATUS_FILE="$WORK_DIR/status.state"
    ENABLED_STATE="$WORK_DIR/enabled.state"
    SAMPLE_CLI="$WORK_DIR/sample.cli"
    CLI_ERR_LOG="$OUT_DIR/cli_error.log"
    : > "$PREV_STATE"; : > "$QSIG_STATE"; : > "$ENABLED_STATE"; : > "$CONFIG_STATE"
    write_programs

    if [ -n "$COUNT" ]; then PLANNED=$COUNT
    elif [ "$DURATION" -gt 0 ]; then PLANNED=$(( (DURATION + INTERVAL - 1) / INTERVAL ))
    else PLANNED=""; fi

    write_header "$OUT_DIR/$F_TS" "$TS_COLS"
    write_header "$OUT_DIR/$F_QUEUED" "$Q_COLS"
    write_header "$OUT_DIR/$F_QSTACK" "$QS_COLS"
    write_columns_csv

    START_STR=$(now_str)
    START_EPOCH=$(date +%s)
    trap on_signal INT TERM HUP
    trap finalize EXIT

    if [ "$MODE" = cli ] && [ "$INTERVAL" -lt 5 ]; then
        log WARN "jboss-cli 方式で ${INTERVAL} 秒間隔は負荷が高くなる可能性があります (5 秒以上 または -m jmx を推奨)"
    fi
    log INFO "収集開始: mode=$MODE interval=${INTERVAL}s samples=${PLANNED:-停止まで} JBOSS_HOME=$JBOSS_HOME controller=$CONTROLLER"
    log INFO "出力先: $OUT_DIR"

    if [ "$MODE" = jmx ]; then collect_jmx; else collect_cli; fi
}

main "$@"
