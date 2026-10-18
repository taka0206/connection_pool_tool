# JBoss EAP 8.1 コネクションプール調査ツール

ECS タスク上の UBI 9.x (9.8) コンテナで稼働する **JBoss EAP 8.1** のデータソース
(IronJacamar) コネクションプール統計を **jboss-cli または JMX** で定期取得して CSV に出力し、
その CSV から **Meiryo UI・モノトーン配色の Excel レポート** を作成するツール一式です。

| ファイル | 動作環境 | 役割 |
|---|---|---|
| `collector/jboss_pool_collector.sh` | ECS 上の UBI 9.x コンテナ (EAP 8.1) | プール統計の収集 → CSV 出力 |
| `excel/pool_csv_to_xlsx.sh` | RHEL 9.x (9.8) | CSV → Excel (.xlsx) |
| `excel/Convert-PoolCsvToXlsx.ps1` | Windows 11 (Windows PowerShell 5.1 / PowerShell 7) | CSV → Excel (.xlsx) |
| `samples/` | ― | 合成データによる出力サンプル (CSV と .xlsx) |

いずれも **1 ファイルで完結** しており、内容をすべてコピー＆貼り付けしてそのまま実行できます。
追加パッケージ (jq / python ライブラリ / Excel 本体 / PowerShell モジュール) は不要です。

---

## 1. 収集スクリプト (`jboss_pool_collector.sh`)

### 1.1 取得する情報

| 区分 | 内容 |
|---|---|
| プール統計 (`statistics=pool`) | InUseCount / ActiveCount / IdleCount / AvailableCount / MaxUsedCount / WaitCount / MaxWaitCount / BlockingFailureCount / TimedOut / Created/DestroyedCount / 待機・取得・使用・作成・滞留時間 (平均/最大/合計) / XA 統計 |
| JDBC 統計 (`statistics=jdbc`) | PreparedStatement キャッシュ (参照・ヒット・ミス・追加・削除・サイズ) |
| 接続待ちスレッド | `dump-queued-threads-in-pool` (接続の空きを待っているスレッドとそのスタック) |
| プール設定 | min/max/initial-pool-size, blocking-timeout, idle-timeout, 検証設定, flush-strategy ほか全属性 (パスワードはマスク) |
| スレッドダンプ (任意) | 枯渇・待機を検出した時点のスレッドダンプ (`-t`) |
| 環境 | 製品バージョン, サーバ状態, ECS タスク ARN / クラスタ / タスク定義 / AZ (ECS メタデータから取得) |

### 1.2 コンテナへの配置

ECS Exec でコンテナに入ります (タスクで ECS Exec が有効になっている必要があります)。

```bash
aws ecs execute-command --cluster <クラスタ> --task <タスクID> --container <コンテナ名> \
    --interactive --command "/bin/bash"
```

コンテナ内で次のように入力し、`jboss_pool_collector.sh` の内容を **すべて** 貼り付けたあと、
最終行に `EOF_POOL` と入力して Enter を押します。

```bash
cat > /tmp/jboss_pool_collector.sh <<'EOF_POOL'
（ここに jboss_pool_collector.sh の内容を貼り付け）
EOF_POOL
```

> コンテナイメージに同梱する、EFS 上に置く、などの方法でも構いません。

### 1.3 実行

```bash
# 10 秒間隔で 1 時間収集 (既定)。ローカル認証のためユーザ/パスワード不要
bash /tmp/jboss_pool_collector.sh

# 5 秒間隔で 24 時間、バックグラウンド実行 (ECS Exec セッションが切れても継続)
bash /tmp/jboss_pool_collector.sh -i 5 -d 86400 -D

# JMX で取得 / 特定データソースのみ / 枯渇時にスレッドダンプ
bash /tmp/jboss_pool_collector.sh -m jmx -s MyDS,MyXADS -t

# 停止 (Ctrl+C または kill)。停止時に集計 CSV を出力し、統計設定を元に戻します
kill $(cat /tmp/pool_report_*/collector.pid)
```

出力先の既定は `/tmp/pool_report_<タスクID またはホスト名>_<日時>/` です (`-o` で変更可)。
実行中も 30 サンプルごとに集計 CSV を更新するため、途中でも結果を確認できます。
強制終了などで集計が出力されなかった場合は `--summarize <出力先>` で再集計できます。

### 1.4 主なオプション

| オプション | 既定値 | 説明 |
|---|---|---|
| `-m, --mode` | `cli` | `cli` = jboss-cli / `jmx` = JMX (remote+http) |
| `-i, --interval` | `10` | 取得間隔 (秒) |
| `-d, --duration` | `3600` | 収集時間 (秒)。`0` は停止するまで |
| `-n, --count` | ― | 取得回数 (指定時は `-d` より優先) |
| `-s, --datasource` | 全て | 対象データソース (カンマ区切り) |
| `-o, --output-dir` | `/tmp/pool_report_...` | 出力先 |
| `-w, --warn-pct` | `80` | 高負荷と判定する使用率 (%) |
| `-b, --bucket-min` | `60` | 時間別集計の単位 (分) |
| `-z, --tz` | `Asia/Tokyo` | タイムスタンプのタイムゾーン (tzdata が無い場合も JST で出力) |
| `-H, --jboss-home` | 自動検出 | `/opt/server`, `/opt/eap` 等、稼働プロセスの `-Djboss.home.dir` から検出 |
| `-c, --controller` | 自動検出 | 稼働中サーバの `-bmanagement` / ポートオフセットから決定 (既定 `remote+http://127.0.0.1:9990`) |
| `-u` / `-p` | ― | 管理ユーザ/パスワード (環境変数 `JBOSS_CLI_USER` / `JBOSS_CLI_PASSWORD` 推奨) |
| `-t, --thread-dump` | 無効 | 枯渇・待機の検出時にスレッドダンプ取得 (`--thread-dump-gap` 秒に 1 回まで、既定 300) |
| `--no-enable-stats` | ― | 統計が無効でも有効化しない |
| `--keep-stats` | ― | 有効化した統計を終了時に元に戻さない |
| `--no-queued-dump` | ― | 接続待ちスレッドを取得しない |
| `-D, --background` | ― | nohup でバックグラウンド実行 |
| `--summarize DIR` | ― | 既存の出力から集計 CSV のみ再生成 |

`bash jboss_pool_collector.sh -h` で全オプションを表示します。

### 1.5 動作上の注意

* **統計の有効化**: EAP の既定では `statistics-enabled=false` のため値がすべて 0 になります。
  本ツールは統計が無効なデータソースについて **実行時のみ** 有効化し (設定ファイルは変更しない・
  再起動不要)、終了時に元へ戻します (`--keep-stats` で維持)。統計収集のオーバーヘッドは一般に小さいものですが、
  高負荷の本番環境では事前に検証環境で影響を確認してください。
* **認証**: コンテナ内から実行すると jboss-cli / JMX ともに **ローカル認証** で接続できるため、
  通常ユーザ/パスワードは不要です。
* **負荷**: `cli` 方式は 1 回の取得ごとに jboss-cli の JVM (ヒープ最大 128MB, SerialGC) を起動し、
  約 1～2 秒 CPU を使います。間隔 5 秒未満で取得したい場合は、JVM を常駐させる `-m jmx` を推奨します。
  コンテナに設定された `JAVA_OPTS` / `JAVA_TOOL_OPTIONS` (javaagent 等) はツールの JVM には引き継ぎません。
* **JMX 方式** は EAP 同梱の `bin/client/jboss-cli-client.jar` と Java のソースファイル実行機能を使います。
  Java に `jdk.compiler` モジュールが無い場合は `cli` 方式を使用してください (エラーで案内します)。
* 必要なコマンドは bash / coreutils / awk / sed / grep / java (EAP 同梱) のみで、`ubi9-minimal` ベースでも動作します。

### 1.6 出力ファイル

| ファイル | 内容 |
|---|---|
| `pool_timeseries.csv` | **時系列** (サンプルごと・データソースごとの全統計 + 使用率・区間増分・状態) |
| `pool_hourly.csv` | **時間別推移** (時間帯ごとの使用中 最小/最大/平均、使用率、待機・タイムアウト件数、最悪状態) |
| `pool_summary.csv` | **データソース別サマリ** (最大・最小・平均・95%値、枯渇回数/時間、判定、所見コメント) |
| `pool_events.csv` | **枯渇イベント** (高負荷・待機・枯渇が連続した期間の開始/終了/継続時間/ピーク) |
| `pool_queued_threads.csv` | 接続待ちスレッド (スタックごとの待機数・呼出元・アプリケーションフレーム) |
| `pool_queued_stacks.csv` | 接続待ちスレッドのスタックトレース (スタックごとに 1 回) |
| `pool_threaddump_threads.csv` + `threaddumps/*.txt` | スレッドダンプ (`-t` 指定時)。接続待ち / JDBC 実行中 / 接続作成中のスレッド一覧 |
| `pool_config.csv` | プール設定 (主要属性は日本語名・既定値つき、パスワードはマスク) |
| `collection_info.csv` | 収集条件・環境 |
| `columns.csv` | 項目定義 (日本語名・単位・書式・説明)。Excel 変換で見出し・書式に使用 |
| `collector.log` / `cli_error.log` | 実行ログ |

CSV は UTF-8 (BOM 付き) のため、Excel でそのまま開いても文字化けしません (`--no-bom` で BOM なし)。

### 1.7 状態判定

| 状態 | 条件 |
|---|---|
| 枯渇 (EXHAUSTED) | 使用中 = max-pool-size / 取得可能数 0 / 前回取得以降に取得タイムアウト (BlockingFailureCount 増加) |
| 待機発生 (WAIT) | 前回取得以降に WaitCount が増加 / 接続待ちスレッドあり |
| 高負荷 (HIGH) | 使用率 (使用中 ÷ max-pool-size) が `--warn-pct` 以上 |
| 正常 (OK) | 上記以外 |
| 統計無効 (NOSTATS) / 取得失敗 (ERROR) | 統計が無効 / 取得エラー |

サマリでは、さらに **コネクションリークの疑い** (30 分以上の収集で使用中接続数が増加傾向かつ
下限も上昇)、長時間保持 (最大使用時間 60 秒以上)、接続作成の遅延 (3 秒以上)、
アイドルタイムアウト以外の接続破棄、サーバ再起動によるカウンタリセットなどを所見として出力します。

### 1.8 CSV の取り出し

* **推奨**: EFS などのボリュームをマウントし、`-o /mnt/efs/...` に出力する。
* ECS Exec で取り出す場合 (base64 で転送しローカルで復元。ローカル側で実行):

```bash
DIR=/tmp/pool_report_xxxx            # コンテナ内の出力先
for f in pool_timeseries pool_hourly pool_summary pool_events pool_config pool_queued_threads pool_queued_stacks collection_info columns; do
  aws ecs execute-command --cluster <クラスタ> --task <タスクID> --container <コンテナ名> --interactive \
      --command "base64 $DIR/$f.csv" \
    | tr -d '\r' | grep -v -e '^$' -e '^Starting session' -e '^Exiting session' | base64 -d > "$f.csv"
done
```

> Session Manager の開始/終了メッセージを除外してから復元しています。
> ファイルが大きい場合は時間がかかるため、ボリュームへの出力を推奨します。

---

## 2. Excel 変換

### 2.1 RHEL 9.x (`pool_csv_to_xlsx.sh`)

RHEL 標準の `python3` (3.6 以降、標準ライブラリのみ) を使って .xlsx を直接生成します。
収集スクリプトと同様に、`cat > pool_csv_to_xlsx.sh <<'EOF_XLSX'` → 内容を貼り付け → `EOF_XLSX` で作成できます。

```bash
bash pool_csv_to_xlsx.sh ./pool_report_task-xxxx_20261018_100000
bash pool_csv_to_xlsx.sh -o /tmp/report.xlsx ./pool_report_task-xxxx_20261018_100000
```

### 2.2 Windows 11 (`Convert-PoolCsvToXlsx.ps1`)

Excel が無い PC でも動作します (Office Open XML を直接生成)。

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\Convert-PoolCsvToXlsx.ps1 -InputDir C:\work\pool_report_task-xxxx_20261018_100000
# 出力先を指定する場合
powershell -NoProfile -ExecutionPolicy Bypass -File .\Convert-PoolCsvToXlsx.ps1 -InputDir C:\work\pool_report_xxx -OutputPath C:\work\report.xlsx
```

> スクリプト本体は ASCII 文字のみで記述しているため、メモ帳に貼り付けてどの文字コードで保存しても
> Windows PowerShell 5.1 で文字化けしません (日本語の見出しは `columns.csv` から読み込みます)。
> インターネットから取得したファイルで実行がブロックされる場合は `Unblock-File .\Convert-PoolCsvToXlsx.ps1` を実行してください。

両スクリプトは同一のレイアウト・書式の Excel を出力します (テストで出力 XML の一致を確認済み)。
オプション: `--max-rows` / `-MaxRows` (時系列シートの最大行数)、`--no-charts` / `-NoCharts` (グラフなし)。

### 2.3 ワークブックの構成

| シート | 内容 |
|---|---|
| サマリ | レポート表題、データソース別サマリ (判定・最大/最小/平均/95%値・枯渇時間・所見)、収集条件・環境、状態の凡例 |
| グラフ | データソースごとに「接続数の推移 (使用中/作成済/アイドル/最大プールサイズ)」「時間別 使用中接続数 (最大/平均/最小)」「時間別 待機・取得タイムアウト発生数」 |
| 時間別推移 | 時間帯ごとの集計表 |
| 枯渇イベント | 高負荷・待機・枯渇の連続期間 |
| 時系列 | 全サンプル (データソースごとに並べ替え) |
| 待機スレッド / 待機スタック | 接続を待っていたスレッドと呼出元 |
| スレッドダンプ | `-t` 指定時のみ |
| プール設定 | 設定値と既定値 |
| 項目説明 | 全列の日本語名・単位・説明 |

書式:

* フォント: **Meiryo UI** (全セル・グラフ・印刷ヘッダ/フッタ)
* 配色: **モノトーン**。見出し=黒地に白字、縞模様=淡いグレー、
  状態は 枯渇=黒地白字 / 待機発生=濃グレー地白字 / 高負荷=グレー地太字 / 統計無効・取得失敗=グレー斜体
* セル: 数値は桁区切り、使用率は `0.0%`、日時は `yyyy/mm/dd hh:mm:ss` の実日付値、罫線、
  見出し行・キー列の固定、オートフィルタ、内容に合わせた列幅、長文の折り返し、閾値以上の使用率を太字
* 印刷: A4 横、見出し行 (と固定列) を各ページに繰り返し、狭いシートは横 1 ページに収める、ヘッダにシート名・フッタにページ番号

`samples/pool_report_sample.xlsx` は合成データ (24 時間・4 データソース: 昼休みの枯渇、
夜間バッチの高負荷、コネクションリーク、統計無効) から作成した出力例です。

---

## 3. 調査の進め方 (例)

1. **サマリ** の「判定」「所見」で、枯渇・待機・リーク疑いのあるデータソースを確認する。
2. **グラフ** で使用中接続数が最大プールサイズに張り付く時間帯を確認する。
3. **枯渇イベント** で発生時刻・継続時間・取得タイムアウト件数を確認し、アプリのログと突き合わせる。
4. **待機スレッド / 待機スタック** の「アプリケーションフレーム」で、接続を待っていた処理を特定する。
5. 枯渇の原因 (接続を保持し続けている処理) は、`-t` で取得した **スレッドダンプ** の
   「JDBC実行中」のスレッド (長時間 SQL) を確認する。接続を保持したまま別処理で止まっている場合は
   スタックに JDBC が現れないため、ダンプ全体 (`threaddumps/*.txt`) も確認する。
6. **プール設定** で max-pool-size、blocking-timeout、idle-timeout、検証設定を確認し、見直しを検討する。

---

## 4. 動作確認について

* WildFly 35.0.1.Final (JBoss EAP 8.1 の基になったアップストリーム) + Java 21 上で、
  非 XA / XA データソース、実際のプール枯渇 (取得タイムアウト・接続待ち)、コネクションリーク、
  無効データソースを発生させ、`cli` / `jmx` 両方式・フォアグラウンド / バックグラウンド・停止時の統計復元を確認しています。
  awk は gawk (UBI 9 標準) と mawk の両方で確認しています。
* Excel 変換は Python 3.9 (RHEL 9 標準) / PowerShell 7.4 で実行し、生成した .xlsx を
  ECMA-376 (Office Open XML) のスキーマで検証、LibreOffice で表示を確認しています。
* 実機の EAP 8.1 コンテナイメージ、Microsoft Excel、Windows PowerShell 5.1 での確認は行っていません。
  本番環境で使用する前に、検証環境で一度実行して結果を確認してください。
