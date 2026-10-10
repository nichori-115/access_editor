# Access AI CLI 実装計画書

Oct 10, 2026 · 設計コンセプト（`Access AI CLI 設計コンセプト.md`）に基づく実装計画

## 1. 本計画の位置づけ

設計コンセプトで定めた原則（非破壊既定・件数一致・範囲必須・1コマンド1操作・機械可読な出力）を、Python 製 CLI `acc` として実装するための計画である。本書では、技術選定、モジュール構成、コマンド仕様、データ形式、安全装置の実装方法、開発フェーズ、テスト方針、未決事項を定める。

スコープは設計コンセプトと同一で、「読む」と「レコードを書き換える」まで。書き換え後の連動処理（既存 VBA）は対象外。

## 2. 技術選定

| 項目 | 選定 | 理由 |
| --- | --- | --- |
| 言語 | Python 3.11 以上 | 設計コンセプトの前提。`tomllib` 標準搭載 |
| CLI フレームワーク | `typer`（内部は `click`） | サブコマンド構成、型ヒントからの引数定義、未知オプションを既定でエラーにできる |
| DB アクセス | `SQLAlchemy 2.x` + `pyodbc` | リフレクションで型・制約を自動取得。`mssql+pyodbc` で SQL Server に直結 |
| 開発用 DB | SQLite（SQLAlchemy 経由） | macOS/Linux 上で ODBC 無しに開発・テスト可能。方言差は Dialect 層で吸収 |
| 設定ファイル | TOML（`acc.toml`） | 権限設定・接続設定を人間が読める形で管理 |
| 出力 | TSV（既定）/ JSONL / JSON | 設計コンセプトのとおり |
| 監査ログ | SQLite（ローカルファイル） | 取り消しに必要な変更前後の値を構造化して保持 |
| パッケージ管理 | `uv` + `pyproject.toml` | Windows 側でも同一手順で環境構築できる |
| テスト | `pytest` | SQLite フィクスチャでコマンド単位のテストを行う |
| 配布 | `pipx install` または `uv tool install` | Windows 側で Python ビット数を Office に合わせて導入 |

### 2.1 Windows 環境に関する前提

- 実運用は Windows + Access ODBC ドライバ。Python と Office のビット数（32/64）を揃える
- 開発・CI は macOS/Linux + SQLite で行い、SQL Server 固有の挙動（`rowversion`、IDENTITY、照合順序）は Windows 上での結合テストで確認する
- SQL Server への接続文字列は `acc.toml` に置き、パスワードは環境変数参照（`${ACC_DB_PASSWORD}`）とする

## 3. リポジトリ構成

```
access_editor/
├── pyproject.toml
├── acc.example.toml              # 設定ファイルの雛形
├── src/acc/
│   ├── __init__.py
│   ├── cli.py                    # typer アプリのエントリ。サブコマンド登録のみ
│   ├── config.py                 # acc.toml の読み込みと検証（pydantic モデル）
│   ├── db.py                     # エンジン生成、接続、トランザクション管理
│   ├── schema/
│   │   ├── reflect.py            # SQLAlchemy リフレクション → 内部モデル
│   │   ├── model.py              # TableDef / ColumnDef / BusinessRule
│   │   ├── render.py             # sqlacodegen 風テキスト出力
│   │   └── rules.py              # 業務ルール（info/comment）のパースと検証
│   ├── io/
│   │   ├── tsv.py                # PostgreSQL COPY 互換のエスケープ/アンエスケープ
│   │   ├── jsonl.py
│   │   └── stdin_keys.py         # 標準入力からの主キー読み取り
│   ├── commands/
│   │   ├── schema_cmd.py         # acc schema
│   │   ├── read_cmd.py           # acc read
│   │   ├── replace_cmd.py        # acc replace
│   │   ├── set_cmd.py            # acc set
│   │   ├── append_cmd.py         # acc append
│   │   ├── delete_cmd.py         # acc delete（設定で許可時のみ有効）
│   │   ├── log_cmd.py            # acc log（監査ログ参照）
│   │   └── undo_cmd.py           # acc undo
│   ├── write/
│   │   ├── target.py             # 対象行の特定（主キー or 条件）
│   │   ├── preview.py            # 変更前→変更後のプレビュー生成
│   │   ├── guard.py              # 件数一致・範囲必須・最大行数・楽観ロック
│   │   ├── validate.py           # 型・NOT NULL・桁数・業務ルール検証
│   │   └── apply.py              # トランザクション内での実反映
│   ├── audit/
│   │   ├── store.py              # 監査ログ SQLite の読み書き
│   │   └── undo.py               # ログからの逆操作生成
│   ├── errors.py                 # 例外階層と終了コードの対応
│   └── output.py                 # 人間向け/JSON 向け出力の切り替え
├── tests/
│   ├── conftest.py               # SQLite フィクスチャ（サンプルテーブル投入）
│   ├── test_tsv.py
│   ├── test_schema.py
│   ├── test_read.py
│   ├── test_replace.py
│   ├── test_set.py
│   ├── test_append.py
│   ├── test_guard.py
│   ├── test_audit_undo.py
│   └── test_cli_errors.py
└── docs/
    ├── Access AI CLI 設計コンセプト.md
    └── Access AI CLI 実装計画書.md
```

## 4. コマンド仕様

コマンド名は `acc`。既存ツールと同名のサブコマンド・オプションは使わない。

### 4.1 共通オプション

| オプション | 意味 |
| --- | --- |
| `--config PATH` | 設定ファイル。既定はカレントの `acc.toml`、無ければ `~/.config/acc/acc.toml` |
| `--json` | 結果を JSON で標準出力に出す（人間向け文言は標準エラーへ） |
| `--profile NAME` | 設定内の接続プロファイルを切り替える（`dev` / `prod` 等） |

未知のオプションや引数は typer の既定動作でエラー終了（終了コード 2）。黙って無視しない。

### 4.2 `acc schema`

```
acc schema                      # テーブル一覧と1行説明
acc schema <table>              # 全列の型・制約を sqlacodegen 風に表示
acc schema <table> <column>     # 1列の型 + 業務ルール
```

- 表示対象は設定の許可リストにあるテーブルのみ
- 書き込み禁止列（IDENTITY、rowversion、計算列、トリガー充填列）は `# readonly: identity` のようにコメントで明示
- 業務ルールは `info={"rule": "区分は01〜05のみ", "allowed": ["01",...,"05"]}` の形で表示
- リフレクション結果はキャッシュ可能（`--refresh` で再取得）

### 4.3 `acc read`

```
acc read <table> --cols 受注番号,区分,得意先名 [--where "区分 = '01'"] [--limit N] [--fmt tsv|jsonl] [--out PATH]
```

- `--cols` は必須。主キー列は指定が無くても先頭に自動付与する
- `--where` は許可された列名と比較演算子のみを受け付ける簡易構文（独自パーサ）。生 SQL は通さない
- `--limit` 省略時は設定の `read.max_rows`（既定 10000）。超過時は打ち切りではなくエラー
- `--out` 指定時はファイルを読み取り専用（0444）で書き出す
- 出力の 1 行目は列名。TSV のエスケープは PostgreSQL `COPY` テキスト形式
- 楽観ロック用に、各行の `rowversion`（無ければ全列ハッシュ）を `--with-version` で末尾列として出せる

### 4.4 `acc replace`（置換）

```
cat keys.txt | acc replace <table> <column> --from "旧" --to "新" [--regex] --expect N [--commit]
acc replace <table> <column> --from "旧" --to "新" --where "..." --expect N [--commit]
```

- 対象は標準入力の主キー、または `--where`。どちらも無ければエラー（範囲必須）
- `--commit` 無しはプレビューのみ。「対象 N 行、変更前 → 変更後」を表示し DB は変更しない
- `--commit` 時は `--expect N` 必須。ヒット件数と不一致なら中止
- `--regex` 指定時は Python `re.sub`。未指定はリテラル置換
- 置換の結果値が変わらない行は対象件数に含めない（無意味な更新を避ける）

### 4.5 `acc set`（上書き）

```
cat keys.txt | acc set <table> <column> --value "固定値" --expect N [--commit]
acc set <table> <column> --null --where "..." --expect N [--commit]
```

- `replace` と同じ対象指定・プレビュー・件数一致ルール
- `--null` で NULL を設定（NOT NULL 列はエラー）
- 複数列を一度に上書きする場合は `--value` を列ごとに繰り返し指定できる（`acc set <table> --col 区分=02 --col 備考=…`）。1 コマンド 1 操作の原則は「操作種別が 1 つ」の意味で維持する

### 4.6 `acc append`（追加）

```
cat rows.tsv | acc append <table> --expect N [--commit]
```

- 標準入力の TSV 1 行目を列名として照合。未知の列名、書き込み禁止列の指定はエラー
- 省略された NOT NULL 列（DB 既定値なし）はエラー
- 行数が `--expect` と一致しなければ中止
- プレビューでは挿入予定行を型変換後の値で表示

### 4.7 `acc delete`（削除）

- 設定 `write.allow_delete = true` のときのみサブコマンドとして登録される。未許可時は「このコマンドは設定で無効です」とエラー
- 対象指定・件数一致・プレビューは他の書き込みと同一
- 削除前の全列値を監査ログに保存し、`undo` で再挿入できるようにする

### 4.8 `acc log` / `acc undo`

```
acc log [--last N] [--table T]      # 監査ログ一覧
acc log show <op_id>                # 1 操作の変更前後を表示
acc undo <op_id> [--commit]         # 逆操作をプレビュー/実行
```

- `undo` は対象行が記録時点の「変更後の値」と一致している場合のみ実行可能。それ以外は中止（別の変更が入っている）
- `undo` 自体も監査ログに記録される

## 5. データ形式

### 5.1 TSV

- 文字コード UTF-8、改行 LF
- NULL は `\N`
- セル内のタブ・改行・バックスラッシュは `\t` `\n` `\\` にエスケープ。読み込み時は逆変換
- 日付・時刻は ISO 8601、小数は `Decimal` の文字列表現。浮動小数の丸めは行わない
- 1 行目は列名。先頭列は主キー（複合主キーは複数列）

### 5.2 JSONL

- 1 行 1 オブジェクト。NULL は `null`、日付は ISO 8601 文字列、`Decimal` は文字列
- 列順は `--cols` の指定順

### 5.3 `--json` 出力（結果）

```json
{
  "ok": true,
  "command": "replace",
  "table": "受注WK",
  "preview": true,
  "matched": 3,
  "changed": 3,
  "rows": [{"pk": {"受注番号": "A001"}, "before": "旧", "after": "新"}],
  "op_id": null
}
```

エラー時:

```json
{"ok": false, "error": {"code": "COUNT_MISMATCH", "message": "expected 3 rows but matched 5"}}
```

## 6. 設定ファイル `acc.toml`

```toml
[profiles.dev]
url = "sqlite:///dev.db"

[profiles.prod]
url = "mssql+pyodbc://user:${ACC_DB_PASSWORD}@host/db?driver=ODBC+Driver+18+for+SQL+Server"

[read]
max_rows = 10000

[write]
allow_delete = false
max_rows_per_command = 500
prefer_work_tables = true

[tables."受注WK"]
description = "受注取込ワーク"
writable_columns = ["区分", "備考", "得意先名"]
readonly_columns = ["ID", "更新日時"]     # リフレクションで検出できない分を補う

[tables."受注WK".rules."区分"]
rule = "区分は01〜05のみ"
allowed = ["01", "02", "03", "04", "05"]

[tables."受注"]
description = "受注本体"
writable = false                           # 本番テーブルへの直接書き込みを禁止

[audit]
path = "~/.local/share/acc/audit.sqlite"
```

- 設定に無いテーブルは読み取り不可・書き込み不可（許可リスト方式）
- `rules` は検証にも `schema` 表示にも同一ソースとして使う

## 7. 安全装置の実装方針

| 装置 | 実装箇所 | 方法 |
| --- | --- | --- |
| プレビュー既定 | `commands/*`、`write/apply.py` | `--commit` が無い限り `apply` を呼ばない。プレビューは `SELECT` のみ |
| 件数一致 | `write/guard.py` | `--commit` 時に `--expect` 必須。`SELECT COUNT(*)` と比較し不一致で `CountMismatch` |
| 範囲必須 | `write/target.py` | 主キーも `--where` も無い場合 `NoTargetSpecified` |
| 最大行数 | `write/guard.py` | `max_rows_per_command` 超過で中止 |
| 楽観ロック | `write/guard.py` | 標準入力に `rowversion`（またはハッシュ）列が含まれていれば `UPDATE ... WHERE pk=? AND version=?` で行ごとに検査。更新件数が期待と違えばロールバック |
| トランザクション | `db.py`、`write/apply.py` | 1 コマンド = 1 トランザクション。例外時は全件ロールバック |
| 監査ログ | `audit/store.py` | `apply` 直前に変更前値、直後に変更後値を同一 `op_id` で保存。コマンドライン文字列・実行者・時刻も記録 |
| 権限設定 | `config.py` | テーブル・列の許可リスト、削除可否、最大行数を起動時に検証 |
| 未対応は止める | `cli.py` | typer の未知オプションエラー。`--where` パーサも未対応構文で即エラー |
| 型・制約検証 | `write/validate.py` | 反映前に Python 側で型変換・桁数・NOT NULL・業務ルールを検査し、違反内容をそのまま返す |

### 7.1 `--where` 簡易構文

生 SQL を受け付けない代わりに、次の構文のみ対応する。

```
<列名> <演算子> <値> [and <列名> <演算子> <値> ...]
演算子: = != < <= > >= like in is null is not null
```

- 列名は設定の許可リストと照合
- 値はバインドパラメータとして渡す（文字列結合しない）
- `or` と括弧は初期スコープ外（必要なら主キーをパイプで渡す）

## 8. 開発フェーズ

### フェーズ 0: 土台（1〜2 日）

- `pyproject.toml`、`src/acc` パッケージ、`typer` エントリ、`acc --version`
- `config.py`（TOML 読み込み・pydantic 検証）
- `db.py`（SQLite / mssql エンジン生成）
- `errors.py` と終了コード対応表
- `tests/conftest.py` にサンプルテーブル（`受注WK`、`受注`）を持つ SQLite フィクスチャ

### フェーズ 1: 読み取り（2〜3 日）

- `io/tsv.py` のエスケープ/アンエスケープと往復テスト
- `acc schema`（リフレクション → 内部モデル → sqlacodegen 風表示、業務ルール合成）
- `acc read`（`--cols` 必須、`--where` パーサ、`--limit`、`--fmt`、`--out` 読み取り専用書き出し）

### フェーズ 2: 書き込みの基盤（2〜3 日）

- `write/target.py`（標準入力主キー / `--where`）
- `write/preview.py`、`write/guard.py`、`write/validate.py`、`write/apply.py`
- 監査ログの保存（`audit/store.py`）

### フェーズ 3: 書き込みコマンド（3〜4 日）

- `acc set` → `acc replace`（リテラル → 正規表現）→ `acc append` の順
- 各コマンドでプレビュー / `--commit` / `--expect` / 件数不一致 / 最大行数のテスト
- 楽観ロック（`--with-version` 出力と、標準入力でのバージョン照合）

### フェーズ 4: 監査・取り消し・削除（2 日）

- `acc log`、`acc log show`、`acc undo`
- `acc delete`（設定で有効化時のみ）

### フェーズ 5: Windows 結合（2〜3 日、Windows 環境が必要）

- `mssql+pyodbc` での接続、ビット数確認手順の文書化
- IDENTITY / rowversion / 計算列のリフレクション確認
- 日本語列名・テーブル名の引用符処理確認
- 実データ（ワークテーブル）での read → set → VBA 連動の通し確認

### フェーズ 6: AI 向け整備（1〜2 日）

- `acc --help` の文言を AI が読んで正しく使える粒度に調整
- 主要テーブル 2〜3 個の概要を出す `acc schema --summary`
- エラーメッセージの自己修正可能性レビュー（違反した制約と与えられた値を必ず含める）
- README（人間向け）と `AGENTS.md`（AI 向けの短い利用ガイド）

## 9. テスト方針

- 単体: TSV エスケープ往復、`--where` パーサ、検証器（型・桁数・業務ルール）、ガード（件数・範囲・最大行数）
- コマンド: `typer.testing.CliRunner` で各サブコマンドをプレビュー / commit / 失敗系で実行し、DB 状態と出力（TSV/JSON）と終了コードを検証
- 監査: 操作 → `undo` → 元に戻ること、第三者の変更後は `undo` が中止すること
- 方言: SQLite で全テストを通し、Windows 上で `mssql` マーカー付きテストのみ追加実行
- 回帰: プレビュー実行後に DB が一切変わっていないことを全書き込みテストで共通アサート

## 10. 終了コード

| コード | 意味 |
| --- | --- |
| 0 | 成功（プレビューを含む） |
| 1 | 実行時エラー（DB 接続失敗など） |
| 2 | 引数・オプションの誤り、未知のオプション |
| 3 | 検証エラー（型・制約・業務ルール違反） |
| 4 | ガードによる中止（件数不一致・範囲なし・最大行数超過・楽観ロック失敗） |
| 5 | 権限エラー（許可されていないテーブル・列・操作） |

## 11. 未決事項と判断期限

| 事項 | 選択肢 | 判断時期 |
| --- | --- | --- |
| 削除操作を CLI に持たせるか | 持たない / 設定で許可時のみ（本計画の案） | フェーズ 4 着手前 |
| 楽観ロックの既定 | 任意（`--with-version` 指定時のみ） / 書き込み時は必須 | フェーズ 3 |
| 複合主キーの標準入力形式 | TSV（列名ヘッダ付き） / `key1\tkey2` ヘッダ無し | フェーズ 2 |
| 監査ログの保管場所 | ローカル SQLite / SQL Server 上の専用テーブル | フェーズ 5 |
| `--where` に `or` を許すか | 許さない（主キーパイプで代替） / 括弧付きで許す | フェーズ 6 |
| 本番テーブルの読み取りを許すか | 読み取りは許可・書き込みのみ禁止（本計画の案） | フェーズ 1 |

## 12. 最初の一歩

1. フェーズ 0 を実施し、`acc --version` と `acc schema`（SQLite フィクスチャ）が動く状態を作る
2. その時点で `acc.example.toml` と `AGENTS.md` の下書きを置き、AI に触らせて `--help` の分かりやすさを確認する
3. フェーズ 1 完了時点で設計コンセプトの作成者にレビューを依頼し、コマンド名・オプション名を確定してから書き込み系に進む
