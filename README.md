# Redmine Gantt Excel Generator

Redmine のチケットを REST API 経由で取得し、ガント風の Excel (`.xlsx`) を自動生成する Python スクリプトです。

- メインシート: `Gantt`
- 生データシート: `Issues_Raw`
- `ProjectPath` は別列出力
- `Subject` は素の件名のまま出力
- 親子チケットは WBS とインデントで表現
- `start_date` / `due_date` を使ってガントバーを描画
- 進捗率 (`done_ratio`) をバー色で反映
- 関連チケット (`relations`) を文字列で併記
- URL / APIキーは INI ファイルから読込可能

---

## 1. 必要環境

- Python 3.10 以上推奨
- Redmine REST API にアクセスできること
- Redmine の API キーを発行済みであること

### 必要ライブラリ

```bash
pip install requests openpyxl
```

---

## 2. 同梱ファイル

- `redmine_gantt_excel.py`
- `redmine_gantt.ini`
- `README_redmine_gantt_excel.md`

---

## 3. INI ファイル設定

`redmine_gantt.ini`

```ini
[redmine]
base_url = https://redmine.example.com
api_key = YOUR_API_KEY
```

### 説明

- `base_url` : Redmine のベース URL
  - 例: `https://redmine.example.com`
- `api_key` : Redmine の個人 API キー

CLI 引数で `--base-url` / `--api-key` を指定した場合は、そちらが優先されます。

---

## 4. 基本的な使い方

### 最小実行

```bash
python redmine_gantt_excel.py \
  --config redmine_gantt.ini \
  --project-id 123 \
  --output redmine_gantt.xlsx
```

### 全ステータス対象

```bash
python redmine_gantt_excel.py \
  --config redmine_gantt.ini \
  --project-id 123 \
  --status-id "*" \
  --output redmine_gantt.xlsx
```

### 週単位で圧縮して出力

```bash
python redmine_gantt_excel.py \
  --config redmine_gantt.ini \
  --project-id 123 \
  --timeline-mode week \
  --output redmine_gantt_weekly.xlsx
```

### 期間を固定して出力

```bash
python redmine_gantt_excel.py \
  --config redmine_gantt.ini \
  --project-id 123 \
  --from-date 2026-04-01 \
  --to-date 2026-06-30 \
  --output redmine_gantt_q2.xlsx
```

### 担当者で絞る

```bash
python redmine_gantt_excel.py \
  --config redmine_gantt.ini \
  --project-id 123 \
  --assigned-to-id me \
  --output my_tasks_gantt.xlsx
```

### 特定バージョンで絞る

```bash
python redmine_gantt_excel.py \
  --config redmine_gantt.ini \
  --project-id 123 \
  --fixed-version-id 5 \
  --output version5_gantt.xlsx
```

### 日付なしチケットも表に含める

```bash
python redmine_gantt_excel.py \
  --config redmine_gantt.ini \
  --project-id 123 \
  --include-no-date-issues \
  --output redmine_gantt_with_nodate.xlsx
```

### relations を個票取得で補完する

```bash
python redmine_gantt_excel.py \
  --config redmine_gantt.ini \
  --project-id 123 \
  --fallback-relations \
  --sleep-sec 0.1 \
  --output redmine_gantt.xlsx
```

---

## 5. 主な引数

| 引数 | 説明 |
|---|---|
| `--config` | INI ファイルパス |
| `--base-url` | Redmine ベース URL |
| `--api-key` | Redmine API キー |
| `--project-id` | 対象プロジェクト ID |
| `--status-id` | ステータス条件。`*` で全件 |
| `--tracker-id` | トラッカー ID をカンマ区切り指定 |
| `--assigned-to-id` | 担当者 ID。`me` 可 |
| `--fixed-version-id` | 対象バージョン ID |
| `--subproject-id` | サブプロジェクト条件 |
| `--limit` | API 1回あたりの取得件数 |
| `--output` | 出力ファイル名 |
| `--sheet-name` | メインシート名 |
| `--from-date` | ガント開始日上書き (`YYYY-MM-DD`) |
| `--to-date` | ガント終了日上書き (`YYYY-MM-DD`) |
| `--timeline-mode` | `day` または `week` |
| `--fallback-relations` | 個票 API で relations を補完 |
| `--sleep-sec` | API 呼出し間隔（秒） |
| `--include-no-date-issues` | 日付なしチケットを表に含める |

---

## 6. 出力内容

### 6.1 メインシート `Gantt`

左側に属性列、右側にガント領域を出します。

#### 固定列

- `WBS`
- `ID`
- `ProjectPath`
- `Subject`
- `Tracker`
- `Status`
- `Priority`
- `Assignee`
- `Version`
- `Start`
- `Due`
- `Progress`
- `Est.Hours`
- `Parent`
- `Relations`

#### 表示仕様

- `ProjectPath` は `親 / 子 / 孫` 形式
- `Subject` は素の件名
- 親子階層は `Subject` にインデント、`WBS` に階層番号を付与
- 親チケット判定された行は件名セルを強調表示
- 土日は薄灰色
- 当日は薄黄色
- 通常バーと進捗バーは色分け
- 親チケット行のバーは別色
- 印刷は横向き A3、ヘッダ行固定

### 6.2 生データシート `Issues_Raw`

取得したチケット情報を一覧で出力します。

列:

- `ID`
- `Project`
- `ProjectPath`
- `Tracker`
- `Status`
- `Priority`
- `Subject`
- `AssignedTo`
- `Version`
- `ParentID`
- `StartDate`
- `DueDate`
- `DoneRatio`
- `EstimatedHours`
- `Relations`
- `CustomFields`

---

## 7. ProjectPath の考え方

`projects.json` を取得し、親子関係から以下のようなパスを組み立てます。

```text
全社PJ / 基盤更改 / API連携
```

この値を `ProjectPath` 列に表示することで、サブプロジェクトが混在していても所属先を見分けやすくしています。

---

## 8. relations の扱い

`Relations` 列には、たとえば以下のような形式で出力します。

```text
precedes #102, blocks #205(+2d), relates #301
```

環境によっては一覧 API の `include=relations` で relations が十分に返らないことがあります。その場合は `--fallback-relations` を付けると、各チケットを個別取得して relations を補完します。

その分 API コール数は増えるため、必要なら `--sleep-sec` も設定してください。

---

## 9. よくある使い方

### 自分の担当だけ毎週出したい

```bash
python redmine_gantt_excel.py \
  --config redmine_gantt.ini \
  --project-id 123 \
  --assigned-to-id me \
  --timeline-mode week \
  --output weekly_my_tasks.xlsx
```

### ある版だけ月次レビュー用に出したい

```bash
python redmine_gantt_excel.py \
  --config redmine_gantt.ini \
  --project-id 123 \
  --fixed-version-id 7 \
  --from-date 2026-05-01 \
  --to-date 2026-05-31 \
  --output release_review.xlsx
```

### 親プロジェクト配下をまとめて見たい

```bash
python redmine_gantt_excel.py \
  --config redmine_gantt.ini \
  --project-id 123 \
  --status-id "*" \
  --output full_project_gantt.xlsx
```

---

## 10. エラー時の確認ポイント

### 401 / 403 が出る

- API キーが誤っている
- API 利用権限がない
- `base_url` が違う

### `No issues found for the given filters.`

- `project-id` が違う
- `status-id` / `tracker-id` / `assigned-to-id` / `fixed-version-id` の条件が厳しすぎる
- 指定プロジェクトにチケットが存在しない

### ガントにバーが出ない

- `start_date` または `due_date` が未設定
- `from-date` / `to-date` の指定範囲外

### ProjectPath が空になる

- API で `projects.json` を取得できていない
- チケットの `project_id` に対する参照先が見つからない

---

## 11. 補足

- このスクリプトは **Redmine 標準ガントをそのまま Excel 化するものではなく**、Redmine のチケット情報から **ガント風レポートを生成する** ものです。
- そのため、Excel 側の見た目は自由に改造しやすい構成になっています。
- 今後の拡張候補としては以下が考えられます。
  - 担当者ごとの色分け
  - 親チケットの期間自動集約
  - マイルストーン列追加
  - 保存済みクエリ対応
  - project identifier 指定対応
  - 複数シート同時出力

---

## 12. 実行例まとめ

```bash
pip install requests openpyxl

python redmine_gantt_excel.py \
  --config redmine_gantt.ini \
  --project-id 123 \
  --status-id "*" \
  --timeline-mode week \
  --output redmine_gantt.xlsx
```

