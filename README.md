# MediPerDiem

医師別の外来日当円収益データ(Excel)を集計し、変化表の更新と各医師配信用ファイルの出力を行う Windows デスクトップアプリ(Tkinter)です。

## ✨ 主な機能

- 月次の医師別収益ファイルを読み込み、変化表(「外来日当円」「外来収益合計」シート)の対象月の列を更新
- 更新前の変化表を自動バックアップ(世代数は設定可能)
- 各医師配信用CSV、常勤医師外来日当円平均の推移(xlsx)を出力
- Thunderbird の Mail Merge 用CSV(各医師向け・医局KPI向け)を出力
- IDが異なる同姓同名の医師は、IDが若い方へ収益と人数を合算
- 診療科目が「健」(健診)の行は集計対象外

## ⚙️ 動作環境

- Windows
- Python 3.13 以上
- [uv](https://docs.astral.sh/uv/)

## 🚀 セットアップ

```bash
git clone https://github.com/yokamoto5742/MediPerDiem
cd MediPerDiem
uv sync
```

`utils/config.ini` の `[Paths]` に、変化表・配信先マスタ・各出力ファイルのパスを設定します。

| キー | 内容 |
| --- | --- |
| `trend_workbook` | 変化表(xlsx) |
| `recipients_csv` | 配信先マスタ(CSV) |
| `distribution_csv` | 各医師配信用CSVの出力先 |
| `average_workbook` | 常勤医師外来日当円平均の推移(xlsx)の出力先 |
| `mail_merge_csv` | 各医師メール差し込み用CSVの出力先 |
| `kpi_mail_merge_csv` | 医局KPIメール差し込み用CSVの出力先 |
| `backup_dir` | 変化表のバックアップ先 |

`[Backup] generations` でバックアップの保持世代数を指定します。

## 📖 使い方

```bash
uv run python main.py
```

1. 「参照...」で医師別収益データ(Excel)を選択する。ファイル名の末尾 `YYYYMM` が対象月になる(例: `医師別収益データ202608.xlsx`)。
2. 「集計を実行」を押す。
3. 完了ダイアログを確認する。変化表に医師の行を追加した場合は、その一覧が表示される。

同じ月のデータを再実行すると対象月の列が上書きされます。詳細は [データ処理再実行時の注意点](docs/データ処理再実行時の注意点.md) を参照してください。

## 📁 プロジェクト構成

```
MediPerDiem/
├── main.py              # エントリーポイント
├── app/                 # UI(Tkinter)と画面表示メッセージ(constants.py)
├── service/             # 業務ロジック
│   ├── revenue_reader.py      # 収益データの読み込み・集計
│   ├── trend_workbook.py      # 変化表の更新・バックアップ
│   ├── distribution_csv.py    # 各医師配信用CSV
│   ├── average_workbook.py    # 平均の推移ブック
│   └── mail_merge_csv.py      # Mail Merge用CSV
├── utils/               # 設定(config.ini)・ログ
├── tests/               # pytest(service/ と同じ構成)
├── docs/                # CHANGELOG、運用手順
└── build.py             # PyInstaller による実行ファイル作成
```

依存の向きは `app/` → `service/` → `utils/` の一方向です。

## 🛠 開発

```bash
uv run pytest tests/ -v --tb=short   # テスト
uv run pyright                       # 型チェック
uv run python build.py               # exe 作成(PyInstaller が必要)
```

## 🔧 トラブルシューティング

- **「ファイルに書き込めません」と表示される**: 変化表などを Excel で開いたままになっています。閉じてから再実行してください。
- **「ファイル名から対象月(YYYYMM)を取得できません」と表示される**: 収益データのファイル名の末尾を `YYYYMM` にしてください。
- **CSV が文字化けする**: 日本語 Windows 由来のファイルは CP932 の場合があります。読み込みは UTF-8 を試してから CP932 にフォールバックします。
- **処理の詳細を確認したい**: `logs/` のログを確認してください。

## ライセンス

このプロジェクトのライセンス情報については、 [LICENSE](docs/LICENSE) を参照してください。

## 更新履歴

更新履歴は [CHANGELOG.md](docs/CHANGELOG.md) を参照してください。
