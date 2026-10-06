# 変更履歴

このプロジェクトのすべての重要な変更は、このファイルに記録されます。

フォーマットは [Keep a Changelog](https://keepachangelog.com/ja/1.1.0/) に基づいており、
バージョン番号は [Semantic Versioning](https://semver.org/lang/ja/) に従っています。

## [Unreleased]

## [0.1.0] - 2026-10-06

### 追加
- 月次の外来収益ファイルを読み込み、変化表(外来日当円・外来収益合計)の対象月の列を更新
- 配信先マスタをもとに各医師配信用CSVを出力
- 更新前の変化表を `backup` フォルダへ自動バックアップ(`config.ini` の `[Backup] generations` で指定した世代数だけ保持、既定12世代)
