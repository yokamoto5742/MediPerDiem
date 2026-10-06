# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.必ず日本語で回答してください。

## 概要

医師別の外来日当円収益データ(Excel/CSV)を集計し、各医師配信用のファイルを出力する Windows デスクトップ GUI アプリ。

## 依存管理

- uv で管理する。追加は `uv add <package>`(開発用は `uv add --dev <package>`)、環境の再現は `uv sync`。

## 型チェック

- 型チェックは pyright で行う。リンター(ruff など)は使わないので導入しない。

## 構成

- `app/`(UI)→ `service/`(業務ロジック)→ `utils/`(設定・ログ)の一方向に依存させる。`main.py` がエントリーポイント。
- import はプロジェクトルート起点の絶対 import(例: `from utils.config_manager import load_config`)。

## データ

- `data/`(入力サンプル)と `logs/` は gitignore 済み。中身は読んでよいがコミットしない。
- 日本語 Windows 由来のファイルは CP932 の可能性がある。読み込みは UTF-8 を試してから CP932 にフォールバックする。
