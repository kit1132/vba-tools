# CLAUDE.md — vba-tools

## Purpose（WHY）

Excel 業務を効率化する VBA マクロ集。シート管理（一覧生成・ナビゲーション）と書式統一（フォント・Zoom 一括適用）の2ツールを提供する。macOS Excel 環境での利用を前提とする。

## Repo Map（WHAT）

```
vba-tools/
├── CLAUDE.md                    ← このファイル（プロジェクト全体のガイド）
├── README.md                    ← プロジェクト概要
├── sheet-manager/               ← シート管理ツール
│   ├── README.md                ← 仕様・導入手順
│   ├── SheetManager.bas         ← メイン: UpdateSheetList（一覧生成・戻るリンク付与）
│   ├── ThisWorkbook.cls         ← イベント: 新シート作成時に自動更新
│   └── old/
│       └── UpdateSheetList_old.bas ← 旧版（参考保存、定数未使用・ハードコード版）
├── format-workbook/             ← 書式統一ツール
│   ├── README.md                ← 仕様・導入手順
│   ├── FormatWorkbook.bas       ← メイン: セル・テーブル・グラフ・図形のフォント統一 + Zoom 80%
│   └── グラフや図形の配置設定をデフォルト.vb ← 配置設定リセット用（未実装）
├── .claude/
│   ├── hooks/                   ← フックスクリプト
│   │   ├── block-main-edit.sh   ← main ブランチ編集ブロック
│   │   └── check-option-explicit.sh ← Option Explicit チェック
│   ├── skills/                  ← Claude Code スキル定義
│   │   ├── code-review/SKILL.md
│   │   ├── refactor/SKILL.md
│   │   ├── debug/SKILL.md
│   │   └── release/SKILL.md
│   └── settings.json            ← フック設定
└── docs/
    ├── architecture.md          ← アーキテクチャ概要
    ├── adr/                     ← Architecture Decision Records
    │   ├── 000-template.md
    │   └── 001-font-meiryo-ui.md
    └── runbooks/
        ├── deploy.md            ← デプロイ（Excel への導入）手順
        └── incident.md          ← 障害対応フロー
```

## Rules & Commands（HOW）

### 開発コマンド

このプロジェクトにはビルド・テスト・lint のコマンドはない。VBA マクロは Excel VBA エディタ上で動作確認する。

### コーディング規約

- VBA ファイルは **UTF-8** で出力する（macOS Excel 向け）
- `Option Explicit` を全モジュールの先頭に記述する
- マジックナンバー・ハードコード文字列は `Private Const` で定数化する
- 変数宣言は `Dim` で型を明示する（`As String`, `As Long` 等）
- パフォーマンス最適化: `Application.ScreenUpdating = False` で画面更新を抑制し、処理後に復元する

### エラーハンドリング

- 動作実績のあるエラーハンドリングパターン（`On Error GoTo` / `On Error Resume Next`）は変更しない
- FormatWorkbook.bas の `ErrHandler` + `Cleanup` パターンは確立済み
- SheetManager.bas の局所的な `On Error Resume Next` → `On Error GoTo 0` パターンは確立済み

### 禁止事項

- PERSONAL.XLSB への直接操作（SheetManager で明示的にブロック済み）
- `old/` フォルダ内のコードを現行版に戻すこと
- フォーマッター・リンターの自動実行（VBA は対応ツールなし）

### レビュー・修正ルール

- レビューと修正は分離する（レビュー時はコード変更しない）
- 動作確認済みのコードへの構造変更はユーザー承認必須
- 各変更にリスク（高/低）と種類を明記する
- 詳細は docs/architecture.md を参照

### 重要モジュール

| モジュール | 注意点 |
|-----------|--------|
| SheetManager.bas | シート削除・行挿入を行うため、実行順序が重要。`On Error Resume Next` の範囲を変更しないこと |
| FormatWorkbook.bas | `ThisWorkbook` 参照で動作。`ActiveWorkbook` に変更すると意図しないブックに適用される可能性あり |
| ThisWorkbook.cls | イベントマクロ。`EnableEvents = False` 中は発火しない |
