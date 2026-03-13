# VBA Tools

Excel業務を効率化するVBAマクロ集です。機能ごとにフォルダを分けて管理しています。

---

## フォルダ構成

```
vba-tools/
├── README.md                   ← このファイル
├── sheet-manager/              ← シート管理ツール
│   ├── README.md
│   ├── SheetManager.bas        ← シート一覧の生成・更新（現行版）
│   ├── AddRowsAndGoHome.bas    ← 全シート先頭に空行追加
│   ├── ThisWorkbook.cls        ← 新シート作成時の自動更新イベント
│   └── UpdateSheetList_old.bas ← UpdateSheetListの旧版（参考保存）
└── format-workbook/            ← 書式統一ツール
    ├── README.md
    ├── FormatWorkbook.bas      ← フォント・Zoom一括統一
    └── グラフや図形の配置設定をデフォルト.vb  ← 配置設定リセット用（未実装）
```

---

## ツール概要

| フォルダ | 用途 | メインマクロ |
|---------|------|------------|
| `sheet-manager/` | シート一覧の自動生成、シート間ナビゲーション | `UpdateSheetList` |
| `format-workbook/` | ブック全体のフォント・表示倍率の統一 | `FormatWorkbook` |

各ツールの詳細は、フォルダ内のREADMEを参照してください。
