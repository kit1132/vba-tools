# Sheet Manager — シート管理ツール

Excelブック内のシートを管理するVBAマクロ集です。
シートが多いブック（月別・案件別など）で、目次のようにシート一覧を生成し、ワンクリックで各シートへジャンプできます。

---

## ファイル構成

| ファイル | 役割 |
|---------|------|
| `SheetManager.bas` | メインモジュール（UpdateSheetList） |
| `AddRowsAndGoHome.bas` | 全シート先頭に空行2行追加＋先頭シートへ移動 |
| `ThisWorkbook.cls` | ブックイベント（新シート作成時に一覧を自動更新） |
| `UpdateSheetList_old.bas` | UpdateSheetListの旧版（定数未使用・戻るリンクのハードコード版） |

---

## マクロ一覧

| マクロ名 | 機能 | いつ使うか |
|---------|------|-----------|
| **`UpdateSheetList`** | 「シート一覧」シートを生成・更新。各シートへのハイパーリンク付き | シートの追加・削除・名前変更をしたとき |
| `AddRowsAndGoHome` | 全シートの先頭に空行2行を追加し、先頭シートへ移動 | 各シートにヘッダー用の空行が必要なとき |

> **注意**: `AddRowsAndGoHome` は実行するたびに2行ずつ増えます。必要なとき以外は実行しないでください。

---

## 導入手順

1. Excelで対象ブックを開き、`Alt + F11` でVBAエディタを起動
2. **「ファイル」→「ファイルのインポート」** で `SheetManager.bas` をインポート
3. 必要に応じて `AddRowsAndGoHome.bas` もインポート
4. **「ファイル」→「名前を付けて保存」** → `.xlsm` 形式で保存

---

## SheetManager.bas の改善点（旧版との差分）

| 項目 | 旧版 (UpdateSheetList_old.bas) | 現行版 (SheetManager.bas) |
|------|------|------|
| シート名の管理 | `"シート一覧"` をハードコード | `Private Const LIST_SHEET_NAME` で定数化 |
| Option Explicit | なし | あり |
| ヘッダー太字 | なし | あり |
