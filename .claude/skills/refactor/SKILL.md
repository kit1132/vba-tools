---
name: refactor
description: >
  VBAリファクタリング、コード整理、構造改善、可読性向上、定数化、モジュール分割、
  「リファクタリングして」「整理して」「きれいにして」「改善して」「定数化して」で発動。
---

# VBA Refactor スキル

## リファクタリング原則

1. **動いているコードを壊さない** — 動作確認済みのコードへの構造変更はユーザー承認必須
2. **1変更1確認** — 複数変更がある場合は1つずつ適用して確認を促す
3. **安全な変更から始める** — 命名改善・定数化 → ロジック変更 → 構造変更の順

## リファクタリング手順

### Phase 1: 分析
1. 対象ファイルを全文読み込む
2. README.md で仕様を確認する
3. リファクタリング対象を洗い出す

### Phase 2: 分類と優先順位付け

変更を以下に分類する：

| 分類 | リスク | 例 |
|------|--------|-----|
| 安全 | 低 | コメント追加、変数名改善、`Private Const` 化 |
| 中程度 | 中 | 処理の共通関数化、引数の追加 |
| 構造変更 | 高 | エラーハンドリングパターン変更、制御フロー変更 |

### Phase 3: 提案

各変更を以下の形式でリスト化する：

```
### 変更 #N: [タイトル]
- **分類**: 安全 / 中程度 / 構造変更
- **該当箇所**: ファイル名:行番号
- **現状**: 現在のコード
- **変更後**: 提案するコード
- **理由**: なぜ変更するか
```

### Phase 4: 実行（ユーザー承認後）

- ユーザーが承認した項目のみ変更する
- リスク「高」の変更は単独で適用し、他の変更と混ぜない
- 変更後のコードを全文表示し、差分を明確にする

## VBA 固有のリファクタリングパターン

### 定数化
```vba
' Before
If ws.Name = "シート一覧" Then
' After
Private Const LIST_SHEET_NAME As String = "シート一覧"
If ws.Name = LIST_SHEET_NAME Then
```

### Application 状態の保存・復元パターン
```vba
' 推奨パターン（このプロジェクトで確立済み）
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
' ... 処理 ...
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
Application.ScreenUpdating = True
```

### 禁止事項
- エラーハンドリングの `On Error GoTo` / `On Error Resume Next` パターンを変更しない
- `ThisWorkbook` → `ActiveWorkbook` の変更は慎重に（FormatWorkbook.bas は `ThisWorkbook` で正しい）
- `old/` フォルダ内のコードは変更対象外
