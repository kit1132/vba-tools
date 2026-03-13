# デプロイ手順（Excel への導入）

VBA マクロを Excel ブックに導入する手順。

## 前提条件

- Excel がインストールされていること（Windows / macOS）
- マクロの実行が許可されていること（セキュリティ設定）

## sheet-manager の導入

### Step 1: VBA エディタを開く
- `Alt + F11`（Windows）
- `Opt + F11` または「ツール → マクロ → Visual Basic Editor」（macOS）

### Step 2: モジュールをインポート
1. 「ファイル」→「ファイルのインポート」
2. `SheetManager.bas` を選択してインポート

### Step 3: イベントマクロを設定（任意）
1. VBA エディタのプロジェクトエクスプローラーで `ThisWorkbook` をダブルクリック
2. `ThisWorkbook.cls` の内容をコピー＆ペースト

### Step 4: 保存
- 「ファイル」→「名前を付けて保存」→ `.xlsm` 形式を選択

### Step 5: 実行
- `Alt + F8` → `UpdateSheetList` を選択 →「実行」

## format-workbook の導入

### Step 1〜2: 上記と同様
- `FormatWorkbook.bas` をインポート

### Step 3: 保存
- `.xlsm` 形式で保存

### Step 4: 実行
- `Alt + F8` → `FormatWorkbook` を選択 →「実行」

## PERSONAL.XLSB への登録（全ブック共通で使う場合）

1. 新規ブックを開き、VBA エディタで `PERSONAL.XLSB` を展開
2. 上記のモジュールをインポート
3. Excel を閉じる際に PERSONAL.XLSB の保存確認で「はい」を選択

> **注意**: SheetManager は `PERSONAL.XLSB` 自体への実行をブロックします。常に対象ブックをアクティブにしてから実行してください。

## ロールバック

1. VBA エディタで該当モジュールを右クリック →「〇〇の解放」
2. `.xlsm` を保存（または `.xlsx` に変換してマクロを完全削除）
