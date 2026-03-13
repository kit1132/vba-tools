#!/bin/bash
# 編集後に .bas/.cls ファイルの Option Explicit チェック
# PostToolUse で実行：警告のみ（ブロックしない）

# stdin から hook input を読む
input=$(cat)
tool_name=$(echo "$input" | python3 -c "import sys,json; d=json.load(sys.stdin); print(d.get('tool_input',{}).get('file_path',''))" 2>/dev/null)

if [[ "$tool_name" == *.bas ]] || [[ "$tool_name" == *.cls ]]; then
  if ! head -5 "$tool_name" | grep -q "Option Explicit"; then
    echo "Warning: $tool_name に Option Explicit がありません。追加を推奨します。" >&2
  fi
fi
exit 0
