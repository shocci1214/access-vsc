#!/bin/sh
# Access VSC: コミット前のコードチェック（Access 不要・数秒）
# エラーがあるとコミットを止めます。どうしてもコミットしたいときは git commit --no-verify
out=$(powershell.exe -NoProfile -ExecutionPolicy Bypass -File "{{TOOLS}}/access.ps1" lint -Src "{{SRC}}" -Text 2>&1)
code=$?
if [ -n "$out" ]; then
  echo "$out"
fi
if [ $code -ne 0 ]; then
  echo ""
  echo "Access VSC: コードチェックでエラーがあります。直してからコミットしてください。"
  exit 1
fi
exit 0
