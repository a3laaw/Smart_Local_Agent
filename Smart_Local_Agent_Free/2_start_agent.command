#!/usr/bin/env bash
TARGET="$1"
if [ -z "$TARGET" ]; then
  echo "Drag and drop ANY folder into this Terminal window and press Enter:"
  read -r TARGET
  TARGET="${TARGET%\'}"
  TARGET="${TARGET#\'}"
fi
cd "$TARGET" || { echo "Cannot open folder: $TARGET"; exit 1; }
echo "Active in: $(pwd)"
echo "Commands: /fillform | /revise | /writeletter | /replyletter | /excelform | /organize"
echo "          /submittals | /pettycash | /invoices | /letters | /contracts | /linklog"
MODEL_ARGS=()
if [ -f "$HOME/.gemini/ai_key.json" ] && command -v python3 >/dev/null 2>&1; then
  KEY_INFO="$(python3 -c "import json,pathlib;g=pathlib.Path.home()/'.gemini';d=json.loads((g/'ai_key.json').read_text('utf-8'));k=(d.get('gemini_api_key') or '').strip();m=(d.get('top_model') or '').strip();sf=g/'settings.json';s=json.loads(sf.read_text('utf-8')) if sf.exists() else {};s.setdefault('security',{}).setdefault('auth',{})['selectedType']='gemini-api-key';sf.write_text(json.dumps(s,indent=2),'utf-8') if k else None;print((k or '__NONE__')+'|'+(m or '__NONE__'))" 2>/dev/null || true)"
  IFS='|' read -r K_VAL M_VAL <<< "$KEY_INFO"
  if [ -n "$K_VAL" ] && [ "$K_VAL" != "__NONE__" ]; then
    export GEMINI_API_KEY="$K_VAL"
  fi
  if [ -n "$M_VAL" ] && [ "$M_VAL" != "__NONE__" ]; then
    MODEL_ARGS=(-m "$M_VAL")
  fi
fi
exec gemini "${MODEL_ARGS[@]}"
