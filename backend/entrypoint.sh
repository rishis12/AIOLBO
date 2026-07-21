#!/bin/bash
# Entrypoint script for debugging Render environment issues

echo "=== ENTRYPOINT DEBUG START ==="
echo "Date: $(date)"
echo ""
echo "=== Command line arguments (\$@) ==="
echo "Number of args: $#"
for i in "$@"; do
    echo "  Arg: '$i'"
done
echo ""
echo "=== Environment Variables ==="
env | sort
echo ""
echo "=== Starting uvicorn ==="

# Run uvicorn
exec uvicorn api:app --host 0.0.0.0 --port 8000
