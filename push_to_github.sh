#!/bin/bash
# ============================================================
# push_to_github.sh
# Первичный пуш монорепо construction-ai-mcp на GitHub
# 
# Использование:
#   1. Положи этот скрипт рядом с папкой construction-ai-mcp/
#   2. Добавь вручную серверные файлы (см. ниже)
#   3. Запусти: bash push_to_github.sh
# ============================================================

REPO_DIR="construction-ai-mcp"
GITHUB_USER="chernetchenko-dev"
GITHUB_REPO="construction-ai-mcp"
REMOTE="https://github.com/${GITHUB_USER}/${GITHUB_REPO}.git"

echo "=================================================="
echo "  construction-ai-mcp → GitHub"
echo "=================================================="
echo ""

# Проверка: папка существует?
if [ ! -d "$REPO_DIR" ]; then
  echo "❌ Папка $REPO_DIR не найдена."
  echo "   Распакуй архив рядом с этим скриптом."
  exit 1
fi

# Напоминание: добавить серверные файлы
echo "📋 Проверь что серверные файлы на месте:"
echo ""
echo "   $REPO_DIR/renga/renga_mcp_server_v2.py"
ls "$REPO_DIR/renga/renga_mcp_server_v2.py" 2>/dev/null \
  && echo "   ✅ renga_mcp_server_v2.py" \
  || echo "   ❌ ОТСУТСТВУЕТ — скопируй вручную!"

echo ""
echo "   $REPO_DIR/msproject/msproject_mcp_server.py"
ls "$REPO_DIR/msproject/msproject_mcp_server.py" 2>/dev/null \
  && echo "   ✅ msproject_mcp_server.py" \
  || echo "   ❌ ОТСУТСТВУЕТ — скопируй вручную!"

echo ""
read -p "Продолжить пуш? (y/n): " CONFIRM
if [ "$CONFIRM" != "y" ]; then
  echo "Отменено."
  exit 0
fi

cd "$REPO_DIR"

# Инициализация git если нужно
if [ ! -d ".git" ]; then
  git init
  echo "✅ git init"
fi

# Привязка remote
git remote remove origin 2>/dev/null
git remote add origin "$REMOTE"
echo "✅ remote → $REMOTE"

# Настройка ветки
git checkout -b main 2>/dev/null || git checkout main

# Добавить все файлы
git add .
git status --short

echo ""
read -p "Commit message (Enter = 'Initial release'): " MSG
MSG=${MSG:-"Initial release: Renga MCP + MSProject MCP + Revit skills"}

git commit -m "$MSG"

echo ""
echo "📤 Пушим на GitHub..."
echo "   (потребуется токен GitHub как пароль)"
echo "   Токен: https://github.com/settings/tokens → New → repo"
echo ""

git push -u origin main

echo ""
echo "=================================================="
echo "✅ Готово!"
echo "   https://github.com/${GITHUB_USER}/${GITHUB_REPO}"
echo "=================================================="
