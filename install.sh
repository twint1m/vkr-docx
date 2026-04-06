#!/bin/bash
# Установка vkr-docx — скилл для оформления ВКР по ГОСТ
# Запуск: curl -fsSL https://raw.githubusercontent.com/twint1m/vkr-docx/main/install.sh | bash

set -e

REPO="https://github.com/twint1m/vkr-docx.git"
INSTALL_DIR="${HOME}/projects/vkr-docx"

echo "=== Установка vkr-docx ==="

# 1. Клонировать или обновить
if [ -d "$INSTALL_DIR" ]; then
    echo "Обновление существующей установки..."
    cd "$INSTALL_DIR" && git pull
else
    echo "Клонирование репозитория..."
    mkdir -p "$(dirname "$INSTALL_DIR")"
    git clone "$REPO" "$INSTALL_DIR"
fi

# 2. Установить Python-пакет
echo "Установка Python-пакета..."
cd "$INSTALL_DIR"
pip install -e . 2>/dev/null || pip install -e . --break-system-packages 2>/dev/null || python3 -m pip install -e . --break-system-packages

# 3. Скопировать скилл Claude Code
echo "Установка скилла /vkr..."
mkdir -p "${HOME}/.claude/commands"
cp "${INSTALL_DIR}/.claude/commands/vkr.md" "${HOME}/.claude/commands/vkr.md"

# 4. Добавить CLAUDE.md (если нет или не содержит ВКР)
if [ ! -f "${HOME}/CLAUDE.md" ]; then
    cp "${INSTALL_DIR}/CLAUDE.md" "${HOME}/CLAUDE.md"
    echo "Создан ~/CLAUDE.md"
elif ! grep -q "vkr_docx" "${HOME}/CLAUDE.md"; then
    echo "" >> "${HOME}/CLAUDE.md"
    cat "${INSTALL_DIR}/CLAUDE.md" >> "${HOME}/CLAUDE.md"
    echo "Добавлено в ~/CLAUDE.md"
else
    echo "~/CLAUDE.md уже содержит настройки ВКР"
fi

echo ""
echo "=== Готово! ==="
echo "  Скилл:  /vkr"
echo "  CLI:    vkr-docx --help"
echo "  Python: from vkr_docx import VKRDocument"
echo ""
echo "Перезапустите Claude Code для активации скилла."
