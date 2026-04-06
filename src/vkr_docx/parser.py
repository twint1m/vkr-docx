"""Парсер markdown-файла главы ВКР в структурированные блоки."""

import os
import re
import glob
from dataclasses import dataclass, field


@dataclass
class Block:
    """Блок контента из markdown."""
    kind: str  # heading1, heading2, heading3, text, figure, table_caption, table, side_by_side
    text: str = ""
    level: int = 0
    headers: list = field(default_factory=list)
    rows: list = field(default_factory=list)
    figure_num: int = 0
    paths: list = field(default_factory=list)  # для side_by_side


def parse_chapter(md_text: str) -> list[Block]:
    """Парсит markdown-текст главы в список блоков.

    Поддерживает:
      # H1 / ## H2 / ### H3
      Обычный текст → параграфы
      *Рисунок N — Описание* → figure caption
      *Таблица N — Описание* → table caption
      | col1 | col2 | ... → таблица
    """
    lines = md_text.split("\n")
    blocks = []
    i = 0

    while i < len(lines):
        line = lines[i].rstrip()

        # Пустая строка — пропускаем
        if not line.strip():
            i += 1
            continue

        # Заголовки
        h_match = re.match(r"^(#{1,3})\s+(.+)$", line)
        if h_match:
            level = len(h_match.group(1))
            text = h_match.group(2).strip()
            kind = f"heading{level}"
            blocks.append(Block(kind=kind, text=text, level=level))
            i += 1
            continue

        # Подпись рисунка: *Рисунок N — Описание*
        fig_match = re.match(r"^\*Рисунок (\d+)\s*[—–-]\s*(.+)\*$", line)
        if fig_match:
            num = int(fig_match.group(1))
            caption = f"Рисунок {num} — {fig_match.group(2)}"
            blocks.append(Block(kind="figure", text=caption, figure_num=num))
            i += 1
            continue

        # Подпись таблицы: *Таблица N — Описание*
        tab_match = re.match(r"^\*Таблица (\d+)\s*[—–-]\s*(.+)\*$", line)
        if tab_match:
            num = int(tab_match.group(1))
            caption = f"Таблица {num} — {tab_match.group(2)}"
            blocks.append(Block(kind="table_caption", text=caption))
            i += 1
            continue

        # Markdown-таблица
        if "|" in line and i + 1 < len(lines) and re.match(r"^\|[-\s|:]+\|$", lines[i + 1].strip()):
            headers = [c.strip() for c in line.strip().strip("|").split("|")]
            i += 2  # пропустить разделитель
            rows = []
            while i < len(lines) and "|" in lines[i]:
                row = [c.strip() for c in lines[i].strip().strip("|").split("|")]
                rows.append(row)
                i += 1
            blocks.append(Block(kind="table", headers=headers, rows=rows))
            continue

        # Обычный текст (параграф) — объединяем последовательные строки
        text_parts = [line.strip()]
        i += 1
        while i < len(lines):
            next_line = lines[i].rstrip()
            if not next_line.strip():
                break  # пустая строка → конец параграфа
            if re.match(r"^#{1,3}\s+", next_line):
                break
            if re.match(r"^\*Рисунок \d+\s*[—–-]", next_line):
                break
            if re.match(r"^\*Таблица \d+\s*[—–-]", next_line):
                break
            if "|" in next_line and i + 1 < len(lines) and re.match(r"^\|[-\s|:]+\|$", lines[i + 1].strip()):
                break
            text_parts.append(next_line.strip())
            i += 1
        blocks.append(Block(kind="text", text=" ".join(text_parts)))

    return blocks


def scan_images_dir(directory: str) -> dict[int, str | tuple]:
    """Автоматически сопоставить файлы из папки с номерами рисунков.

    Поддерживаемые форматы имён:
      fig-1.png, fig-2.jpg           → {1: "path/fig-1.png"}
      рисунок-1.png, рисунок-2.png   → {1: "path/рисунок-1.png"}
      1.png, 2.png, 23.png           → {1: "path/1.png"}
      fig-14a.png + fig-14b.png      → {14: ("path/fig-14a.png", "path/fig-14b.png")}
      01-login.png                   → сортируется по порядку, но не маппится автоматически

    Если файлы имеют числовой префикс/имя, маппинг автоматический.
    Файлы с суффиксами a/b объединяются в side-by-side пары.
    """
    if not os.path.isdir(directory):
        return {}

    # Собрать все изображения
    extensions = ("*.png", "*.jpg", "*.jpeg", "*.gif", "*.bmp", "*.webp")
    files = []
    for ext in extensions:
        files.extend(glob.glob(os.path.join(directory, ext)))
    files.sort()

    result = {}
    side_by_side = {}  # num -> [path_a, path_b]

    for filepath in files:
        basename = os.path.splitext(os.path.basename(filepath))[0]

        # Паттерн: fig-14a / fig-14b → side-by-side
        m_side = re.match(r"(?:fig|рисунок|рис)[_-]?(\d+)([ab])$", basename, re.IGNORECASE)
        if m_side:
            num = int(m_side.group(1))
            suffix = m_side.group(2).lower()
            if num not in side_by_side:
                side_by_side[num] = {}
            side_by_side[num][suffix] = filepath
            continue

        # Паттерн: fig-1, fig-23, рисунок-5, рис_12
        m_named = re.match(r"(?:fig|рисунок|рис)[_-]?(\d+)$", basename, re.IGNORECASE)
        if m_named:
            num = int(m_named.group(1))
            result[num] = filepath
            continue

        # Паттерн: просто число — 1.png, 02.png, 23.png
        m_num = re.match(r"^(\d+)$", basename)
        if m_num:
            num = int(m_num.group(1))
            result[num] = filepath
            continue

    # Собрать side-by-side пары
    for num, paths in side_by_side.items():
        if "a" in paths and "b" in paths:
            result[num] = (paths["a"], paths["b"])
        elif "a" in paths:
            result[num] = paths["a"]
        elif "b" in paths:
            result[num] = paths["b"]

    return result


def parse_figure_map(map_text: str) -> dict[int, str | tuple]:
    """Парсит маппинг рисунков из текста.

    Формат (каждая строка):
      N: path/to/image.png
      N: path1.png + path2.png  (side by side)

    Или markdown-таблица:
      | Рисунок N | path.png |
    """
    result = {}
    for line in map_text.strip().split("\n"):
        line = line.strip()
        if not line:
            continue

        # Формат "N: path" или "N: path1 + path2"
        m = re.match(r"(\d+)\s*:\s*(.+)", line)
        if m:
            num = int(m.group(1))
            paths = m.group(2).strip()
            if "+" in paths:
                parts = [p.strip() for p in paths.split("+")]
                result[num] = tuple(parts[:2])
            else:
                result[num] = paths
            continue

        # Формат markdown-таблицы "| Рисунок N | path |"
        m2 = re.match(r"\|\s*Рисунок (\d+).*?\|\s*(.+?)\s*\|", line)
        if m2:
            num = int(m2.group(1))
            path = m2.group(2).strip().strip("`")
            if "+" in path:
                parts = [p.strip() for p in path.split("+")]
                result[num] = tuple(parts[:2])
            else:
                result[num] = path

    return result
