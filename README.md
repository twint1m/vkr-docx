# vkr-docx

Инструмент для оформления выпускной квалификационной работы (ВКР) в формате `.docx` по ГОСТ 7.32-2017.

Автоматизирует: стили, поля, заголовки, рисунки, таблицы, перекрёстные ссылки, содержание, версионирование.

## Установка

### Одной командой (рекомендуется)

```bash
curl -fsSL https://raw.githubusercontent.com/twint1m/vkr-docx/main/install.sh | bash
```

Скрипт автоматически: клонирует репо, установит пакет, скопирует скилл `/vkr` и настроит Claude Code.

### Вручную

```bash
git clone https://github.com/twint1m/vkr-docx.git ~/projects/vkr-docx
cd ~/projects/vkr-docx
pip install -e .
cp -r .claude/commands/ ~/.claude/commands/
cp CLAUDE.md ~/CLAUDE.md
```

## Быстрый старт

### CLI

```bash
# Полное исправление: стили + ссылки + содержание
vkr-docx full-fix "ВКР 0.2.7.docx"

# Только стили
vkr-docx fix-styles "ВКР 0.2.7.docx"

# Только перекрёстные ссылки
vkr-docx crossrefs "ВКР 0.2.7.docx"

# Только содержание
vkr-docx toc "ВКР 0.2.7.docx"

# Перезаписать файл (без создания новой версии)
vkr-docx full-fix --overwrite "ВКР 0.2.7.docx"

# Автопоиск последнего файла ВКР в текущей директории
vkr-docx full-fix
```

### Python API

```python
from vkr_docx import VKRDocument

vkr = VKRDocument("ВКР 0.2.7.docx")

# Настройка
vkr.fix_page_setup()       # поля по ГОСТ
vkr.fix_styles()           # стили заголовков
vkr.fix_image_sizes()      # ограничить высоту картинок

# Добавление контента
vkr.add_heading1("2 ПРОЕКТИРОВАНИЕ ВЕБ-ПРИЛОЖЕНИЯ")
vkr.add_heading2("2.1 Алгоритм работы")
vkr.add_body_text("Текст параграфа...")
vkr.add_figure("diagram.png", "Рисунок 1 — Архитектура системы")
vkr.add_table_caption("Таблица 1 — Описание полей")
vkr.add_data_table(
    ["Поле", "Тип", "Описание"],
    [["id", "UUID", "Первичный ключ"], ["name", "VARCHAR", "Имя"]]
)

# Ссылки и навигация
vkr.add_crossrefs()        # двусторонние ссылки
vkr.update_toc()           # обновить содержание

# Сохранение
vkr.save()                 # → ВКР 0.2.8.docx (автоинкремент)
```

### Claude Code

Скопируйте файлы для интеграции с Claude Code:

```bash
# Скилл (вызывается через /vkr)
cp -r ~/projects/vkr-docx/.claude/commands/ ~/.claude/commands/

# Триггер (автоопределение задач ВКР)
cp ~/projects/vkr-docx/CLAUDE.md ~/CLAUDE.md
```

После этого Claude Code будет автоматически использовать `vkr-docx` при работе с ВКР.

## Конфигурация

Скопируйте `config.default.yaml` как `config.yaml` в рабочую директорию и измените нужные параметры:

```yaml
page:
  margin_left_cm: 3.0      # поле слева (для подшивки)
  margin_right_cm: 1.5
  margin_top_cm: 2.0
  margin_bottom_cm: 2.0

font:
  name: "Times New Roman"
  size: 14                  # основной текст
  size_table: 12            # текст в таблицах

image:
  max_height_cm: 23.0       # ограничение высоты

# Рабочая директория (где лежат файлы ВКР)
working_dir: "~/Downloads"
```

Поиск конфига: `./config.yaml` → `~/.config/vkr-docx/config.yaml` → дефолтные значения.

Можно указать явно: `vkr-docx --config /path/to/config.yaml full-fix`

## Версионирование

| Версия | Значение |
|---|---|
| `ВКР 0.A.B` | В работе. A = номер главы, B = правки (0..999) |
| `ВКР 1.0.0` | ВКР полностью готова |
| `ВКР 1.0.N` | Правки после готовности |

```python
vkr.set_chapter(3)  # переключить на главу 3 → ВКР 0.3.0
vkr.mark_ready()    # пометить готовой → ВКР 1.0.0
vkr.save()          # автоинкремент → ВКР X.Y.Z+1
```

## Правила оформления (ГОСТ 7.32-2017)

- **Страница**: A4, поля лево=3, право=1.5, верх/низ=2 см
- **Текст**: Times New Roman 14pt, выравнивание по ширине, отступ 1.25 см, межстрочный 1.5
- **Heading 1** (главы): Bold, ПРОПИСНЫЕ, по центру, с новой страницы
- **Heading 2/3** (подразделы): Bold, по ширине, отступ 1.25
- **Рисунки**: по центру, подпись снизу «Рисунок X — ...», max высота 23 см
- **Таблицы**: подпись сверху «Таблица X — ...», границы, 12pt
- **Перекрёстные ссылки**: двусторонние кликабельные (текст ↔ подпись)
- **Содержание**: обновлять после каждой главы, leader dots

## Лицензия

MIT
