"""CLI-интерфейс для vkr-docx."""

import glob
import os
import re
import click

from .config import load_config
from .document import VKRDocument, VKRDocumentError


def _parse_version_tuple(filepath: str) -> tuple:
    """Извлечь числовую версию из имени файла ВКР."""
    basename = os.path.basename(filepath)
    m = re.search(r"ВКР (\d+)\.(\d+)\.?(\d*)", basename)
    if m:
        major = int(m.group(1))
        minor = int(m.group(2))
        patch = int(m.group(3)) if m.group(3) else 0
        return (major, minor, patch)
    return (0, 0, 0)


def _find_latest_vkr(directory: str = ".") -> str | None:
    """Найти файл ВКР с наибольшей версией в директории."""
    pattern = os.path.join(directory, "ВКР *.docx")
    files = glob.glob(pattern)
    if not files:
        return None
    # Сортировка по числовой версии (major, minor, patch)
    files.sort(key=_parse_version_tuple)
    return files[-1]


def _resolve_file(file: str, config: dict) -> str:
    """Найти файл: явный путь → CWD → working_dir."""
    if file:
        if os.path.exists(file):
            return file
        working = config.get("working_dir", ".")
        candidate = os.path.join(working, file)
        if os.path.exists(candidate):
            return candidate
        raise click.ClickException(f"Файл не найден: {file}")

    # Автопоиск
    for directory in [".", config.get("working_dir", "")]:
        if not directory:
            continue
        found = _find_latest_vkr(directory)
        if found:
            return found

    raise click.ClickException(
        "Файл ВКР не найден. Укажите путь явно или настройте working_dir в config.yaml"
    )


@click.group()
@click.option("--config", "config_path", default=None, help="Путь к config.yaml")
@click.pass_context
def main(ctx, config_path):
    """vkr-docx — оформление ВКР по ГОСТ 7.32-2017."""
    ctx.ensure_object(dict)
    ctx.obj["config"] = load_config(config_path)


@main.command("fix-styles")
@click.argument("file", required=False)
@click.option("--overwrite", is_flag=True, help="Перезаписать файл")
@click.pass_context
def fix_styles(ctx, file, overwrite):
    """Исправить стили, поля и размеры изображений."""
    cfg = ctx.obj["config"]
    path = _resolve_file(file, cfg)
    vkr = VKRDocument(path, config=cfg)

    vkr.fix_page_setup()
    vkr.fix_styles()
    n = vkr.fix_image_sizes()

    out = vkr.save_same() if overwrite else vkr.save()
    click.echo(f"Стили исправлены. Изображений масштабировано: {n}")
    click.echo(f"Сохранено: {out}")


@main.command("crossrefs")
@click.argument("file", required=False)
@click.option("--overwrite", is_flag=True)
@click.pass_context
def crossrefs(ctx, file, overwrite):
    """Создать перекрёстные ссылки для рисунков и таблиц."""
    cfg = ctx.obj["config"]
    path = _resolve_file(file, cfg)
    vkr = VKRDocument(path, config=cfg)

    n = vkr.add_crossrefs()

    out = vkr.save_same() if overwrite else vkr.save()
    click.echo(f"Создано {n} пар перекрёстных ссылок")
    click.echo(f"Сохранено: {out}")


@main.command("toc")
@click.argument("file", required=False)
@click.option("--overwrite", is_flag=True)
@click.pass_context
def toc(ctx, file, overwrite):
    """Обновить содержание."""
    cfg = ctx.obj["config"]
    path = _resolve_file(file, cfg)
    vkr = VKRDocument(path, config=cfg)

    n = vkr.update_toc()

    out = vkr.save_same() if overwrite else vkr.save()
    click.echo(f"Содержание обновлено: {n} записей")
    click.echo(f"Сохранено: {out}")


@main.command("full-fix")
@click.argument("file", required=False)
@click.option("--overwrite", is_flag=True)
@click.pass_context
def full_fix(ctx, file, overwrite):
    """Полное исправление: стили + ссылки + содержание."""
    cfg = ctx.obj["config"]
    path = _resolve_file(file, cfg)
    vkr = VKRDocument(path, config=cfg)

    vkr.fix_page_setup()
    vkr.fix_styles()
    imgs = vkr.fix_image_sizes()
    refs = vkr.add_crossrefs()
    toc_n = vkr.update_toc()

    out = vkr.save_same() if overwrite else vkr.save()
    click.echo(f"Стили ✓ | Изображения: {imgs} | Ссылки: {refs} | Содержание: {toc_n}")
    click.echo(f"Сохранено: {out}")


@main.command("add-chapter")
@click.argument("chapter_md")
@click.argument("file", required=False)
@click.option("--images", "images_dir", default=None, help="Папка с изображениями")
@click.option("--overwrite", is_flag=True)
@click.pass_context
def add_chapter(ctx, chapter_md, file, images_dir, overwrite):
    """Добавить главу из markdown-файла.

    Изображения сопоставляются автоматически по именам:
    fig-1.png, рис-2.png, 3.png, fig-14a.png+fig-14b.png
    """
    cfg = ctx.obj["config"]
    path = _resolve_file(file, cfg)
    vkr = VKRDocument(path, config=cfg)

    stats = vkr.add_chapter(chapter_md, images_dir=images_dir)

    out = vkr.save_same() if overwrite else vkr.save()
    click.echo(
        f"Глава добавлена: {stats.get('headings', 0)} заголовков, "
        f"{stats.get('paragraphs', 0)} параграфов, "
        f"{stats.get('figures', 0)} рисунков, "
        f"{stats.get('tables', 0)} таблиц"
    )
    if "crossrefs" in stats:
        click.echo(f"Ссылки: {stats['crossrefs']} | Содержание: {stats.get('toc_entries', 0)}")
    click.echo(f"Сохранено: {out}")


if __name__ == "__main__":
    main()
