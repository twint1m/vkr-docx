"""CLI-интерфейс для vkr-docx."""

import glob
import os
import click

from .config import load_config
from .document import VKRDocument


def _find_latest_vkr(directory: str = ".") -> str | None:
    """Найти файл ВКР с наибольшей версией в директории."""
    pattern = os.path.join(directory, "ВКР *.docx")
    files = glob.glob(pattern)
    if not files:
        return None
    # Сортировка по имени (версия в имени → лексикографический порядок)
    files.sort()
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


if __name__ == "__main__":
    main()
