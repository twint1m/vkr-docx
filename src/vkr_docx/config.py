"""Загрузка и мержинг конфигурации."""

import os
import yaml
from pathlib import Path


def _deep_merge(base: dict, override: dict) -> dict:
    """Рекурсивный мерж: override перезаписывает только указанные ключи."""
    result = base.copy()
    for key, value in override.items():
        if key in result and isinstance(result[key], dict) and isinstance(value, dict):
            result[key] = _deep_merge(result[key], value)
        else:
            result[key] = value
    return result


def _find_config_file() -> str | None:
    """Ищет config.yaml: CWD → ~/.config/vkr-docx/ → None."""
    candidates = [
        Path.cwd() / "config.yaml",
        Path.home() / ".config" / "vkr-docx" / "config.yaml",
    ]
    for path in candidates:
        if path.exists():
            return str(path)
    return None


def _validate_config(cfg: dict):
    """Проверяет корректность значений конфигурации."""
    page = cfg.get("page", {})
    for key in ("width_cm", "height_cm", "margin_left_cm", "margin_right_cm",
                "margin_top_cm", "margin_bottom_cm"):
        val = page.get(key)
        if val is not None and (not isinstance(val, (int, float)) or val <= 0):
            raise ValueError(f"page.{key} должно быть положительным числом, получено: {val!r}")

    font = cfg.get("font", {})
    font_size = font.get("size")
    if font_size is not None and (not isinstance(font_size, (int, float)) or font_size <= 0):
        raise ValueError(f"font.size должен быть > 0, получено: {font_size!r}")

    font_name = font.get("name")
    if font_name is not None and not isinstance(font_name, str):
        raise ValueError(f"font.name должен быть строкой, получено: {font_name!r}")


def load_config(path: str = None) -> dict:
    """Загрузить конфигурацию. Мержит пользовательскую поверх дефолтной.

    Порядок поиска:
      1. Явно переданный path
      2. config.yaml в текущей директории
      3. ~/.config/vkr-docx/config.yaml
      4. Только дефолтные значения
    """
    # Дефолтный конфиг из пакета
    default_path = Path(__file__).parent / "config.default.yaml"

    defaults = {}
    if default_path.exists():
        with open(default_path, "r", encoding="utf-8") as f:
            defaults = yaml.safe_load(f) or {}

    # Пользовательский конфиг
    if path is None:
        path = _find_config_file()

    user_config = {}
    if path and os.path.exists(path):
        with open(path, "r", encoding="utf-8") as f:
            user_config = yaml.safe_load(f) or {}

    config = _deep_merge(defaults, user_config)

    # Раскрыть ~ в working_dir
    if "working_dir" in config:
        config["working_dir"] = os.path.expanduser(config["working_dir"])

    _validate_config(config)

    return config
