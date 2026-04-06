"""vkr-docx — инструмент оформления ВКР по ГОСТ 7.32-2017."""

from importlib.metadata import version, PackageNotFoundError

from .document import VKRDocument, VKRDocumentError
from .config import load_config

try:
    __version__ = version("vkr-docx")
except PackageNotFoundError:
    __version__ = "0.1.0"

__all__ = ["VKRDocument", "VKRDocumentError", "load_config"]
