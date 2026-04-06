"""vkr-docx — инструмент оформления ВКР по ГОСТ 7.32-2017."""

from .document import VKRDocument
from .config import load_config

__version__ = "0.1.0"
__all__ = ["VKRDocument", "load_config"]
