#!/usr/bin/env python
import sys
from pathlib import Path
import importlib.util
import importlib.machinery

# Obtener directorio actual
current_dir = Path(__file__).resolve().parent

# Registrar el paquete bot_financiero_telegram dinámicamente si no está en sys.modules.
# Esto es crítico para despliegues en la nube (como Render) donde la carpeta clonada
# puede tener un nombre con guiones (ej. "telegram-bot") y no se puede importar
# directamente como un paquete de Python, lo que causaría fallas por importaciones relativas.
if "bot_financiero_telegram" not in sys.modules:
    spec = importlib.machinery.ModuleSpec("bot_financiero_telegram", None, is_package=True)
    pkg = importlib.util.module_from_spec(spec)
    pkg.__path__ = [str(current_dir)]
    pkg.__file__ = str(current_dir / "__init__.py")
    sys.modules["bot_financiero_telegram"] = pkg

from bot_financiero_telegram.bot import main as main_func

if __name__ == "__main__":
    main_func()
