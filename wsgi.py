#!/usr/bin/env python
import sys
from pathlib import Path

# Agregar el directorio padre al sys.path para poder importar este directorio como un paquete
current_dir = Path(__file__).resolve().parent
parent_dir = current_dir.parent
if str(parent_dir) not in sys.path:
    sys.path.insert(0, str(parent_dir))

# Importar dinámicamente la función principal usando el nombre de la carpeta como paquete
package_name = current_dir.name
try:
    bot_module = __import__(f"{package_name}.bot", fromlist=["main"])
    main_func = bot_module.main
except ImportError as e:
    # Si por alguna razón no se puede importar como paquete, intentamos importar directamente
    sys.path.insert(0, str(current_dir))
    import bot
    main_func = bot.main

if __name__ == "__main__":
    main_func()
