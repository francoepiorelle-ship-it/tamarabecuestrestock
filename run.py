import os
import sys
from streamlit.web.cli import main

if __name__ == "__main__":
    # Si estamos corriendo dentro del .exe empaquetado
    if getattr(sys, 'frozen', False):
        base_path = sys._MEIPASS
    else:
        base_path = os.path.dirname(__file__)

    # Ruta al archivo principal de tu app
    target_script = os.path.join(base_path, "main.py")
    
    # Iniciar Streamlit apuntando al script interno
    sys.argv = ["streamlit", "run", target_script, "--global.developmentMode=false"]
    sys.exit(main())