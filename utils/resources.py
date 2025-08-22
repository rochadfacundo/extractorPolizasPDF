# utils/resources.py
import os, sys

def asset_path(*relative_parts: str) -> str:
    """
    Devuelve la ruta absoluta a un asset empaquetado.
    - En exe (PyInstaller): carpeta donde está el .exe
    - En código fuente: raíz del repo (../ desde utils/)
    Uso: asset_path("assets", "marcas.json")
    """
    if getattr(sys, "frozen", False):  # ejecutable PyInstaller (one-folder u one-file)
        base_dir = os.path.dirname(sys.executable)
    else:
        # estás en .../utils/resources.py -> subo a raíz del proyecto
        base_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
    return os.path.join(base_dir, *relative_parts)

def documents_dir() -> str:
    """Devuelve la carpeta Documentos real (incluye OneDrive si aplica)."""
    if sys.platform == "win32":
        try:
            import ctypes
            from ctypes import wintypes
            CSIDL_PERSONAL = 5  # Mis Documentos
            SHGFP_TYPE_CURRENT = 0
            buf = ctypes.create_unicode_buffer(wintypes.MAX_PATH)
            if ctypes.windll.shell32.SHGetFolderPathW(
                None, CSIDL_PERSONAL, None, SHGFP_TYPE_CURRENT, buf
            ) == 0 and buf.value:
                return buf.value
        except Exception:
            pass
    # Fallback multiplataforma
    return os.path.join(os.path.expanduser("~"), "Documents")

def excel_output_path(filename: str) -> str:
    """Crea (si no existe) y devuelve Documentos/polizasExtraidas/filename."""
    base = os.path.join(documents_dir(), "polizasExtraidas")
    os.makedirs(base, exist_ok=True)
    return os.path.join(base, filename)
