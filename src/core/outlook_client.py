# core/outlook_client.py
import win32com.client as win32
import os, shutil, gc

def _clear_gen_py_dirs():
    candidates=[
        os.path.join(os.environ.get("TEMP", ""), "gen_py"),
        os.path.join(os.environ.get("LOCALAPPDATA", ""), "Temp", "gen_py"),
        os.path.join(os.environ.get("TEMP", ""), "pywin32_gen_py"),
        os.path.join(os.environ.get("LOCALAPPDATA", ""), "Temp", "pywin32_gen_py"),
    ]
    
    for d in candidates:
        if d and os.path.isdir(d):
            shutil.rmtree(d, ignore_errors=True)

def get_outlook_app(rebuild_on_error: bool = True):
    # 1) Intento estable (late-binding)
    try:
        return win32.Dispatch("Outlook.Application")
    except Exception:
        pass

    # 2) Intento con EnsureDispatch
    try:
        return win32.gencache.EnsureDispatch("Outlook.Application")
    except Exception as e:
        if not rebuild_on_error:
            raise

        # 3) Repair en caliente del gen_py
        try:
            _clear_gen_py_dirs()
            try:
                win32.gencache.is_readonly = False
            except Exception:
                pass
            try:
                win32.gencache.Rebuild()
            except Exception:
                # si Rebuild falla, igual probamos Dispatch
                pass

            gc.collect()
            # probamos otra vez
            try:
                return win32.gencache.EnsureDispatch("Outlook.Application")
            except Exception:
                return win32.Dispatch("Outlook.Application")
        except Exception:
            # último recurso: re-lanzar el error original
            raise e

def get_mapi_namespace(outlook_app):
    return outlook_app.GetNamespace("MAPI")
 