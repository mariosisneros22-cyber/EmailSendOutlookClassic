# core/outlook_client.py
import win32com.client as win32
import os, shutil, gc
import threading, pythoncom

_thread_local = threading.local()

def _ensure_com_initialized():
    if getattr(_thread_local, "com_initialized", False):
        return
    pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)
    _thread_local.com_initialized = True


def _clear_gen_py_dirs():
    candidates = [
        os.path.join(os.environ.get("TEMP", ""), "gen_py"),
        os.path.join(os.environ.get("LOCALAPPDATA", ""), "Temp", "gen_py"),
        os.path.join(os.environ.get("TEMP", ""), "pywin32_gen_py"),
        os.path.join(os.environ.get("LOCALAPPDATA", ""), "Temp", "pywin32_gen_py"),
    ]
    
    for d in candidates:
        if d and os.path.isdir(d):
            shutil.rmtree(d, ignore_errors=True)

def _dispatch_outlook():
    return win32.Dispatch("Outlook.Application")


def _ensure_dispatch_outlook():
    return win32.gencache.EnsureDispatch("Outlook.Application")


def _repair_gencache() -> None:
    _clear_gen_py_dirs()
    win32.gencache.is_readonly = False

    try:
        win32.gencache.Rebuild()
    except Exception:
        pass

    gc.collect()


def get_outlook_app(rebuild_on_error: bool = True):
    _ensure_com_initialized()
    try:
        return _dispatch_outlook()
    except Exception:
        pass

    try:
        return _ensure_dispatch_outlook()
    except Exception:
        if not rebuild_on_error:
            raise

    _repair_gencache()

    try:
        return _ensure_dispatch_outlook()
    except Exception:
        return _dispatch_outlook()


def get_mapi_namespace(outlook_app):
    _ensure_com_initialized()
    return outlook_app.GetNamespace("MAPI")