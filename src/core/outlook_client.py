# core/outlook_client.py
import win32com.client as win32


def get_outlook_app():
    try:
        return win32.gencache.EnsureDispatch("Outlook.Application")
    except Exception:
        return win32.Dispatch("Outlook.Application")


def get_mapi_namespace(outlook_app):
    return outlook_app.GetNamespace("MAPI")
