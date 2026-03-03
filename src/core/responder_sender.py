# core/responder_sender.py
import os 
import time
from datetime import datetime

from core.common import _asegurar_leible, _copiar_a_temp_corto, safe_join_file
from core.outlook_client import get_outlook_app

MAILITEM_CLASS=43

def _iter_mailitems(folder):
    # COM collection: iterar directo a veces falla si cambia mientras iteras
    # Mejor: snapshot por índice
    items = folder.Items
    n = items.Count
    for i in range(1, n + 1):
        try:
            it = items.Item(i)
            if getattr(it, "Class", None) == MAILITEM_CLASS:
                yield it
        except Exception:
            continue
        
def find_latest_in_conversation(folder, conversation_id:str):
    latest=None
    latest_time=None
    
    for it in _iter_mailitems(folder):
        try:
            if getattr(it, "ConversationID", None) != conversation_id:
                continue
            rt = getattr(it, "ReceivedItem", None)
            #fallback si no hay ReceivedTime
            if rt is None:
                continue
            if (latest_time is None) or (rt > latest_time):
                latest = it
                latest_time = rt
        except Exception:
            continue
        
    return latest

def reply_all_with_attachment(mail_item, pdf_path: str, html_body: str | None = None, delay_segundos: float = 1.0,):
    
    """
    Responde con ReplyAll al mail_item (idealmente el último del hilo), adjunta PDF, envía.
    """
    _asegurar_leible(pdf_path)
    
    #ReplyAll
    reply=mail_item.ReplyAll()
    
    #cuerpo (opcional): si no quieres tocar, comenta este bloque
    if html_body is not None:
        reply.HTMLBody =html_body + "<br><br>" + (reply.HTMLBody or "")
        
    #adjunto (copia corta)
    tmp = _copiar_a_temp_corto(pdf_path)
    reply.Attachments.Add(tmp)
    
    reply.Send()
    time.sleep(delay_segundos)
    
    #limpiar temp
    try:
        os.remove(tmp)
    except Exception:
        pass
    
def move_mail(mail_item, target_folder):
    """
    Mueve el mail_item a target_folder.
    Importante: después de Move(), el objeto original deja de ser válido en algunas situaciones.
    """
    mail_item.Move(target_folder)
    