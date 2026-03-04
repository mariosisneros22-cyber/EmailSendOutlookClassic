# core/responder_sender.py
import os 
import time

from core.common import _asegurar_leible, _copiar_a_temp_corto

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
            rt = getattr(it, "ReceivedTime", None)
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


def get_item_by_entry_id(ns, entry_id: str, store_id: str | None = None):
    if store_id:
        return ns.GetItemFromID(entry_id, store_id)
    return ns.GetItemFromID(entry_id)

def iter_conversation_items(conversation):
    stack = []
    roots = conversation.GetRootItems()
    for r in roots:
        stack.append(r)
        
    while stack:
        it = stack.pop()
        yield it
        try:
            children = conversation.GetChildren(it)
            for ch in children:
                stack.append(ch)
        except Exception:
            pass
        
def find_latest_in_conversation_by_anchor(ns, anchor_mail, only_mailitems=True):
    conv = anchor_mail.GetConversation()
    if conv is None:
        return None
    
    latest = None
    latest_dt = None
    
    for it in iter_conversation_items(conv):
        try:
            if only_mailitems and getattr(it, "Class", None) != MAILITEM_CLASS:
                continue
            dt= getattr(it, "ReceivedTime", None)
            if dt is None: 
                continue 
            if latest_dt is None or dt > latest_dt:
                latest = it
                latest_dt = dt
        except Exception:
            continue
    return latest


def move_mail(mail_item, target_folder):
    """
    Mueve el mail_item a target_folder.
    Importante: después de Move(), el objeto original deja de ser válido en algunas situaciones.
    """
    mail_item.Move(target_folder)

