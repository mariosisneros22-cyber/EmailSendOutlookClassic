import os
from datetime import datetime
import pandas as pd

from core.outlook.outlook_client import get_outlook_app
from core.outlook.outlook_folders import (
    get_or_create_subfolder,
    get_saved_outlook_folder,
    load_selected_folder_ids,
)
from core.responder.responder_sender import (
    get_item_by_entry_id,
    move_mail,
    reply_all_with_attachment,
    find_latest_in_conversation_by_anchor
)
from core.shared.file_utils import safe_join_file
from core.shared.naming_rules import build_filename_from_subject
from core.config.app_dirs import RESPONDER_CONTROL_PATH

CONTROL_PATH = str(RESPONDER_CONTROL_PATH)

MAILITEM_CLASS=43

def fetch_last_messages_by_conversation(logger=None):
    folder = get_saved_outlook_folder()
    if folder is None:
        raise RuntimeError("No hay carpeta seleccionada.")
    
    if callable(logger):
        try:
            logger(f"Carpeta seleccionada: {getattr(folder, 'Name','(sin nombre)')}")
        except Exception:
            logger("Carpeta seleccionada: (no se pudo leer Name)")
            

    items = folder.Items
    items.Sort("[ReceivedTime]", True)

    total_items = items.Count  # ← Outlook ya sabe cuántos hay
    mail_items = 0
    seen_conv = set()
    results = []
    
    # iterar ya ordenado desc
    for item in items:
        try:
            if item.Class != MAILITEM_CLASS:
                continue
            
            mail_items += 1
            conv = str(item.ConversationID or "").strip()
            
            if not conv or conv in seen_conv:
                continue

            seen_conv.add(conv)
            results.append({
                "conversation_id": conv,
                "last_entry_id": item.EntryID,
                "subject": item.Subject,
                "received_time": item.ReceivedTime,
            })
        except Exception:
            continue
    if callable(logger):
        logger(f"Items={total_items} | MailItem={mail_items} | Conversaciones únicas={len(results)}")
    return results

def ensure_control_file():
    os.makedirs(os.path.dirname(CONTROL_PATH), exist_ok=True)
    if not os.path.exists(CONTROL_PATH):
        df = pd.DataFrame(columns=[
            "conversation_id",
            "last_entry_id",
            "subject",
            "nombre_sugerido",
            "nombre_archivo",
            "estado",
            "fecha_envio",
        ])
        df.to_excel(CONTROL_PATH, index=False)
        return
    
    #Migracion suave
    df = pd.read_excel(CONTROL_PATH).fillna("")
    changed = False

    for col in ["nombre_sugerido", "nombre_archivo", "estado", "fecha_envio", "error"]:
        if col not in df.columns:
            df[col] = ""
            changed = True

    if changed:
        df.to_excel(CONTROL_PATH, index=False)
        
        
def update_control_from_outlook():
    ensure_control_file()
    existing = pd.read_excel(CONTROL_PATH)

    if "conversation_id" not in existing.columns:
        raise RuntimeError("El control no tiene columna conversation_id. Recréalo o migra el archivo.")

    existing_conv = set(existing["conversation_id"].astype(str))

    messages = fetch_last_messages_by_conversation()

    new_rows = []

    
    for msg in messages:
        conv = str(msg["conversation_id"])
        if conv in existing_conv:
            continue
        
        suggest = build_filename_from_subject(msg["subject"], year_mode="current") or ""
    
        new_rows.append({
            "conversation_id": conv,
            "last_entry_id": msg["last_entry_id"],
            "subject": msg["subject"],
            "nombre_sugerido": suggest,
            "nombre_archivo": "",
            "estado": "Pendiente",
            "fecha_envio": "",
        })

    if new_rows:
        df_new = pd.DataFrame(new_rows)
        df_final = pd.concat([existing, df_new], ignore_index=True)
        df_final.to_excel(CONTROL_PATH, index=False)

    return len(new_rows)

def _same_folder(mail_item, folder) -> bool:
    try:
        return str(mail_item.Parent.EntryID) == str(folder.EntryID)
    except Exception:
        return False

def process_pending_responses(carpeta_archivos: str, html_body: str | None = None, delay_segundos: float = 1.0, only_first_n: int | None = None, logger=None):
    """
    Lee CONTROL_PATH y procesa filas con estado Pendiente y nombre_archivo no vacío.
    """
    ensure_control_file()
    df =pd.read_excel(CONTROL_PATH).fillna("")
    
    # normaliza columnas esperadas
    if "error" not in df.columns:
        df["error"] = ""
        
    folder = get_saved_outlook_folder()
    if folder is None:
        raise RuntimeError("No hay carpeta seleccionada.")

    outlook = get_outlook_app()
    if outlook is None:
        raise RuntimeError("No se pudo acceder a Outlook.")
    ns = outlook.GetNamespace("MAPI")
    store_id, _folder_entry_id = load_selected_folder_ids()
    
    folder_done = get_or_create_subfolder(folder,"Procesados")
    
    pendientes = df[df["estado"].astype(str).str.strip().str.lower().eq("pendiente")].copy()
    if only_first_n is not None:
        pendientes = pendientes.head(int(only_first_n))
        
    if callable(logger):
        logger(f"Pendientes detectados: {len(pendientes)}")
        
    procesados = 0
    
    for idx in pendientes.index:
        conv_id = str (df.at[idx, "conversation_id"]).strip()
        nombre_archivo = str(df.at[idx, "nombre_archivo"]).strip()
        
        if callable(logger):
            logger(f"Procesando conv={conv_id[:8]}... archivo= '{nombre_archivo}'")
            
        if not conv_id:
            df.at[idx, "estado"] = "Error"
            df.at[idx, "error"] = "conversation_id vacío"
            continue

        if not nombre_archivo:
            df.at[idx, "estado"] = "Error"
            df.at[idx, "error"] = "nombre_archivo vacío"
            continue
        
        sent_ok = False
        try:
            pdf_path = safe_join_file(carpeta_archivos, nombre_archivo)

            anchor_entry_id = str(df.at[idx, "last_entry_id"]).strip()
            if not anchor_entry_id:
                raise RuntimeError("last_entry_id vacío en el control.xlsx")
            
            anchor_mail=get_item_by_entry_id(ns, anchor_entry_id, store_id)            
            if anchor_mail is None:
                raise RuntimeError("No se pudo obtener el correo ancla desde EntryID (puede haber sido movido/eliminado).")
            
            
            #last_mail = anchor_mail
            last_mail=find_latest_in_conversation_by_anchor(ns, anchor_mail)
            
            if last_mail is None:
                raise RuntimeError("No se encontró el último mail del hilo en la carpeta seleccionada.")
            
            reply_all_with_attachment(
                last_mail,
                pdf_path=pdf_path,
                html_body=html_body,
                delay_segundos=delay_segundos
            )
            
            sent_ok = True
            df.at[idx, "estado"]="Enviado"
            df.at[idx, "fecha_envio"] =datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            df.at[idx, "error"] =""
            procesados += 1
            
           
            #mover el mail "ultimo" a Procesados
            try:
                if _same_folder(last_mail,folder):
                    move_mail(last_mail, folder_done)
                else:
                    if callable(logger):
                        try:
                            logger(f"Skip move: está en otra carpeta ({getattr(last_mail.Parent, 'Name','?')})")
                        except Exception:
                            logger("Skip move: está en otra carpeta")
            except Exception as e:
                if callable (logger):
                    logger(f"WARN move: {e}")
                df.at[idx,"error"] = f"WARN move: {e}"

        except Exception as e:
            if sent_ok:
                df.at[idx, "error"] = f"WARN post-send: {e}"
            else:
                df.at[idx, "estado"] = "Error"
                df.at[idx, "error"] = str(e)
    
    df.to_excel(CONTROL_PATH, index=False)
    
    if callable(logger):
        logger(f"Procesados OK: {procesados}")
        
    return procesados

def fill_suggested_names(year_mode: str = "current", only_if_empty: bool = True) -> int:
    ensure_control_file()
    df = pd.read_excel(CONTROL_PATH).fillna("")
    
    if "subject" not in df.columns:
        return 0
    
    if "nombre_sugerido" not in df.columns:
        df["nombre_sugerido"] = ""
    
    changed = 0
    for i in df.index:
        if only_if_empty and str(df.at[i, "nombre_sugerido"]).strip():
            continue
        
        subject= str(df.at[i, "subject"]).strip()
        suggest = build_filename_from_subject(subject, year_mode=year_mode)
        if suggest:
            df.at[i, "nombre_sugerido"]= suggest
            changed +=1
    
    if changed:
        df.to_excel(CONTROL_PATH, index = False)
    return changed

def apply_suggested_to_nombre_archivo(only_if_empty: bool = True, add_pdf_ext:bool = False) -> int:
    ensure_control_file()
    df= pd.read_excel(CONTROL_PATH).fillna("")
    
    for col in ["nombre_sugerido", "nombre_archivo"]:
        if col not in df.columns:
            df[col]= ""
            
    changed = 0
    for i in df.index:
        sugerido = str(df.at[i, "nombre_sugerido"]).strip()
        if not sugerido:
            continue
            
        actual = str(df.at[i, "nombre_archivo"]).strip()
        if only_if_empty and actual:
            continue
        
        v = sugerido
        if add_pdf_ext and not v.lower().endswith(".pdf"):
            v += ".pdf"
        
        if v != actual:
            df.at[i, "nombre_archivo"] = v
            changed += 1
        
    df.to_excel(CONTROL_PATH, index=False)
    return changed    
