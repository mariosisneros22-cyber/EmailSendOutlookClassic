import os, datetime
import pandas as pd

from core.outlook_folders import get_saved_outlook_folder, get_or_create_subfolder
from core.responder_sender import find_latest_in_conversation, reply_all_with_attachment, move_mail,safe_join_file

CONTROL_PATH = os.path.abspath(os.path.join(os.path.dirname(__file__),"..","data", "responder_control.xlsx"))


MAILITEM_CLASS=43

def fetch_last_messages_by_conversation():
    folder = get_saved_outlook_folder()
    if folder is None:
        raise RuntimeError("No hay carpeta seleccionada.")

    items = folder.Items
    # Ordenar (muy importante) para que el primero que veamos de cada conversación sea el más reciente
    items.Sort("[ReceivedTime]", True)

    seen_conv = set()
    results = []

    # iterar ya ordenado desc
    for item in items:
        try:
            if item.Class != MAILITEM_CLASS:
                continue
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

    return results

def ensure_control_file():
    os.makedirs(os.path.dirname(CONTROL_PATH), exist_ok=True)
    if not os.path.exists(CONTROL_PATH):
        df = pd.DataFrame(columns=[
            "conversation_id",
            "last_entry_id",
            "subject",
            "nombre_archivo",
            "estado",
            "fecha_envio",
        ])
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

        new_rows.append({
            "conversation_id": conv,
            "last_entry_id": msg["last_entry_id"],
            "subject": msg["subject"],
            "nombre_archivo": "",
            "estado": "Pendiente",
            "fecha_envio": "",
        })

    if new_rows:
        df_new = pd.DataFrame(new_rows)
        df_final = pd.concat([existing, df_new], ignore_index=True)
        df_final.to_excel(CONTROL_PATH, index=False)

    return len(new_rows)



def process_pending_responses(carpeta_archivos: str, html_body: str | None = None, delay_segundos: float = 1.0, only_first_n: int | None = None):
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
    folder_done = get_or_create_subfolder(folder,"Procesados")
    
    pendientes = df[df["estado"].astype(str).str.strip().str.lower().eq("pendiente")].copy()
    if only_first_n is not None:
        pendientes = pendientes.head(int(only_first_n))
        
    procesados = 0
    
    for idx in pendientes.index:
        conv_id = str (df.at[idx, "conversation_id"]).strip()
        nombre_archivo = str(df.at[idx, "nombre_archivo"]).strip()
        
        if not conv_id:
            df.at[idx, "estado"] = "Error"
            df.at[idx, "error"] = "nombre_archivo vacío"
            continue
        
        try:
            pdf_path = safe_join_file(carpeta_archivos, nombre_archivo)
            
            last_mail=find_latest_in_conversation(folder,conv_id)
            if last_mail is None:
                raise RuntimeError("No se encontró el último mail del hilo en la carpeta seleccionada.")
            reply_all_with_attachment(
                last_mail,
                pdf_path=pdf_path,
                html_body=html_body,
                delay_segundos=delay_segundos
            )
            
            #mover el mail "ultimo" a Procesados
            move_mail(last_mail, folder_done)
            
            df.at[idx, "estado"]="Enviado"
            df.at[idx, "fecha_envio"] =datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            df.at[idx, "error"] =""
            procesados += 1
            
        except Exception as e:
            df.at[idx, "estado"]="Error"
            df.at[idx,"error"] = str(e)
    
    df.to_excel(CONTROL_PATH, index=False)
    return procesados