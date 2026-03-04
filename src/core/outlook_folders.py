# core/outlook_folders.py
from core.outlook_client import get_outlook_app
from core.config_store import load_config, save_config



    
def save_selected_folder(store_id: str, entry_id: str):
    data = load_config()
    data["queue_folder_store_id"] = store_id
    data["queue_folder_entry_id"] = entry_id
    save_config(data)

def load_selected_folder_ids():
    data = load_config()
    return data.get("queue_folder_store_id"), data.get("queue_folder_entry_id")

def pick_outlook_folder_and_save():
    outlook = get_outlook_app()
    if outlook is None:
        raise RuntimeError("No se pudo acceder a Outlook.")
    ns = outlook.GetNamespace("MAPI")
    folder = ns.PickFolder()
    if folder is None:
        return None
    save_selected_folder(folder.StoreID, folder.EntryID)
    return folder

def get_saved_outlook_folder():
    outlook = get_outlook_app()
    if outlook is None:
        return None
    ns = outlook.GetNamespace("MAPI")
    store_id, entry_id = load_selected_folder_ids()
    if not store_id or not entry_id:
        return None
    return ns.GetFolderFromID(entry_id, store_id)

# EXTRA
def get_or_create_subfolder(parent_folder, name:str):
    #Outlook Folders collection: parent_folder.Folders
    for f in parent_folder.Folders:
        try:
            if f.Name == name:
                return f
        except Exception: 
            continue
    return parent_folder.Folders.Add(name)

