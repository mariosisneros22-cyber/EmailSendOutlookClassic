from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
import base64
import mimetypes

from core.graph.client import GraphClient

def _to_b64(path: str) -> str:
    raw = Path(path).read_bytes()
    return base64.b64encode(raw).decode("ascii")

def _guess_content_type(path:str) ->str:
    ctype, _ = mimetypes.guess_type(path)
    return ctype or "application/octet-stream"

@dataclass
class GraphProvider:
    client: GraphClient
    
    #Auth/account
    def get_current_user(self) -> dict:
        return self.client.get_me()
    
    #Folders/messages
    def list_folders(self,top: int=100) -> list[dict]:
        return self.client.list_mail_folders(top=top)
    
    def list_messages(self, folder_id: str, top: int = 50) -> list[dict]:
        return self.client.list_messages(folder_id=folder_id, top=top)
    
    def find_latest_in_conversation(self, folder_id: str, conversation_id: str) -> dict | None:
        return self.client.find_latest_by_conversation(folder_id=folder_id,conversation_id=conversation_id)
    
    
    #Sende new mail
    def send_mail(
        self,
        *,
        to: str,
        subject: str,
        html_body: str,
        attachments_paths: list[str] | None = None,
        save_to_sent_items: bool = True,
    ) -> None:
        attachments = []
        for p in attachments_paths or []:
            path = Path(p)
            if not path.is_file():
                raise FileNotFoundError(f"No existe adjunto: {p}")
            
            attachments.append(
                {
                    "name":path.name,
                    "contentType": _guess_content_type(str(path)),
                    "contentBytes": _to_b64(str(path)),
                }
            )
        
        self.client.send_mail(
            to_recipients=[to],
            subject=subject,
            html_body=html_body,
            file_attachments=attachments,
            save_to_sent_items=save_to_sent_items,
        )
        
        #reply all + attachment
    def reply_all_with_attachment(
        self,
        *,
        message_id: str,
        attachment_path: str,
        html_body: str | None = None,
    ) -> str:
        path= Path(attachment_path)
        if not path.is_file():
            raise FileNotFoundError(f"No existe adjunto: {attachment_path}")
        
        draft = self.client.create_reply_all_draft(message_id=message_id)
        draft_id = str(draft.get("id", "")).strip()
        if not draft_id:
            raise RuntimeError("Graph no devolvió id del borrador de reply-all.")
        
        if html_body is not None:
            self.client.patch_message(message_id=draft_id, html_body=html_body)
            
        self.client.add_file_attachment_to_message(
            message_id=draft_id,
            name=path.name,
            content_bytes_b64=_to_b64(str(path)),
            content_type=_guess_content_type(str(path)),
        )
        
        self.client.send_draft(message_id=draft_id)
        return draft_id
    
    #move
    def move_message(self, *, message_id: str, destination_folder_id: str) -> dict:
        return self.client.move_message(
            message_id=message_id,
            destination_folder_id=destination_folder_id,
        )