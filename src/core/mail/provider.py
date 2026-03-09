from __future__ import annotations

from typing import Protocol, runtime_checkable, Any

@runtime_checkable
class MailProvider(Protocol):
    #Auth/account
    def get_current_user(self) ->dict[str, Any]:
        ...
    
    #Folders/messages
    def list_folders(self, top: int = 50) -> list[dict[str,Any]]:
        ...
    
    def list_messages(self,folder_id: str, top: int = 50) -> list[dict[str,Any]]:
        ...
    
    def find_latest_in_conversation(self,folder_id: str, conversation_id:  str) -> dict[str, Any] | None:
        ...
    
    #Send new mail
    def send_mail(
        self,
        *,
        to: str,
        subject: str,
        html_body: str,
        attachments_paths: list[str] | None = None,
        save_to_sent_items: bool = True,
    ) -> None:
        ...
        
    # reply all + attachments
    def reply_all_with_attachment(
        self,
        *,
        message_id: str,
        attachment_path: str,
        html_body: str | None = None,
    ) ->str:
        ...
        
    #move
    def move_message(self, *, message_id: str, destination_folder_id: str) -> dict[str, Any]:
        ...