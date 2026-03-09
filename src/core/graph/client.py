# src/core/graph/client.py
from __future__ import annotations

from dataclasses import dataclass
from typing import Any
import time
import requests

from core.graph.auth import get_access_token, GraphAuthError

GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0"

class GraphClientError(RuntimeError):
    pass

@dataclass
class GraphClient:
    timeout_sec: float = 30.0
    max_retries: int = 3
    backoff_base_sec: float = 1.0
    
    def _headers(self) -> dict[str, str]:
        token = get_access_token()
        return {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json",
        }
        
    def _request(
        self,
        method: str,
        path: str,
        *,
        params: dict[str, Any]  | None = None,
        json_body: dict[str, Any] | None = None,
        expected_status: set[int] | None = None,
    ) -> dict[str, Any] | None:
        if not path.startswith("/"):
            path = "/" + path
            
        url = f"{GRAPH_BASE_URL}{path}"
        expected = expected_status or {200,201,202,204}
        
        last_err: Exception | None = None
        
        for attempt in range(self.max_retries + 1):
            resp = requests.request(
                method=method.upper(),
                url=url,
                headers=self._headers(),
                params=params,
                json=json_body,
                timeout=self.timeout_sec,
            )
            
            if resp.status_code in expected:
                if resp.status_code == 204 or not resp.text.strip():
                    return None
                return resp.json()
            
            if resp.status_code in {429, 500, 502, 503, 504} and attempt < self.max_retries:
                retry_after = resp.headers.get("Retry-After")
                if retry_after and retry_after.isdigit():
                    sleep_s = float(retry_after)
                else:
                    sleep_s = self.backoff_base_sec * (2 * attempt)
                time.sleep(sleep_s)
                continue
            
            req_id = resp.headers.get("request-id") or resp.headers.get("client-request-id")
            detail = ""
            
            try:
                detail_json = resp.json()
                detail = str(detail_json.get("error") or detail_json)
            except Exception:
                detail = resp.text[:500]
                
            raise GraphClientError(
                f"Graph {method.upper()} {path} -> {resp.status_code}."
                f"request_id={req_id}. detail={detail}"
            )
            
            
    def get_me(self) -> dict[str, Any]:
        return self._request("GET", "/me", params={"$select": "id, displayName, userPrincipalName,mail"})
    
    def list_mail_folders(self, top: int = 50) -> list[dict[str, Any]]:
        data = self._request(
            "GET",
            "/me/mailFolders",
            params={"$top": top, "$select": "id,displayName,parentFolderId,totalItemCount,unreadItemCount"},
        ) or {}
        return list(data.get("value", []))
    
    def list_messages(
        self,
        folder_id: str,
        top: int = 50,
        select: str = "id,subject,conversationId,receivedDateTime,from",
        orderby: str = "receivedDateTime desc",
    ) -> list[dict[str, Any]]:
        data = self._request(
            "GET",
            f"/me/mailFolders/{folder_id}/messages",
            params={"$top": top, "$select": select, "$orderby": orderby},
        ) or {}
        return list(data.get("value", []))
    
    def send_mail(
        self,
        *,
        to_recipients: list[str],
        subject: str,
        html_body: str,
        file_attachments: list[dict[str, str]] | None = None,
        save_to_sent_items: bool = True,
    ) -> None:
        # file_attachments: [{"name":"a.pdf","contentBytes":"<base64>","contentType":"application/pdf"}]
        recipients = [{"emailAddress": {"address": x}} for x in to_recipients if x]
        attachments = []
        for a in file_attachments or []:
            attachments.append(
                {
                    "@odata.type": "#microsoft.graph.fileAttachment",
                    "name": a["name"],
                    "contentType": a.get("contentType", "application/octet-stream"),
                    "contentBytes": a["contentBytes"],
                }
            )
        body = {
            "message": {
                "subject": subject or "",
                "body": {"contentType": "HTML", "content": html_body or ""},
                "toRecipients": recipients,
                "attachments": attachments,
            },
            "saveToSentItems": bool(save_to_sent_items),
        }
        
        self._request("POST", "/me/sendMail", json_body=body, expected_status={202})
        
    def create_reply_all_draft(self, message_id: str) -> dict[str, Any]:
        return self._request(
            "POST",
            f"/me/messages/{message_id}/createReplyAll",
            expected_status={201},
        ) or {}
        
    def add_file_attachment_to_message(
        self,
        *,
        message_id: str,
        name: str,
        content_bytes_b64: str,
        content_type: str = "applitacion/octect-stream",
        is_inline: bool = False,
        content_id: str | None = None,
    ) -> dict[str, Any]:
        payload: dict[str, Any] = {
            "@odata.type": "#microsoft.graph.fileAttachment",
            "name": name,
            "contentType": content_type,
            "contentBytes": content_bytes_b64,
            "isInline": bool(is_inline)
        }
        if content_id:
            payload["contentId"] = content_id
            
            
        return self._request(
            "POST",
            f"/me/messages/{message_id}/attachments",
            json_body=payload,
            expected_status={201},
        ) or {}
        
    def patch_message(
        self,
        message_id: str,
        *,
        subject: str | None = None,
        html_body: str | None = None,
    ) -> None:
        payload: dict[str, Any] = {}
        if subject is not None:
            payload["subject" ] = subject
        if html_body is not None:
            payload ["body"] = {"contentType": "HTML", "content": html_body}    
        
        if not payload:
            return
        
        self._request("PATCH", f"/me/messages/{message_id}", json_body=payload, expected_status={200})
    
    def send_draft(self, message_id: str) -> None:
        self._request("POST", f"/me/messages/{message_id}/send", expected_status={202})
        
    def move_message(self, message_id: str, destination_folder_id: str) -> dict[str, Any]:
        return self._request(
            "POST",
            f"/me/messages/{message_id}/move",
            json_body={"destinationId": destination_folder_id},
            expected_status={201},
        ) or {}
        
    def find_latest_by_conversation(self, folder_id: str, conversation_id: str) -> dict[str,Any] | None:
        data = self._request(
            "GET",
            f"/me/mailFolders/{folder_id}/messages",
            params={
                "$top": 1,
                "$filter": f"conversationId eq '{conversation_id}'",
                "$orderby": "receivedDateTime desc",
                "$select": "id,subject, conversationId,receivedDateTime",
            },
        ) or {}
        items = list(data.get("value", []))
        return items[0] if items else None
        