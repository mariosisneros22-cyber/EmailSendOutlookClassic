#src/core/graph/auth.py
from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
import json
import msal

from core.config.app_dirs import data_dir
from core.config.config_store import load_config

SCOPES = ["User.Read", "Mail.Read", "Mail.Send", "Mail.ReadWrite"]
AUTHORITY_TEMPLATE = "https://login.microsoftonline.com/{tenant_id}"
REDIRECT_URI = "http://localhost"

@dataclass
class GraphAuthConfig:
    client_id: str
    tenant_id: str
    scopes: list[str]
    
class GraphAuthError(RuntimeError):
    pass

def _token_cache_path() -> Path:
    p = data_dir() / "graph_token_cache.bin"
    p.parent.mkdir(parents=True, exist_ok=True)
    return p

def _active_account_path() -> Path:
    p = data_dir() / "graph_active_account.json"
    p.parent.mkdir(parents=True, exist_ok=True)
    return p

def _load_cfg() -> GraphAuthConfig:
    cfg = load_config()
    client_id = str(cfg.get("graph_client_id", "")).strip()
    tenant_id = str(cfg.get("graph_tenant_id", "")).strip()
    scopes = cfg.get("graph_scopes") or SCOPES
    
    if not client_id:
        raise GraphAuthError("Falta graph_client_id en config.")
    if not tenant_id:
        raise GraphAuthError("Falta graph_tenant_id en config.")

    return GraphAuthConfig(client_id= client_id, tenant_id=tenant_id,scopes=list(scopes))

def _load_cache() ->msal.SerializableTokenCache:
    cache = msal.SerializableTokenCache()
    cache_path = _token_cache_path()
    if cache_path.exists():
        cache.deserialize(cache_path.read_text(encoding="utf-8"))
    return cache

def _save_cache(cache: msal.SerializableTokenCache) -> None:
    if cache.has_state_changed:
        _token_cache_path().write_text(cache.serialize(), encoding="utf-8")
        

def _build_app(cfg: GraphAuthConfig, cache: msal.SerializableTokenCache) -> msal.PublicClientApplication:
    return msal.PublicClientApplication(
        client_id=cfg.client_id,
        authority = AUTHORITY_TEMPLATE.format(tenant_id=cfg.tenant_id),
        token_cache=cache,
    )

def _save_active_account(home_account_id: str) -> None:
    _active_account_path().write_text(
        json.dumps({"home_account_id": home_account_id}, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    

def _load_active_account_id() -> str | None:
    p = _active_account_path()
    if not p.exists():
        return None
    
    data = json.loads(p.read_text(encoding="utf-8"))
    v = str(data.get("home_account_id", "")).strip()
    return v or None

def _clear_active_account() -> None:
    p = _active_account_path()
    if p.exists():
        p.unlink(missing_ok=True)
        
def _pick_account(app: msal.PublicClientApplication) -> dict | None:
    wanted = _load_active_account_id()
    accounts = app.get_accounts()
    if not accounts:
        return None
    
    if wanted:
        for acc in accounts:
            if acc.get("home_account_id") == wanted:
                return acc
            
    # fallback: primera cuenta cacheada
    return accounts[0]


def login_interactive() -> dict:
    cfg = _load_cfg()
    cache = _load_cache()
    app = _build_app(cfg, cache)
    
    result = app.acquire_token_interactive(
        scopes=cfg.scopes,
        redirect_uri = REDIRECT_URI,
        prompt="select_account",
    )
    
    if "access_token" not in result:
        raise GraphAuthError(f"Login falló: {result.get('error_description') or result}")
    
    home_account_id = result.get("account", {}).get("home_account_id")
    
    if home_account_id:
        _save_active_account(home_account_id)
        
    _save_cache(cache)
    return result
    
    
def get_access_token() -> str:
    cfg = _load_cfg()
    cache = _load_cache()
    app = _build_app(cfg, cache)
    
    account = _pick_account(app)
    if not account:
        raise GraphAuthError("No hay sesion iniciada, ejecuta login_interactive().")
    
    result = app.acquire_token_silent(scopes=cfg.scopes, account=account)
    if not result or  "access_token" not in result:
        raise GraphAuthError("No se pudo renovar token en modo silencioso. Inicia otra vez")
    
    _save_cache(cache)
    return str(result["access_token"])
    
    
    
    
def get_active_account() -> dict | None:
    cfg = _load_cfg()
    cache = _load_cache()
    app = _build_app(cfg, cache)
    return _pick_account(app)

def logout_local() -> None:
    """
    Cierra sesión local de la app (borra cache/tokens locales).
    No cierra sesion global de  Microsoft en el navegador
    """
    p_cache = _token_cache_path()
    if p_cache.exists():
        p_cache.unlink(missing_ok=True)
    _clear_active_account()
    