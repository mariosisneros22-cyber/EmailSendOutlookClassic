from __future__ import annotations

import customtkinter as ctk
from tkinter import messagebox

from core.graph.auth import get_active_account, login_interactive, logout_local
from ui.theme import BTN_PRIMARY, BTN_SECONDARY, CARD_FG_COLOR,FONT_BODY, FONT_TITLE

def _account_label(account:dict | None) -> str:
    if not account:
        return "No autenticado"
    user = account.get("username") or account.get("name") or "(sin cuenta)"
    return f"Autenticado como: {user}"


def mount(parent, on_authenticated):
    frame = ctk.CTkFrame(parent, fg_color = "transparent")
    frame.pack(fill="both", expand=True, padx=20, pady=20)
    
    card = ctk.CTkFrame(frame, fg_color=CARD_FG_COLOR, corner_radius=12)
    card.pack(fill="both", expand=True, padx=80, pady=80)
    
    ctk.CTkLabel(card, text="Inicio de sesion Microsoft", font=FONT_TITLE).pack(
        anchor="center", pady=(30, 16)
    )
    
    status_var = ctk.StringVar(value="Verificando sesion...")
    ctk.CTkLabel(card, textvariable=status_var, font=FONT_BODY).pack(
        anchor="center", pady=(0, 20)
    )
    
    buttons=ctk.CTkFrame(card, fg_color="transparent")
    buttons.pack(anchor="center", pady=(0, 20))
    
    def refresh_status() ->dict | None:
        account = get_active_account()
        status_var.set(_account_label(account))
        return account

    def on_login():
        try:
            login_interactive()
            account = refresh_status()
            if account:
                frame.destroy()
                on_authenticated(account)
        except Exception as exc:
            messagebox.showerror("Login Microsoft", str(exc))
    
    def on_logout():
        try:
            logout_local()
            refresh_status()
        except Exception as exc:
            messagebox.showerror("Cerrar sesion", str(exc))
            
    def on_continue():
        account = get_active_account()
        if not account:
            messagebox.showwarning("Sesion", "Primero debes iniciar sesion")
            return
        frame.destroy()
        on_authenticated(account)
        
    btn_login = ctk.CTkButton(
        buttons,
        text="Iniciar sesion",
        command=on_login,
        width=170,
        **BTN_PRIMARY,
    )
    btn_login.pack(side="left", padx=6)
    
    btn_continue = ctk.CTkButton(
        buttons,
        text="Continuar",
        command=on_continue,
        widht=170,
        **BTN_SECONDARY,
    )
    btn_continue.pack(side="left",padx=6)
    
    btn_logout = ctk.CTkButton(
        buttons,
        text = "Cerrar sesion",
        command=on_logout,
        width=170,
        **BTN_SECONDARY,
    )
    btn_logout.pack(side="left", padx=6)
    
    refresh_status()
    return frame