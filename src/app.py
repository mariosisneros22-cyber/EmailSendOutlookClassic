# app.py
import customtkinter as ctk

from ui.bulk_tab import mount as mount_bulk
from ui.login_gate import mount as mount_login_gate
from ui.responder_tab import mount as mount_responder
from ui.theme import BTN_PRIMARY, BTN_SECONDARY, CARD_FG_COLOR


def _mount_main_tabs(root):
    tabs = ctk.CTkTabview(
fg_color=CARD_FG_COLOR,
        segmented_button_fg_color=BTN_SECONDARY["fg_color"],
        segmented_button_selected_color=BTN_PRIMARY["fg_color"],
        segmented_button_selected_hover_color=BTN_PRIMARY["hover_color"],
        segmented_button_unselected_color=BTN_SECONDARY["fg_color"],
        segmented_button_unselected_hover_color=BTN_SECONDARY["hover_color"],
    )

def main():
    ctk.set_appearance_mode("System")
    ctk.set_default_color_theme("blue")

    root = ctk.CTk()
    root.title("Automatizacion Outlook")
    root.geometry("1200x950")

    tabs = ctk.CTkTabview(
        root,
        fg_color=CARD_FG_COLOR,
        segmented_button_fg_color=BTN_SECONDARY["fg_color"],
        segmented_button_selected_color=BTN_PRIMARY["fg_color"],
        segmented_button_selected_hover_color=BTN_PRIMARY["hover_color"],
        segmented_button_unselected_color=BTN_SECONDARY["fg_color"],
        segmented_button_unselected_hover_color=BTN_SECONDARY["hover_color"],
    )
    tabs.pack(fill="both", expand=True, padx=10, pady=10)
    
    
    tabs._segmented_button.configure(text_color=("black", "black"))

    tab_bulk = tabs.add("Envio masivo")
    tab_resp = tabs.add("Responder")

    mount_bulk(tab_bulk)
    mount_responder(tab_resp)

    
    
    root.mainloop()

if __name__ == "__main__":
    main()
