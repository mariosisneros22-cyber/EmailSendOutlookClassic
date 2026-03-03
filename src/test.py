import os
from core.responder_queue import ensure_control_file, update_control_from_outlook, process_pending_responses
from core.responder_queue import CONTROL_PATH
print("CONTROL_PATH =", CONTROL_PATH)
# 1) Asegura archivo control
ensure_control_file()

# 2) (Opcional) Trae correos nuevos desde Outlook a tu control
# update_control_from_outlook()

# 3) Define carpeta donde está el PDF
CARPETA_PDFS = r"C:\Users\Usuario\OneDrive - CLOUD PERU\Escritorio\Carpeta de Archivos test"  # <-- cambia esto

# 4) Ejecuta SOLO 1 (modo seguro)
n = process_pending_responses(
    carpeta_archivos=CARPETA_PDFS,
    html_body="Hola,<br><br>Adjunto documento.<br><br>Saludos.",
    delay_segundos=1.0,
    only_first_n=1,
)

print("Procesados:", n)