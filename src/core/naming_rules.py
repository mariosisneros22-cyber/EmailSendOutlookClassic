import re
from datetime import datetime

TIPO_MAP = {
    "OBRA": "01",
    "OFICINA CENTRAL": "02",
    "GERENCIA": "03",
    "PROVEEDORES": "04",
}

def _norm_subject(subject: str) -> str:
    s = (subject or "").upper()
    s = re.sub(r"\s+"," ", s).strip()
    return s 

def detect_tipo_code(norm_s: str, tipo_map: dict) -> str | None:
    for key in sorted(tipo_map.keys(), key=len, reverse=True):
        pattern = rf"\b{re.escape(key)}\b"
        if re.search(pattern, norm_s):
            return tipo_map[key]
    
    return None

def build_filename_from_subject(subject: str, year_mode: str = "current") -> str | None:
    s = _norm_subject(subject)
    if not s:
        return None
    
    # N° OBRA
    m_obra = re.search(r"\bOBRA\s+(\d{1,4})\b", s)
    if not m_obra:
        return None
    obra_num = m_obra.group(1).zfill(3)
    
    # OS/OC + N°ORDEN
    m_doc   = re.search(r"\bO?\s*\.?\s*([SC])\s*\.?\s*(\d{1,10})\b", s)
    
    if not m_doc:
        return None
    
    sc = m_doc.group(1) # 'S' o 'C'   
    orden_raw = m_doc.group(2)
    orden_num = str(int(orden_raw)) if orden_raw else None
    doc = "OS" if sc == "S" else "OC"
    
    #TIPO
    tipo_code = detect_tipo_code(s, TIPO_MAP)
    if not tipo_code:
        return None
    
    # AÑO
    if year_mode == "current":
        year = datetime.now().year
    else:
        m_year = re.search(r"\b(20\d{2})\b", s)
        if not m_year:
            return None
        y = int(m_year.group(1))
        year = y if year_mode == "subject" else (y + 1)
        
    return f"{obra_num}-{doc}-{tipo_code}-{orden_num}-{year}.pdf"