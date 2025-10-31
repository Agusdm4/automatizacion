
import sys
import os
import re
import pandas as pd

# Dependencies: PyPDF2 (if not available, install it in your environment)
try:
    import PyPDF2
except Exception as e:
    print("ERROR: Necesitás tener instalado PyPDF2. Instalalo con: pip install PyPDF2")
    sys.exit(1)

MASTER_PATH = "Master_Envios.xlsx"  # Ubicá este script en la misma carpeta del Excel o ajustá la ruta

COLUMNS = [
    "Customer Order Number",
    "Número de B/L",
    "Contenedores (uno por línea)",
    "Producto",
    "Toneladas Netas (kg)",
    "Precio Final (USD)",
    "Número de Factura",
]

def extract_text(pdf_path):
    text = ""
    with open(pdf_path, "rb") as f:
        reader = PyPDF2.PdfReader(f)
        for page in reader.pages:
            try:
                page_text = page.extract_text() or ""
                text += page_text + "\n"
            except Exception:
                continue
    return text

# ---------- Helpers ----------

def find_first(pattern, text, flags=0):
    m = re.search(pattern, text, flags)
    return m.group(1).strip() if m else ""

def find_all(pattern, text, flags=0):
    return re.findall(pattern, text, flags)

# ---------- Parsers ----------

def parse_invoice_number(text):
    # Strict: line that starts with 'Invoice ' followed by digits
    m = re.search(r"(?:^|\n)\s*Invoice\s+(\d{6,})\b", text, flags=re.IGNORECASE)
    if m:
        return m.group(1)
    # Fallback: 'Invoice:' + digits anywhere
    m = re.search(r"Invoice\s*[:#]?\s*(\d{6,})\b", text, flags=re.IGNORECASE)
    return m.group(1) if m else ""

def parse_customer_order(text):
    # Strict format per tu regla: CE-xxxx-xx
    m = re.search(r"\bCE-\d{4}-\d{2}\b", text)
    if m:
        return m.group(0)
    # Fallback general y limpieza para cortar cualquier letra pegada
    m = re.search(r"Customer\s+Order\s+Number\s*[:#]?\s*([A-Za-z0-9\-\/]+)", text, flags=re.IGNORECASE)
    if m:
        val = m.group(1)
        # Si empieza con CE-, recortamos hasta el patrón esperado
        m2 = re.search(r"\bCE-\d{4}-\d{2}\b", val)
        if m2:
            return m2.group(0)
        # En última instancia, cortamos cualquier sufijo no permitido después de números/letras y guiones
        val = re.match(r"[A-Za-z0-9\-\/]+", val).group(0)
        return val
    return ""

def parse_bl_number(text):
    # 1) Primario: etiqueta explícita
    m = re.search(r"(?:^|\n)\s*(?:BILL\s+OF\s+LADING\s+No\.?|B\/?L\s*No\.?)\s*[:#]?\s*([A-Z0-9]{8,20})", text, flags=re.IGNORECASE)
    if m and re.search(r"\d", m.group(1)):
        return m.group(1)

    # 2) Buscar cerca de encabezados típicos del B/L para evitar BOOKING REF
    for kw in ("BILL OF LADING", "RIDER PAGE", "BILL OF LADING No", "RIDER"):
        for km in re.finditer(kw, text, flags=re.IGNORECASE):
            start = max(0, km.start() - 300)
            end = min(len(text), km.end() + 300)
            window = text[start:end]
            # códigos alfanum largos con dígitos, excluir EBKG* (booking refs) y tokens adyacentes a 'BOOKING REF'
            candidates = re.findall(r"\b(?!EBKG)[A-Z]{3,6}[A-Z0-9]{6,12}\b", window)
            candidates = [c for c in candidates if re.search(r"\d", c)]
            # descartar candidatos con 'BOOKING' cerca
            filtered = []
            for c in candidates:
                # posición de c en ventana
                pos = window.find(c)
                near = window[max(0, pos - 40): pos + len(c) + 40]
                if re.search(r"BOOKING\s+REF", near, flags=re.IGNORECASE):
                    continue
                filtered.append(c)
            if filtered:
                # elegir el más largo (suele ser el BL impreso solo)
                filtered.sort(key=len, reverse=True)
                return filtered[0]

    # 3) Fallback global: patrón general excluyendo EBKG y evitando proximidad a BOOKING REF
    for m in re.finditer(r"\b(?!EBKG)[A-Z]{3,6}[A-Z0-9]{6,12}\b", text):
        token = m.group(0)
        if not re.search(r"\d", token):
            continue
        near = text[max(0, m.start()-40): m.end()+40]
        if re.search(r"BOOKING\s+REF", near, flags=re.IGNORECASE):
            continue
        return token
    return ""

def parse_containers(text):
    # ISO 6346-like: 4 letters + 7 digits
    conts = set(find_all(r"\b[A-Z]{4}\d{7}\b", text))
    return sorted(conts)

def parse_product(text):
    # Captura línea con AGILITY y LDPE
    m = re.search(r"(AGILITY[^\n]{0,120}?LDPE)", text, flags=re.IGNORECASE)
    if m:
        product = m.group(1)
        product = product.replace("™", "").strip()
        product = " ".join(product.split())
        return product.upper()
    return ""

def parse_total_net_weight(text):
    # Preferente: TOTAL NET WEIGHT
    m = re.search(r"TOTAL\s+NET\s+WEIGHT\s*[:#]?\s*([\d\.,]+)\s*KG", text, flags=re.IGNORECASE)
    if m:
        val = m.group(1).replace(",", "")
        try:
            return f"{float(val):.3f}"
        except:
            pass

    # Fallback: sumar una sola vez por contenedor
    total = 0.0
    seen = set()
    for cm in re.finditer(r"\b([A-Z]{4}\d{7})\b", text):
        code = cm.group(1)
        if code in seen:
            continue
        start = max(0, cm.start() - 400)
        end = min(len(text), cm.end() + 400)
        window = text[start:end]
        m2 = re.search(r"Item\s+Net\s+Weight\s*:\s*([\d\.,]+)\s*KG", window, flags=re.IGNORECASE)
        if m2:
            val = m2.group(1).replace(",", "")
            try:
                total += float(val)
                seen.add(code)
            except:
                pass
    if total > 0:
        return f"{total:.3f}"
    return ""

def parse_total_amount(text):
    # Evita capturar totales de peso/volumen
    m = re.search(r"(?:^|\n)\s*Total\s+([\d\.,]{2,})\b(?![^\n]{0,40}\b(KG|CD3|WEIGHT|VOLUME)\b)", text, flags=re.IGNORECASE)
    if m:
        val = m.group(1).replace(",", "")
        try:
            return f"{float(val):.2f}"
        except:
            pass
    m = re.search(r"Total\s+([\d\.,]{2,})", text, flags=re.IGNORECASE)
    if m:
        val = m.group(1).replace(",", "")
        try:
            return f"{float(val):.2f}"
        except:
            pass
    return ""

def parse_pdf(text):
    customer_order = parse_customer_order(text)
    bl = parse_bl_number(text)
    containers = parse_containers(text)
    product = parse_product(text)
    net_kg = parse_total_net_weight(text)
    total_amount = parse_total_amount(text)
    invoice = parse_invoice_number(text)

    return {
        "Customer Order Number": customer_order,
        "Número de B/L": bl,
        "Contenedores (uno por línea)": "\n".join(containers),
        "Producto": product,
        "Toneladas Netas (kg)": net_kg,
        "Precio Final (USD)": total_amount,
        "Número de Factura": invoice,
    }

def append_to_excel(row_dict, master_path=MASTER_PATH):
    # Create file if not exists
    if not os.path.exists(master_path):
        df = pd.DataFrame(columns=COLUMNS)
        with pd.ExcelWriter(master_path, engine="xlsxwriter") as writer:
            df.to_excel(writer, index=False, sheet_name="Datos")

    # Append row
    df = pd.read_excel(master_path, sheet_name="Datos")
    for c in COLUMNS:
        if c not in df.columns:
            df[c] = ""
    df = df[COLUMNS]

    df = pd.concat([df, pd.DataFrame([row_dict])], ignore_index=True)

    with pd.ExcelWriter(master_path, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="Datos")

def main():
    if len(sys.argv) < 2:
        print("Uso: python append_from_pdf.py <ruta_al_pdf>")
        sys.exit(1)
    pdf_path = sys.argv[1]
    if not os.path.exists(pdf_path):
        print(f"ERROR: No existe el archivo: {pdf_path}")
        sys.exit(1)

    text = extract_text(pdf_path)
    row = parse_pdf(text)
    append_to_excel(row)
    print("Listo: fila agregada a", MASTER_PATH)

if __name__ == "__main__":
    main()
