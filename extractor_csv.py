# extractor.py
import re
import warnings
import pdfplumber

warnings.filterwarnings("ignore", message="CropBox missing from /Page")

COLUMNS = ['Orden', 'Apellido y Nombre', 'Puntos']

# Patrones: principal, alternativo, y genérico
registro_pattern = re.compile(r"^(\d+)\s+(\d{1,3}(?:\.\d{3}){1,2})\s+(.+)")
alternate_pattern = re.compile(r"^(\d+)\s+(\d+)\s+(.+)")
fallback_pattern = re.compile(r"^(\d+)\s+(.+)")

patron_titulo = re.compile(r"^\(\s*(\d+)\)\s*(.+)")


def normalizar_puntos(valor: str) -> str:
    valor = valor.strip().replace('.', '')
    if re.match(r"^\d{1,3},\d{2}$", valor):
        return valor
    if valor.isdigit():
        if len(valor) >= 3:
            return f"{int(valor[:-2])},{valor[-2:]}"
        else:
            return f"{valor},00"
    return valor


def extraer_por_codigo(pdf_path: str, codigo: str) -> tuple[str, list[dict]]:
    nombre_materia = None
    listado = []
    header_detectado = False
    buscando_codigo = False
    orden_esperado = 1

    print(f"📄 Buscando código ({codigo}) en {pdf_path}...")

    with pdfplumber.open(pdf_path) as pdf:
        for page_num, page in enumerate(pdf.pages, start=1):
            texto = page.extract_text() or ""
            for linea in texto.splitlines():
                linea_strip = linea.strip()

                m = patron_titulo.match(linea_strip)
                if m:
                    if m.group(1) == codigo:
                        nombre_materia = m.group(2).strip()
                        buscando_codigo = True
                        print(f"✅ Sección ({codigo}) encontrada: {nombre_materia} en página {page_num}")
                        continue
                    elif buscando_codigo:
                        print(f"⏹️ Fin de sección ({codigo}) al detectar nueva sección ({m.group(1)})")
                        return nombre_materia, listado

                if buscando_codigo:
                    if not header_detectado and "Apellido y Nombre" in linea_strip:
                        header_detectado = True
                        print("🔍 Encabezado detectado. Procesando registros...")
                        continue

                    if header_detectado:
                        match = registro_pattern.match(linea_strip)
                        if not match:
                            match = alternate_pattern.match(linea_strip)
                        if not match:
                            match = fallback_pattern.match(linea_strip)

                        if match:
                            orden = match.group(1)

                            # Chequear si el orden no es correlativo
                            if int(orden) != orden_esperado:
                                print(f"🔁 Salto de orden esperado {orden_esperado} → {orden}, usando patrón alternativo")

                            resto = match.group(len(match.groups())).strip().split()

                            # Encontrar el índice donde comienza el nombre (empieza con letra)
                            idx_letra = next((i for i, x in enumerate(resto) if x[0].isalpha()), 0)
                            campos_sin_dni = resto[idx_letra:] if idx_letra < len(resto) else resto

                            puntos_raw = campos_sin_dni[-1] if campos_sin_dni else ""
                            puntos = normalizar_puntos(puntos_raw)

                            # Buscar desde nombre al primer campo numérico para cortar apellido y nombre
                            idx_num = next((i for i, v in enumerate(campos_sin_dni)
                                            if re.match(r"^\d+(?:,\d{2})?$", v)), None)

                            nombre = " ".join(campos_sin_dni[:idx_num]) if idx_num is not None else " ".join(campos_sin_dni)

                            fila = {
                                'Orden': orden,
                                'Apellido y Nombre': nombre.strip(),
                                'Puntos': puntos
                            }
                            listado.append(fila)
                            orden_esperado += 1
                            print(f"➕ Registro agregado: {fila}")
                        else:
                            print(f"⚠️ Línea ignorada (no coincide con ningún patrón): {linea_strip}")

    print(f"🔚 Fin del archivo. Registros extraídos: {len(listado)}")
    return nombre_materia or "", listado
