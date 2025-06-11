import sys
import os
import csv
import json
from tabulate import tabulate
from extractor import extraer_por_codigo # Usa el extractor que retorna orden, nombre y puntos

def save_extracted_data(codigo, nombre_materia, registros, output_dir="data"):
    """
    Guarda los datos extraídos en archivos CSV y JSON.

    Args:
        codigo (str): Código de la materia.
        nombre_materia (str): Nombre de la materia.
        registros (list[dict]): Lista de registros de estudiantes.
        output_dir (str): Directorio donde se guardarán los archivos.

    Returns:
        tuple[str, str]: Rutas a los archivos CSV y JSON guardados.
    """
    os.makedirs(output_dir, exist_ok=True)

    headers = ["Orden", "Apellido y Nombre", "Puntos"]
    rows = [[r.get("Orden", ""), r.get("Apellido y Nombre", ""), r.get("Puntos", "")] for r in registros]

    # Guardar CSV
    csv_filename = f"{codigo}_resumido.csv"
    csv_path = os.path.join(output_dir, csv_filename)
    with open(csv_path, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(headers)
        writer.writerows(rows)
    print(f"\n✅ CSV guardado en: {csv_path}")

    # Guardar JSON
    salida_json = {
        "codigo": codigo,
        "materia": nombre_materia,
        "listado": registros
    }
    json_filename = f"{codigo}_listado.json"
    json_path = os.path.join(output_dir, json_filename)
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(salida_json, f, indent=2, ensure_ascii=False)
    print(f"✅ JSON guardado en: {json_path}")

    return csv_path, json_path

def main():
    if len(sys.argv) < 3:
        print("Uso: python main.py archivo.pdf codigo")
        sys.exit(1)

    pdf_path = sys.argv[1]
    codigo = sys.argv[2]

    print(f"📂 Archivo recibido: {pdf_path}")
    print(f"🔢 Código recibido: {codigo}")
    print("Procesando...")

    # Asegurarse de que extractor.py está accesible y se importa correctamente
    # (ya está importado al inicio del script como `from extractor import extraer_por_codigo`)
    nombre_materia, registros = extraer_por_codigo(pdf_path, codigo)

    if nombre_materia is None:
        print(f"⚠️ No se encontró la materia con el código '{codigo}' en el PDF.")
        # Decidir si salir o continuar sin registros
        # Por ahora, se asume que si no hay nombre_materia, no hay registros válidos o la sección no se encontró.
        sys.exit(1)

    print(f"🗂️  Materia encontrada: {nombre_materia}")
    print(f"📝 Registros encontrados: {len(registros)}")

    if not registros:
        print(f"⚠️ No se encontraron registros para el código '{codigo}' en la materia '{nombre_materia}'.")
        # Se podría optar por guardar archivos vacíos o simplemente salir
        # Por ahora, salimos si no hay registros, pero después de encontrar la materia.
        sys.exit(1)

    print(f"\n({codigo}) {nombre_materia}\n")

    # Encabezados deseados para tabulate
    headers_display = ["Orden", "Apellido y Nombre", "Puntos"]
    # Usar .get() para evitar KeyError si alguna clave falta en algún diccionario de registro
    rows_display = [[r.get("Orden",""), r.get("Apellido y Nombre",""), r.get("Puntos","")] for r in registros]

    # Mostrar en consola
    print(tabulate(rows_display, headers=headers_display, tablefmt="github"))

    # Guardar los datos usando la nueva función
    csv_path, json_path = save_extracted_data(codigo, nombre_materia, registros)

    # Los mensajes de confirmación ya están en save_extracted_data
    # print(f"\n✅ CSV guardado en: {csv_path}")
    # print(f"✅ JSON guardado en: {json_path}")

if __name__ == "__main__":
    main()
