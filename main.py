import sys
import os
import csv
import json
from tabulate import tabulate
from extractor import extraer_por_codigo  # Usa el extractor que retorna orden, nombre y puntos

def main():
    if len(sys.argv) < 3:
        print("Uso: python main.py archivo.pdf codigo")
        sys.exit(1)

    pdf_path = sys.argv[1]
    codigo = sys.argv[2]

    print(f"📂 Archivo recibido: {pdf_path}")
    print(f"🔢 Código recibido: {codigo}")
    print("Procesando...")

    nombre_materia, registros = extraer_por_codigo(pdf_path, codigo)

    print(f"🗂️  Materia encontrada: {nombre_materia}")
    print(f"📝 Registros encontrados: {len(registros)}")

    if not registros:
        print(f"⚠️ No se encontraron registros para el código '{codigo}'.")
        sys.exit(1)

    print(f"\n({codigo}) {nombre_materia}\n")

    # Encabezados deseados
    headers = ["Orden", "Apellido y Nombre", "Puntos"]
    rows = [[r["Orden"], r["Apellido y Nombre"], r["Puntos"]] for r in registros]

    # Mostrar en consola
    print(tabulate(rows, headers=headers, tablefmt="github"))

    # Crear carpeta de salida
    os.makedirs("data", exist_ok=True)

    # Guardar CSV
    csv_path = os.path.join("data", f"{codigo}_resumido.csv")
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
    json_path = os.path.join("data", f"{codigo}_listado.json")
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(salida_json, f, indent=2, ensure_ascii=False)
    print(f"✅ JSON guardado en: {json_path}")

if __name__ == "__main__":
    main()
