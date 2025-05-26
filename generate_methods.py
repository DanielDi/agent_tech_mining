#!/usr/bin/env python3
import os
import json
import argparse
import pandas as pd

def main():
    parser = argparse.ArgumentParser(
        description="Generate methods.json from an Excel file of co-word analysis."
    )
    parser.add_argument(
        "--input",
        "-i",
        required=True,
        help="Path to the Excel file (e.g. 'Tendencias en la aplicación práctica.xlsx')"
    )
    parser.add_argument(
        "--output",
        "-o",
        default="methods.json",
        help="Path to output JSON file (default: methods.json)"
    )
    args = parser.parse_args()

    if not os.path.isfile(args.input):
        print(f"Error: Excel file not found: {args.input}", file=sys.stderr)
        return

    # Lee todo el DataFrame
    df = pd.read_excel(args.input)

    # Mapeo de columnas de Excel a las claves internas
    columns_map = {
        "AnalyzedFields": "¿Qué se analiza?",
        "TermPreprocessing": "Preprocesamiento de Términos",
        "Clustering": "Clustering",
        "ClusterAnalysis": "Análisis de Clusters"
    }

    methods = {}
    for key, col in columns_map.items():
        if col not in df.columns:
            methods[key] = []
            continue
        # Divide por ';', aplana y limpia
        items = (
            df[col]
            .dropna()
            .astype(str)
            .str.split(";")
            .explode()
            .str.strip()
        )
        unique_items = sorted([i for i in items.unique() if i])
        methods[key] = unique_items

    # Escribe el JSON
    with open(args.output, "w", encoding="utf-8") as f:
        json.dump(methods, f, ensure_ascii=False, indent=2)

    print(f"✅ Generated {args.output} with the following content:")
    print(json.dumps(methods, ensure_ascii=False, indent=2))

if __name__ == "__main__":
    main()
