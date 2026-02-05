#!/usr/bin/env python3
"""
Script de conversion XLS vers XLSX
Utilise pandas et openpyxl
"""

import sys
import pandas as pd

def convert_xls_to_xlsx(input_file, output_file):
    """Convertit un fichier .xls en .xlsx"""
    try:
        print(f"📖 Lecture de {input_file}...")
        # Lire le fichier .xls
        df = pd.read_excel(input_file, engine='xlrd', header=None)
        
        print(f"✓ {len(df)} lignes trouvées")
        
        # Écrire en .xlsx
        print(f"💾 Écriture vers {output_file}...")
        df.to_excel(output_file, index=False, header=False, engine='openpyxl')
        
        print(f"✅ Conversion réussie !")
        return True
        
    except Exception as e:
        print(f"❌ Erreur: {e}")
        return False

if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("Usage: python3 convert_xls.py input.xls output.xlsx")
        sys.exit(1)
    
    input_file = sys.argv[1]
    output_file = sys.argv[2]
    
    success = convert_xls_to_xlsx(input_file, output_file)
    sys.exit(0 if success else 1)
