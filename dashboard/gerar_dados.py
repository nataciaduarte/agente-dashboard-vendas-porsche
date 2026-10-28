"""Gera data.js para o dashboard a partir da planilha tratada pelo agente.

Uso: python gerar_dados.py ../dados/Porsche_data_base_sanitized.xlsx
Entram no painel apenas vendas com data valida; canceladas vao no arquivo
com o status, e o painel as exclui dos indicadores.
"""
import json
import sys
import pandas as pd
from openpyxl import load_workbook

src = sys.argv[1] if len(sys.argv) > 1 else "../dados/Porsche_data_base_sanitized.xlsx"
df = pd.read_excel(src, sheet_name=0, dtype=str)
rep = pd.read_excel(src, sheet_name="Relatorio_Qualidade", dtype=str)
info = dict(zip(rep["Verificacao"], rep["Detalhe"]))

valid = df[df["SaleDateSanitized"] != "INVALID"]
rows = [{
    "id": int(r.sale_id), "date": r.SaleDateSanitized, "model": r.PorscheModelSanitized,
    "year": r.ModelYearSanitized, "price": float(r.SalesPriceSanitized),
    "city": r.CitySanitized, "state": r.StateSanitized,
    "payment": r.PayMethodSanitized, "status": r.DeliveryStatusSanitized,
} for r in valid.itertuples()]

quality = {
    "received": len(df),
    "impossible": int(info["Datas de venda impossiveis ou ilegiveis"]),
    "future": int(info["Datas de venda futuras"].split()[0]),
    "reference": info["Datas de venda futuras"].split()[-1].rstrip(")"),
    "cancelled": sum(1 for r in rows if r["status"] == "Cancelled"),
}
quality["used"] = len(rows) - quality["cancelled"]

with open("data.js", "w", encoding="utf-8") as f:
    f.write("window.PORSCHE_DATA=" + json.dumps(rows, ensure_ascii=False) + ";\n")
    f.write("window.DATA_QUALITY=" + json.dumps(quality) + ";\n")
print(quality)
