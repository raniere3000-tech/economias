"""Gera um cadastro atual sintético (como "Base Cadastro - 11.11.2026.xlsx") a partir da base sintética do estudo,
para testar o filtro "Só último contrato (SIM)" (v1.37).
Uso: python tests/gerar_cadastro_atual_teste.py base_teste.xlsx "Base Cadastro - 11.11.2026.xlsx"
De cada 10 matrículas da base: 7 com SIM, 1 com NÃO, 1 com duas linhas (NÃO e SIM, fica no filtro) e 1 fora do cadastro.
Acrescenta também matrículas que só existem no cadastro atual."""
import sys
from openpyxl import Workbook, load_workbook

base, saida = sys.argv[1], sys.argv[2]
mats = set()
for ws in load_workbook(base, read_only=True).worksheets:
    cab = None
    for row in ws.iter_rows(values_only=True):
        if cab is None:
            if row and 'NUM_LIGACAO' in row: cab = row.index('NUM_LIGACAO')
            continue
        if row[cab] is not None: mats.add(row[cab])
wb = Workbook(); ws = wb.active; ws.title = 'Query1'
ws.append(['NUM_LIGACAO', 'NOM_CLIENTE', 'CIDADE', 'SIT_LIG', 'ULTIMO_CONTRATO'])
n = {'SIM': 0, 'NÃO': 0, 'duas': 0, 'fora': 0}
for i, m in enumerate(sorted(mats, key=str)):
    r = i % 10
    if r == 9: n['fora'] += 1; continue
    if r == 8: ws.append([m, 'X', 'Y', 'ATIVA', 'NÃO']); ws.append([m, 'X', 'Y', 'ATIVA', 'SIM']); n['duas'] += 1; continue
    u = 'NÃO' if r == 7 else 'SIM'; ws.append([m, 'X', 'Y', 'ATIVA', u]); n[u] += 1
for k in range(50): ws.append([90000000 + k, 'SO CADASTRO', 'Y', 'ATIVA', 'SIM'])
wb.save(saida)
print(f'{len(mats)} matrículas da base: {n["SIM"]} SIM, {n["NÃO"]} NÃO, {n["duas"]} com NÃO e SIM, {n["fora"]} fora do cadastro atual; +50 só no cadastro')
