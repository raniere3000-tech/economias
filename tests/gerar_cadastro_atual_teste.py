"""Gera um cadastro atual sintético (como "Base Cadastro - 11.11.2026.xlsx") a partir da base sintética do estudo,
para testar o filtro "Só último contrato (SIM)" (v1.37).
Uso: python tests/gerar_cadastro_atual_teste.py base_teste.xlsx "Base Cadastro - 11.11.2026.xlsx"
Como no arquivo real, toda matrícula do cadastro tem uma linha SIM. De cada 10 matrículas da base: 6 com o último contrato
ATIVO, 1 com o último contrato ENCERRADO, 1 com o último contrato ANULADO, 1 com duas linhas (NÃO ENCERRADO e SIM ATIVO, fica no
filtro) e 1 fora do cadastro. O filtro "Só último contrato ativo" (v1.38) deixa as que têm a linha SIM com SIT_CONTRATO = ATIVO.
TIPO_FATURAMENTO da linha SIM (v1.39, filtro "Só AGUA E ESGOTO"): AGUA em 1 de cada 3 dezenas de matrículas, AGUA E ESGOTO nas outras;
a linha NÃO é sempre AGUA E ESGOTO, para conferir que ela não conta.
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
ws.append(['NUM_LIGACAO', 'NOM_CLIENTE', 'CIDADE', 'SIT_LIG', 'SIT_CONTRATO', 'TIPO_FATURAMENTO', 'ULTIMO_CONTRATO'])
n = {'ATIVO': 0, 'ENCERRADO': 0, 'ANULADO': 0, 'duas': 0, 'fora': 0}
for i, m in enumerate(sorted(mats, key=str)):
    r = i % 10
    if r == 9: n['fora'] += 1; continue
    tf = 'AGUA' if (i // 10) % 3 == 0 else 'AGUA E ESGOTO'   # v1.39: tipo da linha SIM (1 em cada 3 dezenas fica AGUA)
    if r == 8: ws.append([m, 'X', 'Y', 'ATIVA', 'ENCERRADO', 'AGUA E ESGOTO', 'NÃO']); ws.append([m, 'X', 'Y', 'ATIVA', 'ATIVO', tf, 'SIM']); n['duas'] += 1; continue
    sc = 'ENCERRADO' if r == 7 else 'ANULADO' if r == 6 else 'ATIVO'
    ws.append([m, 'X', 'Y', 'INATIVA' if sc != 'ATIVO' else 'ATIVA', sc, tf, 'SIM']); n[sc] += 1
for k in range(50): ws.append([90000000 + k, 'SO CADASTRO', 'Y', 'ATIVA', 'ATIVO', 'AGUA E ESGOTO', 'SIM'])
wb.save(saida)
print(f'{len(mats)} matrículas da base: último contrato ATIVO {n["ATIVO"]}, ENCERRADO {n["ENCERRADO"]}, ANULADO {n["ANULADO"]}, NÃO + SIM ATIVO {n["duas"]}, fora do cadastro atual {n["fora"]}; +50 só no cadastro')
