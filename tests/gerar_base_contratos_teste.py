"""Acrescenta à base sintética o caso da troca de titularidade no mês: o contrato antigo (NÃO) faturou e o novo (SIM) veio zerado.
Uso: python tests/gerar_base_contratos_teste.py base_teste.xlsx base_multi.xlsx (40 casos em dez/25 e 20 em mar/26)."""
import sys, openpyxl, collections, copy
wb = openpyxl.load_workbook(sys.argv[1]); 
linhas = []; cab = None
for ws in wb.worksheets:
    for r in ws.iter_rows(values_only=True):
        if cab is None and r and 'NUM_LIGACAO' in r: cab = {c: i for i, c in enumerate(r)}; continue
        if r and r[0] is not None and r[0] != 'NUM_LIGACAO': linhas.append((ws, list(r)))
por = collections.defaultdict(dict)
for ws, r in linhas:
    k = (r[cab['Referência']].year, r[cab['Referência']].month); por[r[0]][k] = (ws, r)
dez = (2025, 12); mar = (2026, 3); n1 = n2 = 0
ws0 = wb.worksheets[0]
for m, meses in sorted(por.items(), key=lambda x: str(x[0])):
    if dez in meses and not meses[dez][1][cab['Qtd_eco_fat_esgoto']] and any(v[1][cab['Qtd_eco_fat_esgoto']] for v in meses.values()) and n1 < 40:
        r = copy.copy(meses[dez][1]); later = max(v[1][cab['Qtd_eco_fat_esgoto']] for v in meses.values())
        r[cab['ULTIMO_CONTRATO']] = 'NÃO'; r[cab['SIT_CONTRATO']] = 'ENCERRADO'; r[cab['Qtd_eco_fat_esgoto']] = later; r[cab['Faturamento_bruto_direta_esgoto']] = 50.0
        meses[dez][0].append(r); n1 += 1
    elif mar in meses and meses[mar][1][cab['Qtd_eco_fat_esgoto']] and meses.get((2026, 2), (0, [0] * 20))[1][cab['Qtd_eco_fat_esgoto']] and n2 < 20:
        ws, r = meses[mar]; r2 = copy.copy(r)
        r2[cab['ULTIMO_CONTRATO']] = 'NÃO'; r2[cab['SIT_CONTRATO']] = 'ENCERRADO'
        # zera o contrato SIM no mês (a linha original fica com 0 economias e sem valor)
        for row in ws.iter_rows(min_row=2):
            if row[0].value == m and row[cab['Referência']].value.year == 2026 and row[cab['Referência']].value.month == 3 and row[cab['ULTIMO_CONTRATO']].value == 'SIM':
                row[cab['Qtd_eco_fat_esgoto']].value = 0; row[cab['Faturamento_bruto_direta_esgoto']].value = 0; break
        else: continue
        ws.append(r2); n2 += 1
wb.save(sys.argv[2]); print('dez/25:', n1, 'mar/26:', n2)
