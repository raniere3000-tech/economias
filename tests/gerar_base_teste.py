"""Gera uma base sintética no mesmo formato da exportação (3 abas, dez/25 a set/26).
Uso: python tests/gerar_base_teste.py [quantidade_de_matriculas] [arquivo_saida]
Inclui ligações novas, saídas, contratos anteriores (ULTIMO_CONTRATO = NÃO) e divergências entre
Qtd_eco_fat_esgoto e Eco_total, para exercitar todas as regras do painel."""
import sys
N = int(sys.argv[1]) if len(sys.argv) > 1 else 3000
SAIDA = sys.argv[2] if len(sys.argv) > 2 else 'base_teste.xlsx'
import random, datetime
from openpyxl import Workbook
random.seed(7)
cidades=['CASIMIRO DE ABREU','RIO DAS OSTRAS','MACAÉ','CABO FRIO','SAQUAREMA']
subs=['RESIDENCIAL','COMERCIAL','INDUSTRIAL','PUBLICA']
refs=[datetime.datetime(2025,12,1)]+[datetime.datetime(2026,m,1) for m in range(1,10)]
abas=[('Dezembro 2025 a Fevereiro 2026',refs[0:3]),('Março 2026 a Maio 2026',refs[3:6]),('Junho 2026 a Setembro 2026',refs[6:10])]
hdr=['NUM_LIGACAO','NOM_CLIENTE','CIDADE','SUB_CATEGORIA','SIT_LIG','SIT_CONTRATO','Referência','ULTIMO_CONTRATO','COD_GRUPO','TIPO_FATURAMENTO','Faturamento_bruto_direta_agua','Faturamento_bruto_direta_esgoto','Qtd_eco_fat_agua','Qtd_eco_fat_esgoto','Eco_total']
mats=[]
for i in range(N):
    eco=random.choice([1,1,1,2,3])
    start=random.choice([None]*6+list(range(0,10)))  # mes em que passa a ter esgoto
    stop=random.choice([None]*15+list(range(3,10)))
    mats.append(dict(n=100607450+i,nome=f'CLIENTE {i} & CIA',cid=random.choice(cidades),sub=random.choice(subs),eco=eco,start=start,stop=stop,grp=random.choice([511,512,513]),nasce=random.choice([0]*12+[2,5,7]),div=random.random()<0.03))
wb=Workbook(write_only=True)
for nome,rs in abas:
    ws=wb.create_sheet(nome); ws.append(hdr)
    for m in mats:
        for r in rs:
            k=refs.index(r)
            if k<m['nasce']: continue
            esg = m['start'] is not None and k>=m['start'] and not (m['stop'] is not None and k>=m['stop'])
            fa=round(random.uniform(40,500),2)
            if random.random()<0.02:
                ws.append([m['n'],m['nome'],m['cid'],m['sub'],'ATIVA','ENCERRADO',r,'NÃO',m['grp'],'AGUA',10,0,1,0,1])
            ws.append([m['n'],m['nome'],m['cid'],m['sub'],'ATIVA','ATIVO',r,'SIM',m['grp'],'AGUA E ESGOTO' if esg else 'AGUA',fa,round(fa*0.8,2) if esg else None,m['eco'],(m['eco'] if esg else 0)+(1 if m['div'] and k==9 else 0),m['eco']])
wb.save(SAIDA)
print('ok')
