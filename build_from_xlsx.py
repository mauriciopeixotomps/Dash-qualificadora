"""
Build the Abril 2026 dashboard from the official Pipedrive Insights xlsx exports.
Source of truth: planilhas exportadas (não API).

Inputs (em /uploads, versão mais recente):
- deals-insights-5808773-1936.xlsx  — 4525 negócios (Aberto/Perdido)
- deals-insights-5808773-1941.xlsx  — 4127 perdidos COM Etiqueta + Motivo + Data de perda
- activities-insights-5808773-1935.xlsx — Qualificação Realizada (791)
- activities-insights-5808773-1937.xlsx — Agendamento de CC1 + BOT (601)
- activities-insights-5808773-1938.xlsx — Reagendamento de CC1 (111)
- activities-insights-5808773-1939.xlsx — CC1 Realizada (433)
- activities-insights-5808773-1940.xlsx — CC1 Não Realizada (229)
"""
import pandas as pd, json, os
from datetime import date
from collections import Counter, defaultdict

UPLOADS = '/sessions/youthful-quirky-heisenberg/mnt/uploads'
WORKDIR = '/sessions/youthful-quirky-heisenberg/mnt/Analista de dados - Qualificadora'
TEMPLATE = os.path.join(WORKDIR, 'dashboard_v2.html')
OUT_HTML = os.path.join(WORKDIR, 'Dashboard_Qualificadora_Abril_2026.html')

MONTH_START = date(2026, 4, 1)
MONTH_END = date(2026, 4, 30)
ALL_DAYS = [date(2026, 4, d).isoformat() for d in range(1, 31)]

# SDR list — escolher conforme o mês.
# ABRIL/2026: 10 SDRs (com Jhully Soares)
# MAIO/2026 em diante: 12 SDRs (sem Jhully, com Amada/Camily/Rafael)
SDR_LIST_ABRIL = ['Mélodi Nirvana','Lorenzo Meneghetti','Leonardo Soares','Glauber Bagetti',
                  'Andriely Oliveira','Jhully Soares','Izadora Zacker','Henri Meneghetti',
                  'João Gabriel Dutra','Jonatans Junior']
SDR_LIST_MAIO = ['Mélodi Nirvana','Lorenzo Meneghetti','Leonardo Soares','Glauber Bagetti',
                 'Andriely Oliveira','Izadora Zacker','Henri Meneghetti',
                 'João Gabriel Dutra','Jonatans Junior',
                 'Amada Oliveira','Camily Nadal','Rafael dallarosa longo']

# Este script gera Abril 2026 (mês fechado) — mantém os 10 SDRs originais
SDR_LIST = SDR_LIST_ABRIL
SDR_SET = set(SDR_LIST)

def load(fname):
    return pd.read_excel(os.path.join(UPLOADS, fname), sheet_name=0)

# ============== LOAD ==============
deals    = load('deals-insights-5808773-1936.xlsx')        # 4525 deals
perdidos = load('deals-insights-5808773-1941.xlsx')        # 4127 perdidos detalhados
qual     = load('activities-insights-5808773-1935.xlsx')   # 791
ag       = load('activities-insights-5808773-1937.xlsx')   # 601
reag     = load('activities-insights-5808773-1938.xlsx')   # 111
re_      = load('activities-insights-5808773-1939.xlsx')   # 433 (filtrado, faltam reuniões com closer fora da lista)
ns       = load('activities-insights-5808773-1940.xlsx')   # 229
re_leo   = load('activities-insights-5808773-1942.xlsx')   # 33 reuniões do Leonardo (correção)

# MERGE: 1939 não inclui 2 reuniões do Leonardo onde closer=Bruna Salazar.
# Adicionar do 1942 as que faltam (deduplicadas por email).
re_emails = set(re_['Negócio - E-mail'].fillna('').str.lower().str.strip())
missing = re_leo[~re_leo['Negócio - E-mail'].fillna('').str.lower().str.strip().isin(re_emails)]
if len(missing):
    # Alinhar colunas — adicionar as faltantes com NaN/defaults
    for c in re_.columns:
        if c not in missing.columns: missing[c] = None
    re_ = pd.concat([re_, missing[re_.columns]], ignore_index=True)
    print(f'+ {len(missing)} reuniões do Leonardo adicionadas (closer fora do recorte original)')

print(f'deals={len(deals)} perdidos_detail={len(perdidos)}')
print(f'qual={len(qual)} ag={len(ag)} reag={len(reag)} re={len(re_)} ns={len(ns)}')

# ============== HELPERS ==============
def daily_count(df, datecol):
    out = defaultdict(int)
    for v in df[datecol].dropna():
        d = pd.to_datetime(v).date()
        if MONTH_START <= d <= MONTH_END:
            out[d.isoformat()] += 1
    return dict(out)

def etq_to_product(etq):
    if pd.isna(etq) or not etq: return None
    s = str(etq).lower()
    if 'agro' in s: return 'Studio Agro'
    if 'fiscal' in s: return 'Studio Fiscal'
    if 'pj360' in s or 'pj 360' in s: return 'PJ360'
    if 'partner' in s: return 'GS PARTNER'
    return None

def funil_to_product(f):
    if f == 'TECNOLOGIA': return 'PJ360'
    if f in ('PARTNER','Partners (Raissa)'): return 'GS PARTNER'
    return None  # FRANQUIA → precisa etiqueta

# ============== BUILD MASTER deal → product MAP ==============
# Estratégia:
# 1) De 1941 (perdidos), pegar Título → Etiqueta direto
# 2) Para deals abertos (sem em 1941), buscar etiqueta nas xlsx de atividade
# 3) Para o resto, fallback no funil; se ainda assim None, distribuir FRANQUIA 55/45
title_to_etq = {}
# 1941 — fonte primária
for _, row in perdidos[['Negócio - Título','Negócio - Etiqueta']].dropna(subset=['Negócio - Título']).iterrows():
    title_to_etq[row['Negócio - Título']] = row['Negócio - Etiqueta']

# Reforço com xlsx de atividades que têm etiqueta + título/email do deal
for src in (qual, ag, reag):
    if 'Negócio - Etiqueta' not in src.columns: continue
    title_col = 'Atividade - Negócio' if 'Atividade - Negócio' in src.columns else None
    if not title_col: continue
    for _, row in src[[title_col,'Negócio - Etiqueta']].dropna().iterrows():
        title_to_etq.setdefault(row[title_col], row['Negócio - Etiqueta'])

# Aplicar produto
deals['_etq'] = deals['Negócio - Título'].map(title_to_etq)
deals['_product'] = deals.apply(
    lambda r: etq_to_product(r['_etq']) or funil_to_product(r['Negócio - Funil']),
    axis=1)

# Distribuir FRANQUIA não classificado entre Agro/Fiscal (proporcional ao já classificado)
unclass_franquia = deals[(deals['Negócio - Funil']=='FRANQUIA') & deals['_product'].isna()].index
n_agro_total = ((deals['Negócio - Funil']=='FRANQUIA') & (deals['_product']=='Studio Agro')).sum()
n_fis_total  = ((deals['Negócio - Funil']=='FRANQUIA') & (deals['_product']=='Studio Fiscal')).sum()
ratio_agro = n_agro_total / (n_agro_total + n_fis_total) if (n_agro_total+n_fis_total) else 0.55
n = len(unclass_franquia)
n_agro = round(n * ratio_agro)
deals.loc[unclass_franquia[:n_agro], '_product'] = 'Studio Agro'
deals.loc[unclass_franquia[n_agro:], '_product'] = 'Studio Fiscal'

print('\nProduct distribution:')
print(deals['_product'].value_counts(dropna=False))

# ============== DAILY SERIES ==============
deals['_created'] = pd.to_datetime(deals['Negócio - Negócio criado em'])
# Lost date — agora temos! (1941)
perdidos['_lost'] = pd.to_datetime(perdidos['Negócio - Data de perda'])

qual['_due'] = pd.to_datetime(qual['Atividade - Data de vencimento'])
ag['_due'] = pd.to_datetime(ag['Atividade - Data de vencimento'])
re_['_due'] = pd.to_datetime(re_['Atividade - Data de vencimento'])
ns['_due'] = pd.to_datetime(ns['Atividade - Data de vencimento'])

daily_leads = daily_count(deals, '_created')
daily_lost  = daily_count(perdidos, '_lost')
daily_qual = daily_count(qual, '_due')
daily_ag   = daily_count(ag, '_due')
daily_re   = daily_count(re_, '_due')
daily_ns   = daily_count(ns, '_due')

series = [{'date': d,
           'leads': daily_leads.get(d, 0),
           'lost':  daily_lost.get(d, 0),
           'qual':  daily_qual.get(d, 0),
           'ag':    daily_ag.get(d, 0),
           're':    daily_re.get(d, 0),
           'ns':    daily_ns.get(d, 0)} for d in ALL_DAYS]

# ============== SDR / CLOSER ==============
def sdr_count(series, only_sdr=True):
    c = Counter(series.dropna())
    return {k:v for k,v in c.items() if (not only_sdr) or k in SDR_SET}

sdr_ag  = sdr_count(ag['Atividade - Usuário responsável'])    # AG → quem foi responsável
sdr_re  = sdr_count(re_['Atividade - Criador'])               # RE → quem CRIOU a meeting
sdr_ns  = sdr_count(ns['Atividade - Criador'])
sdr_qual = sdr_count(qual['Atividade - Usuário responsável'])
sdr_reag = sdr_count(reag['Atividade - Criador'])

# Closer view
closer_re = sdr_count(re_['Atividade - Usuário responsável'], only_sdr=False)
closer_ns = sdr_count(ns['Atividade - Usuário responsável'], only_sdr=False)

# Per-day per-SDR breakdown (filter-aware no front)
def daily_by_person(df, datecol, person_col, only_sdr=True):
    out = defaultdict(lambda: defaultdict(int))
    df = df.copy()
    df['_d'] = pd.to_datetime(df[datecol]).dt.date
    for _, row in df.iterrows():
        d = row['_d']
        if d is None or pd.isna(d): continue
        d_iso = d.isoformat()
        if d_iso not in {dd for dd in ALL_DAYS}: continue
        nm = row.get(person_col)
        if pd.isna(nm) or not nm: continue
        nm = str(nm).strip()
        if only_sdr and nm not in SDR_SET: continue
        out[d_iso][nm] += 1
    return {k: dict(v) for k, v in out.items()}

sdr_daily = {
    're': daily_by_person(re_, 'Atividade - Data de vencimento', 'Atividade - Criador'),
    'ag': daily_by_person(ag, 'Atividade - Data de vencimento', 'Atividade - Usuário responsável'),
    'ns': daily_by_person(ns, 'Atividade - Data de vencimento', 'Atividade - Criador'),
    'future': {},  # Abril fechado, sem reuniões futuras
}
closer_daily = {
    're': daily_by_person(re_, 'Atividade - Data de vencimento', 'Atividade - Usuário responsável', only_sdr=False),
    'ns': daily_by_person(ns, 'Atividade - Data de vencimento', 'Atividade - Usuário responsável', only_sdr=False),
    'future': {},
}

# ============== PRODUCT METRICS ==============
def product_for_title(t):
    return etq_to_product(title_to_etq.get(t))

def acts_by_product(df, titlecol):
    if titlecol not in df.columns: return {}
    o = Counter()
    for t in df[titlecol].dropna():
        p = product_for_title(t)
        if p: o[p] += 1
    return dict(o)

# Em alguns xlsx o título do deal vem em "Atividade - Negócio" (email),
# em outros vem em "Negócio - Nome" (nome da pessoa). Tentar ambos.
# Mapa email→produto a partir das xlsx que têm etiqueta (mais confiável)
email_to_product = {}
for src in (perdidos, ag, qual, reag):
    if 'Negócio - Etiqueta' not in src.columns: continue
    if 'Negócio - E-mail' not in src.columns: continue
    for _, row in src[['Negócio - E-mail','Negócio - Etiqueta']].dropna().iterrows():
        e = str(row['Negócio - E-mail']).lower().strip()
        p = etq_to_product(row['Negócio - Etiqueta'])
        if p and e: email_to_product.setdefault(e, p)

# Para 1941 (perdidos) o título também serve como fallback
title_to_product = {}
for _, row in perdidos[['Negócio - Título','Negócio - Etiqueta']].dropna().iterrows():
    p = etq_to_product(row['Negócio - Etiqueta'])
    if p: title_to_product[row['Negócio - Título']] = p

EQUIPE_TO_PRODUCT = {
    'STUDIO AGRO': 'Studio Agro',
    'STUDIO FISCAL': 'Studio Fiscal',
    'PJ360': 'PJ360',
    'PARCERIAS': 'GS PARTNER',
}

def acts_by_product_smart(df):
    o = Counter()
    for _, row in df.iterrows():
        p = None
        # 1) Etiqueta direto se existir
        if 'Negócio - Etiqueta' in df.columns:
            p = etq_to_product(row.get('Negócio - Etiqueta'))
        # 2) Email match
        if not p and 'Negócio - E-mail' in df.columns:
            e = row.get('Negócio - E-mail')
            if pd.notna(e):
                p = email_to_product.get(str(e).lower().strip())
        # 3) Equipe
        if not p and 'Negócio - Equipe' in df.columns:
            p = EQUIPE_TO_PRODUCT.get(str(row.get('Negócio - Equipe','')).strip().upper())
        # 4) Funil
        if not p:
            p = funil_to_product(row.get('Negócio - Funil'))
        if p: o[p] += 1
    return dict(o)

ag_p   = acts_by_product_smart(ag)
re_p   = acts_by_product_smart(re_)
ns_p   = acts_by_product_smart(ns)
qual_p = acts_by_product_smart(qual)
print(f'Produtos: ag={ag_p} re={re_p} ns={ns_p}')

leads_p = dict(deals['_product'].value_counts())
# Perdidos por produto: usar 1941 que tem Etiqueta direto
perdidos['_product'] = perdidos.apply(
    lambda r: etq_to_product(r['Negócio - Etiqueta']) or funil_to_product(r['Negócio - Funil']),
    axis=1)
lost_p = dict(perdidos['_product'].value_counts(dropna=False))
# Distribuir perdidos FRANQUIA não classificado também
# (já caímos no funil → None → vira NaN)

loss_rate = {p: round(lost_p.get(p,0)/leads_p[p]*100, 1)
             for p in leads_p if leads_p[p]}
mql_rate = {p: round(ag_p.get(p,0)/leads_p[p]*100, 1)
            for p in leads_p if leads_p[p] and p in ag_p}

# ============== LOSS REASONS (verdadeiro agora) ==============
reason_product = defaultdict(lambda: defaultdict(int))
for _, row in perdidos[['Negócio - Motivo da perda','_product']].dropna(subset=['Negócio - Motivo da perda']).iterrows():
    r = str(row['Negócio - Motivo da perda']).strip()
    p = row['_product']
    if r and p:
        reason_product[r][p] += 1

# Perdidos por campanha (UTM campaign)
reason_camp = defaultdict(lambda: defaultdict(int))
camp_counter = Counter()
for _, row in perdidos[['Negócio - Motivo da perda','Negócio - UTM campaign']].dropna().iterrows():
    r = str(row['Negócio - Motivo da perda']).strip()
    c = str(row['Negócio - UTM campaign']).strip()
    reason_camp[r][c] += 1
    camp_counter[c] += 1
top_camps = [c for c,_ in camp_counter.most_common(15)]

# ============== CROSS-FUNNEL: Origem (etiqueta) × Funil da reunião ==============
# Argumento: leads de FRANQUIA naturalmente viram reunião de PARTNER (vazamento).
# Queremos ver se PARTNER → FRANQUIA acontece pra balancear, ou se é assimétrico.
def origin_label(etq):
    if not etq or pd.isna(etq): return None
    s = str(etq).lower()
    if 'agro' in s: return 'Studio Agro'
    if 'fiscal' in s: return 'Studio Fiscal'
    if 'pj360' in s or 'pj 360' in s: return 'PJ360'
    if 'partner' in s: return 'Partner'
    return None

# Mapa email → etiqueta original
email_to_etq = {}
for src in (perdidos, ag, qual):
    if 'Negócio - Etiqueta' not in src.columns or 'Negócio - E-mail' not in src.columns: continue
    for _, r in src[['Negócio - E-mail','Negócio - Etiqueta']].dropna().iterrows():
        e = str(r['Negócio - E-mail']).lower().strip()
        if e and e not in email_to_etq:
            email_to_etq[e] = str(r['Negócio - Etiqueta']).strip()

cross_funnel = defaultdict(lambda: defaultdict(int))
for _, row in re_.iterrows():
    email = str(row.get('Negócio - E-mail','') or '').lower().strip()
    if not email: continue
    etq = email_to_etq.get(email)
    origin = origin_label(etq)
    if not origin:
        # Fallback pelo funil atual (apenas para PARTNER e TECNOLOGIA)
        f = row.get('Negócio - Funil')
        if f == 'TECNOLOGIA': origin = 'PJ360'
        elif f == 'PARTNER': origin = 'Partner'
        else: continue
    funil_meeting = row.get('Negócio - Funil') or 'Indef'
    if funil_meeting in ('FRANQUIA','PARTNER','TECNOLOGIA'):
        cross_funnel[origin][funil_meeting] += 1

# Métricas-chave do argumento
cf = {k: dict(v) for k,v in cross_funnel.items()}
fr_origins = ('Studio Agro','Studio Fiscal')
fr_total = sum(sum(cf.get(o,{}).values()) for o in fr_origins)
fr_to_franquia = sum(cf.get(o,{}).get('FRANQUIA',0) for o in fr_origins)
fr_to_partner = sum(cf.get(o,{}).get('PARTNER',0) for o in fr_origins)
pt_total = sum(cf.get('Partner',{}).values())
pt_to_partner = cf.get('Partner',{}).get('PARTNER',0)
pt_to_franquia = cf.get('Partner',{}).get('FRANQUIA',0)

cross_summary = {
    'franquia_origin_total': fr_total,
    'franquia_to_franquia': fr_to_franquia,
    'franquia_to_partner': fr_to_partner,
    'partner_origin_total': pt_total,
    'partner_to_partner': pt_to_partner,
    'partner_to_franquia': pt_to_franquia,
}
print('\nCross-funnel (origem × reunião):')
print(f'  FRANQUIA → FRANQUIA: {fr_to_franquia}, FRANQUIA → PARTNER: {fr_to_partner}')
print(f'  PARTNER → PARTNER: {pt_to_partner}, PARTNER → FRANQUIA: {pt_to_franquia}')

# ============== TOTALS ==============
totals = {
    'leads': len(deals),
    'lost':  int((deals['Negócio - Status']=='Perdido').sum()),
    'qual':  len(qual),
    'ag':    len(ag),
    're':    len(re_),
    'ns':    len(ns),
    'future_re': 0,
}

dashboard = {
    'totals': totals,
    'daily_series': series,
    'sdr': {
        'ag': sdr_ag,
        're': sdr_re,
        'ns': sdr_ns,
        'future': {},
    },
    'sdr_daily': sdr_daily,
    'closer': {
        're': closer_re,
        'future': {},
        'ns': closer_ns,
    },
    'closer_daily': closer_daily,
    'product': {
        'ag': ag_p,
        're': re_p,
        'ns': ns_p,
        'leads': leads_p,
        'lost': lost_p,
        'loss_rate': loss_rate,
        'mql_rate': mql_rate,
    },
    'loss': {
        'by_reason_product': {r: dict(d) for r, d in reason_product.items()},
        'by_reason_campaign': {r: dict(d) for r, d in reason_camp.items()},
        'top_campaigns': top_camps,
    },
    'cross_funnel': cf,
    'cross_summary': cross_summary,
    'meta': {
        'sdr_goal': 70,
        'closer_goal': 40,
        'sdr_pace_target': 70.0,
        'closer_pace_target': 40.0,
        'working_days_elapsed': 22,
        'working_days_total': 22,
        'last_data_date': 'Abril 2026 — fechado, dados oficiais do Pipedrive Insights',
    },
}

print('\nTotais:', totals)
print('\nLeads por produto:', leads_p)
print('Perdidos por produto:', lost_p)
print('Top motivos de perda:')
for r, d in sorted(reason_product.items(), key=lambda x:-sum(x[1].values()))[:10]:
    print(f'  {r}: {sum(d.values())}  →  {dict(d)}')

# Convert numpy types
def to_native(o):
    if isinstance(o, dict): return {str(k): to_native(v) for k, v in o.items()}
    if isinstance(o, list): return [to_native(x) for x in o]
    if hasattr(o, 'item'): return o.item()
    return o
dashboard = to_native(dashboard)

with open(TEMPLATE) as f:
    html = f.read()
html = html.replace('__DATA_PLACEHOLDER__', json.dumps(dashboard, ensure_ascii=False))
html = html.replace('__MONTH_LABEL__', 'Abril 2026')
html = html.replace('__ACTIVE_ABRIL__', 'active')
html = html.replace('__ACTIVE_MAIO__', '')
html = html.replace('__TODAY_ISO__', '2026-04-30')  # Abril fechado — TODAY = último dia útil
with open(OUT_HTML, 'w') as f:
    f.write(html)

# JSON snapshot
with open(os.path.join(WORKDIR, 'dashboard_data.json'), 'w') as f:
    json.dump(dashboard, f, ensure_ascii=False, indent=2)

print(f'\n✓ Dashboard: {OUT_HTML}')
