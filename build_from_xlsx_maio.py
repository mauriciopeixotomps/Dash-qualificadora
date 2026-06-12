"""
Build the Maio 2026 dashboard from Pipedrive Insights xlsx exports.
Mês em andamento — atualizar conforme novas planilhas chegam.

Inputs:
- deals-insights-5808773-1957.xlsx  — 697 negócios (Aberto/Perdido)
- deals-insights-5808773-1960.xlsx  — 418 perdidos detalhados (Etiqueta + Motivo)
- activities-insights-5808773-1958.xlsx — 136 Qualificações Realizadas
- activities-insights-5808773-1959.xlsx — 111 CC1 Agendadas (CC1 + BOT)
- activities-insights-5808773-1962.xlsx — 25 Reagendamentos
- activities-insights-5808773-1963.xlsx — 61 CC1 Realizadas (concluídas)
- activities-insights-5808773-1961.xlsx — 76 CC1 A fazer (agendadas futuras)
- activities-insights-5808773-1964.xlsx — 14 CC1 Não Realizadas
"""
import pandas as pd, json, os
from datetime import date, timedelta, datetime, timezone
from collections import Counter, defaultdict

_BRT = timezone(timedelta(hours=-3))
last_update = datetime.now(_BRT).strftime('%d/%m/%Y %H:%M')

WORKDIR = os.path.dirname(os.path.abspath(__file__))
UPLOADS = os.path.join(os.path.dirname(WORKDIR), 'uploads')
TEMPLATE = os.path.join(WORKDIR, 'dashboard_v2.html')
OUT_HTML = os.path.join(WORKDIR, 'publish', 'maio-2026.html')
METAS = os.path.join(WORKDIR, 'metas_maio_2026.json')

MONTH_START = date(2026, 5, 1)
MONTH_END = date(2026, 5, 31)
TODAY = date.today()  # dinâmico — pega a data real do build
ALL_DAYS = [date(2026, 5, d).isoformat() for d in range(1, 32)]

# Maio: 12 SDRs (sem Jhully, com Amanda/Camily/Rafael)
SDR_LIST = ['Mélodi Nirvana','Lorenzo Meneghetti','Leonardo Soares','Glauber Bagetti',
            'Andriely Oliveira','Izadora Zacker','Henri Meneghetti',
            'João Gabriel Dutra','Jonatans Junior',
            'Amanda Oliveira','Camily Nadal','Rafael dallarosa longo']
SDR_SET = set(SDR_LIST)

# Closer teams (alinhado com os painéis Pipedrive Insights)
# Andresa Caldas é Franquia (confirmação do gestor — apesar de aparecer nos 2 painéis do Pipedrive)
# GS Partner Franqueados — funil próprio (closers Raissas)
CLOSER_GS_FRANQUEADOS = {'Raissa Antunes','Raíssa Nobre'}
CLOSER_PARTNER = {'Maria Luísa','Djulia Silva','Haynnã','Samuel Almeida',
                  'Laura Feijó Junqueira','Manuella Oliveira','Tomaz Fabres',
                  'Tom Mansur','Clayton Bastos'}
CLOSER_FRANQUIA = {'William Dias','Thuany Ghabril','Adalberto Neto','Andresa Caldas',
                   'Émerson Cavitchoni','Suzanny Mauren Dihelem','Leonardo Ribeiro',
                   'Felippe Porcella','Lorenzo Coronel','Samuel Brião',
                   'Fred','Gabriel Meneses'}

with open(METAS) as f:
    metas_cfg = json.load(f)

PRODUTOS_ESCOPO = metas_cfg['produtos_em_escopo']  # ['Studio Agro','Studio Fiscal','GS PARTNER']

def load(f): return pd.read_excel(os.path.join(UPLOADS, f), sheet_name=0)

# Fonte: API Pipedrive (atualizada via fetch_pipedrive.py)
# Caso prefira xlsx manuais, troque o caminho abaixo
DATA_DIR = os.path.join(WORKDIR, 'data_pipedrive')
def load_pd(name): return pd.read_excel(os.path.join(DATA_DIR, name))

_deals_all   = load_pd('deals.xlsx')
_perd_all    = load_pd('perdidos.xlsx')

# Filtra deals e perdidos ao mês específico (deals.xlsx agora cobre Maio+Junho)
def _filter_month(df, date_col, m_start, m_end):
    s = pd.to_datetime(df[date_col], errors='coerce')
    # Usa .dt.date para comparar apenas a data, ignorando a hora — inclui o dia inteiro de m_end
    return df[(s.dt.date >= m_start) & (s.dt.date <= m_end)].copy()

deals        = _filter_month(_deals_all, 'Negócio - Negócio criado em', MONTH_START, MONTH_END)
# Perdidos: só negócios CRIADOS e PERDIDOS dentro do mês
def _filter_criado_e_perdido(df, m_start, m_end):
    cri  = pd.to_datetime(df['Negócio - Negócio criado em'], errors='coerce').dt.date
    perd = pd.to_datetime(df['Negócio - Data de perda'], errors='coerce').dt.date
    mask = (cri >= m_start) & (cri <= m_end) & (perd >= m_start) & (perd <= m_end)
    return df[mask].copy()

if 'Negócio - Data de perda' in _perd_all.columns and 'Negócio - Negócio criado em' in _perd_all.columns:
    perdidos_raw = _filter_criado_e_perdido(_perd_all, MONTH_START, MONTH_END)
elif 'Negócio - Data de perda' in _perd_all.columns:
    perdidos_raw = _filter_month(_perd_all, 'Negócio - Data de perda', MONTH_START, MONTH_END)
else:
    perdidos_raw = _perd_all.copy()
qual         = load_pd('qual.xlsx')
ag           = load_pd('ag.xlsx')
reag         = load_pd('reag.xlsx')
re_          = load_pd('re.xlsx')
re_fut       = load_pd('re_fut.xlsx')
ns           = load_pd('ns.xlsx')
inv_path = os.path.join(DATA_DIR, 'inv.xlsx')
inv          = pd.read_excel(inv_path) if os.path.exists(inv_path) else pd.DataFrame()

# Pipedrive filtra perdidos por motivos "reais" (exclui Teste, NaN, Postergou, etc.)
# Lista alinhada com o painel de Pipedrive (Perdidos por motivo)
MOTIVOS_VALIDOS = {'Sem perfil','Sem interesse','Sem Retorno','Sem Contato','Perfil Fake','Buscando Parceria'}
if 'Negócio - Motivo da perda' in perdidos_raw.columns:
    perdidos = perdidos_raw[perdidos_raw['Negócio - Motivo da perda'].isin(MOTIVOS_VALIDOS)].copy()
else:
    perdidos = perdidos_raw.copy()  # xlsx do painel já filtrado

print(f'deals={len(deals)} perdidos_raw={len(perdidos_raw)} perdidos_validos={len(perdidos)}')
print(f'qual={len(qual)} ag={len(ag)} reag={len(reag)} re={len(re_)} re_fut={len(re_fut)} ns={len(ns)} inv={len(inv)}')

# ============== HELPERS ==============
def daily_count(df, datecol):
    out = defaultdict(int)
    for v in df[datecol].dropna():
        d = pd.to_datetime(v).date()
        if MONTH_START <= d <= MONTH_END:
            out[d.isoformat()] += 1
    return dict(out)

def etq_to_prod(e):
    if pd.isna(e): return None
    s = str(e).lower()
    if 'agro' in s: return 'Studio Agro'
    if 'fiscal' in s: return 'Studio Fiscal'
    if 'pj360' in s or 'pj 360' in s: return 'PJ360'
    if 'partner' in s: return 'GS PARTNER'
    return None

EQUIPE_TO_PRODUCT = {
    'STUDIO AGRO':'Studio Agro',
    'STUDIO FISCAL':'Studio Fiscal',
    'PJ360':'PJ360',
    'PARCERIAS':'GS PARTNER',
    'GS PARTNERS':'GS PARTNER',
}

# Build email→produto a partir de TODOS os arquivos que têm etiqueta
email2prod = {}
for src in (perdidos, ag, qual, reag):
    if 'Negócio - Etiqueta' not in src.columns: continue
    for col_email in ('Negócio - E-mail','Pessoa - E-mail'):
        if col_email not in src.columns: continue
        for _, r in src[[col_email,'Negócio - Etiqueta']].dropna().iterrows():
            e = str(r[col_email]).lower().strip()
            p = etq_to_prod(r['Negócio - Etiqueta'])
            if p and e: email2prod.setdefault(e, p)

# Também construir mapa por Nome do Negócio (caso e-mail esteja vazio nas realizadas)
name2prod = {}
for src in (perdidos, ag, qual, reag):
    if 'Negócio - Etiqueta' not in src.columns: continue
    for col_name in ('Negócio - Nome','Negócio - Título'):
        if col_name not in src.columns: continue
        for _, r in src[[col_name,'Negócio - Etiqueta']].dropna().iterrows():
            n = str(r[col_name]).strip().lower()
            p = etq_to_prod(r['Negócio - Etiqueta'])
            if p and n: name2prod.setdefault(n, p)

print(f'email2prod: {len(email2prod)} entries | name2prod: {len(name2prod)} entries')

def assunto_to_prod(a):
    """Extrai produto do título/assunto da atividade — sinal mais forte que Equipe.
    Ex: 'Studio Agro - José', 'Reunião Studio Fiscal - Marcos', 'GS Partner - Reztech'.
    """
    if not a or pd.isna(a): return None
    s = str(a).lower()
    if 'agro' in s: return 'Studio Agro'
    if 'fiscal' in s: return 'Studio Fiscal'
    if 'pj360' in s or 'pj 360' in s: return 'PJ360'
    if 'partner' in s or 'parceir' in s or 'parceri' in s: return 'GS PARTNER'
    return None

def row_prod(row, df_cols, key='ag'):
    """Classifica produto da row. Regra calibrada vs Pipedrive:
    1) Negócio - Etiqueta (lead source — quando df tem)
    2) Equipe STUDIO AGRO / STUDIO FISCAL é definitiva (não sobrescreve)
    3) Equipe PARCERIAS / GS Partners: Assunto sobrescreve se disser Agro/Fiscal explicitamente
       (caso em que o deal foi roteado pelo time de Partners mas a reunião é de outro produto)
    4) Equipe QUALIFICADORA ou vazia: usa Assunto, depois email lookup, depois Funil
    """
    if 'Negócio - Etiqueta' in df_cols:
        p = etq_to_prod(row.get('Negócio - Etiqueta'))
        if p: return p
    a = assunto_to_prod(row.get('Atividade - Assunto')) if 'Atividade - Assunto' in df_cols else None
    if 'Negócio - Equipe' in df_cols:
        eq = str(row.get('Negócio - Equipe','')).strip().upper()
        if eq == 'STUDIO AGRO': return 'Studio Agro'
        if eq == 'STUDIO FISCAL': return 'Studio Fiscal'
        if eq in ('PARCERIAS','GS PARTNERS'):
            # Override se o assunto disser Agro/Fiscal explicitamente
            if a in ('Studio Agro','Studio Fiscal'): return a
            return 'GS PARTNER'
        if eq == 'PJ360': return 'PJ360'
    # Equipe QUALIFICADORA ou vazia → usa Assunto
    if a: return a
    for col_email in ('Negócio - E-mail','Pessoa - E-mail'):
        if col_email in df_cols:
            e = row.get(col_email)
            if pd.notna(e):
                p = email2prod.get(str(e).lower().strip())
                if p: return p
    for col_name in ('Negócio - Nome','Negócio - Título'):
        if col_name in df_cols:
            n = row.get(col_name)
            if pd.notna(n):
                p = name2prod.get(str(n).strip().lower())
                if p: return p
    f = row.get('Negócio - Funil')
    if f == 'TECNOLOGIA': return 'PJ360'
    if f in ('PARTNER','Partners (Raissa)'): return 'GS PARTNER'
    return None

def row_prod_with_franquia_partner(row, df_cols, key='ag'):
    """Classifica por TIME DO CLOSER (quem realizou a reunião) — alinhado ao relatório
    'REALIZADAS FRANQUIA' do Pipedrive (filtra por Usuário responsável).
    - Responsável = closer Franquia → produto pela etiqueta do lead
      (Agro→Studio Agro, Fiscal→Studio Fiscal, Partner/Raissa→Franquia Partner)
    - Responsável = closer Partner → GS PARTNER
    - Responsável não-closer (ex: SDR em agendamento) → fallback por etiqueta (row_prod)
    Garante que Total Franquia = soma das CC1 dos closers do time Franquia.
    """
    responsavel = str(row.get('Atividade - Usuário responsável', '')).strip()
    etq = str(row.get('Negócio - Etiqueta', '') or '').lower() if 'Negócio - Etiqueta' in df_cols else ''
    if responsavel in CLOSER_GS_FRANQUEADOS:
        return 'GS Partner Franqueados'
    if responsavel in CLOSER_FRANQUIA:
        if 'agro' in etq:   return 'Studio Agro'
        if 'fiscal' in etq: return 'Studio Fiscal'
        if 'partner' in etq or 'raissa' in etq: return 'Franquia Partner'
        # sem etiqueta clara, mas closer Franquia: usa assunto/funil; senão fica em Franquia Partner
        p = row_prod(row, df_cols, key)
        return p if p in ('Studio Agro', 'Studio Fiscal') else 'Franquia Partner'
    if responsavel in CLOSER_PARTNER:
        return 'GS PARTNER'
    return row_prod(row, df_cols, key)

# ============== DEALS — produto ==============
title_to_etq = {}
if 'Negócio - Etiqueta' in perdidos.columns and 'Negócio - Título' in perdidos.columns:
    for _, row in perdidos[['Negócio - Título','Negócio - Etiqueta']].dropna(subset=['Negócio - Título']).iterrows():
        title_to_etq[row['Negócio - Título']] = row['Negócio - Etiqueta']

deals['_etq'] = deals['Negócio - Título'].map(title_to_etq) if 'Negócio - Título' in deals.columns else None
def deal_prod(r):
    p = etq_to_prod(r['_etq'])
    if p: return p
    # Equipe é mais confiável que Funil para FRANQUIA
    if 'Negócio - Equipe' in deals.columns:
        eq = str(r.get('Negócio - Equipe','')).strip().upper()
        p = EQUIPE_TO_PRODUCT.get(eq)
        if p: return p
    f = r['Negócio - Funil']
    if f == 'TECNOLOGIA': return 'PJ360'
    if f in ('PARTNER','Partners (Raissa)'): return 'GS PARTNER'
    return None
deals['_product'] = deals.apply(deal_prod, axis=1)

print('\nProduct distribution:')
print(deals['_product'].value_counts(dropna=False))

# ============== DAILY SERIES ==============
deals['_created'] = pd.to_datetime(deals['Negócio - Negócio criado em'])
# Perdidos: usa Data de perda se preenchida, senão fallback p/ data de criação
def _pick_lost_date(p):
    for col in ('Negócio - Data de perda','Negócio - Data de fechamento esperada','Negócio - Negócio criado em'):
        if col in p.columns:
            s = pd.to_datetime(p[col], errors='coerce')
            if s.notna().sum() > 0: return s
    return pd.Series([pd.NaT]*len(p))
perdidos['_lost'] = _pick_lost_date(perdidos)
# Qual/AG/RE/NS usam "Marcado como feito em" (espelha filtro Pipedrive Insights)
# RE futuro usa "Data de vencimento" (painel CC1 Para Acontecer usa due_date)
_mcf = 'Atividade - Marcado como feito em'
_due = 'Atividade - Data de vencimento'
qual['_due'] = pd.to_datetime(qual[_mcf] if _mcf in qual.columns else qual[_due])
ag['_due']   = pd.to_datetime(ag[_mcf]   if _mcf in ag.columns   else ag[_due])
re_['_due']  = pd.to_datetime(re_[_mcf]  if _mcf in re_.columns  else re_[_due])
ns['_due']   = pd.to_datetime(ns[_mcf]   if _mcf in ns.columns   else ns[_due])
if not inv.empty and _mcf in inv.columns:
    inv['_due'] = pd.to_datetime(inv[_mcf], errors='coerce')
elif not inv.empty and _due in inv.columns:
    inv['_due'] = pd.to_datetime(inv[_due], errors='coerce')

# Filtra atividades ao mês corrente (fetch puxa Maio+Junho; cada dash conta só seu mês)
def _filter_mes(df):
    if df.empty or '_due' not in df.columns: return df
    d = df['_due'].dt.date
    return df[(d >= MONTH_START) & (d <= MONTH_END)].copy()

qual = _filter_mes(qual)
ag   = _filter_mes(ag)
re_  = _filter_mes(re_)
ns   = _filter_mes(ns)
inv  = _filter_mes(inv)

# re_fut usa Data de vencimento (atividades futuras) — filtra ao mês corrente
if not re_fut.empty and 'Atividade - Data de vencimento' in re_fut.columns:
    _rf = pd.to_datetime(re_fut['Atividade - Data de vencimento'], errors='coerce').dt.date
    re_fut = re_fut[(_rf >= MONTH_START) & (_rf <= MONTH_END)].copy()

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

# Per-day per-product breakdown (para filtro de período no dashboard)
def daily_product(df, datecol, key='ag', use_franquia_partner=False):
    out = {d: {p: 0 for p in ['Studio Agro','Studio Fiscal','Franquia Partner','PJ360','GS PARTNER']} for d in ALL_DAYS}
    df = df.copy()
    df['_d'] = pd.to_datetime(df[datecol]).dt.date
    classifier = row_prod_with_franquia_partner if use_franquia_partner else row_prod
    for _, row in df.iterrows():
        d = row['_d']
        if d is None or pd.isna(d): continue
        d_iso = d.isoformat()
        if d_iso not in out: continue
        p = classifier(row, df.columns, key)
        if p: out[d_iso][p] = out[d_iso].get(p, 0) + 1
    return out

# Para CC1 Realizada e Não Realizada: aplicar a classificação que separa "Franquia Partner"
# Usa _due (Marcado como feito em) para consistência com daily_series
# re_fut continua usando Data de vencimento (atividades futuras ainda não concluídas)
daily_ag_p = daily_product(ag, '_due', 'ag', use_franquia_partner=True)
daily_re_p = daily_product(re_, '_due', 're', use_franquia_partner=True)
daily_re_fut_p = daily_product(re_fut, 'Atividade - Data de vencimento', 're_fut', use_franquia_partner=True)
daily_ns_p = daily_product(ns, '_due', 'ns', use_franquia_partner=True)

# ============== SDR / CLOSER ==============
def sdr_count(series, only_sdr=True):
    c = Counter(series.dropna())
    return {k:v for k,v in c.items() if (not only_sdr) or k in SDR_SET}

def closer_count(series):
    """Conta apenas closers reais — exclui SDRs (que às vezes aparecem como responsável)."""
    c = Counter(series.dropna())
    return {k: v for k, v in c.items() if k not in SDR_SET}

sdr_ag  = sdr_count(ag['Atividade - Usuário responsável'])
sdr_re  = sdr_count(re_['Atividade - Criador'])
sdr_ns  = sdr_count(ns['Atividade - Criador'])
sdr_future = sdr_count(re_fut['Atividade - Criador']) if 'Atividade - Criador' in re_fut.columns else {}
sdr_inv = sdr_count(inv['Atividade - Criador']) if not inv.empty and 'Atividade - Criador' in inv.columns else {}
closer_re = closer_count(re_['Atividade - Usuário responsável'])
closer_ns = closer_count(ns['Atividade - Usuário responsável'])
closer_future = closer_count(re_fut['Atividade - Usuário responsável']) if 'Atividade - Usuário responsável' in re_fut.columns else {}
closer_inv = closer_count(inv['Atividade - Usuário responsável']) if not inv.empty and 'Atividade - Usuário responsável' in inv.columns else {}

# Per-day per-SDR breakdown (filter-aware no front)
def daily_by_person(df, datecol, person_col, only_sdr=True):
    out = defaultdict(lambda: defaultdict(int))  # {date: {name: count}}
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
    'future': daily_by_person(re_fut, 'Atividade - Data de vencimento', 'Atividade - Criador') if 'Atividade - Criador' in re_fut.columns else {},
}
def daily_by_closer(df, datecol, person_col):
    """Igual ao daily_by_person mas só inclui não-SDRs (ou seja, closers de fato)."""
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
        if nm in SDR_SET: continue  # exclui SDRs do ranking de closer
        out[d_iso][nm] += 1
    return {k: dict(v) for k, v in out.items()}

closer_daily = {
    're': daily_by_closer(re_, 'Atividade - Data de vencimento', 'Atividade - Usuário responsável'),
    'ns': daily_by_closer(ns, 'Atividade - Data de vencimento', 'Atividade - Usuário responsável'),
    'future': daily_by_closer(re_fut, 'Atividade - Data de vencimento', 'Atividade - Usuário responsável') if 'Atividade - Usuário responsável' in re_fut.columns else {},
}

# ============== PRODUCT METRICS ==============
def acts_by_product_smart(df, key='ag', use_franquia_partner=False):
    o = Counter()
    classifier = row_prod_with_franquia_partner if use_franquia_partner else row_prod
    for _, row in df.iterrows():
        p = classifier(row, df.columns, key)
        if p: o[p] += 1
    return dict(o)

ag_p = acts_by_product_smart(ag, 'ag', use_franquia_partner=True)
re_p = acts_by_product_smart(re_, 're', use_franquia_partner=True)
ns_p = acts_by_product_smart(ns, 'ns', use_franquia_partner=True)

leads_p = dict(deals['_product'].value_counts())

# Leads por funil (100% confiável, soma o total de leads do mês)
_FUNIL_NOME = {'FRANQUIA':'Franquia','PARTNER':'GS Partner',
               'Partners GS- FRANQUEADOS':'GS Partner Franqueados',
               'Partner - Franqueados':'GS Partner Franqueados'}
def _map_funil(fun):
    return _FUNIL_NOME.get(str(fun).strip(), str(fun))

leads_por_funil = {}
if 'Negócio - Funil' in deals.columns:
    for fun, cnt in deals['Negócio - Funil'].value_counts().items():
        leads_por_funil[_map_funil(fun)] = leads_por_funil.get(_map_funil(fun), 0) + int(cnt)

ag_por_funil = {}
if 'Negócio - Funil' in ag.columns:
    for fun, cnt in ag['Negócio - Funil'].value_counts().items():
        ag_por_funil[_map_funil(fun)] = ag_por_funil.get(_map_funil(fun), 0) + int(cnt)

perdidos_por_funil = {}
if 'Negócio - Funil' in perdidos.columns:
    for fun, cnt in perdidos['Negócio - Funil'].value_counts().items():
        perdidos_por_funil[_map_funil(fun)] = perdidos_por_funil.get(_map_funil(fun), 0) + int(cnt)
taxa_perda_funil = {k: round(perdidos_por_funil.get(k,0)/leads_por_funil[k]*100, 1)
                    for k in leads_por_funil if leads_por_funil[k]}

import unicodedata as _ud
def _norm_prof(p):
    if pd.isna(p): return None
    s = str(p).strip()
    if not s: return None
    s = ''.join(c for c in _ud.normalize('NFD', s) if _ud.category(c) != 'Mn')  # tira acentos p/ juntar variantes
    return s[:1].upper() + s[1:].lower()
profissao_por_funil = {}
if 'Negócio - Profissão' in deals.columns and 'Negócio - Funil' in deals.columns:
    tmp = defaultdict(lambda: defaultdict(int))
    for _, r in deals[['Negócio - Funil','Negócio - Profissão']].iterrows():
        prof = _norm_prof(r['Negócio - Profissão'])
        if not prof: continue
        tmp[_map_funil(r['Negócio - Funil'])][prof] += 1
    for fun, profs in tmp.items():
        profissao_por_funil[fun] = sorted(profs.items(), key=lambda x:-x[1])[:8]

# ── RESUMO SEMANAL: leads (por funil) · AG · RE · perdidos ──────────────────
import datetime as _dtm
def _week_monday(d): return d - _dtm.timedelta(days=d.weekday())
_wk = defaultdict(lambda: {'leads': defaultdict(int), 'ag':0, 're':0, 'perdidos':0})
# Leads por funil (data de criação)
_cri = pd.to_datetime(deals['Negócio - Negócio criado em'], errors='coerce')
for _d, _fun in zip(_cri.dt.date, deals['Negócio - Funil'] if 'Negócio - Funil' in deals.columns else [None]*len(deals)):
    if pd.isna(_d): continue
    if not (MONTH_START <= _d <= MONTH_END): continue
    _wk[_week_monday(_d)]['leads'][_map_funil(_fun)] += 1
# AG e RE (data marcada = _due, já filtrado ao mês)
for _col_df, _key in ((ag,'ag'), (re_,'re')):
    if '_due' in _col_df.columns:
        for _d in pd.to_datetime(_col_df['_due'], errors='coerce').dt.date.dropna():
            if MONTH_START <= _d <= MONTH_END: _wk[_week_monday(_d)][_key] += 1
# Perdidos (data de perda)
if 'Negócio - Data de perda' in perdidos.columns:
    for _d in pd.to_datetime(perdidos['Negócio - Data de perda'], errors='coerce').dt.date.dropna():
        if MONTH_START <= _d <= MONTH_END: _wk[_week_monday(_d)]['perdidos'] += 1
weekly_resumo = []
for i, _k in enumerate(sorted(_wk), 1):
    _ini = max(_k, MONTH_START); _fim = min(_k + _dtm.timedelta(days=6), MONTH_END)
    _v = _wk[_k]
    weekly_resumo.append({
        'semana': f'Sem {i}',
        'periodo': f"{_ini.strftime('%d/%m')}–{_fim.strftime('%d/%m')}",
        'leads_por_funil': dict(_v['leads']),
        'leads_total': sum(_v['leads'].values()),
        'ag': _v['ag'], 're': _v['re'], 'perdidos': _v['perdidos'],
    })
def _perdido_prod(r):
    etq = r.get('Negócio - Etiqueta') if 'Negócio - Etiqueta' in perdidos.columns else None
    p = etq_to_prod(etq) if etq is not None else None
    if p: return p
    f = r.get('Negócio - Funil')
    if f == 'TECNOLOGIA': return 'PJ360'
    if f in ('PARTNER','Partners (Raissa)'): return 'GS PARTNER'
    return None
perdidos['_product'] = perdidos.apply(_perdido_prod, axis=1)
lost_p = dict(perdidos['_product'].value_counts(dropna=False))

loss_rate = {p: round(lost_p.get(p,0)/leads_p[p]*100, 1) for p in leads_p if leads_p[p]}
mql_rate = {p: round(ag_p.get(p,0)/leads_p[p]*100, 1) for p in leads_p if leads_p[p] and p in ag_p}

# ============== LOSS REASONS ==============
reason_product = defaultdict(lambda: defaultdict(int))
if 'Negócio - Motivo da perda' in perdidos.columns:
    for _, row in perdidos[['Negócio - Motivo da perda','_product']].dropna(subset=['Negócio - Motivo da perda']).iterrows():
        r = str(row['Negócio - Motivo da perda']).strip()
        p = row['_product']
        if r and p: reason_product[r][p] += 1

reason_camp = defaultdict(lambda: defaultdict(int))
camp_counter = Counter()
if 'Negócio - UTM campaign' in perdidos.columns and 'Negócio - Motivo da perda' in perdidos.columns:
    for _, row in perdidos[['Negócio - Motivo da perda','Negócio - UTM campaign']].dropna().iterrows():
        r = str(row['Negócio - Motivo da perda']).strip()
        c = str(row['Negócio - UTM campaign']).strip()
        reason_camp[r][c] += 1
        camp_counter[c] += 1
top_camps = [c for c,_ in camp_counter.most_common(15)]

# ============== % PERDA POR CAMPANHA + % AGENDADOS ==============
# Para cada UTM campaign: total de leads, perdidos, agendados, % perda, % AG, saldo
camp_leads = Counter()
if 'Negócio - UTM campaign' in deals.columns:
    for v in deals['Negócio - UTM campaign'].dropna():
        camp_leads[str(v).strip()] += 1

camp_lost = Counter()
if 'Negócio - UTM campaign' in perdidos.columns:
    for v in perdidos['Negócio - UTM campaign'].dropna():
        camp_lost[str(v).strip()] += 1

# Agendados por campanha — precisa de UTM campaign na df de agendamentos
camp_ag = Counter()
if 'Negócio - UTM campaign' in ag.columns:
    for v in ag['Negócio - UTM campaign'].dropna():
        camp_ag[str(v).strip()] += 1

# Motivo predominante por campanha (do topo)
camp_top_motivo = {}
camp_motivo_map = defaultdict(Counter)
if 'Negócio - UTM campaign' in perdidos.columns:
    for _, row in perdidos[['Negócio - Motivo da perda','Negócio - UTM campaign']].dropna().iterrows():
        c = str(row['Negócio - UTM campaign']).strip()
        m = str(row['Negócio - Motivo da perda']).strip()
        camp_motivo_map[c][m] += 1
    for c, mots in camp_motivo_map.items():
        camp_top_motivo[c] = mots.most_common(1)[0][0]

# Lista final de campanhas (com volume mínimo de 5 leads pra filtrar ruído)
campaign_stats = []
for c, leads_n in camp_leads.items():
    if leads_n < 5: continue
    lost_n = camp_lost.get(c, 0)
    ag_n = camp_ag.get(c, 0)
    loss_pct = round(lost_n / leads_n * 100, 1) if leads_n else 0
    ag_pct = round(ag_n / leads_n * 100, 1) if leads_n else 0
    saldo = round(ag_pct - loss_pct, 1)  # positivo = campanha boa, negativo = ruim
    campaign_stats.append({
        'campaign': c,
        'leads': leads_n,
        'lost': lost_n,
        'ag': ag_n,
        'open': leads_n - lost_n,
        'loss_pct': loss_pct,
        'ag_pct': ag_pct,
        'saldo': saldo,
        'top_motivo': camp_top_motivo.get(c, '—'),
    })

# ============== % PERDA POR MOTIVO ==============
# Para cada motivo: count, % do total de perdidos, % do total de leads
total_leads = len(deals)
total_perdidos = len(perdidos)
motivo_stats = []
motivo_counter = Counter()
if 'Negócio - Motivo da perda' in perdidos.columns:
    for v in perdidos['Negócio - Motivo da perda'].dropna():
        motivo_counter[str(v).strip()] += 1

for m, n in motivo_counter.items():
    pct_leads = round(n / total_leads * 100, 1) if total_leads else 0
    pct_perdidos = round(n / total_perdidos * 100, 1) if total_perdidos else 0
    motivo_stats.append({
        'motivo': m,
        'count': n,
        'pct_total_leads': pct_leads,    # % do total de leads gerados
        'pct_perdidos': pct_perdidos,    # % do total dos perdidos
    })

# ============== CROSS-FUNNEL: Origem (etiqueta) × Funil da reunião ==============
def origin_label(etq):
    if not etq or pd.isna(etq): return None
    s = str(etq).lower()
    if 'agro' in s: return 'Studio Agro'
    if 'fiscal' in s: return 'Studio Fiscal'
    if 'pj360' in s or 'pj 360' in s: return 'PJ360'
    if 'partner' in s: return 'Partner'
    return None

email_to_etq = {}
for src in (perdidos, ag, qual, reag):
    if 'Negócio - Etiqueta' not in src.columns or 'Negócio - E-mail' not in src.columns: continue
    for _, r in src[['Negócio - E-mail','Negócio - Etiqueta']].dropna().iterrows():
        e = str(r['Negócio - E-mail']).lower().strip()
        if e and e not in email_to_etq:
            email_to_etq[e] = str(r['Negócio - Etiqueta']).strip()

cross_funnel = defaultdict(lambda: defaultdict(int))
re_cols = set(re_.columns)
for _, row in re_.iterrows():
    # Origin: etiqueta direta da row (mesma fonte do card Franquia Partner)
    etq = row.get('Negócio - Etiqueta') if 'Negócio - Etiqueta' in re_cols else None
    if not etq:
        email = str(row.get('Negócio - E-mail','') or '').lower().strip()
        etq = email_to_etq.get(email) if email else None
    origin = origin_label(etq)
    if not origin:
        f = row.get('Negócio - Funil')
        if f == 'TECNOLOGIA': origin = 'PJ360'
        else: continue  # sem etiqueta e não é PJ360: ignora (funil Partner pode mudar)
    # Time da reunião: time do closer (não funil, que pode mudar)
    responsavel = str(row.get('Atividade - Usuário responsável', '')).strip()
    if responsavel in CLOSER_FRANQUIA:
        team_meeting = 'FRANQUIA'
    elif responsavel in CLOSER_PARTNER:
        team_meeting = 'PARTNER'
    else:
        continue
    cross_funnel[origin][team_meeting] += 1

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
print(f'\nCross-funnel: FR→FR={fr_to_franquia} FR→PT={fr_to_partner} | PT→PT={pt_to_partner} PT→FR={pt_to_franquia}')

# ============== TOTALS (filtrados para o mês corrente — consistência com daily_series) ==============
def _may(df, datecol='_due'):
    """Retorna contagem de linhas com datecol dentro de MONTH_START..MONTH_END."""
    if df.empty or datecol not in df.columns: return 0
    s = pd.to_datetime(df[datecol], errors='coerce').dt.date
    return int(((s >= MONTH_START) & (s <= MONTH_END)).sum())

totals = {
    'leads': max(len(deals), len(perdidos)),  # negócios: contagem total (Pipedrive padrão)
    'lost':  len(perdidos),
    'qual':  _may(qual),
    'ag':    _may(ag),
    're':    _may(re_),
    'ns':    _may(ns),
    'future_re': len(re_fut),
    'inv':   _may(inv, '_due') if not inv.empty and '_due' in inv.columns else len(inv),
}

# ============== PACE (working days) ==============
HOLIDAYS_MAIO = {date(2026,5,1)}  # Dia do Trabalho
def count_workdays(start, end):
    n = 0; cur = start
    while cur <= end:
        if cur.weekday() < 5 and cur not in HOLIDAYS_MAIO:
            n += 1
        cur += timedelta(days=1)
    return n
elapsed = count_workdays(MONTH_START, min(TODAY, MONTH_END))
total_workdays = count_workdays(MONTH_START, MONTH_END)

# ============== METAS ==============
metas_p = metas_cfg['metas_por_produto']
meta_time = metas_cfg['meta_time']

# Pace (linear): meta esperada hoje
pace_target_re_total = round(meta_time['realizadas'] * elapsed / total_workdays, 1)
pace_target_ag_total = round(meta_time['agendadas'] * elapsed / total_workdays, 1)

# Realizado escopo (excluindo PJ360)
re_em_escopo = sum(re_p.get(p,0) for p in PRODUTOS_ESCOPO)
ag_em_escopo = sum(ag_p.get(p,0) for p in PRODUTOS_ESCOPO)

print(f'\nPace ({elapsed}/{total_workdays} dias úteis):')
print(f'  RE total escopo: {re_em_escopo}/{meta_time["realizadas"]} (pace alvo: {pace_target_re_total})')
print(f'  AG total escopo: {ag_em_escopo}/{meta_time["agendadas"]} (pace alvo: {pace_target_ag_total})')
for p in PRODUTOS_ESCOPO:
    re_v = re_p.get(p, 0)
    ag_v = ag_p.get(p, 0)
    meta = metas_p[p]
    print(f'  {p}: RE {re_v}/{meta["realizadas"]} · AG {ag_v}/{meta["agendadas"]}')

# Lista de dias úteis (para cálculo de pace por período no front)
working_days_iso = []
cur = MONTH_START
while cur <= MONTH_END:
    if cur.weekday() < 5 and cur not in HOLIDAYS_MAIO:
        working_days_iso.append(cur.isoformat())
    cur += timedelta(days=1)

dashboard = {
    'totals': totals,
    'daily_series': series,
    'daily_product': {'ag': daily_ag_p, 're': daily_re_p, 're_fut': daily_re_fut_p, 'ns': daily_ns_p},
    'working_days': working_days_iso,
    'sdr': {'ag': sdr_ag, 're': sdr_re, 'ns': sdr_ns, 'future': sdr_future, 'inv': sdr_inv},
    'sdr_daily': sdr_daily,
    'closer': {'re': closer_re, 'future': closer_future, 'ns': closer_ns, 'inv': closer_inv},
    'closer_team': {'franquia': sorted(CLOSER_FRANQUIA), 'partner': sorted(CLOSER_PARTNER)},
    'closer_daily': closer_daily,
    'weekly_resumo': weekly_resumo,
    'product': {
        'ag': ag_p, 're': re_p, 'ns': ns_p,
        'leads': leads_p, 'lost': lost_p,
        'loss_rate': loss_rate, 'mql_rate': mql_rate,
        'leads_por_funil': leads_por_funil,
        'ag_por_funil': ag_por_funil,
        'perdidos_por_funil': perdidos_por_funil,
        'taxa_perda_funil': taxa_perda_funil,
        'profissao_por_funil': profissao_por_funil,
    },
    'loss': {
        'by_reason_product': {r: dict(d) for r, d in reason_product.items()},
        'by_reason_campaign': {r: dict(d) for r, d in reason_camp.items()},
        'top_campaigns': top_camps,
        'campaign_stats': campaign_stats,
        'motivo_stats': motivo_stats,
        'total_leads': total_leads,
        'total_perdidos': total_perdidos,
    },
    'cross_funnel': cf,
    'cross_summary': cross_summary,
    'meta': {
        'sdr_goal': metas_cfg['meta_individual']['sdr_realizadas_meta'],
        'closer_goal': metas_cfg['meta_individual']['closer_realizadas_meta'],
        'sdr_pace_target': round(70 * elapsed / total_workdays, 1),
        'closer_pace_target': round(40 * elapsed / total_workdays, 1),
        'working_days_elapsed': elapsed,
        'working_days_total': total_workdays,
        'last_data_date': f'Maio 2026 — em andamento, dia útil {elapsed} de {total_workdays}',
        'last_update': last_update,
        'team_total_re_goal': meta_time['realizadas'],
        'team_total_ag_goal': meta_time['agendadas'],
        'team_pace_target_re': pace_target_re_total,
        'team_pace_target_ag': pace_target_ag_total,
        'metas_por_produto': metas_p,
        'produtos_em_escopo': PRODUTOS_ESCOPO,
    },
}

def to_native(o):
    if isinstance(o, dict): return {str(k): to_native(v) for k, v in o.items()}
    if isinstance(o, list): return [to_native(x) for x in o]
    if hasattr(o, 'item'): return o.item()
    return o
dashboard = to_native(dashboard)

with open(TEMPLATE) as f:
    html = f.read()
html = html.replace('__DATA_PLACEHOLDER__', json.dumps(dashboard, ensure_ascii=False))
html = html.replace('__MONTH_LABEL__', 'Maio 2026')
html = html.replace('__ACTIVE_ABRIL__', '')
html = html.replace('__ACTIVE_MAIO__', 'active')
html = html.replace('__ACTIVE_JUNHO__', '')
html = html.replace('__TODAY_ISO__', TODAY.isoformat())
with open(OUT_HTML, 'w') as f:
    f.write(html)

print(f'\n✓ Salvo: {OUT_HTML}')
print(f'  Totais: {totals}')
