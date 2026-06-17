"""
Gera publish/sala.html (TV mode) a partir dos dados já embutidos em publish/maio-2026.html
+ dados extras lidos diretamente dos xlsx (reagend, leads por etiqueta, perdidos por motivo).
Roda DEPOIS do build_from_xlsx_maio.py.
"""
import re, json, os
import pandas as pd
from datetime import datetime, date as _date, timedelta as _timedelta
from collections import defaultdict

WORKDIR   = os.path.dirname(os.path.abspath(__file__))
DATA_DIR  = os.path.join(WORKDIR, 'data_pipedrive')
DASH_HTML = os.path.join(WORKDIR, 'publish', 'maio-2026.html')
SALA_TPL  = os.path.join(WORKDIR, 'sala_template.html')
OUT_SALA  = os.path.join(WORKDIR, 'publish', 'sala.html')

def load(fname):
    p = os.path.join(DATA_DIR, fname)
    return pd.read_excel(p) if os.path.exists(p) else pd.DataFrame()

# ── 1) Extrai D e today do dashboard já gerado ──────────────────────────────
with open(DASH_HTML) as f:
    dash = f.read()

m_data  = re.search(r'const D = (\{.*?\});', dash, re.DOTALL)
m_today = re.search(r"const today = '([^']+)';", dash)
if not m_data or not m_today:
    raise RuntimeError('Não foi possível extrair D ou today de maio-2026.html')

data_json = m_data.group(1)
today_str = m_today.group(1)

# ── 2) Dados extras para DS ──────────────────────────────────────────────────

today_date_dt = pd.to_datetime(today_str).date()

# Reagendamentos (total mês + hoje)
reag_df = load('reag.xlsx')
reag_count = len(reag_df)
reag_hoje = 0
if not reag_df.empty:
    col_reag_dt = next((c for c in reag_df.columns if 'marcado' in c.lower()), None)
    if col_reag_dt:
        reag_df['_dt'] = pd.to_datetime(reag_df[col_reag_dt], errors='coerce')
        reag_hoje = int((reag_df['_dt'].dt.date == today_date_dt).sum())

# Closer por Funil (separação por lista fixa de nomes, igual ao build_from_xlsx_maio.py)
CLOSER_GS_FRANQUEADOS = {'Raissa Antunes','Raíssa Nobre'}
CLOSER_PARTNER  = {'Maria Luísa','Djulia Silva','Haynnã','Samuel Almeida',
                   'Laura Feijó Junqueira','Manuella Oliveira','Tomaz Fabres',
                   'Tom Mansur','Clayton Bastos'}
CLOSER_FRANQUIA = {'William Dias','Thuany Ghabril','Thuany','Adalberto Neto','Andresa Caldas',
                   'Émerson Cavitchoni','Suzanny Mauren','Suzanny Mauren Dihelem','Leonardo Ribeiro',
                   'Felippe Porcella','Lorenzo Coronel','Samuel Brião','Fred','Gabriel Meneses',
                   'Ariane','Gabriela Oliveira','Luciana Krewer Issler','Mauricio Peixoto'}
# Closers que saíram — contam nos TOTAIS mas NÃO aparecem no ranking
CLOSER_SAIRAM = {'Tomaz Fabres','Haynnã','Thuany','Thuany Ghabril','Andresa Caldas'}

# Data de início do funil Partner-Franqueados sob gestão da equipe
PARTNER_START = _date(2026, 5, 27)

re_df2 = load('re.xlsx')
closer_by_funil = {'FRANQUIA': {}, 'PARTNER': {}}
if not re_df2.empty:
    col_resp2 = next((c for c in re_df2.columns if 'responsável' in c.lower()), None)
    if col_resp2:
        counts = re_df2[col_resp2].dropna().value_counts()
        for name, cnt in counts.items():
            name = str(name)
            if name in CLOSER_SAIRAM: continue  # saiu — fora do ranking (total conta no produto)
            if name in CLOSER_FRANQUIA:
                closer_by_funil['FRANQUIA'][name] = int(cnt)
            elif name in CLOSER_PARTNER:
                closer_by_funil['PARTNER'][name] = int(cnt)

# SDR stats do dia atual
def sdr_today_by(fname, date_hint, user_hint):
    df = load(fname)
    res = defaultdict(int)
    if df.empty: return {}
    col_d = next((c for c in df.columns if date_hint.lower() in c.lower()), None)
    col_u = next((c for c in df.columns if user_hint.lower() in c.lower()), None)
    if not col_d or not col_u: return {}
    df['_dt'] = pd.to_datetime(df[col_d], errors='coerce')
    for _, row in df[df['_dt'].dt.date == today_date_dt].iterrows():
        u = str(row.get(col_u, '') or '').strip()
        if u: res[u] += 1
    return dict(res)

sdr_today = {
    'ag':  sdr_today_by('ag.xlsx',     'marcado',    'criador'),
    're':  sdr_today_by('re.xlsx',     'marcado',    'criador'),
    'ns':  sdr_today_by('ns.xlsx',     'marcado',    'criador'),
    'fut': sdr_today_by('re_fut.xlsx', 'vencimento', 'criador'),
    'inv': sdr_today_by('inv.xlsx',    'marcado',    'criador'),
}

# CC1 Inválida — total mês + hoje + rankings
inv_df = load('inv.xlsx')
inv_count = len(inv_df)
inv_hoje = 0
closer_inv = {'FRANQUIA': {}, 'PARTNER': {}}
sdr_inv = {}
if not inv_df.empty:
    col_inv_dt   = next((c for c in inv_df.columns if 'marcado' in c.lower()), None)
    col_inv_resp = next((c for c in inv_df.columns if 'responsável' in c.lower()), None)
    col_inv_cri  = next((c for c in inv_df.columns if 'criador' in c.lower()), None)
    if col_inv_dt:
        inv_df['_dt'] = pd.to_datetime(inv_df[col_inv_dt], errors='coerce')
        inv_hoje = int((inv_df['_dt'].dt.date == today_date_dt).sum())
    if col_inv_resp:
        for name, cnt in inv_df[col_inv_resp].dropna().value_counts().items():
            name = str(name)
            if name in CLOSER_FRANQUIA:
                closer_inv['FRANQUIA'][name] = int(cnt)
            elif name in CLOSER_PARTNER:
                closer_inv['PARTNER'][name] = int(cnt)
    if col_inv_cri:
        from collections import Counter as _Ctr
        sdr_inv = {str(k): int(v) for k, v in inv_df[col_inv_cri].dropna().value_counts().items()}

# Leads por etiqueta (negócios iniciados por modelo)
# Combina ag.xlsx (ativos/ganhos) + perdidos.xlsx para total iniciado por modelo
leads_by_etiqueta = defaultdict(int)
for fname in ('ag.xlsx', 'perdidos.xlsx'):
    df_etq = load(fname)
    if not df_etq.empty:
        col_etq = next((c for c in df_etq.columns if 'etiqueta' in c.lower() or 'label' in c.lower()), None)
        if col_etq:
            for val, cnt in df_etq[col_etq].dropna().value_counts().items():
                key = str(val).strip()
                if key and key != 'nan':
                    leads_by_etiqueta[key] += int(cnt)
leads_by_etiqueta = dict(leads_by_etiqueta)

# Perdidos por motivo × produto
perdidos_df = load('perdidos.xlsx')
# Mantém só negócios CRIADOS e PERDIDOS dentro de maio
if not perdidos_df.empty and 'Negócio - Data de perda' in perdidos_df.columns \
   and 'Negócio - Negócio criado em' in perdidos_df.columns:
    _cri = pd.to_datetime(perdidos_df['Negócio - Negócio criado em'], errors='coerce').dt.date
    _pdt = pd.to_datetime(perdidos_df['Negócio - Data de perda'], errors='coerce').dt.date
    perdidos_df = perdidos_df[(_cri >= _date(2026,5,1)) & (_cri <= _date(2026,5,31)) &
                              (_pdt >= _date(2026,5,1)) & (_pdt <= _date(2026,5,31))].copy()
perdidos_por_motivo = defaultdict(lambda: defaultdict(int))
MOTIVOS_VALIDOS = {'Sem perfil','Sem interesse','Sem Retorno','Sem Contato','Perfil Fake','Buscando Parceria'}
PRODUTOS_MAP = {
    'studio agro': 'Studio Agro',
    'studio fiscal': 'Studio Fiscal',
    'lead studio agro': 'Studio Agro',
    'lead studio fiscal': 'Studio Fiscal',
    'partner': 'GS PARTNER',
    'lead partner': 'GS PARTNER',
    'franquia': 'Franquia Partner',
}

if not perdidos_df.empty:
    col_motivo  = next((c for c in perdidos_df.columns if 'motivo' in c.lower()), None)
    col_etq2    = next((c for c in perdidos_df.columns if 'etiqueta' in c.lower() or 'label' in c.lower()), None)
    col_funil   = next((c for c in perdidos_df.columns if 'funil' in c.lower()), None)

    for _, row in perdidos_df.iterrows():
        motivo = str(row.get(col_motivo, '') or '').strip() if col_motivo else ''
        if motivo not in MOTIVOS_VALIDOS:
            continue
        # determina produto pela etiqueta ou funil
        etq = str(row.get(col_etq2, '') or '').lower().strip() if col_etq2 else ''
        funil = str(row.get(col_funil, '') or '').lower().strip() if col_funil else ''
        produto = None
        for key, val in PRODUTOS_MAP.items():
            if key in etq or key in funil:
                produto = val
                break
        if not produto:
            continue
        perdidos_por_motivo[motivo][produto] += 1

# Mapa de calor RE — hora de início × dia da semana (Seg–Sáb)
# Usa "Atividade - Hora de início" (horário agendado da reunião)
# e "Atividade - Data de vencimento" (dia da reunião)
# Filtra apenas atividades marcadas como feitas (já estão em re.xlsx)
re_hm_df = load('re.xlsx')
re_heatmap = {}  # {'Seg': {'9': 5, '10': 12, ...}, ...}
DAY_NAMES_HM = ['Seg', 'Ter', 'Qua', 'Qui', 'Sex', 'Sáb']
if not re_hm_df.empty:
    col_hora  = next((c for c in re_hm_df.columns if c == 'Atividade - Hora de início'), None)
    col_data  = next((c for c in re_hm_df.columns if 'data de vencimento' in c.lower()), None)
    col_dt_inicio = next((c for c in re_hm_df.columns if 'data e hora de início' in c.lower()), None)
    if col_hora:
        # Pipedrive armazena due_time em UTC; Brasil = UTC-3
        _dt_utc = pd.to_datetime(
            re_hm_df[col_data].astype(str) + ' ' + re_hm_df[col_hora].astype(str),
            errors='coerce'
        ) if col_data else pd.to_datetime(re_hm_df[col_dt_inicio], errors='coerce')
        re_hm_df['_dt_hm'] = _dt_utc - pd.Timedelta(hours=3)
        for _, hm_row in re_hm_df.iterrows():
            dt_hm = hm_row['_dt_hm']
            if pd.isna(dt_hm):
                continue
            dow = dt_hm.weekday()   # 0=Seg … 6=Dom
            if dow > 5:             # pula domingo
                continue
            hour = dt_hm.hour
            day_n = DAY_NAMES_HM[dow]
            if day_n not in re_heatmap:
                re_heatmap[day_n] = {}
            key_h = str(hour)
            re_heatmap[day_n][key_h] = re_heatmap[day_n].get(key_h, 0) + 1

# Taxa de perda pós CC1 (por funil × motivo)
perdidos_pos_cc1 = {'FRANQUIA': {}, 'PARTNER': {}}
re_cc1_df = load('re.xlsx')
if not re_cc1_df.empty:
    col_status_cc1 = next((c for c in re_cc1_df.columns if 'negócio' in c.lower() and 'status' in c.lower()), None)
    col_motivo_cc1 = next((c for c in re_cc1_df.columns if 'motivo' in c.lower()), None)
    col_funil_cc1  = next((c for c in re_cc1_df.columns if 'funil'  in c.lower()), None)
    if col_status_cc1 and col_motivo_cc1 and col_funil_cc1:
        for fk in ('FRANQUIA', 'PARTNER'):
            sub = re_cc1_df[re_cc1_df[col_funil_cc1] == fk]
            total = len(sub)
            perdidos_sub = sub[sub[col_status_cc1] == 'Perdido']
            motivos = perdidos_sub[col_motivo_cc1].dropna().value_counts()
            perdidos_pos_cc1[fk] = {
                'total_cc1': int(total),
                'total_perdidos': int(len(perdidos_sub)),
                'motivos': {str(k): int(v) for k, v in motivos.items()}
            }

# Ritmo semanal RE + projeção do mês
SEMANAS_MAIO = [
    ('Sem 1', _date(2026,5,4),  _date(2026,5,9)),
    ('Sem 2', _date(2026,5,11), _date(2026,5,16)),
    ('Sem 3', _date(2026,5,18), _date(2026,5,23)),
    ('Sem 4', _date(2026,5,25), _date(2026,5,30)),
]
re_ritmo = []
if not re_hm_df.empty and '_dt_hm' in re_hm_df.columns:
    re_hm_df['_date_s'] = re_hm_df['_dt_hm'].dt.date
    col_resp_s = next((c for c in re_hm_df.columns if 'responsável' in c.lower()), None)
    for label, ws, we in SEMANAS_MAIO:
        mask = (re_hm_df['_date_s'] >= ws) & (re_hm_df['_date_s'] <= we)
        sub   = re_hm_df[mask]
        total = len(sub)
        franq_re = int(sub[col_resp_s].isin(CLOSER_FRANQUIA).sum()) if col_resp_s else 0
        partn_re = int(sub[col_resp_s].isin(CLOSER_PARTNER).sum())  if col_resp_s else 0
        completa     = we < today_date_dt
        em_andamento = ws <= today_date_dt <= we
        proj = None
        dias_passados = 0
        if em_andamento:
            dias_passados = sum(1 for i in range((today_date_dt - ws).days + 1)
                                if (ws + _timedelta(days=i)).weekday() <= 5)
        re_ritmo.append({
            'label': label,
            'dates': f'{ws.strftime("%d/%m")}–{we.strftime("%d/%m")}',
            'total': total,
            'franquia': franq_re,
            'partner':  partn_re,
            'completa': completa,
            'em_andamento': em_andamento,
            'dias_passados': dias_passados,
            'proj': proj,  # calculado após ter avg_last2
        })

# Média das últimas 2 semanas completas (base para projeções)
semanas_ok = [s['total'] for s in re_ritmo if s['completa'] and s['total'] > 0]
avg_last2  = round(sum(semanas_ok[-2:]) / len(semanas_ok[-2:])) if len(semanas_ok) >= 1 else 0

for s in re_ritmo:
    if s['completa']:
        continue
    dias = s['dias_passados']
    total_s = s['total']
    if s['em_andamento']:
        if dias >= 3 and total_s > 0:
            # Ritmo confiável: puro extrapolation
            proj_pace = round(total_s / dias * 6)
        elif dias > 0 and total_s > 0:
            # Dias insuficientes: blend 30% ritmo atual + 70% histórico
            proj_pace = round(total_s / dias * 6 * 0.3 + avg_last2 * 0.7)
        else:
            proj_pace = avg_last2
        s['proj'] = proj_pace
    else:
        # Semana futura
        s['proj'] = avg_last2

proj_mes = sum(
    s['total'] if s['completa'] else (s['proj'] or avg_last2)
    for s in re_ritmo
)
re_semanas = {'semanas': re_ritmo, 'proj_mes': proj_mes, 'avg_last2': avg_last2}

# ── Partner-Franqueados panel (gestão iniciada em PARTNER_START) ─────────────
def _partner_closer_stats(fname, date_col_hint, resp_col_hint, funil_col_hint='funil'):
    """Retorna {closer: {mes: N, hoje: N}} para o funil PARTNER."""
    df = load(fname)
    result = {}
    if df.empty: return result
    col_dt   = next((c for c in df.columns if date_col_hint.lower() in c.lower()), None)
    col_resp = next((c for c in df.columns if resp_col_hint.lower() in c.lower()), None)
    col_fun  = next((c for c in df.columns if funil_col_hint.lower() in c.lower()), None)
    if not (col_dt and col_resp and col_fun): return result
    df = df[df[col_fun] == 'PARTNER'].copy()
    df['_dt'] = pd.to_datetime(df[col_dt], errors='coerce')
    for _, row in df.iterrows():
        name = str(row.get(col_resp, '') or '').strip()
        if not name or name not in CLOSER_PARTNER: continue
        dt = row['_dt']
        if pd.isna(dt): continue
        d = dt.date()
        if name not in result:
            result[name] = {'mes': 0, 'hoje': 0, 'inicio': 0}
        result[name]['mes'] += 1
        if d == today_date_dt:
            result[name]['hoje'] += 1
        if d >= PARTNER_START:
            result[name]['inicio'] += 1
    return result

partner_re_stats  = _partner_closer_stats('re.xlsx',   'marcado', 'responsável')
partner_ns_stats  = _partner_closer_stats('ns.xlsx',   'marcado', 'responsável')

# AG PARTNER por SDR (desde o início da gestão = PARTNER_START)
def _partner_ag_sdr():
    df = load('ag.xlsx')
    result = {}
    if df.empty: return result
    col_dt  = next((c for c in df.columns if 'marcado' in c.lower()), None)
    col_cri = next((c for c in df.columns if 'criador'  in c.lower()), None)
    col_fun = next((c for c in df.columns if 'funil'    in c.lower()), None)
    if not (col_dt and col_cri and col_fun): return result
    df = df[df[col_fun] == 'PARTNER'].copy()
    df['_dt'] = pd.to_datetime(df[col_dt], errors='coerce')
    for _, row in df.iterrows():
        dt = row['_dt']
        if pd.isna(dt): continue
        d = dt.date()
        name = str(row.get(col_cri, '') or '').strip()
        if not name: continue
        if name not in result:
            result[name] = {'mes': 0, 'hoje': 0, 'inicio': 0}
        result[name]['mes'] += 1
        if d == today_date_dt:
            result[name]['hoje'] += 1
        if d >= PARTNER_START:
            result[name]['inicio'] += 1
    return result

partner_ag_sdr = _partner_ag_sdr()

# Deals PARTNER: iniciados e perdidos desde PARTNER_START
deals_df = load('deals.xlsx')
partner_deals = {'iniciados_mes': 0, 'iniciados_inicio': 0,
                 'perdidos_mes': 0,  'perdidos_inicio': 0}
MONTH_START = _date(2026, 5, 1)
MONTH_END   = _date(2026, 5, 31)

if not deals_df.empty:
    col_fun_d = next((c for c in deals_df.columns if 'funil' in c.lower()), None)
    col_cr_d  = next((c for c in deals_df.columns if 'criado em' in c.lower()), None)
    if col_fun_d and col_cr_d:
        pd_df = deals_df[deals_df[col_fun_d] == 'PARTNER'].copy()
        pd_df['_dt'] = pd.to_datetime(pd_df[col_cr_d], errors='coerce')
        _d = pd_df['_dt'].dt.date
        partner_deals['iniciados_mes']    = int(((_d >= MONTH_START) & (_d <= MONTH_END)).sum())
        partner_deals['iniciados_inicio'] = int((_d >= PARTNER_START).sum())

if not perdidos_df.empty:
    col_fun_p = next((c for c in perdidos_df.columns if 'funil' in c.lower()), None)
    col_lp    = next((c for c in perdidos_df.columns if 'data de perda' in c.lower()), None)
    if col_fun_p and col_lp:
        pp_df = perdidos_df[perdidos_df[col_fun_p] == 'PARTNER'].copy()
        pp_df['_dt'] = pd.to_datetime(pp_df[col_lp], errors='coerce')
        _d = pp_df['_dt'].dt.date
        partner_deals['perdidos_mes']    = int(((_d >= MONTH_START) & (_d <= MONTH_END)).sum())
        partner_deals['perdidos_inicio'] = int((_d >= PARTNER_START).sum())

# Média diária de AGs PARTNER desde PARTNER_START (proxy: leads de outros funis direcionados)
dias_desde_inicio = max(1, (today_date_dt - PARTNER_START).days + 1)
ag_inicio_total = sum(v['inicio'] for v in partner_ag_sdr.values())
partner_media_ag = round(ag_inicio_total / dias_desde_inicio, 1)

partner_panel = {
    'deals': partner_deals,
    'ag_sdr': partner_ag_sdr,
    'closer_re': partner_re_stats,
    'closer_ns': partner_ns_stats,
    'ag_mes': sum(v['mes'] for v in partner_ag_sdr.values()),
    'ag_inicio': ag_inicio_total,
    'media_ag_dia': partner_media_ag,
    'start_date': PARTNER_START.strftime('%d/%m/%Y'),
}

# Serializa (converte defaultdict para dict normal)
perdidos_por_motivo_clean = {
    motivo: dict(prods)
    for motivo, prods in sorted(
        perdidos_por_motivo.items(),
        key=lambda x: -sum(x[1].values())
    )
}

DS = {
    'reag': reag_count,
    'reag_hoje': reag_hoje,
    'inv': inv_count,
    'inv_hoje': inv_hoje,
    'closer_inv': closer_inv,
    'sdr_inv': sdr_inv,
    'leads_by_etiqueta': leads_by_etiqueta,
    'perdidos_por_motivo': perdidos_por_motivo_clean,
    'perdidos_pos_cc1': perdidos_pos_cc1,
    'closer_by_funil': closer_by_funil,
    'sdr_today': sdr_today,
    're_heatmap': re_heatmap,
    're_semanas': re_semanas,
    'partner_panel': partner_panel,
}

# ── 3) Gera sala.html ────────────────────────────────────────────────────────
from datetime import timezone, timedelta as _td
_BRT = timezone(_td(hours=-3))
last_update = datetime.now(_BRT).strftime('%d/%m/%Y %H:%M')

with open(SALA_TPL) as f:
    sala = f.read()

sala = sala.replace('__DATA_JSON__',       data_json)
sala = sala.replace('__SALA_JSON__',       json.dumps(DS, ensure_ascii=False))
sala = sala.replace('__TODAY__',           today_str)
sala = sala.replace('__LAST_UPDATE__',     last_update)
sala = sala.replace('__MES_LABEL__',       'MAIO 2026')
sala = sala.replace('__ACTIVE_SALA_MAIO__',  'sala-tab-active')
sala = sala.replace('__ACTIVE_SALA_JUNHO__', '')

with open(OUT_SALA, 'w') as f:
    f.write(sala)

print(f'✓ sala.html gerado — {last_update}')
print(f'  reag={reag_count} | etiquetas={len(leads_by_etiqueta)} | motivos={len(perdidos_por_motivo_clean)}')
