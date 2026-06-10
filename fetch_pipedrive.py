"""
Fetch dados do Pipedrive via API e salva no formato esperado pelo build_from_xlsx_maio.py.
Substitui as exportações xlsx manuais.

Mapeia para 8 arquivos xlsx com nomes estáveis em /data_pipedrive/:
- deals.xlsx          — todos negócios criados em Maio
- perdidos.xlsx       — negócios perdidos com data de perda em Maio
- qual.xlsx           — Qualificação Realizada (concluídas)
- ag.xlsx             — Agendamento de CC1 + BOT (concluídas)
- reag.xlsx           — Reagendamento de CC1 (concluídas)
- re.xlsx             — CC1 Realizada (concluídas)
- re_fut.xlsx         — CC1 Realizada (A fazer)
- ns.xlsx             — CC1 Não realizada (concluídas)
- inv.xlsx            — CC1 Inválida (concluídas)
"""
import urllib.request, urllib.parse, urllib.error, json, os, time
import pandas as pd
from datetime import date, timedelta
import concurrent.futures

WORKDIR = os.path.dirname(os.path.abspath(__file__))
TOKEN_FILE = os.path.join(WORKDIR, '.pipedrive_token')
OUT_DIR = os.path.join(WORKDIR, 'data_pipedrive')
os.makedirs(OUT_DIR, exist_ok=True)

# Aceita token via variável de ambiente (GitHub Actions) ou arquivo local (Mac)
TOKEN = os.environ.get('PIPEDRIVE_TOKEN') or (
    open(TOKEN_FILE).read().strip() if os.path.exists(TOKEN_FILE) else None
)
if not TOKEN:
    raise RuntimeError('PIPEDRIVE_TOKEN não encontrado (env var ou .pipedrive_token)')

BASE = 'https://api.pipedrive.com/v1'

# ============= CONSTANTES MAPEADAS =============
# Pipeline ID → Nome (para coluna "Negócio - Funil")
PIPELINE_MAP = {42:'FRANQUIA', 44:'PARTNER', 43:'TECNOLOGIA', 27:'GS PARTNER',
                57:'GS Educação- Partners', 54:'TREINAMENTO QUALIFICAÇÃO',
                53:'Imersão - Supply Tax (Freitag)', 52:'Imersão - Holding (Maria Eduarda Pavan)',
                51:'Imersão - Holding (Marcelo Borba)', 41:"Jvs'",
                48:'GS EDUCAÇÃO - PLATAFORMA', 46:'GS EDUCAÇÃO - VENDAS',
                47:'GS EDUCAÇÃO - DISPAROS', 45:'GS EDUCAÇÃO', 49:'Upsell',
                50:'Partners GS- FRANQUEADOS', 55:'Supply Tax - Cliente Final',
                56:'Supply Tax - Intermediário'}

# Label ID → Nome (para coluna "Negócio - Etiqueta")
LABEL_MAP = {96:'Lead Studio Agro', 97:'Lead Studio Fiscal', 98:'Lead PJ360º',
             110:'Lead GS PARTNER', 113:'Lead Partner', 114:'Lead Partners (Raissa)'}

# Custom field keys
CF_UTM_SOURCE   = '0ee0949cd3cc7323a52fd79ab3d6500da973198f'
CF_UTM_CAMPAIGN = 'ee93fbc62c1804b82721486630058f3467651cec'
CF_UTM_MEDIUM   = 'f7273db64fbf4cf022c70cc7554c33231653641e'
CF_UTM_CONTENT  = '6673bb1656201df819dcd47be1cf751a966d2229'
CF_UTM_TERM     = '1f0cdcf753543bcca8b7a39baf0805c12db7f3bb'
CF_TELEFONE     = 'a9385cd4cdbac0b87d8f4dcb91db53032c9cb20b'
CF_PROFISSAO    = '2152139aa2df5612ce72ccd79079878f9974742c'
CF_CIDADE       = '4d3dbd8d685e667e12558c0dd485885c9c95eb05'
CF_ESTADO       = '874d1c7c1bd4765ade327c54801061f98983b3c4'
CF_SOBRE        = 'ce4a34ebbcd7dbadf198d9059b1610b74ee41362'
CF_SOBRE_FULL   = '9024574101b03ea317f00c2268fbeda4a93e435f'
CF_DEAL_EMAIL   = 'ddc48802be95495bede3bc57ac2579ae3d6018ae'
CF_LOST_DATE    = 'afca450738e19acb2739a5f4bd8018cf524c2c6b'

# Tipos de atividade (Pipedrive activity types)
ACT_QUAL_REAL_A = 'qualificacao_realizada1'  # id 48
ACT_QUAL_REAL_B = 'qualificacao_realizada'   # id 47
ACT_AG_CC1      = 'agendamento_de_cc1'       # id 7
ACT_AG_BOT      = 'agendamento_cc1___bot'    # id 58
ACT_REAG_CC1    = 'reagendamento_de_cc1'     # id 24
ACT_CC1_REAL    = 'meeting'                  # id 2 (CC1 Realizada)
ACT_CC1_NS      = 'cc1_nao_realizada'        # id 15
ACT_NS_IMPREV   = 'no_show_imprevisto'       # id 59 (novo)
ACT_NS_DESINT   = 'no_show_desinteresse'     # id 60 (novo)
ACT_CC1_INV     = 'cc1_invalida'             # id 61 (novo)

# Stages da Qualificação (etapas iniciais onde a equipe atua)
QUALIF_STAGES = {408, 434, 464, 409, 424, 432, 466, 425}
# 408 Lead Franquia, 434 Pré-Qualif (FR), 464 Em atendimento (FR), 409 MQL FRANQUIA
# 424 Lead (PT), 432 Pré-Qualif (PT), 466 Em atendimento (PT), 425 MQL (PT)

MONTH_START = date(2026, 5, 1)
MONTH_END = date(2026, 6, 30)   # Cobre Maio + Junho (cada build filtra seu mês)
# Fetch com janela mais ampla para capturar atividades com due_date em meses anteriores
# mas marcadas como feitas no mês atual
FETCH_START = date(2026, 1, 1)  # Janela ampla: pega atividades reagendadas de meses anteriores
FETCH_END   = date(2026, 6, 30)

# SDRs da equipe Qualificadora (espelha o que está no build_from_xlsx_maio.py)
SDR_LIST = {'Mélodi Nirvana','Lorenzo Meneghetti','Leonardo Soares','Glauber Bagetti',
            'Andriely Oliveira','Izadora Zacker','Henri Meneghetti',
            'João Gabriel Dutra','Jonatans Junior',
            'Amanda Oliveira','Camily Nadal','Rafael dallarosa longo'}

# Pessoas que NÃO devem ser consideradas como criadoras ou responsáveis pelas atividades
# (espelha o filtro "Usuário responsável não pode ser" do painel Pipedrive Insights "Qualificação Realizada")
EXCLUDE_USERS = {
    'Ana Cristina Wist','André Maciel','Carlos Alberto','Daniela Rodrigues','Eduarda Ferner',
    'Gabriel Meneses','Gabriela Cunha','Gustavo Oliveira','Leonardo Ribeiro','Keven',
    'Marina da Rocha','Mauricio Peixoto','Raquel Arioli','Suzanny Mauren Dihelem','Thuany Ghabril',
    'Veridiana Altman','Yuri Ghabril','Anderson Muller','Bruna Mota','Bruna Salazar',
    'Bruno Leitemperger','Bruno Torves','Darlan','Eduarda Linhares','Felipe Gomes',
    'Gabriely Vianna','Gustavo','Gustavo Armani','Haynnã','Helena Robim',
    'Jéssica Maia','Laura Feijó Junqueira','Lorenzo Coronel','Luana Nardi',"Lucas D'Andrea",
    'Lucas Delpino','Ramiro Ilha Moreira Cardoso','Samuel Brião','Thiago Viana Lucas',
    'Alexandre Furtado','Adalberto Neto','Andresa Caldas','Bárbara','Djulia Silva',
    'Jeferson Ludke','Marcos Tadeu','Maria Luísa','Maria Eduarda Figueiredo','Márcio Nunes',
    'Raissa Antunes','Raíssa Nobre','William Dias','Willian Zanatta','Angelo Fontella Leite',
    'Felippe Porcella','Guilherme Vargas','Jéssica Carvalho','Larissa Kihs','Émerson Cavitchoni',
    'Alécio Carneiro','Claudia','Heverton Matiello','Juan Victor S Suxo','Pablo Medeiros',
    'Rodrigo Burkert','Samuel Almeida','Guilherme Freitag','Renê Callegaro de Morais',
    'Daniela Souza','Fred','Edlaine dos Anjos','Ernesto Pozzebon','Manuella Oliveira',
    'Maria Eduarda Pavan','Aline','Karla Bordin Domingos','Vítor Levi',
}

# Etiquetas (labels) excluídas — são etiquetas antigas/internas que o painel ignora
# IDs mapeados via Pipedrive (LABEL_MAP tem só as relevantes; estes são EXCLUDED)
EXCLUDE_LABEL_NAMES = {
    'Enrolando (talvez não feche)','Ultima tentativa de contato','Mil Reais Treinamento',
    'Lead bom','Lead medio','Lead ruim','DECISÃO POSTERGADA','Prioridade Follow Up',
    'BOT IA Conversando','Studio Law','Studio Agro','Studio Fiscal','Interessado',
    'OPORTUNIDADES ABF','CONFIRMAR LIVE','PAR','Potencial Alto de Fechamento',
    'Em espera','JOB','EXPOINTER','FECHAMENTO ESTE MÊS','CERTIFICAÇÃO',
    'BLACK FRIDAY','Fechado','PJ360','Carlos','Funil Aceleradora',
    'Live Agro 06/11','Perdido Agro - Fiscal','LEAD DA ANA','fenalaw24',
}
# Mapeia esses nomes pra IDs usando o LABEL_MAP_FULL — vai ser construído após fetch
EXCLUDE_LABEL_IDS = set()  # preenche depois

# ============= HELPERS =============
def fetch_all(endpoint, params=None):
    """Itera paginação até pegar todos os resultados."""
    params = params or {}
    params['api_token'] = TOKEN
    params['limit'] = 500
    start = 0
    all_data = []
    while True:
        params['start'] = start
        qs = urllib.parse.urlencode(params, doseq=True)
        url = f'{BASE}/{endpoint}?{qs}'
        body = None
        for attempt in range(10):
            try:
                with urllib.request.urlopen(url, timeout=120) as r:
                    body = json.loads(r.read())
                break
            except urllib.error.HTTPError as e:
                if e.code == 429:  # rate limit
                    wait = 15 * (attempt + 1)
                    print(f'     ⚠️  429 rate limit, esperando {wait}s...')
                    time.sleep(wait)
                elif attempt == 9: raise
                else: time.sleep(3)
            except Exception:
                if attempt == 9: raise
                time.sleep(3)
        if body is None: raise Exception(f'Falha ao buscar {endpoint} após 10 tentativas')
        data = body.get('data') or []
        all_data.extend(data)
        pag = body.get('additional_data', {}).get('pagination', {})
        if not pag.get('more_items_in_collection'): break
        start = pag.get('next_start', start + len(data))
        if len(data) == 0: break
    return all_data

# Cache de users (user_id → name) — incluindo inativos (Jhully, etc.)
print('→ Fetching users...')
USERS = {u['id']: u['name'] for u in fetch_all('users')}
print(f'   {len(USERS)} users carregados (ativos + inativos)')

# Cache de labels (label_id → name) — busca todas as opções de etiqueta
print('→ Fetching labels...')
def get(endpoint):
    url = f'{BASE}/{endpoint}'
    sep = '&' if '?' in endpoint else '?'
    url += f'{sep}api_token={TOKEN}'
    return json.loads(urllib.request.urlopen(url, timeout=60).read())
ALL_LABELS = {}
for f in get('dealFields').get('data', []) or []:
    if f.get('key') == 'label':
        for opt in f.get('options') or []:
            ALL_LABELS[opt['id']] = opt['label']
        break
# Constrói o set de IDs a excluir
EXCLUDE_LABEL_IDS = {lid for lid, lname in ALL_LABELS.items() if lname in EXCLUDE_LABEL_NAMES}
print(f'   {len(ALL_LABELS)} labels totais, {len(EXCLUDE_LABEL_IDS)} excluídos: {sorted(EXCLUDE_LABEL_IDS)}')

def user_name(uid):
    if isinstance(uid, dict): uid = uid.get('id')
    return USERS.get(uid)

# ============= FETCH DEALS =============
# Usar /deals/timeline com field_key=add_time pra pegar só os criados em Maio
def fetch_timeline(field_key, start_dt, days=31):
    """Pega deals via /deals/timeline filtrando por field_key (add_time, lost_time, etc.)."""
    params = {
        'api_token': TOKEN,
        'start_date': start_dt.isoformat(),
        'interval': 'day',
        'amount': days,
        'field_key': field_key,
    }
    url = f'{BASE}/deals/timeline?' + urllib.parse.urlencode(params)
    body = None
    for attempt in range(6):
        try:
            with urllib.request.urlopen(url, timeout=120) as r:
                body = json.loads(r.read())
            break
        except Exception as exc:
            if attempt == 5: raise
            wait = 10 * (attempt + 1)
            print(f'     ⚠️  timeline timeout/erro (tentativa {attempt+1}), esperando {wait}s... [{exc}]')
            time.sleep(wait)
    out = []
    for period in body.get('data', []) or []:
        for d in period.get('deals') or []:
            out.append(d)
    return out

def _fetch_open_won():
    # O /deals/timeline retorna dados INCOMPLETOS (perde muitos deals). Usa /deals
    # paginado (add_time DESC), coletando tudo criado a partir de MONTH_START.
    result = []
    start = 0
    while True:
        params = {'api_token': TOKEN, 'limit': 500, 'start': start, 'sort': 'add_time DESC'}
        url = f'{BASE}/deals?' + urllib.parse.urlencode(params)
        body = None
        for attempt in range(6):
            try:
                with urllib.request.urlopen(url, timeout=120) as r:
                    body = json.loads(r.read())
                break
            except Exception as exc:
                if attempt == 5: raise
                wait = 10 * (attempt + 1)
                print(f'     ⚠️  deals page erro (tentativa {attempt+1}), esperando {wait}s... [{exc}]')
                time.sleep(wait)
        data = body.get('data') or []
        stop = False
        for d in data:
            at = (d.get('add_time') or '')[:10]
            if not at:
                continue
            if at < MONTH_START.isoformat():
                stop = True
                continue
            if at > MONTH_END.isoformat():
                continue
            result.append(d)
        pag = (body.get('additional_data') or {}).get('pagination') or {}
        if stop or not pag.get('more_items_in_collection'):
            break
        start = pag.get('next_start', start + 500)
    print(f'   {len(result)} deals criados em Maio+Junho (/deals paginado)')
    return result

def _fetch_lost():
    lost = []
    start = 0
    pages = 0
    while True:
        params = {'api_token': TOKEN, 'status': 'lost', 'limit': 500, 'start': start,
                  'sort': 'add_time DESC'}
        url = f'{BASE}/deals?' + urllib.parse.urlencode(params)
        body = None
        for attempt in range(6):
            try:
                with urllib.request.urlopen(url, timeout=120) as r:
                    body = json.loads(r.read())
                break
            except Exception as exc:
                if attempt == 5: raise
                wait = 10 * (attempt + 1)
                print(f'     ⚠️  lost page timeout/erro (tentativa {attempt+1}), esperando {wait}s...')
                time.sleep(wait)
        data = body.get('data') or []
        pages += 1
        lost.extend([d for d in data if (d.get('add_time') or '').startswith('2026-05')])
        if data and all((d.get('add_time') or '') < '2026-05-01' for d in data): break
        pag = body.get('additional_data', {}).get('pagination', {})
        if not pag.get('more_items_in_collection'): break
        start = pag.get('next_start', start + len(data))
        if len(data) == 0: break
        if pages > 20: break
    print(f'   {len(lost)} deals LOST criados em Maio ({pages} páginas)')
    return lost

print('→ Fetching deals em paralelo (open/won + lost)...')
with concurrent.futures.ThreadPoolExecutor(max_workers=2) as executor:
    f_ow  = executor.submit(_fetch_open_won)
    f_lost = executor.submit(_fetch_lost)
    deals_maio_open_won = f_ow.result()
    deals_lost_raw      = f_lost.result()

# Combina open+won (timeline) + lost = total Negócios Iniciados em Maio
deals_maio_raw = deals_maio_open_won + deals_lost_raw
# Dedup
seen_ids = set()
deduped = []
for d in deals_maio_raw:
    if d['id'] not in seen_ids:
        deduped.append(d)
        seen_ids.add(d['id'])
deals_maio_raw = deduped
print(f'   {len(deals_maio_raw)} deals únicos criados em Maio (open+won+lost)')

# Junta tudo num dict por id pra ter referência rápida
deals_all_dict = {d['id']: d for d in deals_maio_raw}
for d in deals_lost_raw:
    deals_all_dict.setdefault(d['id'], d)
deals_all = list(deals_all_dict.values())
print(f'   {len(deals_all)} deals únicos no escopo')

def deal_row(d, include_lost_fields=False):
    """Converte deal API → linha xlsx (mesmo formato dos exports manuais)."""
    label_id = d.get('label')
    etiqueta = None
    if label_id is not None:
        try: etiqueta = LABEL_MAP.get(int(label_id))
        except: pass
    row = {
        'Negócio - Título': d.get('title'),
        'Negócio - Valor do negócio': d.get('value'),
        'Negócio - Funil': PIPELINE_MAP.get(d.get('pipeline_id')),
        'Negócio - Proprietário': user_name(d.get('user_id')),
        'Negócio - Status': {'open':'Aberto','won':'Ganho','lost':'Perdido'}.get(d.get('status'), d.get('status')),
        'Negócio - Negócio criado em': d.get('add_time'),
        'Negócio - Data de fechamento esperada': d.get('expected_close_date'),
        'Negócio - Etapa': d.get('stage_id'),
        'Negócio - Sobre você': d.get(CF_SOBRE),
        'Negócio - Sobre você (Info completa)': d.get(CF_SOBRE_FULL),
        'Negócio - Profissão': d.get(CF_PROFISSAO),
        'Negócio - UTM campaign': d.get(CF_UTM_CAMPAIGN),
        'Negócio - UTM content': d.get(CF_UTM_CONTENT),
        'Negócio - UTM medium': d.get(CF_UTM_MEDIUM),
        'Negócio - UTM source': d.get(CF_UTM_SOURCE),
        'Negócio - UTM term': d.get(CF_UTM_TERM),
        'Negócio - Telefone': d.get(CF_TELEFONE),
        'Pessoa - Telefone': None,
    }
    if include_lost_fields:
        row['Negócio - Etiqueta'] = etiqueta
        row['Negócio - Motivo da perda'] = d.get('lost_reason')
        row['Negócio - Data de perda'] = d.get('lost_time')
    return row

# ÓTICA dos painéis Pipedrive:
# - Negócios Iniciados (painel): criado em Maio + Etiqueta ∈ {Lead Studio Agro/Fiscal, Lead Partner}
#   (não filtra por proprietário — pega 3.277)
# - Novo relatório (perdidos): criado em Maio + perdido em Maio + Proprietário SDR (1.515)
ETIQUETAS_VALIDAS_DEALS = {'Lead Studio Agro','Lead Studio Fiscal','Lead Partner','Lead Partners (Raissa)'}

def deal_label_name(d):
    lid = d.get('label')
    if lid is None: return None
    try: return LABEL_MAP.get(int(lid))
    except: return None

def is_qualif_owned(d):
    return user_name(d.get('user_id')) in SDR_LIST

# Deals iniciados: criado em Maio + etiqueta válida + Funil FRANQUIA ou PARTNER
# (espelha filtro Pipedrive: Etiqueta ∈ {Lead Studio Agro/Fiscal/Partner} + Funil ∈ {FRANQUIA,PARTNER})
deals_maio = [deal_row(d, include_lost_fields=False) for d in deals_maio_raw
              if deal_label_name(d) in ETIQUETAS_VALIDAS_DEALS
              and d.get('pipeline_id') in (42, 44, 50)]

# Perdidos: criado em Maio + perdido em Maio + Funil FR/PT + Etapa ∈ etapas iniciais (Qualificação)
# Espelha o painel "Novo relatório de desempenho de negócio"
# Itera sobre o conjunto COMPLETO (deals_maio_raw já inclui lost via /deals paginado),
# não só deals_lost_raw — garante captura de deals criados E perdidos no mês.
perdidos_maio = []
_perd_seen = set()
for d in deals_maio_raw + deals_lost_raw:
    if d['id'] in _perd_seen: continue
    if d.get('status') != 'lost': continue
    lt = d.get('lost_time')
    if not lt: continue
    lt_dt = pd.to_datetime(lt).date()
    if not (MONTH_START <= lt_dt <= MONTH_END): continue
    pid = d.get('pipeline_id')
    if pid not in (42, 44): continue           # Perdidos: só FRANQUIA / PARTNER (espelha filtro Pipedrive)
    if d.get('stage_id') not in QUALIF_STAGES: continue  # etapas de Qualificação
    _perd_seen.add(d['id'])
    perdidos_maio.append(deal_row(d, include_lost_fields=True))

print(f'   {len(deals_maio)} deals (criados+etiqueta válida) para deals.xlsx')
print(f'   {len(perdidos_maio)} perdidos (criado+perdido Maio + Funil FR/PT + etapa Qualificação) para perdidos.xlsx')

# Salva xlsx
pd.DataFrame(deals_maio).to_excel(os.path.join(OUT_DIR,'deals.xlsx'), index=False)
pd.DataFrame(perdidos_maio).to_excel(os.path.join(OUT_DIR,'perdidos.xlsx'), index=False)

# ============= FETCH ACTIVITIES =============
# Vou pegar todas as atividades dos tipos relevantes em Maio
# /activities aceita start_date e end_date (due_date)

# Map deal_id → deal info (para enriquecer activities com Etiqueta/Funil/Equipe)
deal_by_id = deals_all_dict

def deal_etiqueta(deal_id):
    d = deal_by_id.get(deal_id)
    if not d: return None
    lid = d.get('label')
    if lid is None: return None
    try: lid = int(lid)
    except: return None
    return LABEL_MAP.get(lid)

def deal_funil(deal_id):
    d = deal_by_id.get(deal_id)
    return PIPELINE_MAP.get(d.get('pipeline_id')) if d else None

def deal_status(deal_id):
    d = deal_by_id.get(deal_id)
    if not d: return None
    return {'open':'Aberto','won':'Ganho','lost':'Perdido'}.get(d.get('status'))

def deal_email(deal_id):
    d = deal_by_id.get(deal_id)
    return d.get(CF_DEAL_EMAIL) if d else None

def deal_lost_reason(deal_id):
    d = deal_by_id.get(deal_id)
    return d.get('lost_reason') if d else None

def deal_lost_time(deal_id):
    d = deal_by_id.get(deal_id)
    return d.get('lost_time') if d else None

def deal_title(deal_id):
    d = deal_by_id.get(deal_id)
    return d.get('title') if d else None

def deal_team(deal_id):
    """A Equipe vem do team_id do user proprietário."""
    d = deal_by_id.get(deal_id)
    if not d: return None
    uid = d.get('user_id')
    if isinstance(uid, dict): uid = uid.get('id')
    # TODO: mapear user.team → team name. Por enquanto, deixar vazio
    return None

print('→ Fetching activities em paralelo (janela ampla Mar-Mai)...')
# Usa threading para buscar os 9 tipos simultaneamente (reduz ~60s → ~15s)
# Janela ampla para capturar atividades com due_date em Março/Abril
# mas marcadas como feitas em Maio (painéis Pipedrive usam "Marcado como feito em")
TIPOS_RELEVANTES = [ACT_QUAL_REAL_A, ACT_QUAL_REAL_B, ACT_AG_CC1, ACT_AG_BOT,
                    ACT_REAG_CC1, ACT_CC1_REAL, ACT_CC1_NS,
                    ACT_NS_IMPREV, ACT_NS_DESINT, ACT_CC1_INV]

def _fetch_tipo(tipo):
    chunk = fetch_all('activities', {
        'type': tipo,
        'user_id': 0,
        'start_date': FETCH_START.isoformat(),
        'end_date': FETCH_END.isoformat(),
    })
    print(f'   {tipo}: {len(chunk)}')
    return chunk

acts_raw = []
with concurrent.futures.ThreadPoolExecutor(max_workers=4) as executor:
    for chunk in executor.map(_fetch_tipo, TIPOS_RELEVANTES):
        acts_raw.extend(chunk)
print(f'   {len(acts_raw)} atividades (janela ampla) dos tipos relevantes')

def deal_label_id(deal_id):
    d = deal_by_id.get(deal_id)
    if not d: return None
    lid = d.get('label')
    if lid is None: return None
    try: return int(lid)
    except: return None

def _in_maio(dt_str):
    """Retorna True se a string de data/datetime cai em Maio 2026."""
    if not dt_str: return False
    try:
        d = pd.to_datetime(dt_str).date()
        return MONTH_START <= d <= MONTH_END
    except: return False

# ── Filtros por tipo — espelham exatamente os filtros dos painéis Pipedrive ──
#
# Qualificação realizada:
#   Marcado como feito em = Maio | Tipo = Qual | Responsável NOT IN EXCLUDE_USERS
#   (sem filtro de criador)
def keep_qual(a):
    if not a.get('done'): return False
    if not _in_maio(a.get('marked_as_done_time')): return False
    responsavel = user_name(a.get('user_id'))
    if responsavel in EXCLUDE_USERS: return False
    lid = deal_label_id(a.get('deal_id'))
    if lid in EXCLUDE_LABEL_IDS: return False
    # Só conta deals em FRANQUIA ou PARTNER (espelha filtro do painel Pipedrive)
    funil = deal_funil(a.get('deal_id'))
    if funil not in ('FRANQUIA', 'PARTNER'): return False
    return True

# Agendamento CC1 (CC1 + BOT):
#   Marcado como feito em = Maio | Status = Concluída | Criador NOT IN EXCLUDE_USERS
#   | Funil = FRANQUIA ou PARTNER (espelha filtro do painel Pipedrive)
def keep_ag(a):
    if not a.get('done'): return False
    if not _in_maio(a.get('marked_as_done_time')): return False
    criador = user_name(a.get('created_by_user_id'))
    if criador in EXCLUDE_USERS: return False
    funil = deal_funil(a.get('deal_id'))
    if funil not in ('FRANQUIA', 'PARTNER'): return False
    return True

# Reagendamento CC1:
#   Marcado como feito em = Maio | Criador NOT IN EXCLUDE_USERS
def keep_reag(a):
    if not a.get('done'): return False
    if not _in_maio(a.get('marked_as_done_time')): return False
    criador = user_name(a.get('created_by_user_id'))
    if criador in EXCLUDE_USERS: return False
    return True

# CC1 Realizada (concluídas):
#   Marcado como feito em = Maio | Status = Concluída | Tipo = CC1 Realizada
#   | Criador NOT IN EXCLUDE_USERS
#   Espelha EXATAMENTE o painel Pipedrive (sem filtro de funil, sem filtro de responsável)
def keep_re(a):
    if not a.get('done'): return False
    if not _in_maio(a.get('marked_as_done_time')): return False
    criador = user_name(a.get('created_by_user_id'))
    if criador in EXCLUDE_USERS: return False
    return True

# CC1 Para Acontecer (A fazer):
#   Data de vencimento = Maio | Status ≠ Concluída | Criador NOT IN EXCLUDE_USERS
def keep_re_fut(a):
    if a.get('done'): return False
    if not _in_maio(a.get('due_date')): return False
    criador = user_name(a.get('created_by_user_id'))
    if criador in EXCLUDE_USERS: return False
    return True

# No Show (CC1 Não realizada + variantes):
#   Marcado como feito em = Maio | Criador NOT IN EXCLUDE_USERS
#   | Responsável NOT IN SDR_LIST (no-show é sempre no side do closer)
def keep_ns(a):
    if not a.get('done'): return False
    if not _in_maio(a.get('marked_as_done_time')): return False
    responsavel = user_name(a.get('user_id'))
    if responsavel in SDR_LIST: return False
    criador = user_name(a.get('created_by_user_id'))
    if criador in EXCLUDE_USERS: return False
    return True

# Enriquece deal_by_id com TODOS os deals open/won via paginação em lote (500/página).
# (lost já foi buscado integralmente por _fetch_lost). Substitui as ~4.000 chamadas
# individuais que estouravam o rate limit (429) do Pipedrive — agora são ~20-40 páginas.
print('→ Buscando deals open/won em lote (paginado) para completar deal_by_id...')

def _fetch_deals_bulk(status):
    out = []
    start = 0
    while True:
        params = {'api_token': TOKEN, 'status': status, 'limit': 500, 'start': start,
                  'sort': 'add_time DESC'}
        url = f'{BASE}/deals?' + urllib.parse.urlencode(params)
        body = None
        for attempt in range(6):
            try:
                with urllib.request.urlopen(url, timeout=120) as r:
                    body = json.loads(r.read())
                break
            except Exception as exc:
                if attempt == 5: raise
                wait = 10 * (attempt + 1)
                print(f'     ⚠️  bulk {status} erro (tentativa {attempt+1}), esperando {wait}s... [{exc}]')
                time.sleep(wait)
        data = body.get('data') or []
        out.extend(data)
        pag = (body.get('additional_data') or {}).get('pagination') or {}
        if not pag.get('more_items_in_collection'): break
        start = pag.get('next_start', start + 500)
    return out

bulk_added = 0
for status in ('open', 'won'):
    for d in _fetch_deals_bulk(status):
        if d['id'] not in deal_by_id:
            deal_by_id[d['id']] = d
            bulk_added += 1
print(f'   +{bulk_added} deals via lote | total deal_by_id: {len(deal_by_id)}')

# Fallback: deals ainda faltando (raro) — busca individual só dos remanescentes
missing_deal_ids = {a.get('deal_id') for a in acts_raw if a.get('deal_id') and a.get('deal_id') not in deal_by_id}
missing_deal_ids.discard(None)
if missing_deal_ids:
    print(f'→ Fallback: {len(missing_deal_ids)} deals restantes (busca individual)...')
    def _fetch_deal(did):
        url = f'{BASE}/deals/{did}?api_token={TOKEN}'
        for attempt in range(3):
            try:
                with urllib.request.urlopen(url, timeout=20) as r:
                    body = json.loads(r.read())
                return did, body.get('data')
            except urllib.error.HTTPError as e:
                if e.code == 429:
                    time.sleep(2 * (attempt + 1))
                    continue
                return did, None
            except Exception:
                return did, None
        return did, None
    with concurrent.futures.ThreadPoolExecutor(max_workers=5) as executor:
        for did, d in executor.map(_fetch_deal, list(missing_deal_ids)):
            if d:
                deal_by_id[did] = d
print(f'   ✓ deal_by_id total: {len(deal_by_id)} deals')

# Aplica filtros por tipo (depois que deal_by_id está completo)
print(f'→ Aplicando filtros por tipo...')

def act_row(a):
    deal_id = a.get('deal_id')
    return {
        'Negócio - Equipe': deal_team(deal_id),
        'Negócio - Funil': deal_funil(deal_id),
        'Atividade - Usuário responsável': user_name(a.get('user_id')),
        'Atividade - Assunto': a.get('subject'),
        'Atividade - Tipo': a.get('type_name') or a.get('type'),
        'Atividade - Criador': user_name(a.get('created_by_user_id')),
        'Atividade - Status': 'Concluída' if a.get('done') else 'A fazer',
        'Atividade - Data de criação': a.get('add_time'),
        'Atividade - Data de vencimento': a.get('due_date'),
        'Negócio - Nome': deal_title(deal_id),
        'Negócio - Título': deal_title(deal_id),
        'Atividade - Data e hora de início': a.get('due_date'),
        'Atividade - Hora de início': a.get('due_time'),
        'Negócio - Etapa': None,
        'Negócio - Status': deal_status(deal_id),
        'Negócio - E-mail': deal_email(deal_id),
        'Negócio - Etiqueta': deal_etiqueta(deal_id),
        'Negócio - Motivo da perda': deal_lost_reason(deal_id),
        'Negócio - Data de perda': deal_lost_time(deal_id),
        'Pessoa - E-mail': None,
        'Negócio - Telefone': None,
        'Pessoa - Telefone': None,
        'Atividade - Marcado como feito em': a.get('marked_as_done_time'),
        'Negócio - Proprietário': user_name((deal_by_id.get(deal_id) or {}).get('user_id')),
    }

# ── Aplica filtros por tipo e separa nas 6 listas ──
# Cada tipo usa o filtro exato do painel Pipedrive correspondente.
TYPES_QUAL    = {ACT_QUAL_REAL_A, ACT_QUAL_REAL_B}
TYPES_AG      = {ACT_AG_CC1, ACT_AG_BOT}
TYPES_REAG    = {ACT_REAG_CC1}
TYPES_CC1_REAL = {ACT_CC1_REAL}
TYPES_CC1_NS  = {ACT_CC1_NS, ACT_NS_IMPREV, ACT_NS_DESINT}
TYPES_CC1_INV = {ACT_CC1_INV}

# CC1 Inválida: mesmas regras de keep_re (feita em Maio, criador não excluído)
def keep_inv(a):
    if not a.get('done'): return False
    if not _in_maio(a.get('marked_as_done_time')): return False
    criador = user_name(a.get('created_by_user_id'))
    if criador in EXCLUDE_USERS: return False
    return True

qual_rows, ag_rows, reag_rows, re_rows, re_fut_rows, ns_rows, inv_rows = [], [], [], [], [], [], []

for a in acts_raw:
    t = a.get('type')
    if t in TYPES_QUAL:
        if keep_qual(a):     qual_rows.append(act_row(a))
    elif t in TYPES_AG:
        if keep_ag(a):       ag_rows.append(act_row(a))
    elif t in TYPES_REAG:
        if keep_reag(a):     reag_rows.append(act_row(a))
    elif t in TYPES_CC1_REAL:
        if keep_re(a):       re_rows.append(act_row(a))
        elif keep_re_fut(a): re_fut_rows.append(act_row(a))
    elif t in TYPES_CC1_NS:
        if keep_ns(a):       ns_rows.append(act_row(a))
    elif t in TYPES_CC1_INV:
        if keep_inv(a):      inv_rows.append(act_row(a))

print(f'   qual={len(qual_rows)} ag={len(ag_rows)} reag={len(reag_rows)}')
print(f'   re={len(re_rows)} re_fut={len(re_fut_rows)} ns={len(ns_rows)} inv={len(inv_rows)}')

pd.DataFrame(qual_rows).to_excel(os.path.join(OUT_DIR,'qual.xlsx'), index=False)
pd.DataFrame(ag_rows).to_excel(os.path.join(OUT_DIR,'ag.xlsx'), index=False)
pd.DataFrame(reag_rows).to_excel(os.path.join(OUT_DIR,'reag.xlsx'), index=False)
pd.DataFrame(re_rows).to_excel(os.path.join(OUT_DIR,'re.xlsx'), index=False)
pd.DataFrame(re_fut_rows).to_excel(os.path.join(OUT_DIR,'re_fut.xlsx'), index=False)
pd.DataFrame(ns_rows).to_excel(os.path.join(OUT_DIR,'ns.xlsx'), index=False)
pd.DataFrame(inv_rows).to_excel(os.path.join(OUT_DIR,'inv.xlsx'), index=False)

print(f'\n✓ Dados salvos em {OUT_DIR}/')
