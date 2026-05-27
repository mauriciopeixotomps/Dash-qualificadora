#!/usr/bin/env python3
"""
Aplica correções mobile + datalabels nos HTMLs gerados automaticamente,
e sincroniza os dados de maio-2026.html para sala.html.

Uso local (rodar após build_from_xlsx.py):
    python patch_html.py

As correções são idempotentes — podem ser aplicadas múltiplas vezes sem problema.
"""
import re, sys, pathlib, json, datetime

# Marcador único que indica que o CSS mobile já foi injetado neste arquivo.
_CSS_MARKER = '/* PATCH:mobile-labels-v3 */'

# Bloco CSS a ser injetado antes do </style> principal.
_MOBILE_CSS = f'''
  {_CSS_MARKER}
  @media (max-width: 640px) {{
    body {{ font-size: 12px; }}
    .container {{ padding: 12px 10px; width: 100%; }}
    h1 {{ font-size: 17px; }}
    .subtitle {{ font-size: 12px; margin-bottom: 14px; }}
    .last-updated {{ font-size: 11px; padding: 5px 10px; }}
    .filter-bar {{ flex-wrap: wrap; gap: 5px; padding: 8px 10px; }}
    .filter-bar input[type="date"] {{ display: none; }}
    .filter-bar .sep, .filter-bar .label {{ display: none; }}
    .filter-bar #filter-info {{ width: 100%; font-size: 11px; }}
    .filter-bar button {{ padding: 5px 10px; font-size: 11px; }}
    .month-nav {{ gap: 6px; flex-wrap: wrap; }}
    .month-nav a {{ padding: 4px 10px; font-size: 12px; }}
    .grid.grid-5 {{ grid-template-columns: repeat(2, 1fr) !important; gap: 8px; }}
    .kpi {{ padding: 12px 10px; }}
    .kpi-value {{ font-size: 22px; }}
    .kpi-label, .kpi-sub {{ font-size: 10px; }}
    .card {{ padding: 12px 10px; }}
    .card h3 {{ font-size: 12px; }}
    .section-title {{ font-size: 12px; margin: 20px 0 8px; }}
    .canvas-wrap       {{ height: 220px; }}
    .canvas-wrap.tall  {{ height: 280px; }}
    .canvas-wrap.short {{ height: 180px; }}
    table {{ font-size: 11px; }}
    th, td {{ padding: 6px; }}
    .grid-5, .grid-4, .grid-3, .grid-2 {{ grid-template-columns: 1fr !important; }}
    .person-row {{ grid-template-columns: 26px 1fr auto auto !important; gap: 6px !important; }}
    .person-bar-track {{ display: none !important; }}
    .person-stat {{ min-width: 0 !important; font-size: 11px; }}
    .person-gap {{ font-size: 11px; }}
  }}
'''


def _apply_mobile_css(html: str) -> str:
    """Injeta o bloco @media mobile antes do primeiro </style>."""
    # Já foi injetado pelo marcador novo?
    if _CSS_MARKER in html:
        return html
    # Já tem o fix de person-row do patch manual anterior? Não duplicar.
    if 'grid-template-columns: 26px 1fr auto auto !important' in html:
        return html
    return html.replace('</style>', _MOBILE_CSS + '\n</style>', 1)


def _apply_datalabels_fix(html: str) -> str:
    """Garante que cumDl e labelDl mostrem label apenas no último ponto."""
    # Padrão original (gerado pelo build_from_xlsx.py)
    html = html.replace(
        'display: ctx => ctx.dataIndex <= lastIdx && ctx.dataset.data[ctx.dataIndex] > 0,',
        'display: ctx => ctx.dataIndex === lastIdx && ctx.dataset.data[ctx.dataIndex] > 0,',
    )
    # Padrão multi-linha (versão de transição com if/return)
    pattern = (
        r'display: ctx => \{\s*'
        r'if \(ctx\.dataIndex > lastIdx \|\| ctx\.dataset\.data\[ctx\.dataIndex\] <= 0\) return false;\s*'
        r'return ctx\.dataIndex === lastIdx.*?;\s*\},'
    )
    repl = 'display: ctx => ctx.dataIndex === lastIdx && ctx.dataset.data[ctx.dataIndex] > 0,'
    html = re.sub(pattern, repl, html, flags=re.DOTALL)
    return html


def _apply_stat_compact(html: str) -> str:
    """Substitui o texto longo de person-stat pelo formato compacto."""
    # Versão longa gerada pelo build script (maio)
    old = (
        '<div class="person-stat" style="min-width:130px; text-align:right;">\n'
        '        <div><span style="color:#22c55e; font-weight:bold;">${r.re}</span>'
        ' realizadas / ${goalFull}</div>\n'
        '        <div style="font-size:11px; color:#a855f7;">+${r.future} para acontecer</div>\n'
        '      </div>'
    )
    new = (
        '<div class="person-stat">\n'
        '        <div><span style="color:#22c55e; font-weight:bold;">${r.re}</span>'
        '<span style="color:#64748b;"> / ${goalFull}</span></div>\n'
        '        <div style="color:#a855f7;">+${r.future} p.</div>\n'
        '      </div>'
    )
    return html.replace(old, new)


def patch_file(path: pathlib.Path) -> bool:
    html = path.read_text(encoding='utf-8')
    patched = html
    patched = _apply_mobile_css(patched)
    patched = _apply_datalabels_fix(patched)
    patched = _apply_stat_compact(patched)
    if patched != html:
        path.write_text(patched, encoding='utf-8')
        print(f'  ✔  Patched: {path.name}')
        return True
    print(f'  –  Skipped (already up-to-date): {path.name}')
    return False


# ── Sincronização sala.html ────────────────────────────────────────────────

_SALA_JS_MARKER = '/* PATCH:sala-autosync-v1 */'

_FETCH_MAIO_IIFE = r"""
// Aplica D fresco do sessionStorage (depositado pelo fetch anterior)
/* PATCH:sala-autosync-v1 */
(function applyStoredData(){
  try {
    const raw = sessionStorage.getItem('_salaD');
    if (!raw) return;
    const f = JSON.parse(raw);
    D = f.D; today = f.today; LAST_UPDATE = f.LAST_UPDATE;
    sessionStorage.removeItem('_salaD');
  } catch(e) {}
})();

// Busca D mais recente de maio-2026.html em background
(function fetchMaioData(){
  fetch('maio-2026.html?t=' + Date.now(), { cache: 'no-store' })
    .then(r => r.text())
    .then(src => {
      const mD = src.match(/^const D\s*=\s*(.+);\s*$/m);
      const mT = src.match(/const today\s*=\s*'([^']+)'/);
      if (!mD || !mT) return;
      let maioData;
      try { maioData = JSON.parse(mD[1]); } catch(e) { return; }
      const maioRe = maioData.totals && maioData.totals.re;
      const curRe  = D.totals && D.totals.re;
      if (mT[1] === today && maioRe === curRe) return;
      const n = new Date();
      const lu = [String(n.getDate()).padStart(2,'0'), String(n.getMonth()+1).padStart(2,'0'), n.getFullYear()].join('/')
               + ' ' + [String(n.getHours()).padStart(2,'0'), String(n.getMinutes()).padStart(2,'0')].join(':');
      sessionStorage.setItem('_salaD', JSON.stringify({
        D: maioData, today: mT[1], LAST_UPDATE: lu
      }));
      location.reload(true);
    })
    .catch(() => {});
})();
"""

_NEW_CHECK_RELOAD = """// ── AUTO-REFRESH: compara D.totals.re de maio-2026.html a cada 2 min ──
function checkAndReload(){
  fetch('maio-2026.html?t=' + Date.now(), { cache: 'no-store' })
    .then(r => r.text())
    .then(src => {
      const mD = src.match(/^const D\s*=\s*(.+);\s*$/m);
      const mT = src.match(/const today\s*=\s*'([^']+)'/);
      if (!mD || !mT) return;
      let maioData;
      try { maioData = JSON.parse(mD[1]); } catch(e) { return; }
      const maioRe = maioData.totals && maioData.totals.re;
      const curRe  = D.totals && D.totals.re;
      if (mT[1] === today && maioRe === curRe) return;
      const n = new Date();
      const lu = [String(n.getDate()).padStart(2,'0'), String(n.getMonth()+1).padStart(2,'0'), n.getFullYear()].join('/')
               + ' ' + [String(n.getHours()).padStart(2,'0'), String(n.getMinutes()).padStart(2,'0')].join(':');
      sessionStorage.setItem('_salaD', JSON.stringify({
        D: maioData, today: mT[1], LAST_UPDATE: lu
      }));
      location.reload(true);
    })
    .catch(() => {})
    .finally(() => setTimeout(checkAndReload, 120000));
}"""


def _apply_sala_js_fixes(html: str) -> str:
    """
    Injeta applyStoredData + fetchMaioData e corrige checkAndReload.
    Idempotente via _SALA_JS_MARKER.
    """
    if _SALA_JS_MARKER in html:
        return html

    # Converte const D/DS/today/LAST_UPDATE → let (para poder reatribuir)
    html = re.sub(r'\bconst (D\b|DS\b|today\b|LAST_UPDATE\b)', r'let \1', html)

    # Insere applyStoredData + fetchMaioData antes de "// ── HELPERS"
    anchor = '// ── HELPERS'
    if anchor in html:
        html = html.replace(anchor, _FETCH_MAIO_IIFE + anchor, 1)

    # Substitui checkAndReload antigo (monitora sala.html) pelo novo (monitora maio).
    # O padrão captura até o setTimeout que segue a função para não deixar restos.
    # Usa lambda porque o replacement contém \s e outros chars especiais.
    new_reload = _NEW_CHECK_RELOAD + '\nsetTimeout(checkAndReload, 120000);'
    html = re.sub(
        r'// ── AUTO-REFRESH:.*?^setTimeout\(checkAndReload,\s*\d+\);',
        lambda _: new_reload,
        html, count=1, flags=re.DOTALL | re.MULTILINE
    )

    return html

def sync_sala(repo: pathlib.Path) -> bool:
    """
    Aplica JS fixes (idempotente) e sincroniza D/today de maio-2026.html para sala.html.
    Usa today + D.totals.re como indicador de mudança de dados (maio não tem LAST_UPDATE).
    """
    maio = repo / 'maio-2026.html'
    sala = repo / 'sala.html'
    if not maio.exists() or not sala.exists():
        return False

    maio_html = maio.read_text(encoding='utf-8')
    sala_html = sala.read_text(encoding='utf-8')
    original  = sala_html

    # 1. Sempre aplica os fixes de JS (idempotente via marcador)
    sala_html = _apply_sala_js_fixes(sala_html)

    m_D  = re.search(r'^const D\s*=\s*(.+);\s*$', maio_html, re.MULTILINE)
    m_DS = re.search(r'^const DS\s*=\s*(.+);\s*$', maio_html, re.MULTILINE)
    m_T  = re.search(r"const today\s*=\s*'([^']+)'", maio_html)

    if not m_D or not m_T:
        print('  !  sync_sala: não encontrou D/today em maio-2026.html')
        if sala_html != original:
            sala.write_text(sala_html, encoding='utf-8')
            return True
        return False

    new_today = m_T.group(1)

    try:
        maio_re = json.loads(m_D.group(1))['totals']['re']
    except Exception:
        maio_re = None

    cur_D_m = re.search(r'(?:const|let) D\s*=\s*(\{.+\});\s*$', sala_html, re.MULTILINE)
    try:
        sala_re = json.loads(cur_D_m.group(1))['totals']['re'] if cur_D_m else None
    except Exception:
        sala_re = None

    cur_T = re.search(r"(?:const|let) today\s*=\s*'([^']+)'", sala_html)
    data_in_sync = (
        cur_T and cur_T.group(1) == new_today and
        maio_re is not None and sala_re is not None and maio_re == sala_re
    )

    # 2. Sincroniza dados se necessário
    if not data_in_sync:
        sala_html = re.sub(
            r'(?:const|let) D\s*=\s*.+;\s*$',
            f'let D  = {m_D.group(1)};',
            sala_html, count=1, flags=re.MULTILINE
        )
        if m_DS:
            sala_html = re.sub(
                r'(?:const|let) DS\s*=\s*.+;\s*$',
                f'let DS = {m_DS.group(1)};',
                sala_html, count=1, flags=re.MULTILINE
            )
        sala_html = re.sub(
            r"(?:const|let) today\s*=\s*'[^']+'",
            f"let today      = '{new_today}'",
            sala_html, count=1
        )
        new_lu = datetime.datetime.now().strftime('%d/%m/%Y %H:%M')
        sala_html = re.sub(
            r"(?:const|let) LAST_UPDATE\s*=\s*'[^']+'",
            f"let LAST_UPDATE = '{new_lu}'",
            sala_html, count=1
        )
        print(f'  ✔  sala.html dados → today: {new_today}, re: {maio_re}, LAST_UPDATE: {new_lu}')
    else:
        print(f'  –  sala.html dados já sincronizados (today={new_today}, re={maio_re})')

    if sala_html != original:
        sala.write_text(sala_html, encoding='utf-8')
        return True
    return False


if __name__ == '__main__':
    repo = pathlib.Path(__file__).parent

    targets = sorted(repo.glob('*-2026.html'))
    if not targets:
        print('Nenhum arquivo *-2026.html encontrado.', file=sys.stderr)
        sys.exit(1)
    changed = sum(patch_file(f) for f in targets)
    print(f'\n{changed}/{len(targets)} arquivo(s) *-2026.html atualizado(s).')

    print()
    sync_sala(repo)
