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

def sync_sala(repo: pathlib.Path) -> bool:
    """
    Copia D (e DS, se presente) de maio-2026.html para sala.html.
    Usa today + D.totals.re como indicador de mudança (maio não tem LAST_UPDATE).
    Gera um novo LAST_UPDATE com o timestamp atual ao sincronizar.
    """
    maio = repo / 'maio-2026.html'
    sala = repo / 'sala.html'
    if not maio.exists() or not sala.exists():
        return False

    maio_html = maio.read_text(encoding='utf-8')
    sala_html = sala.read_text(encoding='utf-8')

    m_D  = re.search(r'^const D\s*=\s*(.+);\s*$', maio_html, re.MULTILINE)
    m_DS = re.search(r'^const DS\s*=\s*(.+);\s*$', maio_html, re.MULTILINE)
    m_T  = re.search(r"const today\s*=\s*'([^']+)'", maio_html)

    if not m_D or not m_T:
        print('  !  sync_sala: não encontrou D/today em maio-2026.html')
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
    already_synced = (
        cur_T and cur_T.group(1) == new_today and
        maio_re is not None and sala_re is not None and maio_re == sala_re
    )
    if already_synced:
        print(f'  –  sala.html já sincronizado (today={new_today}, re={maio_re})')
        return False

    original = sala_html

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

    if sala_html != original:
        sala.write_text(sala_html, encoding='utf-8')
        print(f'  ✔  sala.html sincronizado → today: {new_today}, re: {maio_re}, LAST_UPDATE: {new_lu}')
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
