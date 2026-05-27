#!/usr/bin/env python3
"""
Aplica correções mobile + datalabels nos HTMLs gerados automaticamente.

Uso local (rodar após build_from_xlsx.py):
    python patch_html.py

As correções são idempotentes — podem ser aplicadas múltiplas vezes sem problema.
"""
import re, sys, pathlib

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


if __name__ == '__main__':
    repo = pathlib.Path(__file__).parent
    targets = sorted(repo.glob('*-2026.html'))
    if not targets:
        print('Nenhum arquivo *-2026.html encontrado.', file=sys.stderr)
        sys.exit(1)
    changed = sum(patch_file(f) for f in targets)
    print(f'\n{changed}/{len(targets)} arquivo(s) atualizado(s).')
