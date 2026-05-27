# Dashboard Qualificadora — GRUPO STUDIO

Painel de acompanhamento mensal da equipe de Qualificação/Pré-venda.

## Páginas
- [Abril 2026](abril-2026.html) — fechado
- [Maio 2026](maio-2026.html) — em andamento

## Como atualizar
1. Exportar planilhas do Pipedrive Insights (deals + atividades).
2. Rodar `build_from_xlsx.py` apontando para os xlsx atualizados.
3. Copiar o HTML gerado pra esta pasta.
4. *(Opcional local)* Rodar `python patch_html.py` para aplicar correções mobile.
5. Commit + push.

> **Automático:** o GitHub Actions (`auto-patch.yml`) detecta qualquer push de
> `*-2026.html` no `main` e aplica as correções automaticamente, sem precisar
> rodar o patch manualmente.

## Deploy
GitHub Pages — basta habilitar Pages nas configurações do repositório (branch `main`, pasta raiz `/`).
