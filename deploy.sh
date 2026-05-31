#!/usr/bin/env bash
# Deploy automatizado: builda dashboards + push para GitHub Pages
set -e

WORKDIR="$(cd "$(dirname "$0")" && pwd)"
TOKEN_FILE="$WORKDIR/.github_token"
PUBLISH="$WORKDIR/publish"
REPO_URL="https://github.com/mauriciopeixotomps/Dash-qualificadora.git"
TMP_REPO="/tmp/deploy-repo"

if [ ! -f "$TOKEN_FILE" ]; then
    echo "✗ Token não encontrado em $TOKEN_FILE"
    exit 1
fi
TOKEN=$(cat "$TOKEN_FILE" | tr -d '[:space:]')

# 1) Build dashboards
echo "→ Build Abril..."
cd "$WORKDIR" && python3 build_from_xlsx.py > /dev/null 2>&1 && cp Dashboard_Qualificadora_Abril_2026.html publish/abril-2026.html

echo "→ Build Maio..."
cd "$WORKDIR" && python3 build_from_xlsx_maio.py > /dev/null 2>&1

echo "→ Build Junho..."
cd "$WORKDIR" && python3 build_from_xlsx_junho.py > /dev/null 2>&1

echo "→ Build Sala TV — Maio..."
cd "$WORKDIR" && python3 build_sala.py > /dev/null 2>&1 || echo "⚠ sala maio falhou (não crítico)"

echo "→ Build Sala TV — Junho..."
cd "$WORKDIR" && python3 build_sala_junho.py > /dev/null 2>&1 || echo "⚠ sala junho falhou (não crítico)"

# 2) Clone & sync
echo "→ Clonando repo..."
rm -rf "$TMP_REPO"
git clone -q "https://x-access-token:${TOKEN}@github.com/mauriciopeixotomps/Dash-qualificadora.git" "$TMP_REPO"

cd "$TMP_REPO"
cp "$PUBLISH/abril-2026.html" .
cp "$PUBLISH/maio-2026.html" .
[ -f "$PUBLISH/junho-2026.html" ] && cp "$PUBLISH/junho-2026.html" .
cp "$PUBLISH/index.html" .
cp "$PUBLISH/README.md" .
[ -f "$PUBLISH/sala.html" ]        && cp "$PUBLISH/sala.html" .
[ -f "$PUBLISH/sala-junho.html" ]  && cp "$PUBLISH/sala-junho.html" .

# 3) Commit & push (só se houver mudanças)
git -c user.name="Mauricio Peixoto" -c user.email="mauricio.peixoto.mps@gmail.com" add -A
if git diff --cached --quiet; then
    echo "= Sem mudanças. Nada a fazer."
else
    MSG="${1:-Atualiza dashboards $(date +'%d/%m/%Y %H:%M')}"
    git -c user.name="Mauricio Peixoto" -c user.email="mauricio.peixoto.mps@gmail.com" \
        commit -q -m "$MSG"
    git push -q origin main
    echo "✓ Push concluído: $MSG"
    echo "  URL: https://mauriciopeixotomps.github.io/Dash-qualificadora/"
fi
