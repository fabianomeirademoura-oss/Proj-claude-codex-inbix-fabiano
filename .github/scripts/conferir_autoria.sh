#!/usr/bin/env bash
# Checagem de autoria em cada pull request (COORDENACAO.md, regra 11).
# Repete no GitHub a regra do gancho .githooks/commit-msg, que pode ser pulado localmente.
# Variáveis: RAMO (branch do PR), BASE e TOPO (commits de início e fim do PR).
set -euo pipefail

case "$RAMO" in
    claude/*)      agente="Claude Code"; autor="Claude Code (agente)" ;;
    codex/*)       agente="Codex";       autor="Codex (agente)" ;;
    coordenador/*) agente="";            autor="" ;;
    *) echo "::error::A branch '$RAMO' não segue o padrão claude/<tarefa>, codex/<tarefa> ou coordenador/<tarefa> (COORDENACAO.md, regra 11)."; exit 1 ;;
esac

total=0; falhas=0
for c in $(git rev-list --no-merges --reverse "$BASE..$TOPO"); do
    total=$((total + 1))
    a=$(git log -1 --format='%an' "$c")
    t=$(git log -1 --format='%B' "$c" | git interpret-trailers --parse | sed -n 's/^Agente: *//p' | tail -n 1)
    s=$(git log -1 --format='%h %s' "$c")
    if [ -z "$autor" ]; then
        # coordenador/*: commit humano, sem agente
        if [[ "$a" == *"(agente)" || -n "$t" ]]; then
            echo "::error::$s: numa branch coordenador/ o commit não pode ser de agente (autor '$a', Agente '${t:-nenhuma}')."
            falhas=$((falhas + 1)); continue
        fi
    elif [ "$a" != "$autor" ] || [ "$t" != "$agente" ]; then
        echo "::error::$s: autor '$a' e linha 'Agente: ${t:-nenhuma}'; esperado autor '$autor' e 'Agente: $agente'."
        falhas=$((falhas + 1)); continue
    fi
    echo "ok   $s  ($a)"
done

if [ "$total" -eq 0 ]; then echo "::error::O pull request não tem commits."; exit 1; fi
if [ "$falhas" -gt 0 ]; then echo "::error::$falhas de $total commits fora da regra de autoria."; exit 1; fi
echo "Todos os $total commits de '$RAMO' identificados: ${agente:-coordenador}."
