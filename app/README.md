# Código da aplicação

O código executável está separado por tipo de arquivo para facilitar manutenção.

- `pages` (em preparação): páginas HTML. Por enquanto, os HTMLs permanecem na raiz para manter as URLs públicas usadas pelo manifesto e pelos links de navegação.
- `styles/`: folhas CSS carregadas pelas páginas HTML.
- `scripts/`: JavaScript carregado pelas páginas HTML.

## Regra de migração
Não mover HTMLs nem alterar URLs de entrada do manifesto até validar as rotas em Preview independente. A navegação JavaScript usa caminhos relativos às URLs das páginas (por exemplo, `taskpane.html`, `email.html`), então esses endpoints precisam continuar disponíveis.