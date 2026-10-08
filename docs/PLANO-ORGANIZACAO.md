# Mapa de dependências e plano de reorganização

## Objetivo
Organizar o código sem alterar o comportamento funcional do suplemento. A reorganização começa exclusivamente em `alpha-2.0-dev`; `main` e `alpha-1.05` permanecem intocadas até validação.

## Dependências atuais verificadas
- `manifest.xml` publica `/taskpane.html`, `/commands.html` e ícones em `/assets/`.
- `taskpane.html` carrega `taskpane.css` e `taskpane.js`.
- `taskpane.js` navega para `novo.html`, `historico.html`, `resumo.html` e `config.html`.
- `novo.html` carrega `novo.css` e `novo.js`; `novo.js` navega para `email.html` e `sms.html`.
- `email.html`, `sms.html`, `historico.html`, `resumo.html` e `config.html` carregam seus arquivos CSS/JS pelo nome relativo na raiz.
- `email.js`, `sms.js`, `historico.js`, `resumo.js` e `config.js` usam APIs do Excel/Office. O histórico também navega entre páginas e usa armazenamento local para duplicação.
- Não foi encontrado `vercel.json` nem `package.json` na raiz da branch consultada. Portanto, não assumir que existe configuração de roteamento/build personalizada.

## Estrutura-alvo proposta
```text
/
├── README.md
├── manifest.xml                 # manifesto estável, manter na raiz até validação
├── taskpane.html                # endpoints atuais mantidos durante a transição
├── commands.html
├── assets/
├── app/
│   ├── pages/
│   │   ├── taskpane.html
│   │   ├── novo.html
│   │   ├── email.html
│   │   ├── sms.html
│   │   ├── historico.html
│   │   ├── resumo.html
│   │   └── config.html
│   ├── styles/
│   └── scripts/
├── alpha2-prototype/
└── docs/
    ├── PLANO-ORGANIZACAO.md
    └── CHECKLIST-VALIDACAO.md
```

## Estratégia segura
1. Primeiro adicionar documentação e testes de referência na branch de desenvolvimento.
2. Fazer a migração em etapas pequenas, preservando temporariamente os endpoints que o manifesto usa.
3. Atualizar links de páginas, referências CSS/JS, ícones e recursos; procurar também URLs hardcoded e `window.location`.
4. Não gerar um manifesto de teste que use o GUID ou a URL de produção.
5. Validar deploy preview próprio antes de importar o manifesto de teste no Excel.
6. Só depois de teste manual no Excel, replicar a organização para cada branch, mantendo o conteúdo próprio daquela versão.
7. Produção e branches estáveis só serão alteradas por último, após confirmação dos testes.

## Bloqueios atuais
- A criação de um projeto Vercel separado para Alpha 2.0 retornou 403; ainda não existe URL de preview independente confirmada.
- Sem preview próprio, não é seguro validar o manifesto do Office nem afirmar que a migração funciona ponta a ponta.
- Este documento é um mapa inicial baseado nos arquivos atualmente presentes em `alpha-2.0-dev`; ele deve ser atualizado conforme os testes revelarem outras referências.
