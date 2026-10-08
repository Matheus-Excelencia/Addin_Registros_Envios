# Alpha 2.0 — Desenvolvimento

A Alpha 2.0 será desenvolvida na branch `alpha-2.0-dev`, criada a partir de `alpha-1.05`. A Alpha 1.05 permanece isolada e estável.

## Primeiro marco: base configurável
- [ ] Inventariar a arquitetura e os pontos de entrada da Alpha 1.05 sem modificá-la.
- [ ] Definir estrutura de modelo versionado.
- [ ] Criar ambiente de teste e dados fictícios.
- [ ] Construir primeiro formulário dinâmico sem dependência de SharePoint administrativo.
- [ ] Demonstrar campos condicionais: “Testemunha do credor?” → Sim exibe nome, e-mail e CPF.
- [ ] Validar CPF por dígitos verificadores.
- [ ] Validar a gravação em uma planilha de teste, por mapeamento explícito de cabeçalho.
- [ ] Só depois selecionar/confirmar armazenamento de configuração e autenticação.

## Regras de segurança do desenvolvimento
- Nunca trabalhar diretamente na branch estável.
- Não usar dados pessoais reais das planilhas anexadas em testes.
- Não gravar na planilha operacional original durante os primeiros testes.
- Não criar recursos pagos sem verificar custo e confirmar com o usuário.
- Não exigir acesso administrativo ao SharePoint.
- Não incluir credenciais, tokens ou chaves privadas no repositório.
- Não publicar a versão de desenvolvimento no endereço de produção sem autorização explícita.

## Critério de conclusão do primeiro marco
Um protótipo abre um modelo configurado, mostra/oculta os campos condicionais corretamente, valida os dados e salva em destino de teste sem sobrescrever colunas não mapeadas.