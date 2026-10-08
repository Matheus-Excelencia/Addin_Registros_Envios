# Alpha 2.0 — Protótipo inicial

## Localização

- Arquivo: `alpha2-prototype/index.html`
- Branch: `alpha-2.0-dev`
- Visualização no GitHub: https://github.com/Matheus-Excelencia/Addin_Registros_Envios/blob/alpha-2.0-dev/alpha2-prototype/index.html

## O que já demonstra

- Identificação visual explícita de ambiente de teste;
- Modelo de cadastro “Responsáveis dos credores”;
- Campos definidos por uma configuração de modelo no JavaScript;
- Campos obrigatórios;
- Validação de formato de e-mail;
- Máscara e validação de CPF por dígitos verificadores;
- Regra condicional: ao selecionar “Testemunha do credor? = Sim”, exibe nome completo, e-mail e CPF da testemunha;
- Prévia de validação, com CPF parcialmente mascarado;
- Botão para limpar o formulário.

## Limitações intencionais

- Não carrega dados reais das planilhas de exemplo;
- Não grava no Excel, SharePoint ou backend;
- Não usa SSO nem afirma uma identidade de usuário;
- Não persiste dados depois de fechar ou atualizar a página;
- Não é ainda um construtor visual de formulários: os campos do modelo são definidos na estrutura de configuração do protótipo;
- Não deve ser usado com dados pessoais reais.

## Próxima etapa

1. Verificar o código e validar a experiência com dados fictícios.
2. Publicar uma implantação Preview isolada, se o projeto Vercel permitir sem alterar produção.
3. Confirmar se a proteção SSO da Vercel permite carregar a página dentro do Excel Web.
4. Somente com URL Preview verificada, criar manifesto separado com GUID novo e todos os endpoints apontando exclusivamente para o Preview.
5. Implementar o primeiro construtor de modelo e, depois, o mapeamento seguro para uma cópia de workbook.

## Proteção da versão estável

Nenhuma alteração foi feita em `main`, `alpha-1.05`, no manifesto estável ou nos endpoints de produção por este protótipo.