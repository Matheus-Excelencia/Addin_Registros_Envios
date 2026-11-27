# Registra — ambiente de homologação Alpha 2.0

- Preview Vercel (alias de branch): https://addin-registros-envios-git-feat-alpha-20-acme-vs-coyote.vercel.app/
- Demo fictícia: https://addin-registros-envios-git-feat-alpha-20-acme-vs-coyote.vercel.app/demo/index.html
- Manifesto separado: [manifest-homologacao.xml](manifest-homologacao.xml).
- Identificador Office exclusivo: c5a78e19-6792-4e98-9c25-c67eb6e7db21. Grupo Excelência - Teste; botão Novo registro (teste).

## Instalação controlada
1. Acesse o preview com sua conta autorizada Vercel. A proteção de deploy permanece ativa; confirme que o Excel Web consegue carregar iframe e ícones. A autenticação no navegador não garante acesso no iframe. Não insira tokens de bypass em XML, URLs ou código.
2. Use uma cópia de teste do workbook com abas Email, SMS e Config, somente dados fictícios, nunca a planilha operacional. Baixe o XML desta pasta no GitHub e carregue em Início > Suplementos > Carregar meu suplemento (se permitido pelo Microsoft 365).
3. Abra Excelência - Teste > Novo registro (teste). O aviso HOMOLOGAÇÃO e identificação manual devem aparecer. A versão de teste não pede token Microsoft SSO; não valida autenticação de produção.
4. Execute [roteiro de homologação](../docs/ALPHA-20-HOMOLOGACAO.md) e registre commit, navegador, resultado e evidência. Não testar em dados reais. Remova apenas a instalação de teste ao finalizar.

## SSO e proteção
O WebApplicationInfo de produção não pode ser reutilizado com o domínio de preview. O teste exige app Entra separado, Resource e consentimento configurados pelo responsável autorizado para validar SSO real em outra origem. Nesta fase não há credenciais fictícias nem alteração no app de produção. Não desativar Deployment Protection globalmente. Se o Office não carregar o preview protegido, avaliar acesso restrito específico com administrador; não promover o preview para produção.

O manifesto de homologação fica fora de public/ para não ser publicado no site oficial. O alias acompanha a branch; para evidência de teste use o SHA do deploy registrado. A URL de produção e o manifesto oficial permanecem os mesmos.
