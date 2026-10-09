# Plano conjunto — Alpha 2.0 + Alpha 2.5 (suplemento e site)

## Objetivo

Desenvolver as duas experiências em paralelo, compartilhando regras de negócio, modelos, validações e histórico sempre que tecnicamente possível:

- **Alpha 2.0 — suplemento Excel:** continua funcionando dentro da pasta de trabalho, usando Office.js para ler/gravar no workbook aberto.
- **Alpha 2.5 — site:** permite entrar pelo navegador, identificar-se quando houver autenticação configurada, selecionar um modelo e uma planilha vinculada, e executar os mesmos fluxos de registro sem precisar abrir o Excel manualmente.

A versão web não substitui o suplemento. Os dois canais devem coexistir e manter a mesma experiência funcional conforme as permissões e capacidades de cada canal.

## Regra essencial: uma fonte de verdade

Modelos, versões, campos, regras condicionais, mapeamentos e permissões devem ser definidos uma vez e consumidos pelos dois canais. Evitar manter uma implementação independente da lógica no site e outra no suplemento.

O registro operacional pode continuar na planilha escolhida. Configurações compartilhadas, vínculos de planilhas, identidade e auditoria devem ficar em armazenamento central protegido, após a escolha e configuração de um backend autorizado.

## Experiência prevista para o site Alpha 2.5

1. Abrir a página principal do site (o endpoint `/` do Preview hoje retorna 404 e precisa ser tratado).
2. Entrar ou autenticar-se, caso a autenticação seja implementada. Não simular identidade nem tratar um login visual como segurança real.
3. Ver apenas os modelos/planilhas autorizados para a pessoa.
4. Selecionar um modelo.
5. Selecionar a planilha de destino vinculada ao modelo.
6. Se a planilha já estiver cadastrada, reutilizar o vínculo e o link previamente salvo; não criar um vínculo duplicado nem pedir a URL novamente sem necessidade.
7. Mostrar o estado do vínculo: validado, precisa reconectar, sem permissão ou requer revisão de mapeamento.
8. Abrir o mesmo menu funcional do suplemento: Novo Registro, Histórico, Resumo e Configurações, respeitando as permissões.
9. Oferecer uma ação clara **Abrir planilha** usando o link registrado, em nova aba.
10. Registrar e exibir sucesso/erro de forma verificável; não mostrar sucesso antes de confirmar a gravação.

Se o usuário escolher uma planilha não cadastrada, oferecer um fluxo para registrar o vínculo, validar acesso e confirmar o mapeamento antes de permitir gravações.

## Limite técnico que precisa ser resolvido antes da implementação

O suplemento usa `Office.js` e opera sobre a pasta de trabalho aberta no Excel. Uma página comum no navegador não tem automaticamente acesso às APIs `Excel.run` nem permissão para editar qualquer planilha só porque conhece seu link.

Para o site gravar no mesmo arquivo, precisamos escolher e validar uma integração suportada, por exemplo Microsoft Graph com autenticação e permissões adequadas, ou outro fluxo explicitamente aprovado. Não assumir que a conta terá permissões administrativas do Microsoft 365/SharePoint. Se a integração não puder ser autorizada, o site deve informar a limitação e oferecer um fluxo alternativo aprovado, sem alegar paridade completa.

## Recursos compartilhados planejados

- Catálogo de modelos e versões publicadas.
- Construtor/configuração de campos, tipos, opções e obrigatoriedade.
- Regras condicionais (ex.: resposta “Sim” exibe campos adicionais).
- Validação de dados, incluindo e-mail e CPF quando aplicável.
- Mapeamento explícito entre campos e cabeçalhos da planilha.
- Detecção de cabeçalhos ausentes, duplicados ou alterados; interromper a gravação e pedir revisão.
- Registro de E-mail e SMS.
- Histórico, filtros, resumo e configurações.
- Vínculos de planilha com URL, identificador estável quando disponível, modelo, estado de validação e data da última conferência.
- Auditoria mínima, com identidade verificada quando disponível e sem expor dados pessoais desnecessários.
- Controle de acesso no backend; ocultar botões na interface não substitui autorização real.

## Pendências por frente

### Alpha 2.0 — suplemento Excel
- [ ] Confirmar manifesto de teste independente: novo GUID e todas as URLs apontando para ambiente de teste.
- [ ] Validar formalmente XML e recursos do manifesto.
- [ ] Corrigir ou definir o comportamento da rota raiz `/` no Preview.
- [ ] Testar Office.js dentro do Excel com o manifesto de teste, sem usar o manifesto estável.
- [ ] Testar leitura e gravação em cópia de planilha, mapeamento de cabeçalhos e prevenção de sobrescrita.
- [ ] Finalizar o primeiro formulário dinâmico e regras condicionais.
- [ ] Testar E-mail, SMS, Histórico, Resumo e Configurações funcionalmente.
- [ ] Verificar navegação, filtros, validações, mensagens de erro e console.
- [ ] Preservar a Alpha 1.05 e a produção; nenhuma promoção sem aprovação explícita.

### Alpha 2.5 — site
- [ ] Definir e implementar a página inicial em `/` (seleção de modelo/planilha e acesso aos módulos).
- [ ] Definir autenticação real e papéis de usuário; não criar login meramente visual como proteção.
- [ ] Selecionar backend e armazenamento de configurações com custo, permissões e políticas de acesso conhecidos.
- [ ] Definir como o site acessará arquivos Excel: Microsoft Graph/autorização adequada ou alternativa aprovada.
- [ ] Criar catálogo de planilhas vinculadas e reutilizar o link salvo quando já existir vínculo.
- [ ] Validar link, acesso e identidade do arquivo antes de habilitar gravações.
- [ ] Implementar os mesmos módulos e regras de negócio compartilhados com a Alpha 2.0.
- [ ] Criar testes para vínculo existente, vínculo ausente, link inválido, permissão removida e cabeçalho alterado.
- [ ] Testar concorrência e erros de gravação; nunca indicar sucesso sem confirmação.
- [ ] Publicar em Preview e validar com dados fictícios/cópias de planilhas.

### Trabalho conjunto (não esperar uma versão terminar para começar a outra)
- [ ] Definir o contrato compartilhado de modelo, campos, regras, vínculo e versão.
- [ ] Separar regras de negócio da interface para que o suplemento e o site reutilizem o mesmo comportamento sempre que possível.
- [ ] Definir a matriz de capacidades: o que funciona em ambos, o que depende de Excel aberto e o que exige permissão web.
- [ ] Construir e testar um módulo de cada vez nos dois canais.
- [ ] Manter dados de teste separados dos dados reais.

## Sequência de entrega paralela

1. **Fundação comum:** modelos, campos, validações, versão e vínculo de planilha.
2. **Primeiro fluxo vertical:** escolher modelo → escolher/vincular planilha → preencher formulário → validar → gravar em destino de teste → confirmar resultado.
3. **Repetir o fluxo nos dois canais:** suplemento com Office.js e site com a integração de arquivo escolhida.
4. **Módulos compartilhados:** E-mail/SMS, depois Histórico e Resumo, depois Configurações.
5. **Autenticação, permissões e auditoria:** antes de armazenar ou expor configurações compartilhadas a usuários.
6. **Testes ponta a ponta e Preview:** planilhas de teste, cenários de erro e revisão de segurança.
7. **Revisão de compatibilidade das branches:** preservar diferenças históricas; estável/produção por último e só com aprovação explícita.

## Critérios de aceite

- O usuário seleciona um modelo e uma planilha vinculada sem ter que redigitar o link quando o vínculo já existe.
- O botão **Abrir planilha** usa o link salvo e não confunde abrir o arquivo com gravar nele.
- O site e o suplemento usam a mesma versão do modelo e as mesmas regras de validação.
- Gravações só ocorrem após validar o vínculo, permissão e mapeamento de cabeçalhos.
- Mudanças de estrutura não escrevem em colunas erradas; a operação para e pede revisão.
- Os fluxos de sucesso e erro são reais e testados com dados fictícios.
- Alpha 1.05 e produção permanecem intactas durante desenvolvimento e validação.

## Estado atual confirmado

- Branch de trabalho: `alpha-2.0-dev`.
- O usuário informou que todas as telas do Preview abriram normalmente.
- Requisições HTTP ao Preview confirmaram `taskpane.html`, `manifest.xml`, CSS/JS locais e ícones com resposta 200.
- A rota raiz `/` do Preview responde 404; precisa ser corrigida para a entrada web Alpha 2.5.
- O manifesto servido no Preview ainda aponta para URLs de produção e usa o GUID estável; não deve ser importado no Excel para testes Alpha 2.0.
- A interface protótipo em `alpha2-prototype/` ainda usa dados fictícios e não persiste nem grava em Excel.
- Não foi ainda comprovada a gravação funcional no Excel nem a edição de uma pasta de trabalho a partir do site.
