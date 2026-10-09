# Roadmap geral — Registros de Envios
**Projeto:** Frm cadastro / Add-in Registros de Envios  
**Documento:** escopo consolidado Alpha 2.0 + Alpha 2.5  
**Data:** 08/10/2026  
**Status:** planejamento; não autoriza publicação em produção.

## 1. Visão do produto

Manter duas formas de acesso ao mesmo produto, com experiência e regras de negócio consistentes:

1. **Suplemento do Excel (Alpha 2.0):** continua disponível dentro do Excel, abrindo o painel lateral e preservando o fluxo já existente.
2. **Site (proposta Alpha 2.5):** permite acessar o sistema pelo navegador sem precisar abrir o Excel Online. O usuário entra, escolhe a planilha/modelo ou contexto de trabalho e utiliza as funções de registro no site.
3. **Abertura opcional da planilha:** quando o usuário quiser trabalhar diretamente no Excel, poderá informar ou selecionar o link da planilha e usar uma ação explícita para abri-la.

As duas interfaces devem compartilhar regras, nomenclaturas e, quando tecnicamente possível, o mesmo serviço de dados. O site não deve ser apenas uma cópia visual do painel: ele precisa ter fluxo próprio de entrada, seleção de contexto e gestão de sessão.

## 2. Princípios e limites

- Não remover nem descontinuar o suplemento para criar o site.
- Não alterar o domínio, manifesto ou deploy estável de produção durante prototipagem e validação.
- Não assumir que login simples, URL de planilha ou seleção de modelo concede acesso aos dados; validar autorização no servidor.
- Nunca armazenar senhas em texto puro ou implementar autenticação artesanal sem controles adequados.
- A seleção de planilha/modelo deve distinguir **modelo de cadastro** de **arquivo real de trabalho**.
- Definir a fonte oficial dos dados antes de permitir que site e suplemento gravem no mesmo conjunto.
- Cada fase precisa de critérios de aceite e testes; não considerar concluído só porque a página abre.
- Preservar compatibilidade com as versões estáveis até haver aprovação explícita para migração.

## 3. Escopo funcional

### A. Suplemento Excel — Alpha 2.0
- Preservar as telas atuais: início, novo registro, e-mail, SMS, histórico, resumo e configurações.
- Validar carregamento de páginas, CSS, JavaScript, navegação e ícones.
- Validar gravação/leitura dos registros no contexto real do Excel.
- Validar cadastros, modelos/mensagens e limite de SMS.
- Verificar requisitos do manifesto e APIs do Office.
- Documentar instalação, atualização, recuperação e solução de problemas.

### B. Site — Alpha 2.5
- Criar página principal útil, substituindo a raiz vazia/404 do Preview.
- Definir fluxo de entrada: login corporativo ou autenticação gerenciada apropriada ao público-alvo.
- Criar área inicial com identificação do usuário e ação para sair.
- Permitir escolher o modelo/contexto de registro disponível para aquele usuário.
- Exibir ações equivalentes às do suplemento: novo registro, histórico, resumo e configurações, respeitando permissões.
- Disponibilizar campo para informar/selecionar o link do arquivo Excel e botão para abri-lo em nova aba.
- Validar formato e domínio do link; não tratar um link como prova de autorização.
- Criar estados de carregamento, vazio, erro, sessão expirada e acesso negado.
- Tornar a interface responsiva para desktop e telas menores.
- Definir como o site acessa, persiste e isola dados sem depender de uma sessão ativa do Excel.

### C. Identidade, usuários e permissões
- Definir público: apenas equipe interna, usuários autenticados da organização ou usuários externos convidados.
- Escolher autenticação gerenciada e compatível com a organização; priorizar identidade corporativa quando disponível.
- Definir papéis mínimos (por exemplo, administrador e usuário) e permissões por modelo/empresa.
- Proteger páginas e operações no servidor, não apenas ocultando botões.
- Planejar expiração de sessão, logout, recuperação de acesso e auditoria básica.
- Não criar cadastro aberto ao público sem decisão explícita sobre aprovação de usuários.

### D. Dados e integração entre site e suplemento
- Mapear onde hoje ficam registros, cadastros, mensagens e configurações.
- Definir fonte única de verdade e contratos de leitura/escrita.
- Se os dados dependem atualmente de planilha ativa/Office.js, definir a alternativa web e suas limitações.
- Definir se o site vai usar uma base de dados central, Microsoft Graph/Excel, ou outro serviço aprovado.
- Tratar conflitos, duplicidade, validação, datas/fuso horário e erros de gravação.
- Garantir separação de dados entre usuários/empresas e controles de acesso.
- Planejar migração e compatibilidade sem perder registros existentes.
- Só permitir gravação cruzada site/suplemento após testes de consistência e autorização.

### E. Infraestrutura e publicação
- Continuar usando Preview para desenvolvimento e testes da Alpha 2.5.
- Definir rotas públicas e rotas autenticadas; a raiz do site deve ter comportamento intencional.
- Revisar variáveis de ambiente, segredos, logs e cabeçalhos de segurança.
- Confirmar URLs corretas de suporte, ícones e manifesto do suplemento.
- Separar ambientes de desenvolvimento, homologação e produção.
- Não alterar domínio estável nem manifesto de produção sem aprovação e plano de reversão.

### F. Documentação e suporte
- Manual de instalação do suplemento com capturas de tela.
- Guia rápido do site e seleção de modelo/planilha.
- Guia de login, permissões e abertura opcional do arquivo.
- FAQ e solução de erros comuns.
- Documento de arquitetura e fluxo dos dados.
- Checklist de testes por versão e registro das decisões.
- Instruções para publicar/voltar versão, com responsáveis e pontos de aprovação.

## 4. Fases e entregáveis

### Fase 0 — Consolidar Alpha 2.0
**Entregáveis**
- Estrutura de arquivos organizada.
- Lista de telas e funções existentes.
- Checklist de regressão preenchido com evidências.
- Lista de problemas conhecidos e dependências.

**Aceite**
- Todas as telas e navegações verificadas no Excel real.
- Gravação e leitura verificadas; falhas conhecidas registradas.
- Nenhuma alteração na produção.

### Fase 1 — Descoberta técnica do site
**Entregáveis**
- Mapa do fluxo atual de dados do suplemento.
- Decisão sobre público, login e permissões.
- Decisão sobre o significado de “modelo” versus “planilha”.
- Proposta de arquitetura de dados para site + suplemento.

**Aceite**
- É possível explicar como o site identifica o usuário, lista os modelos autorizados, lê/grava registros e abre a planilha sem confiar somente no link.

### Fase 2 — Protótipo navegável Alpha 2.5
**Entregáveis**
- Página inicial na raiz do Preview.
- Fluxo demonstrativo de login (sem credenciais reais se for protótipo).
- Seleção de modelo/contexto.
- Área principal com ações de novo registro, histórico, resumo e configurações.
- Campo e botão para abrir a planilha.
- Estados responsivos e de erro.

**Aceite**
- Usuário consegue percorrer o fluxo de ponta a ponta com dados de demonstração.
- O protótipo deixa claro o que é simulado e não grava dados reais inadvertidamente.

### Fase 3 — Autenticação e autorização
**Entregáveis**
- Provedor de identidade aprovado e configurado em ambiente de teste.
- Sessões e logout.
- Papéis e regras de acesso no servidor.
- Proteção de endpoints e tratamento de sessão expirada.

**Aceite**
- Usuário não autenticado não acessa áreas protegidas.
- Usuário não consegue consultar ou alterar dados de outro usuário/empresa sem permissão.
- Nenhum segredo é exposto no código do cliente.

### Fase 4 — Serviço de dados e integração
**Entregáveis**
- Serviço/API ou integração Microsoft aprovada.
- Persistência central ou estratégia definida de acesso à planilha.
- Validação de entradas, tratamento de erros e auditoria básica.
- Testes de consistência entre site e suplemento.

**Aceite**
- Criação, consulta e atualização são persistentes e autorizadas.
- Site e suplemento apresentam os mesmos dados quando usam a mesma fonte.
- Conflitos e falhas não causam perda silenciosa de dados.

### Fase 5 — Paridade funcional
**Entregáveis**
- Fluxos web de e-mail e SMS.
- Histórico com filtros/ordenação.
- Resumo e filtros por período.
- Configurações, cadastros, mensagens e limite de SMS conforme regras aprovadas.
- Abertura opcional da planilha.

**Aceite**
- Casos de uso definidos para cada tela passam em teste.
- Validações e permissões são equivalentes às regras aprovadas do produto.
- Recursos não disponíveis no navegador ficam claramente identificados.

### Fase 6 — Segurança, qualidade e homologação
**Entregáveis**
- Testes funcionais, de permissões, responsividade e regressão.
- Revisão de privacidade e retenção de dados.
- Testes em navegadores suportados e Excel desktop/online.
- Manual e checklist de homologação.

**Aceite**
- Sem defeitos críticos abertos.
- Testes de autorização aprovados.
- Backup/recuperação e procedimento de reversão documentados.

### Fase 7 — Piloto controlado
**Entregáveis**
- Preview/homologação acessível a um grupo pequeno.
- Coleta de feedback e lista priorizada de correções.
- Plano de suporte e monitoramento.

**Aceite**
- Usuários-piloto concluem os fluxos definidos.
- Problemas de dados, acesso e usabilidade foram tratados ou aceitos formalmente.

### Fase 8 — Publicação gradual
**Entregáveis**
- Aprovação explícita de produção.
- Plano de publicação e rollback.
- Verificação de URLs, manifesto, domínio e monitoramento.
- Comunicação de versão e documentação final.

**Aceite**
- Publicação aprovada pelo responsável.
- Suplemento continua funcionando.
- Site e suplemento coexistem sem mudança não autorizada no fluxo estável.

## 5. Prioridade recomendada

1. **P0:** consolidar e validar o suplemento Alpha 2.0; mapear a persistência atual.
2. **P1:** desenhar arquitetura e decisões de autenticação/dados para o site.
3. **P2:** criar protótipo navegável Alpha 2.5, sem gravação real.
4. **P3:** implementar identidade e permissões.
5. **P4:** implementar dados e integração segura.
6. **P5:** completar paridade funcional e testes.
7. **P6:** piloto, correções e só então decisão de produção.

## 6. Decisões pendentes

- O site será exclusivo para colaboradores da empresa ou também para usuários externos?
- O login deve usar a conta corporativa Microsoft ou outro provedor aprovado?
- “Modelo/planilha” significa escolher um template lógico, um arquivo Excel específico, ou ambos?
- Os registros hoje são armazenados dentro de uma planilha, em outra fonte, ou em ambas?
- O site e o suplemento devem gravar na mesma base imediatamente, ou a sincronização pode ficar para uma fase posterior?
- Quem poderá criar usuários, modelos, cadastros e mensagens?
- Quais navegadores e dispositivos precisam ser suportados?
- Qual o processo de aprovação para publicar em produção?

## 7. Fora de escopo até aprovação
- Substituir o suplemento existente pelo site.
- Alterar o manifesto ou o domínio estável de produção.
- Migrar registros reais sem plano de migração e backup.
- Abrir o cadastro de usuários publicamente.
- Habilitar escrita no Excel/serviço real antes de definir identidade, permissões e fonte de verdade.

## 8. Regra de conclusão
Uma fase só será marcada como concluída com: entregáveis disponíveis, critérios de aceite verificados, evidências registradas e riscos conhecidos documentados. Deploy “READY” ou página carregando, isoladamente, não significa que a fase foi homologada.
