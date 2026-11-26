# Próximas atualizações — Registros de Envios

Requisitos registrados em 26/11/2026. Alpha 1.07 é uma proposta de identificação, não uma versão publicada.

## Identidade visual
- Replicar no suplemento a identidade da landing page: nova logo, verde petróleo #087f72, azul #124b60, menta #a7f3d0 e dourado #f8d58a, estilo e transições suaves.
- Light/Dark Mode com preferência persistente; atualizar ícones e README conforme a identidade. Preservar SSO, manifesto funcional e dados.

## Demonstração web
- Reutilizar UI e regras reais com adaptador Excel simulado isolado, sem dados corporativos ou login real.
- Paridade de telas, mensagens, cadastro/edição de mensagens, perguntas, configurações, histórico, duplicação, resumo e planilha simulada. Validar cada fluxo contra suplemento real. Demo atual simplificada não atende paridade.

## Resumo
- Alternar cards, lista e gráfico, persistindo a escolha.
- Card opcional de Valor da pergunta, com agregação numérica coerente e formatação monetária quando aplicável.

## Histórico
- Exibir Valor nos cards quando aplicável; duplicação deve preencher o valor original, preservando zero e distinguindo vazio.

## Perguntas personalizadas
- Configurar visibilidade independente no Histórico e Resumo; definir opção de exibição opcional ou obrigatória em cada destino, separada da obrigatoriedade de resposta no formulário. Confirmar semântica e agregações antes de implementar.
- Compatibilidade com registros antigos e configuração existente; testar edição, desativação, duplicação, persistência e formatos.

## Critérios de entrega
- Branch separada, testes automatizados, homologação Excel Web e comparação da demo. Não considerar concluído sem evidência. Não alterar dados reais.
