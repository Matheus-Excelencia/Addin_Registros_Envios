# Revisão de segurança — Alpha 2.0
Revisão estática parcial: histórico escapa campos antes de interpolar HTML; duplicação passa a usar sessionStorage também no suplemento real, reduzindo persistência de dados operacionais. Não é isolamento entre workbooks na mesma sessão. CI usa permissões de leitura e não persiste credenciais de checkout. Manifesto e escopos SSO não alterados.

Pendências: verificar fórmulas introduzidas por texto em células Excel, isolamento entre workbooks, logs de erro, comportamento do SSO real e testes de navegador. Não constitui auditoria completa ou certificação.
