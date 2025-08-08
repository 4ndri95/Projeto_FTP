# Automação de Extração FTP e Consulta Oracle

Automatiza a extração parcial de documentos PDF de diretórios em uma rede FTP, realiza consultas em banco de dados Oracle PL/SQL utilizando os nomes dos documentos como palavras-chave, gera uma planilha organizada
com os dados extraídos e envia o arquivo automaticamente por e-mail (Outlook). O processo inclui tratamento de erros e relatórios, funcionando de forma simples e automatizada via execução de scripts.

# Funcionalidades
- Conexão e download automático de arquivos PDF via FTP
- Processamento e validação dos arquivos baixados
- Consulta automatizada em banco de dados Oracle com PL/SQL
- Geração de planilha Excel organizada com resultados
- Envio automático do relatório por e-mail via Outlook
- Logs detalhados para monitoramento e diagnóstico
- Exclusão automática dos arquivos processados no servidor FTP (opcional)

# Requisitos
- Python 3.x
- Bibliotecas Python: ftplib, logging, os, pandas, openpyxl, oracledb, win32com.client
- Cliente Oracle instalado e configurado para conexão via oracledb
- Outlook instalado para envio de e-mails
- Acesso válido à rede FTP e ao banco Oracle
- Arquivos de configuração: credenciais FTP e banco (login_sql.ini), planilhas de destinatários (Emails.xlsx)

# Como executar
1. Configure as credenciais e caminhos nos arquivos de configuração (login_sql.ini, diretórios no código etc.)
2. Execute o script principal que automatiza todo o fluxo (download, consulta, geração da planilha e envio do e-mail):
3. Verifique os logs (ftp_log.log, planilha_log.log, email_log.log) para acompanhar o processo
4. Abra o e-mail gerado automaticamente no Outlook para revisão e envio final

