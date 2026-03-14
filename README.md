 (cd "$(git rev-parse --show-toplevel)" && git apply --3way <<'EOF' 
diff --git a/README.md b/README.md
new file mode 100644
index 0000000000000000000000000000000000000000..55ed55ec710271558f58bf8d5a0e5d89883b5ead
--- /dev/null
+++ b/README.md
@@ -0,0 +1,56 @@
+# NexusLab Ponto Pro
+
+Software desktop para **gerenciamento completo de ponto eletrônico** com foco em uso com equipamentos **Control iD Access** e fluxo compatível com **AFD**.
+
+## Funcionalidades entregues
+
+- Cadastro de colaboradores com **nome, matrícula, departamento e PIS**
+- Edição e exclusão de colaborador
+- Registro manual de ponto
+- Importação de batidas por CSV
+- Importação de AFD/TXT (parser flexível para layouts comuns)
+- Histórico completo de batidas
+- Resumo mensal por colaborador
+- Exportação de relatório mensal em CSV
+- Geração de relatório mensal em HTML para impressão/PDF
+- Geração de espelho individual de ponto em HTML para impressão/PDF
+- Configurações gerais da empresa (nome/CNPJ)
+- Configuração da integração Control iD (host/porta/token/endpoint)
+- Envio do colaborador selecionado para endpoint HTTP do integrador Control iD
+
+## Observação legal
+
+O sistema foi estruturado para apoiar processos de ponto eletrônico e AFD, porém a conformidade final depende da configuração do REP/Control iD, das regras da empresa e de validação jurídica/contábil conforme **Portaria 671/MTP** e normas coletivas.
+
+## Executar o sistema
+
+```bash
+python3 -m venv .venv
+source .venv/bin/activate
+python -m pip install -r requirements.txt
+python src/main.py
+```
+
+## Formatos de importação
+
+### CSV
+
+Separado por `;` com cabeçalho:
+
+```csv
+matricula;tipo
+123;Entrada
+123;Saída
+```
+
+### AFD/TXT
+
+Suporte inicial a:
+- linhas `matricula;YYYYMMDD;HHMM`
+- linhas numéricas de layout fixo (extração automática de data/hora/matrícula)
+
+Se o seu equipamento usar layout diferente, ajuste o parser no método `_parse_afd_line`.
+
+---
+
+Desenvolvido com assinatura **NexusLab**.
 
EOF
)
