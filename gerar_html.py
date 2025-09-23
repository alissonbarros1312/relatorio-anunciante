import gspread
from oauth2client.service_account import ServiceAccountCredentials

# Caminho do arquivo de credenciais
CREDENTIALS_FILE = "credentials.json"
HTML_OUTPUT = "relatorio.html"

# IDs das planilhas (substitua pelos reais)
SHEET_ID_VISITANTES = "10WR4hOcQMlfMUHPg2t5TSigzfoS_VgO6y_1d_BecUlk"
SHEET_ID_IGREJAS = "1VmcqvfnZacUGT4zEcPYIhrFmVvRDoilHEKJ_lkQddHE"

# Autenticação no Google
scope = ["https://spreadsheets.google.com/feeds",
         "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_FILE, scope)
client = gspread.authorize(creds)

# Abre a planilha Visitantes
sheet_visitantes = client.open_by_key(SHEET_ID_VISITANTES).sheet1
visitantes_data = sheet_visitantes.get_all_records()

# Abre a planilha Igrejas
sheet_igrejas = client.open_by_key(SHEET_ID_IGREJAS).sheet1
igrejas_data = sheet_igrejas.get_all_records()

# ---- Construindo HTML estilizado e responsivo ----
html = """
<!DOCTYPE html>
<html lang="pt-br">
<head>
    <meta charset="UTF-8">
    <title>Relatório</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; background: #f9f9f9; }
        h1 { color: #333; border-bottom: 2px solid #444; padding-bottom: 5px; }
        .table-container { overflow-x: auto; margin-bottom: 40px; }
        table { width: 100%; min-width: 600px; border-collapse: collapse; }
        th, td { border: 1px solid #ccc; padding: 10px; text-align: left; }
        th { background: #444; color: #fff; }
        tr:nth-child(even) { background: #f2f2f2; }
        tr:hover { background: #e9f5ff; }
        @media (max-width: 768px) {
            body { font-size: 14px; margin: 10px; }
            th, td { padding: 8px; }
        }
    </style>
</head>
<body>
    <h1>Relatório de Visitantes</h1>
    <div class="table-container">
        <table>
            <tr>
                <th>Igreja</th>
                <th>Nome</th>
                <th>Acompanhantes</th>
                <th>Observações</th>
            </tr>
"""

# Adiciona visitantes
for v in visitantes_data:
    acompanhantes = ", ".join(
        [v.get(f"Acompanhante {i}", "") for i in range(1, 5)]
    ).strip(", ")
    html += f"""
            <tr>
                <td>{v.get('Qual igreja?', '')}</td>
                <td>{v.get('Qual o seu nome?', '')}</td>
                <td>{acompanhantes}</td>
                <td>{v.get('Observações', '')}</td>
            </tr>
    """

# Tabela Igrejas
html += """
        </table>
    </div>
    <h1>Relatório de Igrejas</h1>
    <div class="table-container">
        <table>
            <tr>
                <th>Igrejas</th>
                <th>Conjunto</th>
                <th>Líderes</th>
                <th>Observações</th>
            </tr>
"""

for i in igrejas_data:
    # Tratamento da pergunta múltipla escolha "Qual Igreja?"
    igrejas = i.get("Qual Igreja?", "")
    if isinstance(igrejas, list):
        igrejas = ", ".join(igrejas)
    elif isinstance(igrejas, str):
        igrejas = igrejas.strip()

    html += f"""
            <tr>
                <td>{igrejas}</td>
                <td>{i.get('Nome do conjunto?', '')}</td>
                <td>{i.get('Nome dos Líderes?', '')}</td>
                <td>{i.get('Observações', '')}</td>
            </tr>
    """

html += """
        </table>
    </div>
</body>
</html>
"""

# Salva o arquivo HTML
with open(HTML_OUTPUT, "w", encoding="utf-8") as f:
    f.write(html)

print(f"✅ Relatório gerado com sucesso: {HTML_OUTPUT}")
