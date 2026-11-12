"""
Alerta de validade por e-mail.

Lê a planilha de controle de validades, identifica os lotes que vencem nos
próximos N dias e envia um e-mail em HTML com a lista, do mais urgente para o
menos urgente. Pensado para rodar todo dia pelo Agendador de Tarefas do Windows
(ou cron, no Linux).

Toda a configuração vem de variáveis de ambiente (ou de um arquivo .env ao lado
do script). Veja o arquivo .env.exemplo.
"""

import os
import smtplib
from datetime import date, datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path

from openpyxl import load_workbook

# Carrega o .env se o python-dotenv estiver instalado (opcional)
try:
    from dotenv import load_dotenv
    load_dotenv(Path(__file__).with_name(".env"))
except ImportError:
    pass

# ============================================================
#  CONFIGURAÇÕES (variáveis de ambiente / .env)
# ============================================================

PASTA_SCRIPT = Path(__file__).resolve().parent

CAMINHO_XLSX = os.environ.get("VALIDADE_XLSX", str(PASTA_SCRIPT / "controle_validade.xlsx"))
NOME_ABA = os.environ.get("VALIDADE_ABA", "Validades")
DIAS_ALERTA = int(os.environ.get("VALIDADE_DIAS_ALERTA", "90"))

SMTP_HOST = os.environ.get("SMTP_HOST", "smtp.gmail.com")
SMTP_PORTA = int(os.environ.get("SMTP_PORTA", "465"))
EMAIL_REMETENTE = os.environ.get("EMAIL_REMETENTE", "")
SENHA_EMAIL = os.environ.get("EMAIL_SENHA", "")
EMAIL_DESTINO = os.environ.get("EMAIL_DESTINO", "")

# Lembretes opcionais exibidos no corpo do e-mail, separados por ";"
LEMBRETES = [l.strip() for l in os.environ.get("VALIDADE_LEMBRETES", "").split(";") if l.strip()]


# ============================================================
#  LÓGICA DO ALERTA
# ============================================================

def _fmt_num(valor):
    """Formata número removendo .0 quando for inteiro (ex.: 61931.0 -> 61931)."""
    if valor is None:
        return ""
    if isinstance(valor, float):
        return str(int(valor)) if valor.is_integer() else str(valor)
    return str(valor).strip()


def verificar_validades():
    if not os.path.exists(CAMINHO_XLSX):
        print(f"[ERRO] Planilha não encontrada: {CAMINHO_XLSX}")
        return

    wb = load_workbook(CAMINHO_XLSX, data_only=True)
    if NOME_ABA not in wb.sheetnames:
        print(f"[ERRO] Aba '{NOME_ABA}' não encontrada.")
        return

    ws = wb[NOME_ABA]
    hoje = date.today()

    # Mapa código de barras -> SKU direto da aba Produtos (não depende de fórmula calculada)
    sku_map = {}
    if "Produtos" in wb.sheetnames:
        for bc_p, sku_p in wb["Produtos"].iter_rows(min_row=2, min_col=1, max_col=2, values_only=True):
            if bc_p and sku_p:
                sku_map[_fmt_num(bc_p)] = _fmt_num(sku_p)

    # Colunas de data: toda célula da linha 1 que contém uma data (a quantidade fica na coluna seguinte)
    colunas_data = [cell.column for cell in ws[1] if cell.value and hasattr(cell.value, "year")]

    alertas = []

    for row in ws.iter_rows(min_row=2):
        sku = row[0].value      # col A (pode estar vazia se a fórmula não foi calculada)
        barcode = row[1].value  # col B
        produto = row[2].value  # col C
        if not barcode:
            continue

        bc_str = _fmt_num(barcode)
        sku_str = _fmt_num(sku) or sku_map.get(bc_str, "")

        for col_data in colunas_data:
            val = ws.cell(row=row[0].row, column=col_data).value
            qtd = ws.cell(row=row[0].row, column=col_data + 1).value
            if not val or not qtd:
                continue
            if isinstance(val, datetime):
                val = val.date()
            if not isinstance(val, date):
                continue

            dias_restantes = (val - hoje).days
            if 0 <= dias_restantes <= DIAS_ALERTA:
                alertas.append({
                    "sku": sku_str,
                    "produto": str(produto or "").strip(),
                    "barcode": bc_str,
                    "data": val.strftime("%d/%m/%Y"),
                    "dias": dias_restantes,
                    "qtd": _fmt_num(qtd),
                })

    if not alertas:
        print(f"[OK] Nenhum produto vencendo nos próximos {DIAS_ALERTA} dias.")
        return

    alertas.sort(key=lambda x: x["dias"])  # mais urgente primeiro
    enviar_email(alertas)


def montar_html(alertas):
    linhas = ""
    for a in alertas:
        cor = "#FFCCCC" if a["dias"] <= 30 else "#FFD580" if a["dias"] <= 60 else "#FFFACD"
        linhas += f"""
        <tr style="background:{cor}">
            <td>{a['sku']}</td>
            <td>{a['produto']}</td>
            <td>{a['barcode']}</td>
            <td><strong>{a['data']}</strong></td>
            <td style="text-align:center"><strong>{a['dias']}</strong></td>
            <td style="text-align:center">{a['qtd']}</td>
        </tr>"""

    lembretes_html = ""
    if LEMBRETES:
        itens = "".join(f"<li>{l}</li>" for l in LEMBRETES)
        lembretes_html = f"<p><strong>Lembretes:</strong></p><ul>{itens}</ul>"

    return f"""
    <h2 style="color:#1A3A5C">Alerta de Controle de Validade</h2>
    <p>Os produtos abaixo vencem nos próximos <strong>{DIAS_ALERTA} dias</strong>:</p>
    {lembretes_html}
    <table border="1" cellpadding="8" cellspacing="0"
           style="border-collapse:collapse;font-family:Arial,sans-serif;font-size:13px">
        <tr style="background:#1A3A5C;color:#fff">
            <th>SKU</th><th>Produto</th><th>Código de barras</th>
            <th>Validade</th><th>Dias restantes</th><th>Quantidade</th>
        </tr>
        {linhas}
    </table>
    <br>
    <p style="color:#888;font-size:12px">
        Gerado automaticamente pelo Controle de Validades em {date.today().strftime('%d/%m/%Y')}.
    </p>
    """


def enviar_email(alertas):
    if not (EMAIL_REMETENTE and SENHA_EMAIL and EMAIL_DESTINO):
        print("[ERRO] Configure EMAIL_REMETENTE, EMAIL_SENHA e EMAIL_DESTINO (veja .env.exemplo).")
        return

    destinatarios = [e.strip() for e in EMAIL_DESTINO.split(",") if e.strip()]

    msg = MIMEMultipart("alternative")
    msg["Subject"] = f"Alerta de validade: {len(alertas)} produto(s) vencendo em até {DIAS_ALERTA} dias"
    msg["From"] = EMAIL_REMETENTE
    msg["To"] = ", ".join(destinatarios)
    msg.attach(MIMEText(montar_html(alertas), "html"))

    try:
        with smtplib.SMTP_SSL(SMTP_HOST, SMTP_PORTA) as smtp:
            smtp.login(EMAIL_REMETENTE, SENHA_EMAIL)
            smtp.sendmail(EMAIL_REMETENTE, destinatarios, msg.as_string())
        print(f"[OK] E-mail enviado com {len(alertas)} alerta(s).")
    except Exception as e:
        print(f"[ERRO] Falha ao enviar e-mail: {e}")


if __name__ == "__main__":
    verificar_validades()
