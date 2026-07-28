"""
Agente de Tratamento de Dados - Projeto Porsche
=================================================

Agente de IA (via API da Anthropic / Claude) que interpreta e sanitiza os
dados brutos de vendas Porsche, entregando um arquivo final limpo e um
relatório de acurácia contra o gabarito (colunas *Sanitized já presentes
na planilha original).

Como usar:
    export ANTHROPIC_API_KEY="sua-chave-aqui"
    python agente_tratamento_dados.py caminho/para/planilha.xlsx

Saída:
    - porsche_dados_tratados.xlsx  -> dados limpos + colunas de auditoria
    - relatorio_acuracia.txt       -> % de acerto do agente por campo
"""

import json
import os
import sys
from datetime import datetime

import pandas as pd
from anthropic import Anthropic

MODEL = "claude-sonnet-5"  # ajuste para o modelo desejado, se necessário
BATCH_SIZE = 10  # nº de registros enviados por chamada de API

# Colunas brutas que o agente vai receber e tratar
COLUNAS_BRUTAS = [
    "sale_id",
    "sale_date",
    "customer_name",
    "porsche_model",
    "model_year",
    "sale_price",
    "vehicle_mileage",
    "payment_method",
    "city",
    "state",
    "salesperson",
    "delivery_status",
]

# Mapeia campo sanitizado do agente -> coluna gabarito na planilha original
# (usada só para validação, quando existir gabarito para o campo)
GABARITO = {
    "sale_date_sanitized": "SaleDateSanitized",
    "porsche_model_sanitized": "PorscheModelSanitized",
    "model_year_sanitized": "ModelYearSanitized",
    "sale_price_sanitized": "SalesPriceSanitized",
    "vehicle_mileage_sanitized": "VehicleMileageSanitized",
    "pay_method_sanitized": "PayMethodSanitized",
    "city_sanitized": "CitySanitized",
    "state_sanitized": "StateSanitized",
    "delivery_status_sanitized": "DeliveryStatusSanitized",
}

SYSTEM_PROMPT = """\
Você é um agente especialista em limpeza e padronização de dados (data \
wrangling) para uma base de vendas de veículos Porsche nos EUA.

Você vai receber um lote de registros brutos, cada um com campos sujos, \
inconsistentes ou mal formatados. Sua tarefa é interpretar cada campo com \
bom senso e devolver a versão sanitizada, seguindo estas regras:

1. sale_date -> formato ISO "YYYY-MM-DD". Se a data não existir no \
   calendário (ex: 30 de fevereiro, 31 de abril, mês/dia > 12/31 sem forma \
   de interpretar), retorne exatamente "INVALID".
2. customer_name -> nome próprio em Title Case, sem hífens artificiais \
   substituindo espaço (ex: "Daniel-Jones" -> "Daniel Jones"), sem CAPS.
3. porsche_model -> mantenha o nome oficial do modelo, apenas normalizando \
   capitalização.
4. model_year -> ano com 4 dígitos (ex: "twenty twenty four" -> "2024", \
   "20-23" ou "20 24" -> interprete como o ano mais provável).
5. sale_price -> número decimal em USD, sem símbolos, vírgulas ou texto \
   (ex: "$79,500.00" -> 79500.0; "188k USD" -> 188000.0; "eighty two \
   thousand USD" -> 82000.0). Cuidado com separador de milhar europeu \
   (ex: "USD 112.750" -> 112750.0, não 112.75).
6. vehicle_mileage -> número inteiro em MILHAS. Se o valor original estiver \
   em quilômetros (contém "KM" ou "km"), converta para milhas \
   (1 km = 0.621371 milhas) e arredonde para inteiro. Valores por extenso \
   (ex: "zero miles", "fifteen thousand miles") devem virar número.
7. payment_method -> padronize para uma destas categorias, com essa grafia \
   exata: "Credit Card", "Debit Card", "Cash", "Wire Transfer", \
   "Bank Transfer", "Financing", "Lease", "Crypto Payment", "ACH Payment".
8. city -> Title Case.
9. state -> sigla de 2 letras (padrão USPS), maiúscula.
10. salesperson -> nome próprio em Title Case.
11. delivery_status -> padronize para uma destas categorias, com essa \
    grafia exata: "Delivered", "Pending", "In Transit", "Cancelled", \
    "Shipped", "Awaiting Delivery", "Awaiting Pickup", "Awaiting Review", \
    "Pending Approval", "Pending Review". Corrija erros de digitação e \
    remova pontuação/ênfase extra (ex: "DELIVERD" -> "Delivered", \
    "pending!!" -> "Pending").

Responda SOMENTE usando a ferramenta `registros_sanitizados` fornecida, \
um item por registro recebido, na mesma ordem, incluindo o sale_id \
original para conferência.
"""

TOOL_SCHEMA = {
    "name": "registros_sanitizados",
    "description": "Devolve a lista de registros sanitizados.",
    "input_schema": {
        "type": "object",
        "properties": {
            "registros": {
                "type": "array",
                "items": {
                    "type": "object",
                    "properties": {
                        "sale_id": {"type": "integer"},
                        "sale_date_sanitized": {"type": "string"},
                        "customer_name_sanitized": {"type": "string"},
                        "porsche_model_sanitized": {"type": "string"},
                        "model_year_sanitized": {"type": "string"},
                        "sale_price_sanitized": {"type": "number"},
                        "vehicle_mileage_sanitized": {"type": "integer"},
                        "pay_method_sanitized": {"type": "string"},
                        "city_sanitized": {"type": "string"},
                        "state_sanitized": {"type": "string"},
                        "salesperson_sanitized": {"type": "string"},
                        "delivery_status_sanitized": {"type": "string"},
                    },
                    "required": [
                        "sale_id",
                        "sale_date_sanitized",
                        "customer_name_sanitized",
                        "porsche_model_sanitized",
                        "model_year_sanitized",
                        "sale_price_sanitized",
                        "vehicle_mileage_sanitized",
                        "pay_method_sanitized",
                        "city_sanitized",
                        "state_sanitized",
                        "salesperson_sanitized",
                        "delivery_status_sanitized",
                    ],
                },
            }
        },
        "required": ["registros"],
    },
}


def montar_lote(df_lote: pd.DataFrame) -> str:
    """Serializa um lote de registros brutos em JSON para enviar ao modelo."""
    registros = df_lote[COLUNAS_BRUTAS].to_dict(orient="records")
    return json.dumps(registros, ensure_ascii=False, default=str)


def chamar_agente(client: Anthropic, lote_json: str) -> list[dict]:
    """Chama a API da Anthropic para sanitizar um lote e retorna os registros."""
    response = client.messages.create(
        model=MODEL,
        max_tokens=4000,
        system=SYSTEM_PROMPT,
        tools=[TOOL_SCHEMA],
        tool_choice={"type": "tool", "name": "registros_sanitizados"},
        messages=[
            {
                "role": "user",
                "content": f"Sanitize este lote de registros:\n{lote_json}",
            }
        ],
    )

    for bloco in response.content:
        if bloco.type == "tool_use" and bloco.name == "registros_sanitizados":
            return bloco.input["registros"]

    raise RuntimeError("O modelo não retornou o tool_use esperado.")


def processar_planilha(caminho_entrada: str) -> None:
    if not os.getenv("ANTHROPIC_API_KEY"):
        sys.exit(
            "Erro: defina a variável de ambiente ANTHROPIC_API_KEY antes de "
            "rodar o agente."
        )

    client = Anthropic()
    df_bruto = pd.read_excel(caminho_entrada, sheet_name="Sanitized")

    resultados = []
    total_lotes = (len(df_bruto) + BATCH_SIZE - 1) // BATCH_SIZE

    for i in range(0, len(df_bruto), BATCH_SIZE):
        lote = df_bruto.iloc[i : i + BATCH_SIZE]
        num_lote = i // BATCH_SIZE + 1
        print(f"Processando lote {num_lote}/{total_lotes}...")

        lote_json = montar_lote(lote)
        registros_sanitizados = chamar_agente(client, lote_json)
        resultados.extend(registros_sanitizados)

    df_agente = pd.DataFrame(resultados)

    # Junta o resultado do agente com o gabarito original para auditoria
    df_final = df_bruto.merge(df_agente, on="sale_id", how="left", suffixes=("", "_agente"))

    gerar_relatorio_acuracia(df_final)

    caminho_saida = "porsche_dados_tratados.xlsx"
    df_final.to_excel(caminho_saida, index=False)
    print(f"\nArquivo gerado: {caminho_saida}")


def gerar_relatorio_acuracia(df_final: pd.DataFrame) -> None:
    linhas_relatorio = ["Relatório de Acurácia do Agente de Tratamento de Dados", "=" * 55, ""]
    total_geral = 0
    acertos_geral = 0

    for campo_agente, coluna_gabarito in GABARITO.items():
        if coluna_gabarito not in df_final.columns or campo_agente not in df_final.columns:
            continue

        esperado = df_final[coluna_gabarito].astype(str).str.strip()
        obtido = df_final[campo_agente].astype(str).str.strip()

        # tolerância numérica para preço/km (evita falso-negativo por 100.0 vs 100)
        try:
            esperado_num = pd.to_numeric(df_final[coluna_gabarito], errors="raise")
            obtido_num = pd.to_numeric(df_final[campo_agente], errors="raise")
            acertos = (esperado_num.round(2) == obtido_num.round(2)).sum()
        except (ValueError, TypeError):
            acertos = (esperado.str.lower() == obtido.str.lower()).sum()

        total = len(df_final)
        pct = 100 * acertos / total
        total_geral += total
        acertos_geral += acertos
        linhas_relatorio.append(f"{campo_agente:35s} {acertos:3d}/{total:3d}  ({pct:5.1f}%)")

    pct_geral = 100 * acertos_geral / total_geral if total_geral else 0
    linhas_relatorio += ["", f"ACURÁCIA GERAL: {acertos_geral}/{total_geral} ({pct_geral:.1f}%)"]

    relatorio = "\n".join(linhas_relatorio)
    print("\n" + relatorio)

    with open("relatorio_acuracia.txt", "w", encoding="utf-8") as f:
        f.write(relatorio)


if __name__ == "__main__":
    if len(sys.argv) != 2:
        sys.exit("Uso: python agente_tratamento_dados.py caminho/para/planilha.xlsx")
    processar_planilha(sys.argv[1])
