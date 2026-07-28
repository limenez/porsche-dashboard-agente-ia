# Agente de Tratamento de Dados — Projeto Porsche

Agente de IA que usa a **API da Anthropic (Claude)** para interpretar e
sanitizar a base de vendas Porsche, entregando um arquivo limpo e um
relatório de acurácia contra o gabarito já presente na planilha original
(colunas `*Sanitized`).

## Como funciona

1. Lê a aba `Sanitized` da planilha e pega apenas as colunas **brutas**
   (`sale_date`, `customer_name`, `porsche_model`, `model_year`,
   `sale_price`, `vehicle_mileage`, `payment_method`, `city`, `state`,
   `salesperson`, `delivery_status`).
2. Envia os registros em **lotes de 10** para o modelo, com um
   `system prompt` detalhando as regras de normalização (formato de data,
   conversão km→milhas, padronização de categorias, etc).
3. Usa **tool use / saída estruturada** (`tool_choice` forçado) para
   garantir que a resposta do modelo sempre venha em JSON válido, no
   formato esperado — sem parsing frágil de texto livre.
4. Cruza o resultado do agente com as colunas `*Sanitized` (gabarito) e
   calcula a **acurácia por campo** e a acurácia geral.
5. Salva:
   - `porsche_dados_tratados.xlsx` — dados originais + colunas tratadas
     pelo agente, lado a lado, prontas para auditoria.
   - `relatorio_acuracia.txt` — relatório de % de acerto por campo.

## Pré-requisitos

```bash
pip install anthropic pandas openpyxl
export ANTHROPIC_API_KEY="sua-chave-aqui"
```

## Uso

```bash
python agente_tratamento_dados.py caminho/para/planilha_porsche.xlsx
```

## Personalização

- **Modelo**: ajuste a constante `MODEL` no topo do script (padrão:
  `claude-sonnet-5`). Para tarefas simples como esta, um modelo mais
  leve (ex: Haiku) também tende a funcionar bem e sai mais barato.
- **Tamanho do lote**: `BATCH_SIZE` controla quantos registros vão por
  chamada de API. Lotes maiores = menos chamadas, porém respostas mais
  longas (atenção ao `max_tokens`).
- **Regras de sanitização**: estão todas no `SYSTEM_PROMPT` — edite
  livremente para adaptar a outras bases ou regras de negócio.

## Observações sobre a base

- As colunas `customer_name` e `salesperson` **não têm gabarito** na
  planilha original — o agente ainda assim as normaliza (Title Case),
  mas essas duas não entram no cálculo de acurácia.
- Datas impossíveis no calendário (ex: 30 de fevereiro) devem ser
  sanitizadas como a string `"INVALID"` — é assim que o gabarito trata
  esses casos.
- Quilometragem informada em KM é convertida para milhas
  (`1 km = 0.621371 mi`).
