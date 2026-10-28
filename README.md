# Agente de tratamento de dados + dashboard de vendas

Projeto que leva uma planilha de vendas com dados bagunçados até um painel interativo. Primeiro, um agente em Python limpa e padroniza a planilha seguindo regras escritas num schema e gera um relatório de qualidade. Depois, o dashboard é montado só com os dados que passaram pelo tratamento.

Os dados são fictícios e foram fornecidos pelos cursos da DIO usados como base do projeto.

**Dashboard publicado:** https://nataciaduarte.github.io/agente-dashboard-vendas-porsche/dashboard/

## O problema

A planilha original tem 100 vendas com datas em vários formatos diferentes, preços escritos como "188k USD" ou "eighty two thousand USD", quilometragem em milhas e em quilômetros, estados com nome completo ou sigla e status de entrega com erros de digitação.

Na primeira versão do dashboard, esses dados entraram sem filtro suficiente: o painel mostrava vendas com data em 2027 e somava vendas canceladas na receita.

## Como funciona

```
dados/Porsche_data_base.xlsx          planilha original
        │
        ▼
agente/porsche_sanitizer_agent.py     aplica as regras de agente/schema.md
        │
        ▼
dados/Porsche_data_base_sanitized.xlsx   planilha tratada + aba Relatorio_Qualidade
        │
        ▼
dashboard/gerar_dados.py              seleciona as vendas válidas e gera data.js
        │
        ▼
dashboard/index.html                  painel interativo
```

O agente mantém todas as colunas originais e adiciona ao lado de cada uma a versão tratada. Quando um valor não pode ser corrigido com segurança, ele vira `INVALID`, em vez de ficar em branco ou ser adivinhado.

## Resultado do tratamento

| Etapa | Registros |
|---|---:|
| Vendas na planilha original | 100 |
| Data impossível (ex.: 30 de fevereiro, mês 15) | 24 |
| Data futura, posterior a 28/10/2026 | 20 |
| Vendas com data válida | 56 |
| Canceladas, fora dos indicadores | 5 |
| Vendas consideradas no dashboard | 51 |

## Regras adicionadas nesta versão

1. **Datas futuras viram `INVALID`.** Uma venda com data posterior ao dia em que o agente roda não pode ser uma venda concluída. O relatório de qualidade conta as datas impossíveis e as futuras separadamente.
2. **Vendas canceladas ficam fora dos indicadores.** Elas continuam na planilha tratada, mas o dashboard não as soma em vendas, receita e ticket médio, e mostra quantas foram excluídas.

## Como rodar

```bash
pip install -r requirements.txt
python agente/porsche_sanitizer_agent.py dados/Porsche_data_base.xlsx dados/Porsche_data_base_sanitized.xlsx --ref 2026-10-28
cd dashboard
python gerar_dados.py ../dados/Porsche_data_base_sanitized.xlsx
```

O parâmetro `--ref` fixa a data de referência para datas futuras. Sem ele, o agente usa a data do dia, e o resultado muda com o tempo.

## Minha participação

O agente e a primeira versão do dashboard foram produzidos com apoio de IA durante os cursos "Criando Agentes de Tratamento de Dados" e "Criando uma Dashboard da Porsche com Agentes de IA", da DIO. Nesta versão, revisei o dashboard publicado com apoio de IA, identificamos os problemas de qualidade dos dados e eu decidi as duas regras novas. A implementação das mudanças no agente e no painel foi feita com o Claude, e eu conferi os números do dashboard com a planilha tratada.

## Tecnologias

Python (openpyxl, pandas), HTML, CSS, JavaScript e GitHub Pages.

---

**Natácia Duarte** · [LinkedIn](https://www.linkedin.com/in/nataciaduarte7a4860381) · [GitHub](https://github.com/nataciaduarte)
