# 📦 Auditoria de Estoque Automatizada (Python + Excel + SQL)

Ferramenta que **compara a contagem física do estoque com o que está registrado no sistema (ERP)** e gera automaticamente um **relatório em Excel** pronto para a gestão.

## 💡 De onde veio a ideia

Este projeto nasceu da minha experiência real:

- **Assistente de Farmácia na Santa Casa de Atibaia** — fazia o controle de estoque físico e digital de medicamentos no ERP hospitalar, onde itens em falta ou vencidos são um problema crítico.
- **Auditor e Inspetor de Logística na MagaLog** — auditava fluxos de mercadorias com coletores de dados e analisava as divergências para encontrar erros de processo.

Nas duas funções, o trabalho de cruzar planilhas e montar relatórios era manual. Este script faz isso em segundos.

## ✅ O que ele faz

| Análise | Para que serve |
|---|---|
| **Divergências** | Mostra cada item onde a contagem física não bate com o sistema (falta ou sobra) e o impacto em R$ |
| **Acurácia do inventário** | % de itens sem divergência — indicador clássico de qualidade de estoque |
| **Estoque crítico** | Itens abaixo do estoque mínimo, com sugestão de quanto repor |
| **Vencimentos** | Lotes que vencem nos próximos N dias (vermelho até 30 dias, amarelo até 60) |
| **Resumo + gráfico** | Painel com os indicadores principais e gráfico do impacto financeiro |

## 📊 Exemplo de resultado

```
Auditoria concluída!
  - Itens auditados: 16
  - Itens sem divergência: 8
  - Acurácia do inventário: 50.0%
  - Itens com falta: 7
  - Itens com sobra: 1
  - Perda financeira (faltas): R$ -777.50
  - Valor total em estoque: R$ 13,319.50
  - Itens em estoque crítico: 3
  - Lotes vencendo em até 60 dias: 5
```

O Excel gerado fica em [`relatorios/relatorio_auditoria.xlsx`](relatorios/relatorio_auditoria.xlsx), com as abas **Resumo**, **Inventário completo**, **Divergências**, **Estoque crítico** e **Vencimentos** — com cores, filtros e cabeçalho congelado.

## 🚀 Como usar

```bash
pip install -r requirements.txt

# Usa os dados de exemplo
python src/auditoria_estoque.py

# Com seus próprios dados e alerta de vencimento de 90 dias
python src/auditoria_estoque.py --arquivo minha_contagem.csv --dias-vencimento 90
```

O CSV de entrada precisa ter as colunas:
`codigo, produto, categoria, lote, validade, estoque_minimo, qtd_sistema, qtd_fisica, custo_unitario`

## 🗄️ Versão em SQL

Em [`sql/auditoria_estoque.sql`](sql/auditoria_estoque.sql) estão as mesmas análises escritas em SQL (divergências, acurácia por categoria, estoque crítico e vencimentos), usando `CASE WHEN`, `GROUP BY`, agregações e cálculo com datas.

```bash
sqlite3 estoque.db < sql/auditoria_estoque.sql
```

## 🛠️ Tecnologias

- **Python** — lógica da auditoria
- **Pandas** — leitura e cálculos sobre os dados
- **openpyxl** — geração do Excel formatado com gráfico
- **SQL** — as mesmas análises direto no banco de dados

## 📁 Estrutura

```
├── dados/contagem_estoque.csv      # dados de exemplo (fictícios)
├── src/auditoria_estoque.py        # script principal
├── sql/auditoria_estoque.sql       # versão em SQL
└── relatorios/relatorio_auditoria.xlsx  # relatório gerado
```

> Os dados deste repositório são fictícios e servem apenas para demonstração.

---
Feito por **Marcos André Sousa Cardoso** · [LinkedIn](https://www.linkedin.com/in/marcsou777) · [GitHub](https://github.com/Marcos-Sousa-EXE)
