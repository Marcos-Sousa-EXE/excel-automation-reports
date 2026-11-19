"""
Auditoria de Estoque - gera um relatório em Excel a partir de uma contagem de inventário.

Inspirado na rotina de controle de estoque de uma farmácia hospitalar e na
auditoria de mercadorias em um centro de distribuição:

1. Compara a quantidade registrada no sistema (ERP) com a contagem física.
2. Aponta as divergências (falta ou sobra) e o impacto financeiro.
3. Lista os itens abaixo do estoque mínimo (estoque crítico).
4. Lista os lotes que vencem nos próximos dias.
5. Gera um arquivo .xlsx formatado, pronto para enviar à gestão.

Uso:
    python src/auditoria_estoque.py
    python src/auditoria_estoque.py --arquivo dados/contagem_estoque.csv --dias-vencimento 60
"""

import argparse
from datetime import date, datetime
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

PASTA_PROJETO = Path(__file__).resolve().parent.parent

# Cores usadas no relatório
AZUL_ESCURO = "1F3864"
VERMELHO_CLARO = "F8D7DA"
AMARELO_CLARO = "FFF3CD"
VERDE_CLARO = "D4EDDA"
CINZA_CLARO = "F2F2F2"

BORDA_FINA = Border(*(Side(style="thin", color="BFBFBF"),) * 4)


# ---------------------------------------------------------------------------
# 1. Leitura e cálculos
# ---------------------------------------------------------------------------

def carregar_contagem(caminho: Path) -> pd.DataFrame:
    """Lê o CSV da contagem e converte a coluna de validade em data."""
    df = pd.read_csv(caminho)
    df["validade"] = pd.to_datetime(df["validade"])
    return df


def calcular_auditoria(df: pd.DataFrame, data_ref: date) -> pd.DataFrame:
    """Adiciona as colunas de divergência, valor, situação e dias até vencer."""
    df = df.copy()
    df["divergencia"] = df["qtd_fisica"] - df["qtd_sistema"]
    df["valor_divergencia"] = (df["divergencia"] * df["custo_unitario"]).round(2)
    df["valor_estoque"] = (df["qtd_fisica"] * df["custo_unitario"]).round(2)
    df["dias_para_vencer"] = (df["validade"] - pd.Timestamp(data_ref)).dt.days

    def situacao(linha):
        if linha["divergencia"] < 0:
            return "Falta"
        if linha["divergencia"] > 0:
            return "Sobra"
        return "OK"

    df["situacao"] = df.apply(situacao, axis=1)
    df["estoque_critico"] = df["qtd_fisica"] < df["estoque_minimo"]
    return df


def montar_resumo(df: pd.DataFrame, dias_vencimento: int) -> dict:
    """Indicadores principais da auditoria."""
    total_itens = len(df)
    itens_ok = int((df["situacao"] == "OK").sum())
    return {
        "Itens auditados": total_itens,
        "Itens sem divergência": itens_ok,
        "Acurácia do inventário": itens_ok / total_itens,
        "Itens com falta": int((df["situacao"] == "Falta").sum()),
        "Itens com sobra": int((df["situacao"] == "Sobra").sum()),
        "Perda financeira (faltas)": float(df.loc[df["divergencia"] < 0, "valor_divergencia"].sum()),
        "Valor total em estoque": float(df["valor_estoque"].sum()),
        "Itens em estoque crítico": int(df["estoque_critico"].sum()),
        f"Lotes vencendo em até {dias_vencimento} dias": int((df["dias_para_vencer"] <= dias_vencimento).sum()),
    }


# ---------------------------------------------------------------------------
# 2. Escrita do Excel
# ---------------------------------------------------------------------------

def estilizar_cabecalho(ws, linha: int, n_colunas: int):
    for col in range(1, n_colunas + 1):
        cel = ws.cell(row=linha, column=col)
        cel.font = Font(bold=True, color="FFFFFF")
        cel.fill = PatternFill("solid", fgColor=AZUL_ESCURO)
        cel.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cel.border = BORDA_FINA


def ajustar_colunas(ws, larguras: dict):
    for col, largura in larguras.items():
        ws.column_dimensions[col].width = largura


def escrever_tabela(ws, df: pd.DataFrame, colunas: dict, linha_inicial: int = 1,
                    formatos: dict | None = None, cor_linha=None):
    """Escreve um DataFrame na planilha.

    colunas: {nome_no_df: titulo_no_excel}
    formatos: {nome_no_df: formato_numero_excel}
    cor_linha: função que recebe a linha e devolve uma cor (ou None)
    """
    formatos = formatos or {}
    for i, titulo in enumerate(colunas.values(), start=1):
        ws.cell(row=linha_inicial, column=i, value=titulo)
    estilizar_cabecalho(ws, linha_inicial, len(colunas))

    for r, (_, linha) in enumerate(df.iterrows(), start=linha_inicial + 1):
        cor = cor_linha(linha) if cor_linha else None
        for c, nome in enumerate(colunas.keys(), start=1):
            valor = linha[nome]
            if isinstance(valor, pd.Timestamp):
                valor = valor.to_pydatetime()
            if hasattr(valor, "item"):  # numpy -> python
                valor = valor.item()
            cel = ws.cell(row=r, column=c, value=valor)
            cel.border = BORDA_FINA
            if nome in formatos:
                cel.number_format = formatos[nome]
            if cor:
                cel.fill = PatternFill("solid", fgColor=cor)

    ws.freeze_panes = ws.cell(row=linha_inicial + 1, column=1)
    if len(df):
        ultima_col = get_column_letter(len(colunas))
        ws.auto_filter.ref = f"A{linha_inicial}:{ultima_col}{linha_inicial + len(df)}"


def aba_resumo(wb: Workbook, resumo: dict, df: pd.DataFrame, data_ref: date):
    ws = wb.active
    ws.title = "Resumo"

    ws["A1"] = "Relatório de Auditoria de Estoque"
    ws["A1"].font = Font(bold=True, size=16, color=AZUL_ESCURO)
    ws["A2"] = f"Data de referência: {data_ref.strftime('%d/%m/%Y')}"
    ws["A2"].font = Font(italic=True, color="595959")

    ws["A4"], ws["B4"] = "Indicador", "Valor"
    estilizar_cabecalho(ws, 4, 2)

    for i, (indicador, valor) in enumerate(resumo.items(), start=5):
        ws.cell(row=i, column=1, value=indicador).border = BORDA_FINA
        cel = ws.cell(row=i, column=2, value=valor)
        cel.border = BORDA_FINA
        if "Acurácia" in indicador:
            cel.number_format = "0.0%"
        elif "Perda" in indicador or "Valor" in indicador:
            cel.number_format = 'R$ #,##0.00'
        if i % 2 == 0:
            for col in (1, 2):
                ws.cell(row=i, column=col).fill = PatternFill("solid", fgColor=CINZA_CLARO)

    ajustar_colunas(ws, {"A": 38, "B": 18})


def gerar_relatorio(df: pd.DataFrame, resumo: dict, destino: Path, data_ref: date, dias_vencimento: int):
    wb = Workbook()
    aba_resumo(wb, resumo, df, data_ref)

    # Aba com todos os itens
    ws = wb.create_sheet("Inventário completo")
    cores = {"Falta": VERMELHO_CLARO, "Sobra": AMARELO_CLARO, "OK": None}
    escrever_tabela(
        ws, df,
        colunas={
            "codigo": "Código", "produto": "Produto", "categoria": "Categoria",
            "lote": "Lote", "validade": "Validade", "qtd_sistema": "Qtd. Sistema",
            "qtd_fisica": "Qtd. Física", "divergencia": "Divergência",
            "valor_divergencia": "Divergência (R$)", "situacao": "Situação",
        },
        formatos={"validade": "DD/MM/YYYY", "valor_divergencia": 'R$ #,##0.00'},
        cor_linha=lambda l: cores[l["situacao"]],
    )
    ajustar_colunas(ws, {"A": 10, "B": 34, "C": 17, "D": 9, "E": 12, "F": 12,
                         "G": 12, "H": 12, "I": 16, "J": 11})

    # Aba de divergências
    ws = wb.create_sheet("Divergências")
    div = df[df["divergencia"] != 0].sort_values("valor_divergencia")
    escrever_tabela(
        ws, div,
        colunas={"codigo": "Código", "produto": "Produto", "lote": "Lote",
                 "qtd_sistema": "Qtd. Sistema", "qtd_fisica": "Qtd. Física",
                 "divergencia": "Divergência", "valor_divergencia": "Impacto (R$)",
                 "situacao": "Situação"},
        formatos={"valor_divergencia": 'R$ #,##0.00'},
        cor_linha=lambda l: cores[l["situacao"]],
    )
    ajustar_colunas(ws, {"A": 10, "B": 34, "C": 9, "D": 12, "E": 12, "F": 12, "G": 14, "H": 11})

    # Gráfico no Resumo usando os dados da aba Divergências
    if len(div):
        grafico = BarChart()
        grafico.type = "bar"
        grafico.title = "Impacto financeiro das divergências (R$)"
        grafico.legend = None
        dados = Reference(ws, min_col=7, min_row=1, max_row=len(div) + 1)
        categorias = Reference(ws, min_col=2, min_row=2, max_row=len(div) + 1)
        grafico.add_data(dados, titles_from_data=True)
        grafico.set_categories(categorias)
        grafico.height, grafico.width = 9, 20
        wb["Resumo"].add_chart(grafico, "D4")

    # Aba de estoque crítico
    ws = wb.create_sheet("Estoque crítico")
    crit = df[df["estoque_critico"]].copy()
    crit["repor"] = crit["estoque_minimo"] - crit["qtd_fisica"]
    escrever_tabela(
        ws, crit,
        colunas={"codigo": "Código", "produto": "Produto", "estoque_minimo": "Estoque mínimo",
                 "qtd_fisica": "Qtd. Física", "repor": "Sugestão de reposição"},
        cor_linha=lambda l: VERMELHO_CLARO,
    )
    ajustar_colunas(ws, {"A": 10, "B": 34, "C": 15, "D": 12, "E": 20})

    # Aba de vencimentos
    ws = wb.create_sheet("Vencimentos")
    venc = df[df["dias_para_vencer"] <= dias_vencimento].sort_values("dias_para_vencer")
    escrever_tabela(
        ws, venc,
        colunas={"codigo": "Código", "produto": "Produto", "lote": "Lote",
                 "validade": "Validade", "dias_para_vencer": "Dias para vencer",
                 "qtd_fisica": "Qtd. Física"},
        formatos={"validade": "DD/MM/YYYY"},
        cor_linha=lambda l: VERMELHO_CLARO if l["dias_para_vencer"] <= 30 else AMARELO_CLARO,
    )
    ajustar_colunas(ws, {"A": 10, "B": 34, "C": 9, "D": 12, "E": 16, "F": 12})

    destino.parent.mkdir(parents=True, exist_ok=True)
    wb.save(destino)


# ---------------------------------------------------------------------------
# 3. Execução
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(description="Gera o relatório de auditoria de estoque em Excel.")
    parser.add_argument("--arquivo", default=PASTA_PROJETO / "dados" / "contagem_estoque.csv",
                        help="CSV com a contagem de estoque")
    parser.add_argument("--saida", default=PASTA_PROJETO / "relatorios" / "relatorio_auditoria.xlsx",
                        help="Caminho do Excel gerado")
    parser.add_argument("--dias-vencimento", type=int, default=60,
                        help="Alertar lotes que vencem em até N dias (padrão: 60)")
    parser.add_argument("--data-referencia", default=None,
                        help="Data da auditoria no formato AAAA-MM-DD (padrão: hoje)")
    args = parser.parse_args()

    data_ref = (datetime.strptime(args.data_referencia, "%Y-%m-%d").date()
                if args.data_referencia else date.today())

    df = calcular_auditoria(carregar_contagem(Path(args.arquivo)), data_ref)
    resumo = montar_resumo(df, args.dias_vencimento)
    gerar_relatorio(df, resumo, Path(args.saida), data_ref, args.dias_vencimento)

    print("Auditoria concluída!")
    for indicador, valor in resumo.items():
        if isinstance(valor, float) and "Acurácia" in indicador:
            valor = f"{valor:.1%}"
        elif isinstance(valor, float):
            valor = f"R$ {valor:,.2f}"
        print(f"  - {indicador}: {valor}")
    print(f"Relatório salvo em: {args.saida}")


if __name__ == "__main__":
    main()
