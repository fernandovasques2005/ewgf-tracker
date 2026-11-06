import re
from datetime import datetime
import pandas as pd
import gspread
from playwright.sync_api import sync_playwright

ARQUIVO_CREDENCIAIS = "credenciais.json"
NOME_DA_PLANILHA = "Historico_Stats_EWGF"
ARQUIVO_EXCEL_JOGADORES = "tekken_players.xlsx"

ATRIBUTOS = ["Attack", "Defense", "Technique", "Spirit", "Appeal"]
COLUNAS = ["Data", "Jogador", "Player_ID"] + ATRIBUTOS

def carregar_jogadores_do_excel(caminho_arquivo):
    df = pd.read_excel(caminho_arquivo, sheet_name="Tekken Players")
    jogadores = {}
    for _, row in df.iterrows():
        nick = str(row["Nick"]).strip()
        tekken_id = str(row["Tekken ID"]).strip()
        jogadores[nick] = tekken_id
    return jogadores

def extrair_stats_jogador(page, nome_referencia, tekken_id):
    id_limpo = re.sub(r"[^A-Za-z0-9]", "", str(tekken_id).strip())
    url = f"https://ewgf.gg/player/{id_limpo}"
    print(f"Buscando dados de: {nome_referencia} ({id_limpo})...")

    page.goto(url, wait_until="domcontentloaded", timeout=60000)
    page.get_by_text("Stat Pentagon", exact=True).wait_for(timeout=20000)

    detalhes = page.get_by_text("Details", exact=True)
    if not detalhes.count():
        raise ValueError("A seção Details não foi encontrada no perfil.")
    detalhes.first.click()

    page.wait_for_function(
        "() => /\\bAttack\\s+\\d{1,3}\\b/i.test(document.body.innerText)",
        timeout=10000,
    )

    texto = page.locator("body").inner_text()
    inicio = texto.find("Stat Pentagon")
    fim = texto.find("Global Statistics", inicio)
    bloco = texto[inicio:fim] if fim > inicio else texto[inicio:inicio + 500]

    dados = {
        "Data": datetime.now().strftime("%Y-%m-%d"),
        "Jogador": nome_referencia,
        "Player_ID": id_limpo,
    }

    for atributo in ATRIBUTOS:
        padrao = rf"\b{re.escape(atributo)}\b\s*[:\-]?\s*(\d{{1,3}})\b"
        match = re.search(padrao, bloco, flags=re.IGNORECASE)
        if not match:
            raise ValueError(
                f"Não encontrei o valor de '{atributo}'. "
                f"Texto do pentágono encontrado: {bloco!r}"
            )
        dados[atributo] = int(match.group(1))

    return [dados[coluna] for coluna in COLUNAS]

def salvar_no_google_sheets(linhas_novas):
    gc = gspread.service_account(filename=ARQUIVO_CREDENCIAIS)
    sh = gc.open(NOME_DA_PLANILHA)
    aba = sh.sheet1

    if not aba.row_values(1):
        aba.append_row(COLUNAS, value_input_option="USER_ENTERED")

    aba.append_rows(linhas_novas, value_input_option="USER_ENTERED")
    print(f"\nSucesso! {len(linhas_novas)} registo(s) guardado(s) na folha '{NOME_DA_PLANILHA}'.")

def main():
    jogadores = carregar_jogadores_do_excel(ARQUIVO_EXCEL_JOGADORES)
    print(f"Total de jogadores carregados: {len(jogadores)}\n")

    linhas_para_inserir = []
    
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
        )
        page = context.new_page()
        
        for nome, tid in jogadores.items():
            try:
                linha = extrair_stats_jogador(page, nome, tid)
                linhas_para_inserir.append(linha)
            except Exception as e:
                print(f"  -> Erro ao recolher {nome} ({tid}): {e}")
                
        browser.close()

    if linhas_para_inserir:
        salvar_no_google_sheets(linhas_para_inserir)

if __name__ == "__main__":
    main()
