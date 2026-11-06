import re
from datetime import datetime
import pandas as pd
import gspread
from playwright.sync_api import sync_playwright

ARQUIVO_CREDENCIAIS = "credenciais.json"
NOME_DA_PLANILHA = "Historico_Stats_EWGF"
ARQUIVO_EXCEL_JOGADORES = "tekken_players.xlsx"

CATEGORIAS = {
    "ATTACK": ["Attack Frequency", "Heavy Damage", "Aggressiveness", "Dominance"],
    "DEFENSE": ["Block", "Evasion", "Throw Escape", "Composure"],
    "TECHNIQUE": ["Accuracy", "Judgement", "Retaliation", "Stage Use"],
    "SPIRIT": ["Close Battles", "Comeback", "Fighting Spirit", "Concentration"],
    "APPEAL": ["Respect", "Ambition", "Fairness", "Versatility"]
}

SUB_STATS = [stat for subs in CATEGORIAS.values() for stat in subs]

COLUNAS = (
    ["Data", "Jogador", "Player_ID"] +
    ["ATTACK"] + CATEGORIAS["ATTACK"] +
    ["DEFENSE"] + CATEGORIAS["DEFENSE"] +
    ["TECHNIQUE"] + CATEGORIAS["TECHNIQUE"] +
    ["SPIRIT"] + CATEGORIAS["SPIRIT"] +
    ["APPEAL"] + CATEGORIAS["APPEAL"]
)

def carregar_jogadores_do_excel(caminho_arquivo):
    df = pd.read_excel(caminho_arquivo, sheet_name="Tekken Players")
    jogadores = {}
    for _, row in df.iterrows():
        nick = str(row["Nick"]).strip()
        tekken_id = str(row["Tekken ID"]).strip()
        jogadores[nick] = tekken_id
    return jogadores

def extrair_stats_jogador(page, nome_referencia, tekken_id):
    id_limpo = tekken_id.replace("-", "").strip()
    url = f"https://ewgf.gg/player/{id_limpo}"
    print(f"A recolher dados de: {nome_referencia} ({id_limpo})...")
    
    # Abre a página e espera a rede acalmar para garantir que o JS executou
    page.goto(url, wait_until="networkidle", timeout=45000)
    
    # Aguarda firmemente o bloco de estatísticas aparecer na tela
    try:
        page.wait_for_selector("text=Attack Frequency", timeout=15000)
    except Exception:
        print(f"  -> Aviso: Elementos demoraram para {nome_referencia}, a tentar recarregar...")
        page.reload(wait_until="networkidle", timeout=30000)
        page.wait_for_selector("text=Attack Frequency", timeout=15000)
    
    texto_pagina = page.inner_text("body")
    
    dados = {
        "Data": datetime.now().strftime("%Y-%m-%d"),
        "Jogador": nome_referencia,
        "Player_ID": id_limpo
    }
    
    # Extrai cada um dos 20 sub-atributos
    for stat in SUB_STATS:
        padrao = rf"{re.escape(stat)}\s+(\d+)\s*/\s*25"
        match = re.search(padrao, texto_pagina)
        dados[stat] = int(match.group(1)) if match else 0

    # Soma os 4 sub-atributos de cada secção para obter a nota sobre 100
    for cat, subs in CATEGORIAS.items():
        dados[cat] = sum(dados[s] for s in subs)

    return [dados[col] for col in COLUNAS]

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
