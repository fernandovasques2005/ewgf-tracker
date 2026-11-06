import re
from datetime import datetime
from pathlib import Path

import gspread
import pandas as pd
from playwright.sync_api import Page, TimeoutError as PlaywrightTimeoutError
from playwright.sync_api import sync_playwright


# ============================================================
# CONFIGURAÇÕES
# ============================================================

BASE_DIR = Path(__file__).resolve().parent

# O workflow do GitHub Actions deve criar este arquivo a partir
# de um Repository Secret antes de executar este script.
ARQUIVO_CREDENCIAIS = BASE_DIR / "credenciais.json"

NOME_DA_PLANILHA = "Historico_Stats_EWGF"
ARQUIVO_EXCEL_JOGADORES = BASE_DIR / "tekken_players.xlsx"

NOME_DA_ABA = None
# None = usar a primeira aba da planilha Google Sheets.
# Para escolher uma aba específica, use, por exemplo:
# NOME_DA_ABA = "Página1"

URL_BASE_PERFIL = "https://ewgf.gg/player"

# Durante o primeiro teste, deixe True para salvar uma imagem
# quando ocorrer erro em um jogador.
SALVAR_SCREENSHOT_EM_ERRO = True
PASTA_DEBUG = BASE_DIR / "debug_ewgf"


# ============================================================
# NOMES DOS STATS
# ============================================================

CATEGORIAS = {
    "ATTACK": [
        "Attack Frequency",
        "Heavy Damage",
        "Aggressiveness",
        "Dominance",
    ],
    "DEFENSE": [
        "Block",
        "Evasion",
        "Throw Escape",
        "Composure",
    ],
    "TECHNIQUE": [
        "Accuracy",
        "Judgement",
        "Retaliation",
        "Stage Use",
    ],
    "SPIRIT": [
        "Close Battles",
        "Comeback",
        "Fighting Spirit",
        "Concentration",
    ],
    "APPEAL": [
        "Respect",
        "Ambition",
        "Fairness",
        "Versatility",
    ],
}

SUB_STATS = [
    stat
    for substats in CATEGORIAS.values()
    for stat in substats
]

COLUNAS = (
    ["Data", "Jogador", "Player_ID"]
    + ["ATTACK"]
    + CATEGORIAS["ATTACK"]
    + ["DEFENSE"]
    + CATEGORIAS["DEFENSE"]
    + ["TECHNIQUE"]
    + CATEGORIAS["TECHNIQUE"]
    + ["SPIRIT"]
    + CATEGORIAS["SPIRIT"]
    + ["APPEAL"]
    + CATEGORIAS["APPEAL"]
)


# ============================================================
# LEITURA DA PLANILHA EXCEL
# ============================================================

def carregar_jogadores_do_excel(caminho_arquivo: Path) -> dict[str, str]:
    """Lê a aba 'Tekken Players' e retorna {Nick: Tekken ID}."""

    if not caminho_arquivo.exists():
        raise FileNotFoundError(
            f"Arquivo Excel não encontrado: {caminho_arquivo}"
        )

    df = pd.read_excel(
        caminho_arquivo,
        sheet_name="Tekken Players",
    )

    colunas_obrigatorias = {"Nick", "Tekken ID"}
    colunas_faltantes = colunas_obrigatorias - set(df.columns)

    if colunas_faltantes:
        raise ValueError(
            "Colunas ausentes no Excel: "
            + ", ".join(sorted(colunas_faltantes))
            + ". As colunas esperadas são 'Nick' e 'Tekken ID'."
        )

    jogadores = {}

    for numero_linha, row in df.iterrows():
        if pd.isna(row["Nick"]) or pd.isna(row["Tekken ID"]):
            print(
                f"Aviso: linha {numero_linha + 2} ignorada porque Nick ou Tekken ID está vazio."
            )
            continue

        nick = str(row["Nick"]).strip()
        tekken_id = str(row["Tekken ID"]).strip()

        if not nick or not tekken_id:
            continue

        jogadores[nick] = tekken_id

    return jogadores


# ============================================================
# EXTRAÇÃO DO EWGF.GG
# ============================================================

def normalizar_tekken_id(tekken_id: str) -> str:
    """Remove hífens, espaços e outros caracteres não alfanuméricos."""

    id_limpo = re.sub(r"[^A-Za-z0-9]", "", str(tekken_id).strip())

    if not id_limpo:
        raise ValueError(f"Tekken ID inválido: {tekken_id!r}")

    return id_limpo


def abrir_detalhes_do_pentagono(page: Page) -> None:
    """Clica no controle Details da seção Stat Pentagon."""

    page.get_by_text("Stat Pentagon", exact=True).wait_for(timeout=20000)

    # Primeiro tenta encontrar um botão com o nome Details.
    try:
        detalhes = page.get_by_role(
            "button",
            name=re.compile(r"^Details$", re.IGNORECASE),
        )

        if detalhes.count() > 0:
            detalhes.last.click()
            return
    except Exception:
        pass

    # Fallback para versões em que Details não é um elemento button.
    detalhes = page.get_by_text("Details", exact=True)

    if detalhes.count() == 0:
        raise RuntimeError(
            "O controle 'Details' não foi encontrado na seção Stat Pentagon."
        )

    detalhes.last.click()


def extrair_bloco_do_pentagono(page: Page) -> str:
    """Abre Details e retorna o texto da seção Stat Pentagon."""

    abrir_detalhes_do_pentagono(page)

    # Aguarda o conteúdo expandido aparecer.
    try:
        page.get_by_text("Attack Frequency", exact=True).wait_for(
            timeout=15000
        )
    except PlaywrightTimeoutError as erro:
        texto_atual = page.locator("body").inner_text()
        raise RuntimeError(
            "Após clicar em Details, 'Attack Frequency' não apareceu. "
            "Os nomes dos subatributos podem ter mudado. "
            f"Texto atual da página: {texto_atual[:2500]!r}"
        ) from erro

    texto_pagina = page.locator("body").inner_text()

    inicio = texto_pagina.find("Stat Pentagon")
    fim = texto_pagina.find("Global Statistics", inicio)

    if inicio == -1:
        raise RuntimeError("A seção 'Stat Pentagon' não foi encontrada.")

    if fim == -1:
        return texto_pagina[inicio:]

    return texto_pagina[inicio:fim]


def extrair_stats_jogador(
    page: Page,
    nome_referencia: str,
    tekken_id: str,
) -> list:
    """Abre o perfil e extrai os 20 subatributos do pentágono."""

    id_limpo = normalizar_tekken_id(tekken_id)
    url = f"{URL_BASE_PERFIL}/{id_limpo}"

    print(f"Buscando dados de: {nome_referencia} ({id_limpo})...")
    print(f"URL: {url}")

    page.goto(
        url,
        wait_until="domcontentloaded",
        timeout=60000,
    )

    try:
        bloco_pentagono = extrair_bloco_do_pentagono(page)
    except Exception:
        print(
            f"Não foi possível extrair o pentágono de {nome_referencia}."
        )
        print("Texto disponível na página:")
        print(page.locator("body").inner_text()[:4000])
        raise

    print("Bloco do pentágono encontrado:")
    print(bloco_pentagono)

    dados = {
        "Data": datetime.now().strftime("%Y-%m-%d"),
        "Jogador": nome_referencia,
        "Player_ID": id_limpo,
    }

    faltantes = []

    for stat in SUB_STATS:
        # Aceita os formatos:
        #   Attack Frequency 25/25
        #   Attack Frequency: 25 / 25
        #   Attack Frequency 25
        padrao = (
            rf"\b{re.escape(stat)}\b"
            rf"\s*[:\-]?\s*"
            rf"(\d{{1,3}})"
            rf"(?:\s*/\s*25)?"
            rf"\b"
        )

        match = re.search(
            padrao,
            bloco_pentagono,
            flags=re.IGNORECASE,
        )

        if match:
            dados[stat] = int(match.group(1))
        else:
            faltantes.append(stat)

    if faltantes:
        raise RuntimeError(
            "Os seguintes stats não foram encontrados: "
            + ", ".join(faltantes)
            + ". Verifique os nomes exibidos pelo site."
        )

    # Calcula a nota agregada de cada categoria.
    for categoria, substats in CATEGORIAS.items():
        dados[categoria] = sum(dados[stat] for stat in substats)

    linha = [dados[coluna] for coluna in COLUNAS]

    if len(linha) != len(COLUNAS):
        raise RuntimeError(
            f"Quantidade de valores ({len(linha)}) diferente da quantidade "
            f"de colunas ({len(COLUNAS)})."
        )

    print(
        f"OK: {nome_referencia} - "
        + ", ".join(
            f"{categoria}={dados[categoria]}"
            for categoria in CATEGORIAS
        )
    )

    return linha


# ============================================================
# GOOGLE SHEETS
# ============================================================

def obter_aba_google_sheets():
    """Abre a planilha e retorna a aba configurada."""

    if not ARQUIVO_CREDENCIAIS.exists():
        raise FileNotFoundError(
            "Arquivo de credenciais não encontrado: "
            f"{ARQUIVO_CREDENCIAIS}. "
            "O workflow precisa criá-lo a partir do Repository Secret."
        )

    gc = gspread.service_account(
        filename=str(ARQUIVO_CREDENCIAIS)
    )

    planilha = gc.open(NOME_DA_PLANILHA)

    if NOME_DA_ABA:
        return planilha.worksheet(NOME_DA_ABA)

    return planilha.sheet1


def salvar_no_google_sheets(linhas_novas: list[list]) -> None:
    """Cria o cabeçalho se necessário e acrescenta as linhas."""

    aba = obter_aba_google_sheets()

    cabecalho_atual = aba.row_values(1)

    if not cabecalho_atual:
        aba.append_row(
            COLUNAS,
            value_input_option="USER_ENTERED",
        )
    elif cabecalho_atual != COLUNAS:
        print(
            "Aviso: o cabeçalho atual da aba é diferente do esperado."
        )
        print(f"Cabeçalho atual: {cabecalho_atual}")
        print(f"Cabeçalho esperado: {COLUNAS}")
        print(
            "As linhas não serão gravadas para evitar deslocamento de colunas. "
            "Use uma aba nova ou ajuste o cabeçalho."
        )
        return

    if not linhas_novas:
        print("Nenhuma linha nova para gravar.")
        return

    aba.append_rows(
        linhas_novas,
        value_input_option="USER_ENTERED",
    )

    print(
        f"Sucesso: {len(linhas_novas)} registro(s) guardado(s) "
        f"na planilha '{NOME_DA_PLANILHA}'."
    )


# ============================================================
# EXECUÇÃO PRINCIPAL
# ============================================================

def salvar_debug(page: Page, nome: str, tekken_id: str) -> None:
    """Salva screenshot e texto da página para diagnosticar mudanças no site."""

    if not SALVAR_SCREENSHOT_EM_ERRO:
        return

    try:
        PASTA_DEBUG.mkdir(parents=True, exist_ok=True)
        nome_seguro = re.sub(r"[^A-Za-z0-9_.-]", "_", nome)

        page.screenshot(
            path=str(PASTA_DEBUG / f"{nome_seguro}_{tekken_id}.png"),
            full_page=True,
        )

        texto = page.locator("body").inner_text()
        (PASTA_DEBUG / f"{nome_seguro}_{tekken_id}.txt").write_text(
            texto,
            encoding="utf-8",
        )

    except Exception as erro:
        print(f"Aviso: não foi possível salvar debug: {erro}")


def main() -> None:
    jogadores = carregar_jogadores_do_excel(
        ARQUIVO_EXCEL_JOGADORES
    )

    print(
        f"Total de jogadores carregados do Excel: {len(jogadores)}"
    )

    if not jogadores:
        raise RuntimeError("Nenhum jogador válido foi encontrado no Excel.")

    linhas_para_inserir = []
    erros = []

    with sync_playwright() as playwright:
        browser = playwright.chromium.launch(headless=True)

        context = browser.new_context(
            user_agent=(
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/131.0.0.0 Safari/537.36"
            ),
            viewport={"width": 1440, "height": 1000},
            locale="en-US",
        )

        page = context.new_page()

        try:
            for nome, tekken_id in jogadores.items():
                try:
                    linha = extrair_stats_jogador(
                        page,
                        nome,
                        tekken_id,
                    )
                    linhas_para_inserir.append(linha)

                except Exception as erro:
                    erro_texto = str(erro)
                    erros.append((nome, tekken_id, erro_texto))

                    print(
                        f"ERRO ao coletar {nome} ({tekken_id}): "
                        f"{erro_texto}"
                    )

                    salvar_debug(
                        page,
                        nome,
                        normalizar_tekken_id(tekken_id),
                    )

        finally:
            context.close()
            browser.close()

    if linhas_para_inserir:
        salvar_no_google_sheets(linhas_para_inserir)
    else:
        print("Nenhum jogador foi coletado com sucesso.")

    print(
        f"Resumo: {len(linhas_para_inserir)} sucesso(s), "
        f"{len(erros)} erro(s)."
    )

    if erros:
        print("Jogadores com erro:")
        for nome, tekken_id, erro in erros:
            print(f"- {nome} ({tekken_id}): {erro}")


if __name__ == "__main__":
    main()

