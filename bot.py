import asyncio
from playwright.async_api import async_playwright, TimeoutError, expect
from playwright._impl._errors import TargetClosedError
import unidecode
import sys
import os
import re
from datetime import datetime

# --- Configuração para PyInstaller e Playwright ---
if getattr(sys, 'frozen', False):
    os.environ['PLAYWRIGHT_BROWSERS_PATH'] = os.path.join(sys._MEIPASS, 'ms-playwright')

# --- Caminho Base para PDFs ---
DECRETOS_DIRECTORY = "C:\\DECRETOS_ETCM"

# --- Seletores CSS e XPaths ---
XPATH_MUNICIPIO_LABEL = '//*[@id="consultaPublicaTabPanel:consultaPublicaPCSearchForm:municipio_label"]'
CSS_UNIDADE_TRIGGER = '#consultaPublicaTabPanel\:consultaPublicaPCSearchForm\:unidadeJurisdicionada > div.ui-selectonemenu-trigger.ui-state-default.ui-corner-right'
XPATH_PESQUISAR_BUTTON = '//*[@id="consultaPublicaTabPanel:consultaPublicaPCSearchForm:searchButton"]'
XPATH_TABELA_RESULTADOS = '//*[@id="consultaPublicaTabPanel:consultaPublicaDataTable_data"]' # Para esperar a tabela carregar
XPATH_COMPETENCIA_BASE = '//*[@id="consultaPublicaTabPanel:consultaPublicaDataTable:{row_index}:j_idt117"]/span'
XPATH_SELECIONAR_PRESTACAO_BASE = '//*[@id="consultaPublicaTabPanel:consultaPublicaDataTable:{row_index}:j_idt92:selecionarPrestacao"]/img'
XPATH_CLASSIFICACAO_LABEL = '//*[@id="consultaPublicaTabPanel:formFiltros:classificacao_label"]'
XPATH_BTN_FILTRAR_DOCUMENTO = '//*[@id="consultaPublicaTabPanel:formFiltros:btnFiltrar"]'
XPATH_TABELA_DOCUMENTOS_DATA = '//*[@id="consultaPublicaTabPanel:tabelaDocumentos_data"]'
XPATH_DOWNLOAD_DOC_BASE = '//*[@id="consultaPublicaTabPanel:tabelaDocumentos:{row_index}:j_idt298:downloadDocBinario"]/img'


# --- Classificações a serem pesquisadas ---
CLASSIFICACOES_XPATH_MAP = {
    "PCMGE011": '//*[@id="consultaPublicaTabPanel:formFiltros:classificacao_7"]',
    "PCMGE012": '//*[@id="consultaPublicaTabPanel:formFiltros:classificacao_8"]',
    "PCMGE013": '//*[@id="consultaPublicaTabPanel:formFiltros:classificacao_9"]',
    "PCMGE014": '//*[@id="consultaPublicaTabPanel:formFiltros:classificacao_10"]'
}
CLASSIFICACOES_DECRETOS = list(CLASSIFICACOES_XPATH_MAP.keys())

# --- Cidades que usam XPath fixo para a Unidade Jurisdicionada ---
CIDADES_FIXED_XPATH_UNIDADE = [
    "CACULÉ", "CÂNDIDO SALES", "CONDEÚBA", "ÉRICO CARDOSO", "ITORORÓ",
    "PIRIPÁ", "SEBASTIÃO LARANJEIRAS", "TANHAÇU"
]

# --- Cidades que também pesquisam a Câmara Municipal ---
CIDADES_COM_CAMARA_MUNICIPAL = [
    "BARRA DA ESTIVA", "BELO CAMPO", "CACULÉ", "CONDEÚBA", "ÉRICO CARDOSO",
    "ITAPETINGA", "JACARACI", "PAU BRASIL", "PIRIPÁ", "PLANALTO", "POÇÕES", "URANDI"
]

SAAE_ERICO_CARDOSO = ["ÉRICO CARDOSO"]
IPREVIB_IBICOARA = ["IBICOARA"]
CAPREVAC_CARAIBAS = ["CARAÍBAS"]
CISVITA_CONQUISTA = ["VITÓRIA DA CONQUISTA"]

# --- Cidades para pular pesquisa de Prefeitura e Câmara ---
CIDADES_SKIP_PREF_CAMARA = ["CARAÍBAS", "VITÓRIA DA CONQUISTA"]

# Função principal modificada para usar a nova lógica
async def pesquisar_e_baixar_decretos_melhorada(page, cidade, unidade, competencia):
    """
    Versão melhorada da função principal
    """
    competencia_sanitized = competencia.replace('/', '-')
    save_dir = os.path.join(DECRETOS_DIRECTORY, unidecode.unidecode(cidade).upper(), competencia_sanitized)
    
    try:
        os.makedirs(save_dir, exist_ok=True)
    except OSError as e:
        print(f"ERRO: Não foi possível criar o diretório {save_dir}. Erro: {e}")
        return

    print(f"--- Iniciando busca de decretos para {unidade} - {competencia} ---")

    for classificacao in CLASSIFICACOES_DECRETOS:
        print(f"--- Pesquisando classificação: {classificacao} ---")
        try:
            # Abrir dropdown e selecionar classificação (código original mantido)
            await page.locator(XPATH_CLASSIFICACAO_LABEL).click()
            await page.wait_for_timeout(500)
            
            classificacao_xpath = CLASSIFICACOES_XPATH_MAP[classificacao]
            await page.locator(classificacao_xpath).click()
            await page.wait_for_timeout(500)
            
            await page.locator(XPATH_BTN_FILTRAR_DOCUMENTO).click()
            
            print("Aguardando tabela de documentos recarregar...")
            await page.wait_for_load_state('networkidle', timeout=10000)
            await page.wait_for_timeout(2000)

            tabela_data = page.locator(XPATH_TABELA_DOCUMENTOS_DATA)
            rows = await tabela_data.locator('tr').all()

            if not rows:
                print(f"Nenhum documento encontrado para a classificação {classificacao}.")
                continue

            print(f"Encontrados {len(rows)} documentos. Verificando...")
            
            for i, row in enumerate(rows):
                doc_name_locator = row.locator('td:nth-child(3)')
                doc_name = await doc_name_locator.text_content() or ""
                doc_name = doc_name.strip()

                if "Declaração de Inexistência" in doc_name:
                    print(f"  - Linha {i+1}: Ignorando '{doc_name}'")
                    continue
                
                print(f"  + Linha {i+1}: Baixando '{doc_name}'...")
                
                # Usar a nova função de download
                sucesso = await baixar_documento_pdf(page, i, doc_name, save_dir, "xpath=//cr-icon-button[@id=\"download\"]")
                
                if sucesso:
                    print(f"    ✓ Documento baixado com sucesso!")
                else:
                    print(f"    ✗ Falha ao baixar o documento")
                
                # Pequena pausa entre downloads
                await page.wait_for_timeout(1000)

        except Exception as e:
            print(f"ERRO inesperado ao processar a classificação '{classificacao}': {e}")

    print(f"--- Busca de decretos concluída para {unidade}. ---")


async def baixar_documento_pdf(page, row_index, doc_name, save_dir, download_button_selector):
    """
    Baixa um documento PDF da nova aba que se abre ao clicar no botão de visualização.
    - Abre a aba do documento
    - Localiza o botão de download usando o seletor informado (inclusive dentro de iframes)
    - Clica e aguarda o download via Playwright
    - Faz fallback para tentativas alternativas se necessário
    """
    doc_name_sanitized = unidecode.unidecode(doc_name).replace(' ', '_').replace('/', '-')
    timestamp = datetime.now().strftime("%H%M%S")
    final_filename = f"{doc_name_sanitized[:100]}_{timestamp}.pdf"
    save_path = os.path.join(save_dir, final_filename)

    new_page = None
    try:
        print(f"    Abrindo documento: {doc_name}")

        # Abrir a nova aba clicando no botão de visualização
        async with page.context.expect_page() as new_page_info:
            viewer_button_xpath = XPATH_DOWNLOAD_DOC_BASE.format(row_index=row_index)
            await page.locator(viewer_button_xpath).click()

        new_page = await new_page_info.value
        print(f"    Nova aba aberta: {new_page.url}")

        # Aguardar a nova página carregar completamente
        await new_page.wait_for_load_state('networkidle', timeout=15000)
        await new_page.wait_for_timeout(500)

        # Configurar o tratamento de download
        async with new_page.expect_download() as download_info:
            print(f"    Procurando botão de download com o seletor: {download_button_selector}")

            # 1) Tentar no documento principal
            button = new_page.locator(download_button_selector)
            try:
                await button.wait_for(state='visible', timeout=5000)
                await button.click()
                print("    Clique no botão de download (main document) realizado.")
            except TimeoutError:
                print("    Botão não visível no documento principal. Tentando em iframes...")
                found = False
                # 2) Procurar em iframes
                for frame in new_page.frames:
                    try:
                        frame_button = frame.locator(download_button_selector)
                        await frame_button.wait_for(state='visible', timeout=2000)
                        await frame_button.click()
                        print(f"    Clique no botão de download realizado no iframe: {frame.url}")
                        found = True
                        break
                    except TimeoutError:
                        continue
                if not found:
                    print("    ❌ Botão de download não encontrado em nenhum contexto com o seletor informado.")
                    # Cancelar o expect_download antes do fallback
                    raise TimeoutError("Botão de download não encontrado")

            # Aguardar o download
            download = await download_info.value
            await download.save_as(save_path)
            print(f"    Arquivo salvo em: {save_path}")
            return True

    except Exception as e:
        print(f"    ERRO ao baixar documento com seletor '{download_button_selector}': {e}")
        # Fallbacks: tentar URL direta e interceptação de rede
        try:
            if new_page:
                if await interceptar_pdf_por_url(new_page, save_path):
                    return True
                if await interceptar_pdf_fallback(new_page, save_path):
                    return True
        except Exception as fe:
            print(f"    Fallbacks também falharam: {fe}")
        return False

    finally:
        if new_page and not new_page.is_closed():
            await new_page.close()
            print("    Nova aba fechada")

async def pesquisar_cidade(page, cidade, unidade_type="PREFEITURA", competencia_filter=None):
    cidade_formatada = unidecode.unidecode(cidade).upper()
    print(f"--- Iniciando pesquisa para: {cidade} ({unidade_type}) ---")
    
    unidade_nome_map = {
        "PREFEITURA": f"Prefeitura Municipal de {cidade}",
        "CAMARA": f"Câmara Municipal de {cidade}",
        "IPREVIB - IBICOARA": f"IPREVIB - {cidade}",
        "SAAE - ÉRICO CARDOSO": f"SAAE - {cidade}",
        "CAPREVAC - CARAÍBAS": f"CAPREVAC - {cidade}",
        "CISVITA - VITÓRIA DA CONQUISTA": f"CISVITA - {cidade}"
    }
    unidade_nome = unidade_nome_map.get(unidade_type, "N/A")

    try:
        await page.goto("https://e.tcm.ba.gov.br/epp/ConsultaPublica/listView.seam", wait_until="networkidle", timeout=10000)

        print(f"Selecionando município: {cidade}")
        await page.locator(XPATH_MUNICIPIO_LABEL).click()
        
        municipio_locator = page.locator(f'li[data-label="{cidade.upper()}"], li[data-label="{unidecode.unidecode(cidade).upper()}"]')
        try:
            await municipio_locator.first.click(timeout=2000)
        except TimeoutError:
            print(f"Não foi possível localizar o município '{cidade}' pelo data-label. Tentando por texto...")
            await page.locator(f'xpath=//li[contains(text(), "{cidade.upper()}")]').first.click()
        
        print(f"Município '{cidade}' selecionado.")
        await page.wait_for_timeout(1000)

        print(f"Selecionando unidade: {unidade_nome}")
        await page.locator(CSS_UNIDADE_TRIGGER).click(force=True, timeout=5000)
        
        # CORREÇÃO 2: Usar um seletor mais robusto para Câmara
        if unidade_type == "CAMARA":
            unidade_locator = page.locator('xpath=//li[contains(text(), "Câmara Municipal")]')
        else:
            unidade_locator = page.locator(f'li[data-label="{unidade_nome}"]')
        
        await unidade_locator.click(timeout=5000)
        print("Unidade selecionada com sucesso.")
        await page.wait_for_timeout(500)

        await page.locator(XPATH_PESQUISAR_BUTTON).click()
        
        # A espera explícita foi removida. A lógica de loop abaixo já lida com a tabela vazia.
        # Adicionando uma pequena espera para dar tempo da página reagir.
        print("Aguardando resposta da pesquisa...")
        await page.wait_for_timeout(3000)
        
        print(f"Buscando competência '{competencia_filter}' na tabela...")
        row_index = 0
        while True:
            try:
                competencia_locator = page.locator(XPATH_COMPETENCIA_BASE.format(row_index=row_index))
                await competencia_locator.wait_for(state='visible', timeout=2000)
                
                current_competencia_full = await competencia_locator.text_content() or ""
                current_competencia_valor = current_competencia_full.split(':', 1)[-1].strip()

                if competencia_filter == current_competencia_valor:
                    print(f"Competência '{competencia_filter}' encontrada na linha {row_index + 1}. Clicando para ver detalhes...")
                    
                    selecionar_button_xpath = XPATH_SELECIONAR_PRESTACAO_BASE.format(row_index=row_index)
                    await page.locator(selecionar_button_xpath).click()
                    
                    await page.wait_for_load_state('networkidle', timeout=15000)
                    print("Página de detalhes da competência carregada.")
                    return True, unidade_nome, competencia_filter

                row_index += 1
            except TimeoutError:
                if row_index == 0:
                    print("AVISO: Tabela de resultados vazia.")
                else:
                    print(f"Fim da tabela. Não foi encontrada a competência '{competencia_filter}'.")
                return False, None, None

    except (TimeoutError, TargetClosedError, Exception) as e:
        print(f"ERRO ao processar a cidade: {cidade}. Detalhes: {e}")
        return False, None, None

async def main():
    full_cidades_para_pesquisar = [
        "ABAÍRA", "BARRA DA ESTIVA", "BELO CAMPO", "BOA NOVA", "BOM JESUS DA SERRA",
        "CACULÉ", "CAETANOS", "CÂNDIDO SALES", "CARAÍBAS", "CONDEÚBA", "CORDEIROS",
        "ENCRUZILHADA", "ÉRICO CARDOSO", "IBICOARA", "ITORORÓ", "ITAPETINGA",
        "JACARACI", "LIVRAMENTO DE NOSSA SENHORA", "MAETINGA", "MORTUGABA", "PAU BRASIL", "PIRIPÁ",
        "PLANALTO", "RIO DO PIRES", "SEBASTIÃO LARANJEIRAS", "TANHAÇU", "URANDI", "VITÓRIA DA CONQUISTA"
    ]

    while True:
        if os.name == 'nt': os.system('cls')
        else: os.system('clear')
        print("AGBOT - DOWNLOAD DE PROCESSOS E-TCM")
        print("Criado por Ericsson Cardoso.")
        print("-" * 40)

        search_option = input("Deseja pesquisar todos os municípios (T) ou apenas alguns (A)? ").upper()
        current_cidades_to_search = full_cidades_para_pesquisar if search_option == "T" else [m.strip() for m in input("Digite os municípios, separados por vírgula: ").split(',') if m.strip()]

        if not current_cidades_to_search:
            print("Nenhuma cidade selecionada. Reiniciando...")
            await asyncio.sleep(2)
            continue

        print(f"Cidades a serem pesquisadas: {', '.join(current_cidades_to_search)}")
        print("-" * 40)
        
        competencia_filter = ""
        while not competencia_filter:
            competencia_filter = input("Digite a COMPETÊNCIA desejada (ex: 05/2025): ").strip()
        
        print("-" * 40)

        try:
            async with async_playwright() as p:
                browser = await p.chromium.launch(headless=False)
                context = await browser.new_context(accept_downloads=True)
                page = await context.new_page()

                for cidade in current_cidades_to_search:
                    cidade_upper = cidade.upper()
                    
                    unidades_a_pesquisar = []
                    if cidade_upper not in CIDADES_SKIP_PREF_CAMARA:
                        if cidade_upper != "PAU BRASIL":
                            unidades_a_pesquisar.append("PREFEITURA")
                        if cidade_upper in CIDADES_COM_CAMARA_MUNICIPAL:
                            unidades_a_pesquisar.append("CAMARA")
                    
                    if cidade_upper in SAAE_ERICO_CARDOSO: unidades_a_pesquisar.append("SAAE - ÉRICO CARDOSO")
                    if cidade_upper in IPREVIB_IBICOARA: unidades_a_pesquisar.append("IPREVIB - IBICOARA")
                    if cidade_upper in CAPREVAC_CARAIBAS: unidades_a_pesquisar.append("CAPREVAC - CARAÍBAS")
                    if cidade_upper in CISVITA_CONQUISTA: unidades_a_pesquisar.append("CISVITA - VITÓRIA DA CONQUISTA")

                    for unidade_key in unidades_a_pesquisar:
                        success, unidade_nome, competencia = await pesquisar_cidade(page, cidade, unidade_key, competencia_filter)
                        if success:
                            await pesquisar_e_baixar_decretos_melhorada(page, cidade, unidade_nome, competencia)
                            
                await browser.close()
                print("\nProcesso finalizado para todas as cidades selecionadas.")

        except (KeyboardInterrupt, asyncio.CancelledError):
            print("\n\nProcesso interrompido pelo usuário.")
        except Exception as e:
            print(f"ERRO INESPERADO NO NÍVEL SUPERIOR: {e}")
        
        search_again = input("Deseja realizar uma nova pesquisa (S/N)? ").upper()
        if search_again != "S":
            print("Encerrando o programa.")
            break



if __name__ == "__main__":
    asyncio.run(main())