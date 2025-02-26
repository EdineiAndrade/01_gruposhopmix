from playwright.sync_api import sync_playwright
import pandas as pd
import time
import locale

locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')

# Função para extrair dados de um produto na página de detalhes
def extract_product_data(page, url_product):
    try:
        page.goto(url_product)
        time.sleep(.1)
        # Extrair informações usando os seletores fornecidos
        produto = page.query_selector('#name').get_attribute("value")
            
        categoria = page.query_selector('#category').get_attribute("value").split(' >')[-1].strip()
        codigo = url_product.split("/")[-1]
        sku = page.query_selector('#sku').get_attribute("value")
        ean = page.query_selector('#ean').get_attribute("value")
        ncm = page.query_selector('#ncm').get_attribute("value")   

        preco = float(page.query_selector('#sale_value').get_attribute("value"))
        preco = locale.format_string("%.2f", preco, grouping=True)
        estoque = page.query_selector('#stock').get_attribute("value")
        peso = page.query_selector('#weight').get_attribute("value") 
        altura = int(float(page.query_selector('#height').get_attribute("value")) * 100)
        largura = int(float(page.query_selector('#width').get_attribute("value"))*100)
        comprimento = int(float(page.query_selector('#length').get_attribute("value"))*100)
        description = page.query_selector('#description').input_value().replace('\n','')

        
        imagens = list(map(lambda el: el.get_attribute("src"), page.query_selector_all('//*[@id="render_images"]/div/img')))
        lista_imagens = ", ".join(map(str, imagens))
        
        return {
            'Categoria': categoria,
            'ID': codigo,
            'SKU': sku,
            'EAN': ean,
            'NCM': ncm,
            'Preço promocional': 0,
            "Tipo": "simple",
            "GTIN UPC EAN ISBN": "",
            'Nome': produto,
            "Publicado": 1,
            "Em Destaque": 0,
            "Visibilidade no Catálogo": "visible",
            "Descrição Curta": "",
            "Descrição": description,
            "Data de Preço Promocional Começa em": "",
            "Data de Preço Promocional Termina em": "",
            "Status do Imposto": "taxable",
            "Classe de Imposto": "",
            "Em Estoque": 1,
            "Estoque": estoque,
            "Quantidade Baixa de Estoque": 3,
            "São Permitidas Encomendas": 0,
            "Vendido Individualmente": 0,
            "Peso (kg)": peso,
            "Comprimento (cm)": comprimento,
            "Largura (cm)": largura,
            "Altura (cm)": altura,
            "Permitir Avaliações de Clientes": 1,
            "Nota de Compra": "",
            "Preço Promocional": "",
            "Preço": preco,
            "Categorias": "",
            "Tags": "",
            "Classe de Entrega": "",
            "Imagens": lista_imagens,
            "Limite de Downloads": "",
            "Dias para Expirar o Download": "",
            "Ascendente": "",
            "Grupo de Produtos": "",
            "Upsells": "",
            "Venda Cruzada": "",
            "URL Externa": "",
            "Texto do Botão": "",
            "Posição": "",
            "Brands": "",
            "Nome do Atributo 1": "",
            "Valores do Atributo 1": "",
            "Visibilidade do Atributo 1": 0,
            "Atributo Global 1": 1,
            "Atributo Padrão 1": "",
            "Nome do Atributo 2": "",
            "Valores do Atributo 2": "",
            "Visibilidade do Atributo 2": 0,
            "Atributo Global 2": "",
            "Atributo Padrão 2": ""
        }
    except Exception as e:
        print(f"Erro ao extrair dados do produto: {e}")
        return None

# Função para salvar os dados em um arquivo Excel
def save_to_excel(data, filename='products.xlsx'):
    if isinstance(data, dict):
        df = pd.DataFrame.from_dict(data, orient="index").T
    else:    
        df = pd.DataFrame(data)
    df.to_excel(filename, index=False)
    print(f"Dados salvos em {filename}")

# Função principal para realizar o scraping
def scrape_categories(base_url):
    products_data = []

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)  # headless=False para ver o navegador em ação
        page = browser.new_page()
        page.goto(base_url)
        cont = 0
        products_data = []
        # Definição do caminho do arquivo
        caminho_arquivo = r"C:\sh\sh.txt"

        # Leitura do arquivo
        with open(caminho_arquivo, "r", encoding="utf-8") as arquivo:
            linha = arquivo.readline().strip()  # Lê a primeira linha e remove espaços extras
            login, senha = linha.split(",")  # Divide os valores separados por vírgula
        # Extrair links das categorias
        page.locator('(//*[@class="form-control bg-transparent "])[1]').fill(login)
        page.locator('(//*[@class="form-control bg-transparent "])[2]').fill(senha)
        page.locator('//*[@class="btn btn-primary"]').click()
        page.locator('(//*[@class="menu-text"])[6]').click()

        
        for n in range(1,129):
            page.goto(f"https://app.gruposhopmix.com.br/dashboard/catalog?page={n}")
            
            product_links = page.query_selector_all('//*[@class="card-body"]//h5/a')
            product_urls = list(map(lambda link: link.get_attribute('href'), product_links))

            for url_product in product_urls:                
                print(f"Processando categoria: {url_product}")
                #url_product = "https://app.gruposhopmix.com.br/dashboard/product/create/152"
                product_data = extract_product_data(page, url_product)
                products_data.append(product_data)
                cont = cont + 1
                if cont >= 20:
                    time.sleep(1)
                    save_to_excel(products_data, 'products.xlsx')
                    time.sleep(1)
                    cont = 0
        browser.close()

        return products_data

# Executar o scraping e salvar os dados
if __name__ == "__main__":
    base_url = 'https://app.gruposhopmix.com.br/login'  # Substitua pela URL base do e-commerce
    products = scrape_categories(base_url)
    save_to_excel(products, 'products.xlsx')