from robocorp.tasks import task
from robocorp import browser
import time
from RPA.HTTP import HTTP
from RPA.Tables import Tables
from RPA.PDF import PDF
from RPA.Archive import Archive



@task
def nivel_medio_rpa():
    """Robô criado no intuito de obter certificado de nível médio em RoboCorp"""
    
    browser.configure(slowmo=1000)
    
    abre_pagina_robot()
    
    preenche_campos_pesquisa(leitura_excel())
    
    zip_arquivos()



def abre_pagina_robot():
    browser.goto(r"https://robotsparebinindustries.com/")
    
    http = HTTP()
    
    http.download(url=r"https://robotsparebinindustries.com/orders.csv", overwrite=True)
    
    page = browser.page()
    
    page.fill("#username", "maria")
    
    page.fill('#password', 'thoushallnotpass')
    
    time.sleep(5)
    
    page.click("button:text('Log in')")
    
    time.sleep(7)
    
    page.locator('//*[@id="root"]/header/div/ul/li[2]/a').click()
    
    time.sleep(2)
    
    page.locator('//*[@id="root"]/div/div[2]/div/div/div/div/div/button[1]').click()    
    
def leitura_excel():
    
    csv = Tables()
    
    planilha = csv.read_table_from_csv("orders.csv")
    
    return planilha
    
def preenche_campos_pesquisa(planilha):
    
    page = browser.page()
    
    placeholder_legs = 'Enter the part number for the legs'
    
    for i,linha in enumerate(planilha):
        corpo = linha['Body']
        
        ordem_ped = linha['Order number']
        
        page.locator('//*[@id="head"]').select_option(linha['Head'])
        
        if corpo == '1':
            page.locator('//*[@id="root"]/div/div[1]/div/div[1]/form/div[2]/div/div[1]').click(click_count=2)
            
        if corpo == '2':
            page.locator('//*[@id="root"]/div/div[1]/div/div[1]/form/div[2]/div/div[2]').click(click_count=2)
            
        if corpo == '3':
            page.get_by_label('D.A.V.E body').click(click_count=2)
            
        if corpo == '4':
            page.locator('//*[@id="root"]/div/div[1]/div/div[1]/form/div[2]/div/div[4]').click(click_count=2)
            
        if corpo == '5':
            page.locator('//*[@id="root"]/div/div[1]/div/div[1]/form/div[2]/div/div[5]').click(click_count=2)
            
        if corpo == '6':
            page.locator('//*[@id="root"]/div/div[1]/div/div[1]/form/div[2]/div/div[6]').click(click_count=2)
            
        
        page.get_by_placeholder('Enter the part number for the legs').click()
        
        if linha['Legs'] == '1':
            page.get_by_placeholder(placeholder_legs).press("ArrowUp")
        if linha['Legs'] == '2':
            page.get_by_placeholder(placeholder_legs).press("ArrowUp")
            page.get_by_placeholder(placeholder_legs).press("ArrowUp")
        if linha['Legs'] == '3':
            page.get_by_placeholder(placeholder_legs).press("ArrowUp")
            page.get_by_placeholder(placeholder_legs).press("ArrowUp")
            page.get_by_placeholder(placeholder_legs).press("ArrowUp")
        if linha['Legs'] == '4':
            page.get_by_placeholder(placeholder_legs).press("ArrowUp")
            page.get_by_placeholder(placeholder_legs).press("ArrowUp")
            page.get_by_placeholder(placeholder_legs).press("ArrowUp")
            page.get_by_placeholder(placeholder_legs).press("ArrowUp")
        if linha['Legs'] == '5':
            page.get_by_placeholder(placeholder_legs).press("ArrowUp")
            page.get_by_placeholder(placeholder_legs).press("ArrowUp")
            page.get_by_placeholder(placeholder_legs).press("ArrowUp")
            page.get_by_placeholder(placeholder_legs).press("ArrowUp")
            page.get_by_placeholder(placeholder_legs).press("ArrowUp")
        if linha['Legs'] == '6':
            page.get_by_placeholder(placeholder_legs).press("ArrowUp")
            page.get_by_placeholder(placeholder_legs).press("ArrowUp")
            page.get_by_placeholder(placeholder_legs).press("ArrowUp")
            page.get_by_placeholder(placeholder_legs).press("ArrowUp")
            page.get_by_placeholder(placeholder_legs).press("ArrowUp")
            page.get_by_placeholder(placeholder_legs).press("ArrowUp")
            
        page.locator('//*[@id="address"]').click()
        page.fill('//*[@id="address"]',linha['Address'])
        
        
        page.locator('//*[@id="preview"]').click()
        time.sleep(2)
        page.locator('//*[@id="order"]').click()
        time.sleep(5)    

        error = page.text_content('//*[@id="root"]/div/div[1]/div/div[1]/div')
        while ('Receipt' not in error):
            time.sleep(5)
            page.locator('//*[@id="order"]').click()
            error = page.text_content('//*[@id="root"]/div/div[1]/div/div[1]/div')
            
            
               
        
        arquivo2 = exportar_para_pdf(ordem_pedido=ordem_ped)
        
        arquivo_png = tira_screenshot(arquivo2[1])
        
        junta_screenshot_arquivo(pdf_file=arquivo2[0], screenshot=arquivo_png)
        
            
        page.locator('//*[@id="order-another"]').click()
        time.sleep(2)
        page.locator('//*[@id="root"]/div/div[2]/div/div/div/div/div/button[1]').click()   

    
def exportar_para_pdf(ordem_pedido: str):
    
    """_summary_
    
    Recebe uma váriavel que nomeará arquivo pdf gerado.

    Args:
        ordem_pedido (str): _description_
    """
    
    page = browser.page()
    
    pdf = PDF()
    
    resultado_pedido = page.locator('//*[@id="order-completion"]').inner_html()
    
    pdf.html_to_pdf(resultado_pedido,fr"output/pedido_{ordem_pedido}.pdf")
    
    arquivo = fr"C:\Users\andre.junior\Documents\Personal\Robocorp\my-rsb-robot\output\pedido_{ordem_pedido}.pdf"
    
    caminho_pasta = r"output/"    
    
    return arquivo ,caminho_pasta



def tira_screenshot(order_number: str):
    page = browser.page()
    
    page.screenshot(path=fr"{order_number}pedido.png")
    
    print1 = r'C:\Users\andre.junior\Documents\Personal\Robocorp\my-rsb-robot\output\pedido.png'
    
    return print1


def junta_screenshot_arquivo(screenshot: bytes, pdf_file: str):
    
    pdf = PDF()
    
    lista_arquivos = [
        pdf_file,
        f'{str(screenshot)}: align=center',
        f'{str(screenshot)}: x=0, y=0'
    ]
    
    pdf.add_files_to_pdf(
        files = lista_arquivos,
        target_document = pdf_file
    )
    
 
def zip_arquivos():
    
    zipagem = Archive()
    
    zipagem.archive_folder_with_zip(folder=r"output" , include='*.pdf',archive_name='Zip_Docs.zip', recursive=True)
    
    
    

    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    



    
    
    
    