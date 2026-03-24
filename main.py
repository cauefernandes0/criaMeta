from playwright.sync_api import sync_playwright
import keyring 
import pandas as pd
from playwright.sync_api import TimeoutError as PlaywrightTimeoutError
import re

'''
PT
Automação em Python realizada utilizando Pandas, Playwright e Keyring (para as senhas e usuário do ambiente).
Essa automação é utilizada para realizar a adição de metas dos consultores de vendas da empresa (via planilha do excel) para o CRM (Microsoft Dynamics), 
utilizei playwright pois esse sistema da Microsoft as esperas de carregamento de telas,salvamento de registros podem variar muito de registro para registro, nisso é 
essencial utilizar uma lib que já possua uma espera nativa e mais dinâmica para os elementos no navegador.

Autor: Cauê Fernandes

EN
Python automation implemented using Pandas, Playwright and Keyring (for environment's user and passwords credentials).
This automation is used to add sales consultants' targets (via an Excel spreadsheet) to the CRM (Microsoft Dynamics), I've used Playwright because,in this Microsoft 
system, screen loadings and record sacing can vary significantly from record to record, therefore, it is essential to use a lib that already has a
native and more dynamic wait for browser elements

Author: Caue Fernandes
'''
with sync_playwright() as pw:
    def esta_vazio_no_excel(campo):
        if campo <= 0 or pd.isna(campo): #isna verifica se há 'NaN/null'
            print("Campo vazio no excel!")
            return True
        return False
    
    #Subfunção da função 'esta_vazio', verifica se o consultor e o mes específico existem no Dynamics
    def existe_meta(termo):      
        try:
            campo_busca = pagina.get_by_role("link", name=termo)
            campo_busca.first.wait_for(state='visible',timeout=5000)
            return True
        except PlaywrightTimeoutError:
            return False
    
    #Verifica se no Dynamics as metas já foram cadastradas
    def esta_vazio(consultor, mes):
        termo_busca = f'{consultor} {mes} 2026'
        # Campo de busca
        pagina.get_by_role("searchbox", name="Meta Filtrar por palavra-chave").click()
        campo_busca = pagina.get_by_role(
        "searchbox", 
        name=re.compile("Meta Filtrar|Aplicar começa com", re.IGNORECASE)
    )
        campo_busca.fill(termo_busca, force=True)
        campo_busca.press("Enter")


        if existe_meta(f'{consultor} {mes}'):
            print(f"🛑 Meta já existente: {termo_busca}")
            status = False
        else:
            print(f"Nenhuma meta encontrada: {termo_busca}")
            status = True
        try:
            pagina.get_by_role("button", name="Limpar pesquisa").click()
        except Exception as e:
            print("Erro ao limpar filtro:", e)

        return status
    
    #Função para criar a meta dentro do Dynamics
    def criar_meta(meta):
        print(meta)
        metaAtt = str(meta)
        pagina.get_by_role("menuitem", name="Criar", exact=True).click()

        pagina.get_by_role("textbox", name="Nome").fill(f"{celula} {mes} 2026")

        pagina.get_by_role("combobox", name="Métrica da Meta, Pesquisa").fill(metrica)

        pagina.keyboard.press("Enter")
        pagina.get_by_role("treeitem", name="Delta, Valor").click()

        pagina.get_by_role("combobox", name="Proprietário da Meta, Pesquisa").fill(celula)

        pagina.keyboard.press("Enter")
        opcao_proprietario = pagina.get_by_label("Painel suspenso").get_by_text(celula, exact=True)
        opcao_proprietario.first.wait_for(state="visible", timeout=3000)
        opcao_proprietario.click()

        pagina.get_by_role("button").get_by_text(f"Excluir {userEmail}").click()
        pagina.get_by_role("combobox", name="Gerente, Pesquisa").fill(gerente)
        pagina.keyboard.press("Enter")

        item_gerente = pagina.get_by_role("treeitem").get_by_text(gerente)
        item_gerente.wait_for(state="visible", timeout=5000)
        item_gerente.click()
        pagina.get_by_role("tab", name="Período").click()

        pagina.get_by_role("combobox", name="Período Fiscal").click()
        pagina.get_by_role("option", name=mes).click()

        pagina.get_by_role("tab", name="Destinos").click()

        campo_valor =pagina.get_by_role("textbox", name="Destino (Money)")
        campo_valor.clear()
        campo_valor.fill(metaAtt)

      

        pagina.get_by_role("menuitem", name="Salvar e Fechar").click()
    
    navegador = pw.chromium.launch(headless=False)
    dyn = '' #link do seu ambiente
    

    doc = pd.read_excel(".xlsx") #Insira o documento aqui
    metrica = 'Delta'
    gerente = '' #Insira o nome do usuário do gerente de vendas, para vincular a meta

    pswd = keyring.get_password("DynamicsUser",'seuUsuario') #antes disso faça o seguinte comando no terminal comando keyring.set_password("DynamicsUser","seuUsuario","suaSenha")
    userEmail = keyring.get_credential("DynamicsUser",'seuUsuario') #pegar o email do seu usuário pelo keyring
    print(doc.head(40))
    pagina = navegador.new_page()
    pagina.goto(dyn)


    #Login
    pagina.get_by_role("textbox", name="Insira o seu email, telefone").fill(userEmail)
    pagina.keyboard.press("Enter")

    pagina.get_by_role("textbox").get_by_text('Insira a senha para').fill(pswd)
    pagina.keyboard.press("Enter")

    pagina.get_by_role("button", name="Sim").click()
    

    for consultores in range(37,40): #Insira o range dos consultores nessa linha
        for i in range(1,13): #Insira o range dos meses aqui
            celula = doc.iloc[consultores,0]
            print("O consultor é",celula)
            mes = doc.iloc[23,i]
            print("O mês é",mes)
            meta = doc.iloc[consultores,i]  
            print("O valor da meta é", meta)
            # campo_vazio = esta_vazio(meta)
            excel_vazio =esta_vazio_no_excel(meta) #Caso a meta de um mês específico não exista, ele não irá adicionar no Dynamics
            meta_vazio = esta_vazio(celula,mes)
            if not excel_vazio and meta_vazio:
                criar_meta(meta)
            else:
                continue
            
    navegador.close()