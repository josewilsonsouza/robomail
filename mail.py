import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from datetime import datetime, timedelta
import time
import traceback

# configs
URL_PLANILHA = "https://docs.google.com/spreadsheets/d/1VeInFHfXwIrWc06CSn6ode2IRsThyWITMrqFIV1cXoU/export?format=csv"
URL_WEBMAIL = "https://webmail.inmetro.gov.br/owa/auth/logon.aspx"
EMAIL_DESTINO = "destinatario@inmetro.gov.br"
EMAILS_COPIA = "copia1@inmetro.gov.br; copia2@inmetro.gov.br; copia3@inmetro.gov.br"

# info dados
DB_BOLSISTAS = {
    "Daniel Marques da Silva": {"linha": "BF-1 - Nilópolis", "func": "Bolsista", "up_uo": "Dmtic/Lainf", "tel": "(21) 97613-6422", "ponto": "Av. Getúlio de moura, 3512"},
    "Davi de Souza Feliz": {"linha": "ZO-3 - Bangu", "func": "Bolsista", "up_uo": "Dmtic/Lainf", "tel": "(21) 99738-6836", "ponto": "Avenida Brasil, passarela 27"},
    "Gabriel Trajano de Almeida": {"linha": "ZN-6 - Andaraí", "func": "Bolsista", "up_uo": "Dmtic/Lainf", "tel": "(21) 97515-6483", "ponto": "Rua Nicarágua, esquina com a Conde de Agrolongo"},
    "Isaque Silva dos Santos Faria": {"linha": "BF-4 - Duque de Caxias", "func": "Bolsista", "up_uo": "Dmtic/Lainf", "tel": "(21) 97609-5731", "ponto": "Rodovia Washington Luiz, passarela Vila Canaã"},
    "Ivan Roberto Meirelles": {"linha": "BF-3- Saracuruna", "func": "Bolsista", "up_uo": "Dmtic/Lainf", "tel": "(21) 97659-6662", "ponto": "Ponto de ônibus Entrada do Barro Branco"},
    "José Wilson Conceição de Souza": {"linha": "ZN-1 - Ilha do Governador", "func": "Bolsista", "up_uo": "Dmtic/Lainf", "tel": "(21) 98307-2973", "ponto": "Av. Brigadeiro Trompovski - Colégio zerohum / Peixaria"},
    "Karla Jacqueline Augusto Machado": {"linha": "ZO-2 - ZONA OESTE", "func": "Bolsista", "up_uo": "Dmtic/Lainf", "tel": "9777", "ponto": "Est. Cel. Pedro Corrêa - RIO 2"},
    "Luan Michel Soares Pereira": {"linha": "BF-4 - Duque de Caxias", "func": "Bolsista", "up_uo": "Dmtic/Lainf", "tel": "(21) 99386-3475", "ponto": "Passarela do bairro Jardim Olimpo"},
    "Malkai dos Santos Pereira Oliveira": {"linha": "ZN-3 - Tijuca", "func": "Bolsista", "up_uo": "Dmtic/Lainf", "tel": "(21) 99683-7128", "ponto": "Praça saens pena"},
    "Mina Lura Mathias Monteiro": {"linha": "BF-3-Saracuruna", "func": "Bolsista", "up_uo": "Dmtic/Lainf", "tel": "(21) 99055-7593", "ponto": "Avenida automóvel club, próximo ao 1611"},
    "Marcus Vinícius Pereira Moreira": {"linha": "ZO-3 - Bangu", "func": "Bolsista", "up_uo": "Dmtic/Lainf", "tel": "(21) 99416-0410", "ponto": "Estrada Intendente Magalhães"},
    "Paulo Eduardo Silva dos Santos": {"linha": "RS-3 - Petrópolis B", "func": "Bolsista", "up_uo": "Dmtic/Lainf", "tel": "(24) 99971-8074", "ponto": "Rua da Imperatriz, em frente ao Museu Imperial"},
    "Stephanie Sousa da Silva": {"linha": "ZN-2 - Abolição", "func": "Bolsista", "up_uo": "Dmtic/Lainf", "tel": "(21) 97672-2335", "ponto": "Avenida Brasil, próximo ao IBGE"},
}

def obter_datas_proxima_semana():
    hoje = datetime.now()
    dias_para_segunda = 7 - hoje.weekday()
    proxima_segunda = hoje + timedelta(days=dias_para_segunda)
    proxima_sexta = proxima_segunda + timedelta(days=4)
    return proxima_segunda.strftime("%d/%m/%Y"), proxima_sexta.strftime("%d/%m/%Y")

def gerar_corpo_email(url_csv, html=False):
    try:
        print("Lendo planilha do Google...")
        df = pd.read_csv(url_csv)
        df['Carimbo de data/hora'] = pd.to_datetime(df['Carimbo de data/hora'], dayfirst=True)

        # Coleta registros a partir de quinta 14:30 da semana anterior até agora
        # e envia solicitação para a PRÓXIMA semana
        hoje = datetime.now()

        # Calcula a quinta-feira anterior às 14:30
        # Se hoje é quinta após 14:30, a janela começa na quinta passada às 14:30
        # Se hoje é antes de quinta 14:30, a janela começa na quinta de 2 semanas atrás
        dias_desde_quinta = (hoje.weekday() - 3) % 7  # 3 = quinta-feira
        quinta_anterior = hoje - timedelta(days=dias_desde_quinta)
        quinta_anterior = quinta_anterior.replace(hour=14, minute=30, second=0, microsecond=0)

        # Se estamos antes de quinta 14:30, volta mais uma semana
        if hoje < quinta_anterior:
            quinta_anterior = quinta_anterior - timedelta(days=7)

        inicio_janela = quinta_anterior
        fim_janela = hoje  # Até o momento atual

        # Calcula a semana-alvo (próxima semana)
        data_ini, data_fim = obter_datas_proxima_semana()

        print("=" * 60)
        print("📅 JANELA DE COLETA (respostas do formulário):")
        print(f"   De: {inicio_janela.strftime('%d/%m/%Y %H:%M')} (quinta anterior)")
        print(f"   Até: {fim_janela.strftime('%d/%m/%Y %H:%M')} (agora)")
        print()
        print("🚌 SEMANA SOLICITADA (transporte):")
        print(f"   De: {data_ini} (próxima segunda)")
        print(f"   Até: {data_fim} (próxima sexta)")
        print("=" * 60)

        # Filtra registros da quinta anterior 14:30 até agora
        df_semana = df[(df['Carimbo de data/hora'] >= inicio_janela) &
                       (df['Carimbo de data/hora'] <= fim_janela)]

        if df_semana.empty:
            print()
            print("⚠️  AVISO: Nenhuma resposta encontrada na janela de coleta.")
            print("   Possíveis causas:")
            print("   - Nenhum bolsista preencheu o formulário neste período")
            print("   - O formulário ainda não recebeu respostas")
            print()
            return "Nenhuma resposta encontrada no período (desde quinta 14:30 anterior)."

        # Mostra total de registros antes da deduplicação
        print(f"\n📋 Total de registros encontrados: {len(df_semana)}")

        # Ordena por data/hora decrescente e mantém apenas o registro mais recente de cada bolsista
        df_semana = df_semana.sort_values('Carimbo de data/hora', ascending=False)
        df_recente = df_semana.drop_duplicates(subset=['Nome'], keep='first')

        # Mostra quantos registros foram removidos por duplicação
        duplicados = len(df_semana) - len(df_recente)
        if duplicados > 0:
            print(f"🔄 Registros duplicados removidos: {duplicados}")
            print(f"   (Bolsistas que preencheram o formulário mais de uma vez)")

        # Reordena por nome para ficar organizado
        df_recente = df_recente.sort_values('Nome')

        print(f"✅ Total de bolsistas únicos: {len(df_recente)}")
        print()

        if html:
            # Versão HTML colorida
            texto = f"""<html><body style="font-family: Arial, sans-serif; font-size: 11pt;">
<p>Prezados(as), boa tarde!</p>
<p>Solicito por gentileza a inclusão dos alunos bolsistas relacionados abaixo para utilização do ônibus na semana de <strong>{data_ini} a {data_fim}</strong>:</p>
<p><strong>Setor: DMTIC/LAINF</strong></p>
"""
            contador = 1
            for index, row in df_recente.iterrows():
                nome = row['Nome']
                dias_selecionados = row['Dias']
                dados = DB_BOLSISTAS.get(nome)

                if dados:
                    # Cor de fundo alternada para cada bolsista
                    bg_color = "#e3f2fd" if contador % 2 == 1 else "#fff3e0"

                    texto += f"""
<div style="background-color: {bg_color}; padding: 10px; margin: 10px 0; border-left: 4px solid #1976d2; border-radius: 4px;">
    <p style="margin: 5px 0;"><strong style="color: #1565c0;">{contador:02d} - LINHA:</strong> <span style="color: #d84315;">{dados['linha']}</span></p>
    <p style="margin: 5px 0;"><strong style="color: #1565c0;">Nome Completo:</strong> {nome}</p>
    <p style="margin: 5px 0;"><strong style="color: #1565c0;">Categoria funcional:</strong> {dados['func']}</p>
    <p style="margin: 5px 0;"><strong style="color: #1565c0;">UP/UO:</strong> {dados['up_uo']}</p>
    <p style="margin: 5px 0;"><strong style="color: #1565c0;">Ramal/Telefone:</strong> {dados['tel']}</p>
    <p style="margin: 5px 0;"><strong style="color: #1565c0;">Dias:</strong> <span style="color: #2e7d32;">{dias_selecionados}</span></p>
    <p style="margin: 5px 0;"><strong style="color: #1565c0;">PONTO DE EMBARQUE:</strong> {dados['ponto']}</p>
</div>
"""
                    contador += 1
                else:
                    print(f"AVISO: Nome '{nome}' não encontrado no banco de dados.")

            texto += """
<p style="margin-top: 20px;">Atenciosamente,<br>José Wilson C. Souza<br>DMTIC/LAINF</p>
</body></html>
"""
        else:
            # Versão texto simples (preview no terminal)
            texto = f"Prezados(as), boa tarde!\n\n"
            texto += f"Solicito por gentileza a inclusão dos alunos bolsistas relacionados abaixo para utilização do ônibus na semana de {data_ini} a {data_fim}:\n"
            texto += "Setor: DMTIC/LAINF\n\n"

            contador = 1
            for index, row in df_recente.iterrows():
                nome = row['Nome']
                dias_selecionados = row['Dias']
                dados = DB_BOLSISTAS.get(nome)

                if dados:
                    texto += f"{contador:02d} - LINHA: {dados['linha']}\n"
                    texto += f"Nome Completo: {nome}\n"
                    texto += f"Categoria funcional: {dados['func']}\n"
                    texto += f"UP/UO: {dados['up_uo']}\n"
                    texto += f"Ramal/Telefone: {dados['tel']}\n"
                    texto += f"Dias: {dias_selecionados}\n"
                    texto += f"PONTO DE EMBARQUE: {dados['ponto']}\n\n"
                    contador += 1
                else:
                    print(f"AVISO: Nome '{nome}' não encontrado no banco de dados.")

            texto += "Atenciosamente, \n José Wilson Conceição de Souza \n DMTIC/LAINF"

        return texto
    except Exception as e:
        return f"Erro ao processar planilha: {e}"

def enviar_email():
    usuario = input("Digite seu usuário: ")
    senha = input("Digite sua senha: ")

    # Gera versão texto para preview
    corpo_texto = gerar_corpo_email(URL_PLANILHA, html=False)
    print("\n--- PRÉVIA DO E-MAIL ---")
    print(corpo_texto)
    print("------------------------")

    # Gera versão HTML para enviar
    corpo_html = gerar_corpo_email(URL_PLANILHA, html=True)

    if "Erro" in corpo_texto or "Nenhuma resposta" in corpo_texto:
        print("Cancelando envio por erro nos dados.")
        return

    confirmacao = input("O texto está correto? (s/n): ")
    if confirmacao.lower() != 's':
        return

    print("\n--- INICIANDO ROBÔ 2.0 ---")
    options = webdriver.ChromeOptions()
    # Remove notificações e popups que podem atrapalhar
    options.add_argument("--disable-notifications")
    options.add_experimental_option("excludeSwitches", ["enable-logging"])

    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    wait = WebDriverWait(driver, 15)

    try:
        # 1. Carrega site
        print("Passo 1: Acessando site...")
        driver.get(URL_WEBMAIL)
        
        # 2. Login
        print("Passo 2: Preenchendo login...")
        wait.until(EC.presence_of_element_located((By.ID, "username"))).send_keys(usuario)
        driver.find_element(By.ID, "password").send_keys(senha)

        # 3. IMPORTANTE: Forçar versão PADRÃO (desmarcar Light) para suportar HTML
        print("Passo 3: Garantindo versão PADRÃO (para suportar HTML formatado)...")
        try:
            chk_light = driver.find_element(By.ID, "chkBsc")
            # Verifica se está marcado e desmarca se necessário
            if chk_light.is_selected():
                driver.execute_script("arguments[0].click();", chk_light)
                #print("   -> Checkbox Light desmarcado - forçando versão Padrão")
            else:
                print("   -> Versão Padrão selecionada")
        except Exception as e:
            print(f"   -> Não foi possível verificar checkbox: {e}")
            print("   -> Continuando com configuração padrão")

        time.sleep(1) 

        # 4. Entrar
        print("Passo 4: Entrando...")
        try:
            driver.find_element(By.CLASS_NAME, "signinbutton").click()
        except:
            driver.find_element(By.CSS_SELECTOR, "input[type='submit']").click()

        # 5. Detectar Versão e Clicar em Novo
        print("Passo 5: Verificando versão carregada...")

        tentativa_sucesso = False

        # TENTA VERSÃO PADRÃO PRIMEIRO (mais comum)
        try:
            # Aguarda o botão "Novo(a)" da versão Padrão estar presente
            btn = wait.until(EC.presence_of_element_located(
                (By.XPATH, "//span[@class='tbLh tbBefore tbAfter' and contains(text(), 'Novo')]")
            ))
            # Clica via JavaScript
            driver.execute_script("arguments[0].click();", btn)
            #print("   -> ✅ Versão PADRÃO detectada (suporta HTML formatado)")
            print("   -> Cliquei no botão 'Novo(a)'")
            tentativa_sucesso = True
            time.sleep(1)
        except:
            # Se não encontrou, tenta versão Light
            try:
                btn_novo = WebDriverWait(driver, 3).until(
                    EC.element_to_be_clickable((By.ID, "lnkNew"))
                )
                btn_novo.click()
                print("   -> ⚠️ Versão LIGHT carregou (HTML pode não funcionar)")
                tentativa_sucesso = True
            except:
                # Último recurso: tenta pelo texto
                try:
                    btn_novo = WebDriverWait(driver, 3).until(
                        EC.element_to_be_clickable((By.LINK_TEXT, "Nova"))
                    )
                    btn_novo.click()
                    print("   -> ⚠️ Versão LIGHT carregou (HTML pode não funcionar)")
                    tentativa_sucesso = True
                except:
                    pass

        # Se nenhum método funcionou, tenta atalho de teclado
        if not tentativa_sucesso:
            try:
                body = driver.find_element(By.TAG_NAME, "body")
                body.send_keys(Keys.CONTROL + 'n')
                print("   -> Enviei comando CTRL+N. Aguardando janela abrir...")
                tentativa_sucesso = True
                time.sleep(3)
            except:
                print("   -> Falha ao tentar atalhos.")

        if not tentativa_sucesso:
            print("   -> ATENÇÃO: Não consegui clicar em 'Novo' automaticamente.")
            time.sleep(10)

        # Gerencia janelas (caso abra pop-up na versão Padrão)
        print("   -> Verificando se abriu nova janela...")
        time.sleep(1)
        if len(driver.window_handles) > 1:
            print(f"   -> Nova janela detectada! Mudando para a janela de composição")
            driver.switch_to.window(driver.window_handles[-1])

        # 6. Preencher E-mail (Tenta seletores de ambas as versões)
        print("Passo 6: Escrevendo e-mail...")

        time.sleep(.5)  # Aguarda carregar formulário completamente

        # Campo "Para"
        campo_para_preenchido = False
        try:
            # Versão Padrão (mais comum) - div contenteditable com id="divTo"
            campo_para = wait.until(EC.presence_of_element_located((By.ID, "divTo")))
            campo_para.send_keys(EMAIL_DESTINO)
            campo_para_preenchido = True
            print("   -> ✅ Campo 'Para' preenchido (Padrão)")
        except:
            try:
                # Tenta versão Light - procura o campo txtto (minúsculo!)
                campo_para = driver.find_element(By.NAME, "txtto")
                campo_para.send_keys(EMAIL_DESTINO)
                campo_para_preenchido = True
                print("   -> ✅ Campo 'Para' preenchido (Light)")
            except Exception as e:
                print(f"   -> AVISO: Campo 'Para' não encontrado: {e}")

        # Campo "Cc"
        try:
            # Tenta Standard version primeiro (mais comum)
            campo_cc = wait.until(EC.presence_of_element_located((By.ID, "divCc")))
            campo_cc.send_keys(EMAILS_COPIA)
            print("   -> ✅ Campo 'Cc' preenchido (Standard version)")
        except:
            try:
                # Tenta Light version
                campo_cc = driver.find_element(By.NAME, "txtcc")
                campo_cc.send_keys(EMAILS_COPIA)
                print("   -> ✅ Campo 'Cc' preenchido (Light version)")
            except:
                print("   -> AVISO: Campo 'Cc' não encontrado")

        # campo "Assunto"
        assunto_preenchido = False
        try:
            # Tenta Standard version primeiro (mais comum)
            campo_assunto = wait.until(EC.presence_of_element_located((By.ID, "txtSubj")))
            campo_assunto.send_keys("Solicitação de utilização do transporte Inmetro")
            assunto_preenchido = True
            print("   -> ✅ Campo 'Assunto' preenchido (Standard version)")
        except:
            try:
                # Tenta Light version
                campo_assunto = driver.find_element(By.NAME, "txtsbj")  # É "txtsbj" não "txtsubj"!
                campo_assunto.send_keys("Solicitação de utilização do transporte Inmetro")
                assunto_preenchido = True
                print("   -> ✅ Campo 'Assunto' preenchido (Light version)")
            except:
                print("   -> AVISO: Campo 'Assunto' não encontrado.")

        # corpo do email (HTML formatado)
        corpo_preenchido = False

        # ESTRATÉGIA: Abrir HTML renderizado no navegador, copiar, colar
        print("   -> Preparando HTML formatado...")

        import os
        html_temp_path = os.path.join(os.path.dirname(__file__), "temp_email.html")

        # Salva o HTML em arquivo temporário
        with open(html_temp_path, "w", encoding="utf-8") as f:
            f.write(corpo_html)

        # Guarda TODAS as janelas abertas ANTES de abrir o HTML
        janelas_antes = driver.window_handles
        janela_email_popup = driver.current_window_handle  # Popup do email
        janela_principal = janelas_antes[0]  # Primeira janela (principal)
        print(f"   -> Janelas abertas antes: {len(janelas_antes)}")
        print(f"   -> Janela principal: {janela_principal[:8]}...")
        print(f"   -> Janela email popup: {janela_email_popup[:8]}...")

        # VOLTA para a janela PRINCIPAL para abrir o HTML (popups bloqueiam window.open)
        driver.switch_to.window(janela_principal)
        print(f"   -> Mudou para janela principal")
        time.sleep(0.5)

        # Converte o caminho do arquivo para URL (multiplataforma)
        import platform
        if platform.system() == "Windows":
            file_path_url = f"file:///{html_temp_path.replace(chr(92), '/')}"
        else:
            file_path_url = f"file://{html_temp_path}"
        print(f"   -> Arquivo HTML: {file_path_url}")

        # alternativa abrir nova aba vazia primeiro, depois navegar
        try:
            # Abre nova aba vazia
            driver.execute_script("window.open('about:blank', '_blank');")
            time.sleep(1)

            # Verifica se abriu
            janelas_temp = driver.window_handles
            print(f"   -> Após abrir aba vazia: {len(janelas_temp)} janelas")

            if len(janelas_temp) > len(janelas_antes):
                # Muda para a nova aba
                nova_aba = janelas_temp[-1]
                driver.switch_to.window(nova_aba)
                print(f"   -> Mudou para nova aba: {nova_aba[:8]}...")

                # Agora navega para o arquivo HTML
                driver.get(file_path_url)
                print(f"   -> Navegou para arquivo HTML")
                time.sleep(1)
            else:
                print(f"   -> ERRO: Nova aba não abriu!")
        except Exception as e:
            print(f"   -> ERRO ao abrir aba: {e}")

        # Identifica qual é a janela HTML (a nova que foi aberta)
        janelas_depois = driver.window_handles
        print(f"   -> Janelas abertas depois: {len(janelas_depois)}")
        print(f"   -> Lista de janelas: {[w[:8] + '...' for w in janelas_depois]}")

        janela_html = None
        for w in janelas_depois:
            if w not in janelas_antes:
                janela_html = w
                break

        if janela_html:
            # Muda para a janela HTML
            driver.switch_to.window(janela_html)
            time.sleep(0.5)

            # Seleciona todo o conteúdo renderizado e copia
            driver.find_element(By.TAG_NAME, "body").send_keys(Keys.CONTROL + 'a')
            time.sleep(0.3)
            driver.find_element(By.TAG_NAME, "body").send_keys(Keys.CONTROL + 'c')
            time.sleep(0.3)
            print("   -> HTML renderizado copiado da área de transferência")

            # Fecha só a aba HTML
            driver.close()
            print(f"   -> Aba HTML fechada")

            # Volta para a janela do email POPUP
            driver.switch_to.window(janela_email_popup)
            print(f"   -> Voltou para janela do email popup")
            time.sleep(0.5)
        else:
            print("   -> AVISO: Não conseguiu abrir janela HTML")
            # Se não conseguiu, volta para o popup do email
            driver.switch_to.window(janela_email_popup)

        # Agora cola o HTML formatado no campo do corpo
        try:
            # Tenta Standard version: iframe com id="ifBdy"
            iframe_corpo = driver.find_element(By.ID, "ifBdy")
            driver.switch_to.frame(iframe_corpo)
            time.sleep(0.5)

            # Dentro do iframe, procura o body e cola
            body_iframe = driver.find_element(By.TAG_NAME, "body")
            body_iframe.click()
            time.sleep(0.3)

            # Cola o conteúdo HTML (Ctrl+V)
            body_iframe.send_keys(Keys.CONTROL + 'v')
            time.sleep(0.5)

            # Volta para o contexto principal
            driver.switch_to.default_content()

            corpo_preenchido = True
            print("   -> ✅ Corpo do email preenchido com HTML formatado (iframe Standard version)")
        except:
            try:
                # Tenta Light version: textarea com id="txtBdy"
                campo_corpo = driver.find_element(By.ID, "txtBdy")
                campo_corpo.click()
                time.sleep(0.5)
                campo_corpo.clear()

                # Cola o conteúdo HTML (Ctrl+V)
                campo_corpo.send_keys(Keys.CONTROL + 'v')

                corpo_preenchido = True
                print("   -> ✅ Corpo do email preenchido com HTML formatado (textarea)")
            except:
                try:
                    # Tenta Light version com name="txtbdy" (minúsculo)
                    campo_corpo = driver.find_element(By.NAME, "txtbdy")
                    campo_corpo.click()
                    time.sleep(0.5)
                    campo_corpo.clear()

                    # Cola o conteúdo HTML (Ctrl+V)
                    campo_corpo.send_keys(Keys.CONTROL + 'v')

                    corpo_preenchido = True
                    print("   -> ✅ Corpo do email preenchido com HTML formatado (Light version)")
                except:
                    # Tenta contenteditable div
                    try:
                        campo_corpo = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "div[contenteditable='true']")))
                        campo_corpo.click()
                        time.sleep(0.5)

                        # Cola o conteúdo HTML (Ctrl+V)
                        campo_corpo.send_keys(Keys.CONTROL + 'v')

                        corpo_preenchido = True
                        print("   -> ✅ Corpo do email preenchido com HTML formatado (contenteditable)")
                    except Exception as e:
                        # Tenta seletores alternativos (editores ricos)
                        selectors_corpo = [
                            (By.CSS_SELECTOR, "div[aria-label*='Corpo da mensagem']"),
                            (By.CSS_SELECTOR, "div[role='textbox']"),
                            (By.XPATH, "//div[@contenteditable='true']"),
                            (By.CSS_SELECTOR, "textarea[aria-label*='Mensagem']"),
                        ]

                        for selector_type, selector_value in selectors_corpo:
                            try:
                                campo = driver.find_element(selector_type, selector_value)
                                campo.click()
                                time.sleep(0.5)

                                # Cola o conteúdo HTML
                                campo.send_keys(Keys.CONTROL + 'v')

                                print(f"   -> ✅ Corpo do email preenchido com HTML formatado (fallback)")
                                corpo_preenchido = True
                                break
                            except:
                                continue

        # Remove o arquivo temporário
        try:
            os.remove(html_temp_path)
        except:
            pass

        if not corpo_preenchido:
            print("   -> AVISO: Campo corpo não encontrado.")
            print("   -> COPIE o texto do email acima e COLE no navegador.")

        # Tira screenshot para debug se algum campo não foi preenchido
        if not (campo_para_preenchido and assunto_preenchido and corpo_preenchido):
            try:
                screenshot_path = "debug_campos_vazios.png"
                driver.save_screenshot(screenshot_path)
                print(f"\n   -> Screenshot salvo em: {screenshot_path}")
            except:
                pass

        # 7. ENVIAR EMAIL
        if campo_para_preenchido and assunto_preenchido and corpo_preenchido:
            print("\n--- PASSO 7: ENVIANDO EMAIL ---")
            confirmacao_envio = input("Deseja ENVIAR o email agora? (s/n): ")

            if confirmacao_envio.lower() == 's':
                try:
                    # Procura o botão de enviar em diferentes formatos
                    print("   -> Procurando botão 'Enviar'...")

                    botao_enviado = False
                    selectors_enviar = [
                        (By.XPATH, "//span[@class='tbLh tbBefore tbAfter' and contains(text(), 'Enviar')]"),
                        (By.XPATH, "//span[contains(text(), 'Enviar')]"),
                        (By.ID, "lnkHdrsend"),  # Link do OWA com id específico
                        (By.XPATH, "//a[@id='lnkHdrsend']"),
                        (By.XPATH, "//a[contains(@onclick, 'send')]"),
                        (By.XPATH, "//a[contains(text(), 'Enviar')]"),
                        (By.LINK_TEXT, "Enviar"),
                        (By.PARTIAL_LINK_TEXT, "Enviar"),
                        (By.NAME, "btnS"),  # Fallback para outras versões
                        (By.ID, "btnS"),
                        (By.XPATH, "//input[@value='Enviar']"),
                        (By.XPATH, "//button[contains(text(), 'Enviar')]"),
                    ]

                    for selector_type, selector_value in selectors_enviar:
                        try:
                            btn_enviar = driver.find_element(selector_type, selector_value)
                            driver.execute_script("arguments[0].click();", btn_enviar)
                            print(f"   -> ✅ Email ENVIADO com sucesso usando {selector_type}!")
                            botao_enviado = True
                            time.sleep(2)  # Aguarda confirmação
                            break
                        except:
                            continue

                    if not botao_enviado:
                        print("   -> ⚠️ AVISO: Botão 'Enviar' não encontrado automaticamente.")
                        print("   -> Por favor, CLIQUE MANUALMENTE no botão 'Enviar' no navegador.")
                        input("Pressione ENTER após enviar o email...")

                except Exception as e:
                    print(f"   -> Erro ao tentar enviar: {e}")
                    print("   -> Por favor, clique manualmente no botão 'Enviar'.")
                    input("Pressione ENTER após enviar o email...")
            else:
                print("   -> Envio cancelado pelo usuário.")

        print("\n------------------------------------------------")
        print("PROCESSO FINALIZADO.")
        if campo_para_preenchido and assunto_preenchido and corpo_preenchido:
            print("✅ Email preenchido e pronto!")
        else:
            print("⚠️ Alguns campos não foram preenchidos. Verifique o navegador.")
        print("------------------------------------------------")
        input("Pressione ENTER para fechar o robô...")
        
    except Exception as e:
        print(f"\nERRO: {e}")
        traceback.print_exc()
    finally:
        driver.quit()

if __name__ == "__main__":
    enviar_email()