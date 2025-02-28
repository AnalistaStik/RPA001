import pyautogui
import subprocess
import time
import pyperclip 
import os
import logging
import pygetwindow as gw
from datetime import datetime, timedelta
import customtkinter as ctk
import tkinter as tk
import os
import subprocess
from tkinter import filedialog, messagebox, simpledialog
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import xml.etree.ElementTree as ET
import fitz
from PIL import Image
import re
import time
import threading
import pytesseract
from pytesseract import image_to_string 
import numpy as np
from PIL import ImageEnhance, ImageFilter

def iniciar_robo():

    """# Configurações do logging 
    base_dir = os.getcwd()  # Obtém o diretório de trabalho atual
    log_dir = os.path.join(base_dir, 'logs')  # Diretório de logs relativo ao diretório de execução

    # Garantir que o diretório de logs exista
    if not os.path.exists(log_dir):
        os.makedirs(log_dir)

    # Configurações do logging 
    data_hoje = datetime.now().strftime("%d-%m-%Y")

    logging.basicConfig(
        filename=os.path.join(log_dir, f"log_{data_hoje}.txt"),
        level=logging.INFO,
        format="%(asctime)s - %(levelname)s - %(message)s",
    )"""

    senha_topmanager = "123456"
    tempo_espera_max = 30

    # Caminho
    diretorio_pdf = r"C://1 - XML e PDF"

    # Tempos de espera
    time_busca = 0.2
    time_login = 20
    time_short = 1
    time_medium = 3
    time_maximizar = 10
    time_long = 8

    # Posições dos elementos na tela       

    coordenadas = {
        "Consulta": (218, 32),
        "Doc. Eletronico": (252, 56),
        "Doc_Eletronico": (494, 51),
        "Filtro": (210, 77),
        "periodo": (279, 91),
        "Empresa": (452, 90),    
        "N. Pedido Base": (631, 129),
        "Pos. Acrobat": (104, 41),
        "Fechar Janela": (1894, 15),
        "Fechar_janela_navegador": (1897, 17)
    }

    logging.info("Programa Iniciado")

    #Função que preenche um campo de texto
    def preencher_campo(texto):
        pyautogui.press("backspace")
        pyautogui.write(texto)
        
    def maximizar_janela(titulo_parcial):
    #Função para maximizar uma janela específica ou clicar no botão de maximizar
        try:
            time.sleep(time_maximizar)
            janelas = gw.getWindowsWithTitle(titulo_parcial)
            if janelas:
                janela = janelas[0]
                if not janela.isMaximized:
                    janela.maximize()
                logging.info(f"Janela '{titulo_parcial}' maximizada.")
            time.sleep(time_short)
        except Exception as e:
            logging.error(f"Erro ao maximizar a janela {titulo_parcial}: {e}")
            
    #Função que baixa o XML
    def baixar_XML():
        try:
            
            time.sleep(time_short)
            pyautogui.press("F8")
            time.sleep(time_short)

            pyautogui.moveTo(845, 603)
            pyautogui.click()
            time.sleep(time_short)
            
            pyautogui.moveTo(903, 460)
            pyautogui.click()
            time.sleep(time_short)

            pyautogui.press("enter")
            time.sleep(time_short)
            logging.info("XML baixado com sucesso.")

        except Exception as e:
            logging.error(f"Erro ao baixar XML: {e}")
            
    #Função que baixa o PDF
    def baixar_pedido(pedido_x, pedido_y):
        try:
            time.sleep(time_short)
            pyautogui.click(button="right") 
            time.sleep(time_short)

            for _ in range(11):
                pyautogui.press("down")
                time.sleep(0.2)  

            pyautogui.press("enter")  
            time.sleep(12)

            pyautogui.moveTo(coordenadas['Pos. Acrobat'])
            pyautogui.click()
            time.sleep(time_long)
            
            pyautogui.hotkey("win", "up")
            time.sleep(time_medium)
            
            pyautogui.hotkey("ctrl", "shift", "s")
            time.sleep(time_medium)

            # Salvar PDF
            pyautogui.moveTo(1297, 797)  
            pyautogui.click()
            time.sleep(time_medium)
            pyautogui.moveTo(571, 48)
            pyautogui.click()
            time.sleep(time_medium)
            preencher_campo("C:/1 - XML e PDF")
            time.sleep(time_medium)
            
            pyautogui.press("enter")
            time.sleep(time_short)
            pyautogui.moveTo(781, 508)
            pyautogui.click()
            time.sleep(time_short)
            
            pyautogui.moveTo(coordenadas['Fechar_janela_navegador'])
            pyautogui.click()
            time.sleep(time_long)
            pyautogui.moveTo(coordenadas['Fechar Janela'])
            pyautogui.click()
            time.sleep(time_long)
            logging.info("PDF baixado com sucesso.")

        except Exception as e:
            logging.error(f"Erro ao baixar PDF: {e}")
        
    # Entrar e logar no TopManager
    try:
        if not gw.getWindowsWithTitle("TopManager"):
            pyautogui.hotkey("ctrl", "shift", "t")
            logging.info("Atalho enviado para abrir o TopManager.")

            # Espera até o TopManager abrir ou o tempo máximo ser atingido
            tempo_inicial = time.time()
            while not gw.getWindowsWithTitle("TopManager"):
                if time.time() - tempo_inicial > tempo_espera_max:
                    logging.error("O TopManager não abriu dentro do tempo esperado.")
                    raise Exception("Tempo limite excedido ao abrir o TopManager.")
                time.sleep(1)
        logging.info("TopManager detectado. Tentando login.")

        # Login no sistema
        time.sleep(5)
        pyautogui.moveTo(398, 346)
        time.sleep(time_medium)
        pyautogui.click()
        pyautogui.write(senha_topmanager)
        pyautogui.press("enter")
        time.sleep(time_login)

        if gw.getWindowsWithTitle("TopManager (Licenciado para Stik)"):
            logging.info("Login no TopManager realizado com sucesso.")
            maximizar_janela("TopManager")
            
        else:
            raise Exception("Falha no login no TopManager.")
        
        # Acesso à aba de documentos eletrônicos
        try:
            pyautogui.moveTo(coordenadas['Consulta'])
            pyautogui.click()
            time.sleep(time_short)

            pyautogui.moveTo(coordenadas['Doc. Eletronico'])
            pyautogui.click()
            time.sleep(time_short)

            pyautogui.moveTo(coordenadas['Doc_Eletronico'])
            pyautogui.click()
            time.sleep(time_short)

            
            pyautogui.moveTo(679, 86)
            pyautogui.click()
            time.sleep(1)

            pyautogui.moveTo(coordenadas['Filtro'])
            pyautogui.click()
            time.sleep(time_short)

            pyautogui.moveTo(coordenadas['periodo'])
            pyautogui.click(button='left')
            time.sleep(time_short)
            pyautogui.write('h')
            time.sleep(time_short)
            pyautogui.press("tab")
            time.sleep(time_short)
            pyautogui.press("tab")
            time.sleep(time_short)
            preencher_campo('2')
            time.sleep(time_short)
            pyautogui.press('tab')
            time.sleep(time_short)
            pyautogui.press('tab')
            time.sleep(time_short)
            pyautogui.press('tab')
            time.sleep(time_short)

            pyautogui.press('enter')
            time.sleep(time_medium)
            preencher_campo('4109')
            time.sleep(time_short)
            pyautogui.press('enter')
            time.sleep(time_short)

            pyautogui.press('F5')
            time.sleep(time_long)
            logging.info("Acesso à aba de documentos eletrônicos realizado com sucesso.")
            
        except Exception as e:
            logging.error(f"Erro ao acessar aba de documentos eletrônicos: {e}")
            raise Exception("Falha ao acessar a aba de documentos eletrônicos.")

        # Configurações e verificação dos pedidos
        pedido_x, pedido_y = coordenadas["N. Pedido Base"]
        incremento_y = 18  # Distância entre as linhas dos pedidos
        deslocamento_x = 100  # Ajuste lateral para clicar na opção "Abrir Pedido"
        deslocamento_y = 10   # Pequeno ajuste vertical se necessário
        numero_pedidos = 0
        pedido_anterior = None  # Variável para armazenar o pedido anterior

        while True:
            try:
                pyautogui.moveTo(pedido_x, pedido_y)
                pyautogui.click()
                time.sleep(time_short)

                # Copiar o conteúdo
                pyautogui.hotkey("ctrl", "c")
                time.sleep(time_short)
                pedido_atual = pyperclip.paste().strip()

                # Verificar se a célula está vazia (último pedido atingido)
                if not pedido_atual:
                    print("Nenhum pedido encontrado ou fim da lista.")
                    break

                # Verificar se o pedido atual é o mesmo que o anterior
                if pedido_atual == pedido_anterior:
                    print("Pedido repetido, encerrando a busca")
                    break

                logging.info(f"Baixando pedido: {pedido_atual}")
                numero_pedidos += 1

                baixar_XML()
                baixar_pedido(pedido_x, pedido_y)

                pedido_anterior = pedido_atual
                pedido_y += incremento_y

            except Exception as e:
                logging.error(f"Erro no processo de download: {e}")

        logging.info(f"Total de pedidos baixados: {numero_pedidos}")

        try:
            time.sleep(time_medium)
            pyautogui.moveTo(1895, 10)
            pyautogui.click()
            logging.info("TopManager fechado com sucesso.")
        except Exception as e:
            logging.error(f"Erro ao fechar o TopManager: {e}")

    except Exception as e:
        logging.error(f"Erro no processo: {e}")


def iniciar_GUI():
    pytesseract.pytesseract.tesseract_cmd = r"C://Users//fnojosa//AppData//Local//Programs//Tesseract-OCR//tesseract.exe"

    pasta_pedidos = "C://1 - XML e PDF"

    selecionados = {"xml": None, "pdf": None}

    pedidos_enviados = set()

    def carregar_nfs_enviadas():
        global nfs_enviadas
        nfs_enviadas = {}
        try:
            if os.path.exists("nfs_enviadas.txt"):
                with open("nfs_enviadas.txt", "r") as f:
                    for line in f:
                        numero_nf, destinatario = line.strip().split(" - ")
                        if numero_nf not in nfs_enviadas:
                            nfs_enviadas[numero_nf] = []
                        nfs_enviadas[numero_nf].append(destinatario)
        except Exception as e:
            print(f"Erro ao carregar NF's enviadas: {e}")

    # Função que envia o Email
    def enviar_email_com_anexo():
        if not selecionados["xml"] or not selecionados["pdf"]:
            messagebox.showwarning("Erro", "Selecione um arquivo XML e um PDF antes de enviar.")
            return

        if not os.path.exists(selecionados["xml"]) or not os.path.exists(selecionados["pdf"]):
            messagebox.showwarning("Erro", "Arquivo XML ou PDF não encontrado.")
            return

        if selecionados["xml"] in pedidos_enviados or selecionados["pdf"] in pedidos_enviados:
            messagebox.showwarning("Aviso", "Este pedido já foi enviado.")
            return

        numero_nf_xml = obter_valor_nf(selecionados["xml"])
        numero_nf_pdf = obter_valor_nf_pdf(selecionados["pdf"])

        if not numero_nf_xml or not numero_nf_pdf or numero_nf_xml != numero_nf_pdf:
            messagebox.showwarning("Erro", "Os números da NF no XML e no PDF não correspondem. O envio foi bloqueado.")
            return

        corpo_email = f"""Prezado(a),

    Segue em anexo a nota em XML e PDF.

    Atenciosamente,
    """
        
        try:
            remetente = "comunicacao@stik.com.br"
            senha = "Mailstk400"
            destinatarios = ["joao.silva@stik.com.br", "bruno.costa@stik.com.br"]
            assunto = f"Danfe - {numero_nf_xml} / Stik Elásticos"

            for destinatario in destinatarios:
                msg = MIMEMultipart()
                msg["From"] = remetente
                msg["To"] = destinatario
                msg["Subject"] = assunto
                msg.attach(MIMEText(corpo_email, "plain"))

                with open(selecionados["xml"], "rb") as arquivo_xml:
                    parte_xml = MIMEBase("application", "octet-stream")
                    parte_xml.set_payload(arquivo_xml.read())
                    encoders.encode_base64(parte_xml)
                    parte_xml.add_header("Content-Disposition", f"attachment; filename={os.path.basename(selecionados['xml'])}")
                    msg.attach(parte_xml)

                with open(selecionados["pdf"], "rb") as arquivo_pdf:
                    parte_pdf = MIMEBase("application", "octet-stream")
                    parte_pdf.set_payload(arquivo_pdf.read())
                    encoders.encode_base64(parte_pdf)
                    parte_pdf.add_header("Content-Disposition", f"attachment; filename={os.path.basename(selecionados['pdf'])}")
                    msg.attach(parte_pdf)

                with smtplib.SMTP("smtp.stik.com.br", 587) as servidor:
                    servidor.starttls()
                    servidor.login(remetente, senha)
                    servidor.sendmail(remetente, destinatario, msg.as_string())

                print(f"E-mail enviado para {destinatario} com sucesso!")

            # Após o envio, marca os arquivos como enviados
            pedidos_enviados.add(selecionados["xml"])
            pedidos_enviados.add(selecionados["pdf"])

            # Renomeia os arquivos para marcar como enviados
            renomear_pedido_enviado(selecionados["xml"])
            renomear_pedido_enviado(selecionados["pdf"])

        except smtplib.SMTPException as e:
            messagebox.showerror("Erro SMTP", f"Erro ao tentar se conectar ao servidor SMTP: {str(e)}")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao enviar e-mail: {str(e)}")


    # Função para agrupar pedidos por NF
    def agrupar_pedidos_por_nf():
        arquivos = os.listdir(pasta_pedidos)
        pedidos_por_nf = {}

        for arquivo in arquivos:
            caminho_completo = os.path.join(pasta_pedidos, arquivo)

            if arquivo.endswith(".xml"):
                numero_nf = obter_valor_nf(caminho_completo)
                if numero_nf:
                    if numero_nf not in pedidos_por_nf:
                        pedidos_por_nf[numero_nf] = {"xml": None, "pdf": None}
                    pedidos_por_nf[numero_nf]["xml"] = caminho_completo
            
            elif arquivo.endswith(".pdf"):
                numero_nf = obter_valor_nf_pdf(caminho_completo)
                if numero_nf:
                    if numero_nf not in pedidos_por_nf:
                        pedidos_por_nf[numero_nf] = {"xml": None, "pdf": None}
                    pedidos_por_nf[numero_nf]["pdf"] = caminho_completo

        return pedidos_por_nf
    
    time.sleep(2)

    # Função para verificar pedidos e enviar e-mail
    def verificar_pedidos_e_enviar_email():
        print("Iniciando verificação de pedidos...") 
        arquivos = os.listdir(pasta_pedidos)
        pedidos_encontrados = any(arquivo.endswith((".xml", ".pdf")) for arquivo in arquivos)

        if pedidos_encontrados:
            print("Pedidos encontrados. Agrupando e enviando e-mail...") 
            pedidos_por_nf = agrupar_pedidos_por_nf()

            pedidos_enviados_com_sucesso = 0

            # Enviar e-mail para cada grupo de pedidos com a mesma NF
            for numero_nf, pedidos in pedidos_por_nf.items():
                if pedidos["xml"] and pedidos["pdf"]:
                    print(f"Enviando e-mail para NF-{numero_nf}")
                    selecionados["xml"] = pedidos["xml"]
                    selecionados["pdf"] = pedidos["pdf"]
                    enviar_email_com_anexo()
                    pedidos_enviados_com_sucesso += 1
                else:
                    print(f"Arquivos incompletos para NF-{numero_nf}, não enviados.")
            
            if pedidos_enviados_com_sucesso > 0:
                messagebox.showinfo("Sucesso", f"{pedidos_enviados_com_sucesso} pedido(s) enviado(s) com sucesso!")
            else:
                messagebox.showwarning("Aviso", "Nenhum pedido enviado.")
            
            time.sleep(2)

            pyautogui.press("enter")
            
            root.destroy()
        else:
            print("Nenhum pedido encontrado para enviar.")

    def abrir_pdf_com_acrobat(pdf_path):
        try:
            acrobat_path = r"C://Program Files//Adobe//Acrobat DC//Acrobat//Acrobat.exe"
            
            if not os.path.exists(acrobat_path):
                print(f"Erro: O arquivo {acrobat_path} não foi encontrado.")
                return
            
            subprocess.Popen([acrobat_path, pdf_path])
            print(f"PDF aberto com sucesso: {pdf_path}")
        
        except Exception as e:
            print(f"Erro ao abrir o PDF com Acrobat: {e}")

    def obter_valor_nf_pdf(pdf_path):
        try:
            extracted_text = ""
            with fitz.open(pdf_path) as pdf_document:
                for page_number in range(len(pdf_document)):
                    page = pdf_document.load_page(page_number)
                    text = page.get_text("text") 
                    extracted_text += text + "\n"

            print(f"Texto extraído do PDF ({pdf_path}):\n{extracted_text}")

            match = re.search(r"(?:N[°ºoO]?[°º]?\s*\.?\s*)?(\d{6,})", extracted_text)
            if match:
                numero_nf_pdf = match.group(1)
                print(f"Valor da NF encontrado no PDF: {numero_nf_pdf}")
                return numero_nf_pdf
            else:
                print("Número da NF não encontrado no PDF.")
                return None
        except Exception as e:
            print(f"Erro ao extrair número da NF do PDF: {e}")
            return None
    
    
    # Função para o duplo clique na Listbox
    def on_double_click(event):
        selecionado = listbox.get(listbox.curselection())
        if not selecionado:
            return
        
        caminho_completo = os.path.join(pasta_pedidos, selecionado)
        
        if selecionado.endswith(".xml"):
            selecionados["xml"] = caminho_completo
        elif selecionado.endswith(".pdf"):
            selecionados["pdf"] = caminho_completo
            abrir_pdf_com_acrobat(caminho_completo)
        
        atualizar_label_selecionados()

    def atualizar_lista():
        listbox.delete(0, tk.END)
        if os.path.exists(pasta_pedidos):
            arquivos = os.listdir(pasta_pedidos)
            for arquivo in arquivos:
                if arquivo.endswith(".xml") or arquivo.endswith(".pdf"):
                    listbox.insert(tk.END, arquivo)
        else:
            messagebox.showerror("Erro", f"O caminho da pasta não foi encontrado {pasta_pedidos}")

    def selecionar_pedido():
        selecionado = listbox.get(tk.ACTIVE)
        if not selecionado:
            messagebox.showwarning("Seleção inválida", "Selecione um arquivo da lista.")
            return
        
        caminho_completo = os.path.join(pasta_pedidos, selecionado)
        
        if selecionado.endswith(".xml"):
            selecionados["xml"] = caminho_completo
        elif selecionado.endswith(".pdf"):
            selecionados["pdf"] = caminho_completo
        
        atualizar_label_selecionados()

    def atualizar_label_selecionados():
        texto = f"XML: {os.path.basename(selecionados['xml']) if selecionados['xml'] else 'Nenhum'}\nPDF: {os.path.basename(selecionados['pdf']) if selecionados['pdf'] else 'Nenhum'}"
        label_selecionados.configure(text=texto, width=40)
        
    # Obtém o valor da Nf do XML
    def obter_valor_nf(caminho_xml):
        try:
            print(f"Abrindo XML: {caminho_xml}")  # Adicionando impressão de depuração
            tree = ET.parse(caminho_xml)
            root = tree.getroot()
            ns = {'ns': 'http://www.portalfiscal.inf.br/nfe'}
            nNF = root.find(".//ns:nNF", ns)
            if nNF is not None:
                return nNF.text
            else:
                print(f"Não foi possível encontrar o número da NF no XML: {caminho_xml}")
                return None
        except Exception as e:
            print(f"Erro ao extrair número da NF do XML: {e}")
            return None

    time.sleep(2)

            
    # Função para verificar os valores da nota fiscal do XML E PDF
    def verificar_correspondencia():
        if not selecionados["xml"] or not selecionados["pdf"]:
            messagebox.showwarning("Erro", "Selecione um arquivo XML e um PDF antes de verificar a correspondência.")
            return

        # Obter valor da NF-e do XML selecionado
        numero_nf_xml = obter_valor_nf(selecionados["xml"])
        if not numero_nf_xml:
            messagebox.showwarning("Erro", "Não foi possível encontrar o número da NF no XML.")
            return

        # Obter valor da NF do PDF
        numero_nf_pdf = obter_valor_nf_pdf(selecionados["pdf"])
        if not numero_nf_pdf:
            messagebox.showwarning("Erro", "Não foi possível encontrar o número da NF no PDF.")
            return

        # Comparar os valores da NF
        if numero_nf_xml != numero_nf_pdf:
            messagebox.showwarning("Erro", "Os números da NF no XML e no PDF não correspondem. O envio foi bloqueado.")
            return

        messagebox.showinfo("Sucesso", "Os números da NF no XML e no PDF correspondem.")

    def renomear_pedido_enviado(caminho_arquivo):
        if not caminho_arquivo.endswith("_enviado.pdf") and not caminho_arquivo.endswith("_enviado.xml"):
            novo_nome = caminho_arquivo.replace(".pdf", "_enviado.pdf").replace(".xml", "_enviado.xml")
            novo_caminho = os.path.join(pasta_pedidos, novo_nome)
            try:
                os.rename(caminho_arquivo, novo_caminho)
                print(f"Arquivo {os.path.basename(caminho_arquivo)} renomeado para {novo_nome}")
            except Exception as e:
                print(f"Erro ao renomear {os.path.basename(caminho_arquivo)}: {e}")

    """def anexar_pedido():
        if selecionados["xml"] and selecionados["pdf"]:
            messagebox.showinfo("Sucesso", "Pedidos anexados com sucesso!")
        else:
            messagebox.showwarning("Erro", "Selecione um XML e um PDF antes de anexar.")"""
            
    # Função que renomeia os arquivos automaticamente
    def renomear_pedidos_automatico(pasta_pedidos):
        arquivos = os.listdir(pasta_pedidos)
        
        for arquivo in arquivos:
            caminho_arquivo = os.path.join(pasta_pedidos, arquivo)

            # Renomeia XMLs
            if arquivo.endswith("-nfe.xml"):
                nf = obter_valor_nf(caminho_arquivo)
                if nf:
                    novo_nome = f"Nf-{nf}.xml"
                    novo_caminho = os.path.join(pasta_pedidos, novo_nome)
                    try:
                        os.rename(caminho_arquivo, novo_caminho)
                        print(f"Arquivo XML renomeado para {novo_nome}")
                    except Exception as e:
                        print(f"Erro ao renomear {arquivo}: {e}")

            # Renomeia PDFs
            elif arquivo.endswith(".pdf"):
                nf = obter_valor_nf_pdf(caminho_arquivo)
                if nf:
                    novo_nome = f"Nf-{nf}.pdf"
                    novo_caminho = os.path.join(pasta_pedidos, novo_nome)
                    try:
                        os.rename(caminho_arquivo, novo_caminho)
                        print(f"Arquivo PDF renomeado para {novo_nome}")
                    except Exception as e:
                        print(f"Erro ao renomear {arquivo}: {e}")

    # Executar a função para renomear os arquivos
    renomear_pedidos_automatico("C://1 - XML e PDF")

    def centralizar_janela(root):
        root.update_idletasks()
        largura = root.winfo_width()
        altura = root.winfo_height()
        x = (root.winfo_screenwidth() // 2) - (largura // 2)
        y = (root.winfo_screenheight() // 2) - (altura // 2)
        root.geometry(f"{largura}x{altura}+{x}+{y}")

    root = ctk.CTk()
    root.title("Enviar Pedidos")
    root.geometry("600x400")
    root.after(100, lambda: centralizar_janela(root))
    root.resizable(False, False)

    main_frame = ctk.CTkFrame(root)
    main_frame.pack(padx=10, pady=10, fill="both", expand=True)

    listbox = tk.Listbox(main_frame, selectmode=tk.SINGLE, height=10)
    listbox.pack(fill="both", expand=True, pady=10)

    button_frame = ctk.CTkFrame(main_frame)
    button_frame.pack(pady=10)

    button_selecionar = ctk.CTkButton(button_frame, text="Anexar Pedidos e Enviar", command=enviar_email_com_anexo, width=120)
    button_selecionar.grid(row=0, column=0, padx=5, pady=5)

    """button_anexar = ctk.CTkButton(button_frame, text="Anexar Pedido", command=anexar_pedido, width=120)
    button_anexar.grid(row=0, column=1, padx=5, pady=5)"""

    """button_enviar_pedido = ctk.CTkButton(button_frame, text="Enviar Pedido", width=120, command=enviar_email_com_anexo)
    button_enviar_pedido.grid(row=0, column=2, padx=5, pady=5)"""

    label_selecionados = ctk.CTkLabel(button_frame, text="XML: Nenhum\nPDF: Nenhum", justify="left", width=40)
    label_selecionados.grid(row=1, column=0, columnspan=3, pady=5)

    renomear_pedidos_automatico(pasta_pedidos)
    root.after(2000, verificar_pedidos_e_enviar_email)
    atualizar_lista()
    listbox.bind("<Double-1>", on_double_click)
    root.mainloop()

def main():
    iniciar_robo()
    
    time.sleep(1)
    
    pedidos = [     
        arquivo
        for arquivo in os.listdir("C://1 - XML e PDF")
        if arquivo.lower().endswith('.pdf') and '_enviado' not in arquivo
    ]
    
    if pedidos:
        print(f"Foram encontrados {len(pedidos)} pedidos. Iniciando a GUI...")
        iniciar_GUI()
    else:
        messagebox.showinfo("Aviso", "Nenhum pedido encontrado para enviar.")

if __name__ == "__main__":
    main()