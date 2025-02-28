import pyautogui
import subprocess
import time
import pyperclip 
import sys
import logging
import pygetwindow as gw
from datetime import datetime, timedelta
import customtkinter as ctk
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog, messagebox, simpledialog
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import fitz
import pyodbc
import threading
from PIL import Image
import os
import pytesseract
import re

pasta_pedidos = "C:/2 - Boletos"

def iniciar_robo():
    # Tempo
    time_busca = 0.2
    time_login = 20
    time_short = 1
    time_medium = 3
    time_maximizar = 10
    time_long = 8

    senha_topmanager = "123456"
    tempo_espera_max = 50

    coordenadas = {
        "Movimentos": (147, 34),
        "Vendas": (166, 171),
        "Titulo Venda": (371, 486),
        #Maximiza aba
        "Filtro": (211, 78),
        "Empresa":(196, 93),
        "Agent. Cobrador": (744, 91),
        "Rest. Titulo": (944, 92),
        "Periodo Emi": (1410, 93),
        # F5
        "N. Pedido Base": (40, 114),
        "Pos. Acrobat": (104, 38),
        "Fechar janela navegador": (1900, 19),
        "Fechar Janela": (1892, 13),
        }

    pedido_x, pedido_y = coordenadas["N. Pedido Base"]
    incremento_y = 16

    def contar_pedidos():
        total_pedidos = 0
        pedidos_encontrados = []  
        temp_y = pedido_y
        last_pedido = None  

        while True:
            pyautogui.moveTo(pedido_x, temp_y)
            pyautogui.doubleClick()
            time.sleep(time_short)

            pyautogui.hotkey("ctrl", "c")
            time.sleep(time_short)
            pedido_atual = pyperclip.paste().strip()

            if pedido_atual:
                if pedido_atual == last_pedido:
                    # Se o pedido atual for o mesmo que o último, significa que chegamos ao fim
                    break
                pedidos_encontrados.append(pedido_atual)  # Armazena o pedido
                last_pedido = pedido_atual  # Atualiza o último pedido
                total_pedidos += 1  # Conta um novo pedido
            else:
                # Se a célula estiver vazia, interrompe
                break

            temp_y += incremento_y  # Move para a próxima linha
            pyautogui.press("down")
            time.sleep(time_short)

        return total_pedidos, pedidos_encontrados

    def baixar_boleto(pedido_x, pedido_y):
        try:
            time.sleep(time_short)
            pyautogui.click(button="right") 
            time.sleep(time_short)

            for _ in range(6):
                pyautogui.press("down")
                time.sleep(0.2)  

            pyautogui.press("enter")  
            time.sleep(time_long)

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
            preencher_campo("C:/2 - Boletos")
            time.sleep(time_medium)
            
            pyautogui.press("enter")
            time.sleep(time_short)
            pyautogui.moveTo(781, 508)
            pyautogui.click()
            time.sleep(time_short)
            
            pyautogui.moveTo(coordenadas['Fechar janela navegador'])
            pyautogui.click()
            time.sleep(time_medium)
            pyautogui.moveTo(coordenadas["Fechar Janela"])
            pyautogui.click()
            time.sleep(time_medium)
            logging.info("PDF baixado com sucesso.")

        except Exception as e:
            logging.error(f"Erro ao baixar PDF: {e}")
            
    def preencher_campo(texto):
        pyautogui.press("backspace")
        pyautogui.write(texto)
        
    #Função para maximizar uma janela específica ou clicar no botão de maximizar
    def maximizar_janela(titulo_parcial):
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
        
    #Entrar e logar no TopManager
    try:
        if not gw.getWindowsWithTitle("TopManager"):
            pyautogui.hotkey("ctrl", "shift", "t")
            time.sleep(5)
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
            
        #Acessar a aba dos Documentos Eletrônicos
        try:
            pyautogui.moveTo(coordenadas['Movimentos'])
            pyautogui.click()
            time.sleep(time_short)

            pyautogui.moveTo(coordenadas["Vendas"])
            pyautogui.click()
            time.sleep(time_short)

            pyautogui.moveTo(coordenadas['Titulo Venda'])
            pyautogui.click()
            time.sleep(5)
            pyautogui.moveTo(731, 84)
            pyautogui.click()
            time.sleep(time_short)
            time.sleep(time_short)
            
            pyautogui.moveTo(coordenadas['Filtro'])
            pyautogui.click()
            time.sleep(time_short)

            pyautogui.moveTo(coordenadas['Empresa'])
            pyautogui.click()
            time.sleep(time_short)
            pyautogui.write("2")
            time.sleep(time_short)
            pyautogui.moveTo(coordenadas['Agent. Cobrador'])
            time.sleep(time_short)
            pyautogui.click()
            
            preencher_campo("53")
            time.sleep(time_short)
            pyautogui.press("enter")   
            time.sleep(time_short)
            
            pyautogui.moveTo(coordenadas['Rest. Titulo'])
            pyautogui.click()
            time.sleep(time_short)
            pyautogui.write("em")
            time.sleep(time_short)
            pyautogui.press("enter")
            time.sleep(time_medium)
            pyautogui.moveTo(coordenadas['Periodo Emi'])
            pyautogui.click()
            time.sleep(time_short)
            pyautogui.write("h")
            time.sleep(time_short)
            pyautogui.press("enter")
            time.sleep(time_short)
            """pyautogui.press("tab")
            preencher_campo("07/02/2025")
            time.sleep(time_medium)
            pyautogui.press("tab")
            preencher_campo("07/02/2025")
            time.sleep(time_medium)
            pyautogui.press("enter")
            time.sleep(time_medium)"""
            
            pyautogui.press('F5')
            time.sleep(time_long)
            logging.info("Acesso à aba de Titulo de Vendas.")
            
        except Exception as e:
            logging.error(f"Erro ao acessar aba de Titulo de Vendas: {e}")
            raise Exception("Falha ao acessar a aba de Títutlos de Vendas")
            
           
        def baixar_boletos(pedidos_encontrados):
            temp_y = pedido_y  
            for pedido in pedidos_encontrados:
                pyautogui.moveTo(pedido_x, temp_y)
                pyautogui.doubleClick()
                time.sleep(time_busca)

                baixar_boleto(pedido_x, temp_y)  

                temp_y += incremento_y  

        logging.info("Contando quantidade de pedidos disponíveis...")
        numero_pedidos, pedidos_encontrados = contar_pedidos()

        if numero_pedidos == 0:
            logging.info("Nenhum pedido disponível para baixar.")
        else:
            logging.info(f"Total de pedidos encontrados: {numero_pedidos}")

            # Iniciar o processo de download dos boletos
            logging.info("Iniciando download dos boletos...")
            baixar_boletos(pedidos_encontrados)
        
        try:
            time.sleep(time_medium)
            pyautogui.moveTo(1898, 14)
            pyautogui.click()
            logging.info("TopManager fechado com sucesso.")
        except Exception as e:
            logging.error(f"Erro ao fechar o TopManager: {e}")

    except Exception as e:
        logging.error(f"Erro no processo: {e}")

def iniciar_GUI():

    pytesseract.pytesseract.tesseract_cmd = r"C://Users//fnojosa//AppData//Local//Programs//Tesseract-OCR//tesseract.exe"

    pasta_pedidos = "C:/2 - Boletos"
    selecionados = {"pdf": None}

    def obter_conexao_sql():
        server = '45.235.240.135'
        database = 'Stik'
        username = 'ti'
        password = 'Stik0123'
        conn_str = f"DRIVER={{SQL Server}};SERVER={server};DATABASE={database};UID={username};PWD={password}"
        try:
            conexao = pyodbc.connect(conn_str)
            return conexao
        except Exception as e:
            messagebox.showerror("Erro de Conexão", f"Erro ao conectar ao banco de dados: {e}")
            return None

    def on_double_click(event):
        selecionados = treeview.get(treeview.curselection())
        caminho_boleto = os.path.join(pasta_pedidos, selecionados)
        if os.path.exists(caminho_boleto):
            os.startfile(caminho_boleto)
        else:
            messagebox.showerror("Erro", "Arquivo não encontrado")


    # Função que renomeia os boletos na pasta
    def renomear_boletos_pasta(pasta):
        documentos_parcelas = {}
        arquivos_nao_renomeados = [arquivo for arquivo in os.listdir(pasta) if arquivo.lower().endswith('.pdf') and '-' not in arquivo]

        if not arquivos_nao_renomeados:
            print("Nenhum pedido não renomeado encontrado.")
            return

        for arquivo in arquivos_nao_renomeados:
            caminho_pdf = os.path.join(pasta, arquivo)
            try:
                doc = fitz.open(caminho_pdf)
                page = doc.load_page(0)
                pix = page.get_pixmap(matrix=fitz.Matrix(300/72, 300/72))
                img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)

                texto = pytesseract.image_to_string(img)
                print(f"Texto extraído de {arquivo}:\n{texto}\n")
                doc.close()

                match = re.search(r"Nr\.\s*do\s*Documento.*?(\d{7,10})", texto, re.IGNORECASE | re.DOTALL)

                if match:
                    document_number = match.group(1)
                    base_number = document_number[:-1]
                    if base_number not in documentos_parcelas:
                        documentos_parcelas[base_number] = 1
                    else:
                        documentos_parcelas[base_number] += 1
                    
                    parcela = documentos_parcelas[base_number]
                    novo_nome = f"Stik.{base_number}-{parcela}.pdf"  
                    novo_caminho = os.path.join(pasta, novo_nome)
                    
                    if os.path.exists(novo_caminho):
                        print(f"Arquivo {novo_nome} já existe. Pulando a renomeação de {arquivo}.")
                    else:
                        os.rename(caminho_pdf, novo_caminho)
                        print(f"Renomeado {arquivo} para {novo_nome}")
                else:
                    print(f"Número de documento não encontrado em {arquivo}")
                    
            except Exception as e:
                print(f"Erro ao processar {arquivo}: {e}")

    # Função que verifica a correspondência de todos os boletos com o mesmo número de documento
    def agrupar_boletos_por_documento():
        boletos_por_documento = {}

        # Renomeia os boletos antes de enviar
        renomear_boletos_pasta(pasta_pedidos)

        boletos_renomeados = [arquivo for arquivo in os.listdir(pasta_pedidos) if arquivo.lower().endswith('.pdf') and '-' in arquivo]

        if not boletos_renomeados:
            messagebox.showwarning("Erro", "Nenhum boleto encontrado para enviar.")
            return

        # Agrupa os boletos pelo número de documento
        for arquivo in boletos_renomeados:
            numero_documento = arquivo.split("-")[0] 
            if numero_documento not in boletos_por_documento:
                boletos_por_documento[numero_documento] = []
            boletos_por_documento[numero_documento].append(arquivo)

        return boletos_por_documento

    # Função para enviar boletos em sequência
    def enviar_boletos_sequenciais():
        boletos_por_documento = agrupar_boletos_por_documento()

        if not boletos_por_documento:
            messagebox.showwarning("Erro", "Nenhum boleto encontrado para enviar.")
            return

        for numero_documento, boletos in boletos_por_documento.items():
            print(f"Enviando boletos para o número de documento: {numero_documento}")
            
            enviar_email_com_anexo(boletos)

            for boleto in boletos:
                marcar_boleto_enviado(boleto)

        messagebox.showinfo("Sucesso", "Todos os boletos foram enviados com sucesso!")
        
    def marcar_boleto_enviado(boleto):
        novo_nome = boleto.replace(".pdf", "_enviado.pdf")
        os.rename(os.path.join(pasta_pedidos, boleto), os.path.join(pasta_pedidos, novo_nome))
        print(f"Boleto {boleto} marcado como enviado.")

    # Função que extrai o número do documento do PDF
    def extrair_numero_documento(pdf):
        try:
            doc = fitz.open(pdf)
            page = doc.load_page(0)
            pix = page.get_pixmap(matrix=fitz.Matrix(300/72, 300/72))
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            
            texto = pytesseract.image_to_string(img)
            doc.close()

            # Procura pelo número do documento no texto extraído
            match = re.search(r"Nr\.\s*do\s*Documento.*?(\d{7,10})", texto, re.IGNORECASE | re.DOTALL)
            if match:
                return match.group(1)
            else:
                return None
        except Exception as e:
            print(f"Erro ao extrair número do documento: {e}")
            return None

    # Função que envia o Email com os boletos encontrados
    def enviar_email_com_anexo(boletos):
        if not boletos:
            messagebox.showwarning("Erro", "Nenhum boleto encontrado para enviar.")
            return

        corpo_email = f"""Prezado(a),

    Segue em anexo os boletos.

    Atenciosamente,
    """

        try:
            remetente = "comunicacao@stik.com.br"
            senha = "Mailstk400"
            destinatarios = ["joao.silva@stik.com.br", "bruno.costa@stik.com.br,", "informatica01@stik.com.br"]
            assunto = f"Boletos - {boletos[0]} / Stik Elásticos"

            for destinatario in destinatarios:
                msg = MIMEMultipart()
                msg["From"] = remetente
                msg["To"] = destinatario
                msg["Subject"] = assunto
                msg.attach(MIMEText(corpo_email, "plain"))

                for boleto in boletos:
                    with open(os.path.join(pasta_pedidos, boleto), "rb") as arquivo_pdf:
                        parte_pdf = MIMEBase("application", "octet-stream")
                        parte_pdf.set_payload(arquivo_pdf.read())
                        encoders.encode_base64(parte_pdf)
                        parte_pdf.add_header(
                            "Content-Disposition",
                            f"attachment; filename={boleto}",
                        )
                        msg.attach(parte_pdf)

                with smtplib.SMTP("smtp.stik.com.br", 587) as servidor:
                    servidor.starttls()
                    servidor.login(remetente, senha)
                    servidor.sendmail(remetente, destinatario, msg.as_string())

        except smtplib.SMTPException as e:
            messagebox.showerror("Erro SMTP", f"Erro ao tentar se conectar ao servidor SMTP: {str(e)}")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao enviar e-mail: {str(e)}")
        
    def selecionar_boletos():
        global boletos_selecionados
        boletos_selecionados = filedialog.askopenfilenames(
            title="Selecione os boletos para enviar",
            filetypes=[("Arquivos PDF", "*.pdf")],
        )
        if boletos_selecionados:
            label_selecionados.configure(text=f"Boletos Selecionados: {len(boletos_selecionados)}")

    # Função para atualizar a lista de arquivos na ListBox
    def atualizar_lista():
        treeview.delete(*treeview.get_children())
        arquivos = os.listdir(pasta_pedidos)
        for arquivo in arquivos:
            if arquivo.endswith(".pdf"):
                pedidos_renomeados = arquivo
                fornecedores = "Angela Maria Oliveira"
                operacao = "Nota Fiscal Eletrônica (p)"
                treeview.insert("", "end", values=(pedidos_renomeados, fornecedores, operacao))
                
    def centralizar_janela(root):
        root.update_idletasks()
        largura = root.winfo_width()
        altura = root.winfo_height()
        x = (root.winfo_screenwidth() // 2) - (largura // 2)
        y = (root.winfo_screenheight() // 2) - (altura // 2)
        root.geometry(f"{largura}x{altura}+{x}+{y}")

    root = ctk.CTk()
    root.title("Enviar Boletos")
    root.geometry("600x400")
    root.resizable(False, False)

    # Criação do frame principal
    main_frame = ctk.CTkFrame(root)
    main_frame.pack(padx=10, pady=10, fill="both", expand=True)

    # Criação do Treeview com as 3 colunas
    treeview = ttk.Treeview(main_frame, columns=("Título", "Cliente", "Operação"), show="headings")
    treeview.heading("Título", text="Título", anchor="w")
    treeview.heading("Cliente", text="Cliente", anchor="w")
    treeview.heading("Operação", text="Operação",anchor="w")
    treeview.pack(pady=10, fill="both", expand=True)

    # Botões e outros componentes
    button_frame = ctk.CTkFrame(main_frame)
    button_frame.pack(pady=10)

    button_selecionar = ctk.CTkButton(button_frame, text="Anexar Pedidos e Enviar", command=enviar_boletos_sequenciais,width=120)
    button_selecionar.grid(row=0, column=0, padx=5, pady=5)

    # Label que mostra a quantidade de boletos selecionados
    label_selecionados = ctk.CTkLabel(main_frame, text="Boletos Selecionados: Nenhum", width=40)
    label_selecionados.pack(pady=10)

    # Função para centralizar a janela
    def centralizar_janela(root):
        root.update_idletasks()
        largura = root.winfo_width()
        altura = root.winfo_height()
        x = (root.winfo_screenwidth() // 2) - (largura // 2)
        y = (root.winfo_screenheight() // 2) - (altura // 2)
        root.geometry(f"{largura}x{altura}+{x}+{y}")

    # Inicialização do layout
    root.after(100, lambda: centralizar_janela(root))
    root.after(100, renomear_boletos_pasta(pasta_pedidos))
    root.after(100, atualizar_lista)
    root.mainloop()

def main():
    iniciar_robo()
    
    time.sleep(1)
    
    pedidos = [
        arquivo
        for arquivo in os.listdir(pasta_pedidos)
        if arquivo.lower().endswith('.pdf') and '_enviado' not in arquivo
    ]
    
    if pedidos:
        print(f"Foram encontrados {len(pedidos)} pedidos. Iniciando a GUI...")
        iniciar_GUI()
    else:
        messagebox.showinfo("Aviso", "Nenhum pedido encontrado para enviar.")

if __name__ == "__main__":
    main()