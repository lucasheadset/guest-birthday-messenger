from datetime import datetime
import fitz  # PyMuPDF
import pywhatkit as kit
import tkinter as tk
from tkcalendar import Calendar
from tkinter import messagebox, ttk
from PIL import Image, ImageTk, ImageSequence
from tkinter import Toplevel, Menu
import re
import pandas as pd
import os
import time
import sys
import ctypes
import threading
import webbrowser
import pystray
from pystray import MenuItem as item
import atexit
import unicodedata
import configparser
import plyer
import random
from plyer import notification
from tkcalendar import DateEntry
import json 

COUNTRIES_DATA = [
    ('Brazil', 'BR', '55'),
    ('United States', 'US', '1'),
    ('United Kingdom', 'GB', '44'),
    ('Afghanistan', 'AF', '93'),
    ('Albania', 'AL', '355'),
    ('Algeria', 'DZ', '213'),
    ('Andorra', 'AD', '376'),
    ('Angola', 'AO', '244'),
    ('Argentina', 'AR', '54'),
    ('Armenia', 'AM', '374'),
    ('Australia', 'AU', '61'),
    ('Austria', 'AT', '43'),
    ('Azerbaijan', 'AZ', '994'),
    ('Bahamas', 'BS', '1242'),
    ('Bahrain', 'BH', '973'),
    ('Bangladesh', 'BD', '880'),
    ('Barbados', 'BB', '1246'),
    ('Belarus', 'BY', '375'),
    ('Belgium', 'BE', '32'),
    ('Belize', 'BZ', '501'),
    ('Benin', 'BJ', '229'),
    ('Bhutan', 'BT', '975'),
    ('Bolivia', 'BO', '591'),
    ('Bosnia and Herzegovina', 'BA', '387'),
    ('Botswana', 'BW', '267'),
    ('Bulgaria', 'BG', '359'),
    ('Burkina Faso', 'BF', '226'),
    ('Burundi', 'BI', '257'),
    ('Cambodia', 'KH', '855'),
    ('Cameroon', 'CM', '237'),
    ('Canada', 'CA', '1'),
    ('Cape Verde', 'CV', '238'),
    ('Central African Republic', 'CF', '236'),
    ('Chad', 'TD', '235'),
    ('Chile', 'CL', '56'),
    ('China', 'CN', '86'),
    ('Colombia', 'CO', '57'),
    ('Comoros', 'KM', '269'),
    ('Congo', 'CG', '242'),
    ('Costa Rica', 'CR', '506'),
    ("Cote d'Ivoire", 'CI', '225'),
    ('Croatia', 'HR', '385'),
    ('Cuba', 'CU', '53'),
    ('Cyprus', 'CY', '357'),
    ('Czech Republic', 'CZ', '420'),
    ('Denmark', 'DK', '45'),
    ('Djibouti', 'DJ', '253'),
    ('Dominica', 'DM', '1767'),
    ('Dominican Republic', 'DO', '1809'),
    ('Ecuador', 'EC', '593'),
    ('Egypt', 'EG', '20'),
    ('El Salvador', 'SV', '503'),
    ('Equatorial Guinea', 'GQ', '240'),
    ('Eritrea', 'ER', '291'),
    ('Estonia', 'EE', '372'),
    ('Ethiopia', 'ET', '251'),
    ('Fiji', 'FJ', '679'),
    ('Finland', 'FI', '358'),
    ('France', 'FR', '33'),
    ('Gabon', 'GA', '241'),
    ('Gambia', 'GM', '220'),
    ('Georgia', 'GE', '995'),
    ('Germany', 'DE', '49'),
    ('Ghana', 'GH', '233'),
    ('Greece', 'GR', '30'),
    ('Greenland', 'GL', '299'),
    ('Guatemala', 'GT', '502'),
    ('Guinea', 'GN', '224'),
    ('Haiti', 'HT', '509'),
    ('Honduras', 'HN', '504'),
    ('Hong Kong', 'HK', '852'),
    ('Hungary', 'HU', '36'),
    ('Iceland', 'IS', '354'),
    ('India', 'IN', '91'),
    ('Indonesia', 'ID', '62'),
    ('Iran', 'IR', '98'),
    ('Iraq', 'IQ', '964'),
    ('Ireland', 'IE', '353'),
    ('Israel', 'IL', '972'),
    ('Italy', 'IT', '39'),
    ('Jamaica', 'JM', '1876'),
    ('Japan', 'JP', '81'),
    ('Jordan', 'JO', '962'),
    ('Kazakhstan', 'KZ', '7'),
    ('Kenya', 'KE', '254'),
    ('Kuwait', 'KW', '965'),
    ('Kyrgyzstan', 'KG', '996'),
    ('Latvia', 'LV', '371'),
    ('Lebanon', 'LB', '961'),
    ('Liberia', 'LR', '231'),
    ('Libya', 'LY', '218'),
    ('Liechtenstein', 'LI', '423'),
    ('Lithuania', 'LT', '370'),
    ('Luxembourg', 'LU', '352'),
    ('Macao', 'MO', '853'),
    ('Madagascar', 'MG', '261'),
    ('Malaysia', 'MY', '60'),
    ('Maldives', 'MV', '960'),
    ('Mali', 'ML', '223'),
    ('Malta', 'MT', '356'),
    ('Mexico', 'MX', '52'),
    ('Monaco', 'MC', '377'),
    ('Mongolia', 'MN', '976'),
    ('Montenegro', 'ME', '382'),
    ('Morocco', 'MA', '212'),
    ('Mozambique', 'MZ', '258'),
    ('Myanmar', 'MM', '95'),
    ('Namibia', 'NA', '264'),
    ('Nepal', 'NP', '977'),
    ('Netherlands', 'NL', '31'),
    ('New Zealand', 'NZ', '64'),
    ('Nicaragua', 'NI', '505'),
    ('Niger', 'NE', '227'),
    ('Nigeria', 'NG', '234'),
    ('North Korea', 'KP', '850'),
    ('Norway', 'NO', '47'),
    ('Oman', 'OM', '968'),
    ('Pakistan', 'PK', '92'),
    ('Panama', 'PA', '507'),
    ('Papua New Guinea', 'PG', '675'),
    ('Paraguay', 'PY', '595'),
    ('Peru', 'PE', '51'),
    ('Philippines', 'PH', '63'),
    ('Poland', 'PL', '48'),
    ('Portugal', 'PT', '351'),
    ('Puerto Rico', 'PR', '1'),
    ('Qatar', 'QA', '974'),
    ('Romania', 'RO', '40'),
    ('Russia', 'RU', '7'),
    ('Rwanda', 'RW', '250'),
    ('San Marino', 'SM', '378'),
    ('Saudi Arabia', 'SA', '966'),
    ('Senegal', 'SN', '221'),
    ('Serbia', 'RS', '381'),
    ('Seychelles', 'SC', '248'),
    ('Sierra Leone', 'SL', '232'),
    ('Singapore', 'SG', '65'),
    ('Slovakia', 'SK', '421'),
    ('Slovenia', 'SI', '386'),
    ('Somalia', 'SO', '252'),
    ('South Africa', 'ZA', '27'),
    ('South Korea', 'KR', '82'),
    ('Spain', 'ES', '34'),
    ('Sri Lanka', 'LK', '94'),
    ('Sudan', 'SD', '249'),
    ('Sweden', 'SE', '46'),
    ('Switzerland', 'CH', '41'),
    ('Syria', 'SY', '963'),
    ('Taiwan', 'TW', '886'),
    ('Tanzania', 'TZ', '255'),
    ('Thailand', 'TH', '66'),
    ('Togo', 'TG', '228'),
    ('Tonga', 'TO', '676'),
    ('Tunisia', 'TN', '216'),
    ('Turkey', 'TR', '90'),
    ('Turkmenistan', 'TM', '993'),
    ('Uganda', 'UG', '256'),
    ('Ukraine', 'UA', '380'),
    ('United Arab Emirates', 'AE', '971'),
    ('Uruguay', 'UY', '598'),
    ('Uzbekistan', 'UZ', '998'),
    ('Venezuela', 'VE', '58'),
    ('Vietnam', 'VN', '84'),
    ('Yemen', 'YE', '967'),
    ('Zambia', 'ZM', '260'),
    ('Zimbabwe', 'ZW', '263'),
]

def iso_to_emoji(iso_code):
    """
    Converte um código de país ISO de duas letras em um emoji de bandeira.
    """
    return ''.join(chr(ord(c.upper()) + 127397) for c in iso_code)

# Função para calcular a idade
def calculate_age(birthdate):
    birthdate = datetime.strptime(birthdate, "%d/%m/%Y")
    today = datetime.today()
    age = today.year - birthdate.year - ((today.month, today.day) < (birthdate.month, birthdate.day))
    return age

def carregar_base_de_dados(caminho_csv):
    """
    Carrega a base de dados completa de um único arquivo CSV.
    """
    try:
        # dtype=str garante que telefones e CPFs não percam o zero à esquerda
        df = pd.read_csv(caminho_csv, dtype=str)
        # Remove linhas onde a coluna NOME está vazia, se houver
        df.dropna(subset=['NOME'], inplace=True)
        return df
    except FileNotFoundError:
        messagebox.showerror("Erro Crítico", f"A base de dados '{caminho_csv}' não foi encontrada!")
        return None
    
def encontrar_aniversariantes(hospedes_df, specific_date=None):
    """
    Encontra aniversariantes em um DataFrame já carregado.
    """
    if hospedes_df is None:
        return []

    birthdays = []
    # Usa a data específica se fornecida, caso contrário usa a data atual
    target_date = specific_date if specific_date else datetime.now().strftime("%d/%m")

    # Garante que a coluna de data de nascimento exista
    if 'DT. NASCIMENTO' not in hospedes_df.columns:
        messagebox.showerror("Erro", "A coluna 'DT. NASCIMENTO' não foi encontrada na base de dados.")
        return []

    # Filtra o DataFrame para encontrar aniversariantes do dia
    aniversariantes_df = hospedes_df[hospedes_df['DT. NASCIMENTO'].str.startswith(target_date, na=False)]

    for index, row in aniversariantes_df.iterrows():
        nome = row['NOME']
        dt_nascimento = row['DT. NASCIMENTO']
        
        # Pega o telefone e formata
        telefone = row.get('TELEFONE', 'Telefone não encontrado')
        if pd.isna(telefone) or not str(telefone).strip():
            telefone = "Telefone não encontrado"
            nome_original = nome # Guarda o nome original
            nome = "INVÁLIDO (SEM TELEFONE)" # Marca como inválido para a lista
        else:
            telefone = str(telefone).replace(" ", "").replace("-", "").replace("(", "").replace(")", "")
            if not telefone.startswith("+"):
                 telefone = "+55" + telefone
        
        idade = calculate_age(dt_nascimento)
        birthdays.append((nome, idade, dt_nascimento, telefone))

    return birthdays

# Função para validar números de telefone brasileiros
def is_valid_number(phone_number):
    # Expressão regular para validar um número de telefone brasileiro no formato +55DDDNXXXXXXX
    pattern = r'^\+55\d{2}\d{8,9}$'
    if re.match(pattern, phone_number):
        return True
    else:
        return False

CONFIG_FILE = "window_position.ini"

def save_window_position(root):
    config = configparser.ConfigParser()
    config['WINDOW'] = {
        'x': root.winfo_x(),
        'y': root.winfo_y()
    }
    with open(CONFIG_FILE, 'w') as configfile:
        config.write(configfile)

def load_window_position(root):
    if os.path.exists(CONFIG_FILE):
        config = configparser.ConfigParser()
        config.read(CONFIG_FILE)
        if 'WINDOW' in config:
            x = config.getint('WINDOW', 'x', fallback=100)
            y = config.getint('WINDOW', 'y', fallback=100)
            root.geometry(f"+{x}+{y}")

# Função para carregar a lista de exceção de um arquivo
def load_exceptions_from_file():
    if os.path.exists("exceções.txt"):
        with open("exceções.txt", "r") as f:
            return [line.strip() for line in f.readlines()]
    return []

# Função para carregar as preferências do usuário (como o modo escuro e opções de configurações)
def load_preferences():
    try:
        with open('config.json', 'r') as f:
            return json.load(f)
    except FileNotFoundError:
        return {
            "dark_mode": False,
            "allow_background_execution": True,
            "enable_notifications": True,
            "default_send_without_age": False,
            "remove_intro": False,
            "disallow_manual_mode": False,
            "block_manual_mode": False,
            "block_calendar": False,
            "send_time_warning": False,
            "password_to_start": False,
            "start_system": False,
            "show_guest_info_on_double_click": False,
            "filter_by_date": False
        }  # Padrão caso o arquivo não exista

# Função para salvar as preferências do usuário
def save_preferences(preferences):
    with open('config.json', 'w') as f:
        json.dump(preferences, f)

# Carregar as preferências do usuário
preferences = load_preferences()
is_dark_mode = preferences.get("dark_mode", False)

def show_birthday_notification(total_birthdays):
    if total_birthdays > 0:
        notification_text = f"Hoje temos {total_birthdays} hóspede(s) que faz(em) aniversário!"
    else:
        notification_text = "Hoje não temos aniversariantes."

    notification.notify(
        title= "Parabenize os aniversariantes diários",
        message=notification_text,
        app_name="Mensageiro de Aniversariantes",
        app_icon="C:\\Users\\Lucas Siqueira\\Downloads\\aniversario atualizado 18-10\\principal.ico",  # Caminho para o ícone do programa
        timeout=10  # Tempo que a notificação fica na tela
    )
    
def format_name(name):
        return name.split()[0].capitalize()
    
def show_custom_dialog(title, message, confirm_action, cancel_action=None):
    dialog = tk.Toplevel()
    dialog.title(title)
    dialog.geometry("350x150")
    dialog.configure(bg="#f0f0f0")
    dialog.grab_set()

    style = ttk.Style()
    style.configure("TLabel", font=("Helvetica", 10), background="#f0f0f0")
    style.configure("TButton", font=("Helvetica", 10), padding=6)
    
    tk.Label(dialog, text=message, wraplength=300, justify="center", bg="#f0f0f0", font=("Helvetica", 12)).pack(pady=20)

    button_frame = tk.Frame(dialog, bg="#f0f0f0")
    button_frame.pack(pady=10)

    ttk.Button(button_frame, text="Sim", command=lambda: [confirm_action(), dialog.destroy()]).pack(side="left", padx=10)
    
    if cancel_action:
        ttk.Button(button_frame, text="Não, enviar agora", command=lambda: [cancel_action(), dialog.destroy()]).pack(side="right", padx=10)
    else:
        ttk.Button(button_frame, text="Fechar", command=dialog.destroy).pack(side="right", padx=10)

    dialog.wait_window()

# Função para remover acentos de uma string
def remover_acentos(texto):
    return ''.join(
        c for c in unicodedata.normalize('NFD', texto)
        if unicodedata.category(c) != 'Mn'
    )

# Função para determinar a saudação baseada no nome
def determine_salutation(name):
    first_name = format_name(name)
    
    # Verifica se já existe um pronome salvo
    saved_pronoun = get_saved_pronoun(first_name)
    if saved_pronoun:
        return f"Prezado {first_name}" if saved_pronoun == "Masculino" else f"Prezada {first_name}"

    if first_name in nomes_femininos:
        return f"Prezada {first_name}"
    elif first_name in nomes_masculinos:
        return f"Prezado {first_name}"
    else:
        return f"Prezado(a) {first_name}"
    
def define_pronoun(treeview, selected_item, on_close_callback):
    # Verifique se já existe uma janela de pronome aberta
    for widget in root.winfo_children():
        if isinstance(widget, tk.Toplevel) and widget.title() == "Definir Pronome":
            return  # Se já existir, não crie outra

    pronoun_window = tk.Toplevel()
    pronoun_window.title("Definir Pronome")
    pronoun_window.grab_set()  # Impede a interação com outras janelas até que essa seja fechada
    pronoun_window.configure(bg="#e0e0e0", padx=20, pady=20)

    title_font = ("Helvetica", 14, "bold")
    label_font = ("Helvetica", 12)
    button_font = ("Helvetica", 12, "bold")

    tk.Label(pronoun_window, text="Selecione o pronome", bg="#e0e0e0", font=title_font).grid(row=0, column=0, padx=10, pady=(0, 10), sticky="w")

    values = treeview.item(selected_item, "values")
    name, phone, age = values
    first_name = format_name(name)

    # Adiciona o nome do hóspede abaixo de "Selecione o pronome"
    tk.Label(pronoun_window, text=f"Hóspede: {first_name}", bg="#e0e0e0", font=label_font).grid(row=1, column=0, padx=10, pady=(0, 10), sticky="w")

    pronoun_var = tk.StringVar(pronoun_window)

    if first_name in nomes_masculinos:
        pronoun_var.set("Masculino")
    elif first_name in nomes_femininos:
        pronoun_var.set("Feminino")
    else:
        pronoun_var.set("Não informado")

    pronoun_options = ["Masculino", "Feminino", "Não informado"]
    pronoun_menu = tk.OptionMenu(pronoun_window, pronoun_var, *pronoun_options)
    pronoun_menu.config(font=label_font, bg="#f7f7f7", fg="#333333", highlightthickness=1, highlightbackground="#b0b0b0")
    pronoun_menu.grid(row=2, column=0, padx=10, pady=10, sticky="ew")

    save_button = tk.Button(pronoun_window, text="Salvar", command=lambda: save_pronoun(pronoun_window, pronoun_var, first_name, on_close_callback), font=button_font, bg="#4caf50", fg="white", relief="flat", padx=10, pady=5)
    save_button.grid(row=3, column=0, padx=10, pady=20, sticky="e")

    pronoun_window.geometry("300x200")
    pronoun_window.resizable(False, False)
    
# Arquivos que armazenam os nomes
FEMININO_FILE = "dic_fem.txt"
MASCULINO_FILE = "dic_mas.txt"

# Listas globais para armazenar os nomes carregados dos arquivos
nomes_femininos = []
nomes_masculinos = []

# Função para carregar os nomes dos arquivos
def load_names():
    global nomes_femininos, nomes_masculinos
    
    if os.path.exists(FEMININO_FILE):
        with open(FEMININO_FILE, 'r', encoding='utf-8') as file:
            nomes_femininos = [name.strip().strip('\"') for name in file.read().split(", ")]

    if os.path.exists(MASCULINO_FILE):
        with open(MASCULINO_FILE, 'r', encoding='utf-8') as file:
            nomes_masculinos = [name.strip().strip('\"') for name in file.read().split(", ")]

# Função para salvar um nome no arquivo apropriado
def save_name_to_file(first_name, pronoun):
    if pronoun == "Masculino":
        if first_name not in nomes_masculinos:
            with open(MASCULINO_FILE, 'a', encoding='utf-8') as file:
                file.write(f'"{first_name}", ')
            nomes_masculinos.append(first_name)
    elif pronoun == "Feminino":
        if first_name not in nomes_femininos:
            with open(FEMININO_FILE, 'a', encoding='utf-8') as file:
                file.write(f'"{first_name}", ')
            nomes_femininos.append(first_name)

# Função para determinar a saudação baseada no nome
def determine_salutation(name):
    first_name = format_name(name)
    
    if first_name in nomes_femininos:
        return f"Prezada {first_name}"
    elif first_name in nomes_masculinos:
        return f"Prezado {first_name}"
    else:
        return f"Prezado(a) {first_name}"

# Função para salvar o pronome e adicionar o nome ao dicionário correspondente
def save_pronoun(pronoun_window, pronoun_var, first_name, on_close_callback):
    selected_pronoun = pronoun_var.get()

    if selected_pronoun == "Masculino":
        if first_name not in nomes_masculinos:
            save_name_to_file(first_name, "Masculino")
    elif selected_pronoun == "Feminino":
        if first_name not in nomes_femininos:
            save_name_to_file(first_name, "Feminino")

    pronoun_window.destroy()
    on_close_callback()

# Função para obter um pronome salvo
def get_saved_pronoun(name):
    config = configparser.ConfigParser()
    if os.path.exists(PRONOUNS_FILE):
        config.read(PRONOUNS_FILE)
        return config.get("PRONOUNS", name, fallback=None)
    return None

def define_pronoun_from_menu(treeview, selected_item):
    define_pronoun(treeview, selected_item, lambda: None)  # on_close_callback é uma função vazia aqui

def determinar_saudacao(nome_completo):
    primeiro_nome = nome_completo.split()[0]  # Pega o primeiro nome
    if primeiro_nome in nomes_femininos:
        return "Prezada"
    elif primeiro_nome in nomes_masculinos:
        return "Prezado"
    else:
        return "Prezado(a)"
    
# Função para determinar o gênero do hóspede com base no nome
def determine_gender(name):
    first_name = format_name(name)
    
    if first_name in nomes_femininos:
        return "Feminino"
    elif first_name in nomes_masculinos:
        return "Masculino"
    else:
        return "Não informado"
    
load_names()

# Função para carregar VIPs do arquivo
def load_vips_from_file():
    vip_list = []
    try:
        with open("vips.txt", "r") as f:
            for line in f:
                vip_name, _ = line.strip().split(',')
                vip_list.append(vip_name)
    except FileNotFoundError:
        pass
    return vip_list

# Função para aplicar o gradiente VIP ao carregar a lista
def update_vip_list_message():
    vip_list = load_vips_from_file()
    if vip_list:
        for frame, _, name, _ in frames:
            if name in vip_list:
                apply_vip_gradient(frame)

# Função para enviar mensagens de aniversário
def send_birthday_messages(birthdays):
    for name, age, dt_nascimento, phone in birthdays:
        if phone != "Telefone não encontrado":
            message = f"Prezado {name}, Nóslhe desejamos um feliz aniversário de {age} anos e muitos anos de saúde e felicidade! Estamos ansioso para lhe receber novamente em nosso hotel !!!🌴​🎉"
            # kit.sendwhatmsg_instantly(phone, message, wait_time=10)
    messagebox.showinfo("Envio Concluído", "Mensagens enviadas com sucesso!")
    
# Carregar lista de exceções do arquivo
exception_list = load_exceptions_from_file()

def processar_fila_de_envio():
    """
    Esta função roda em uma thread separada para enviar as mensagens
    sem congelar a interface do programa.
    """
    # Desativa o botão para evitar cliques duplos
    send_button.config(state=tk.DISABLED)
    
    contador_enviados = 0
    total_aniversariantes = len(birthdays)

    # Itera sobre a lista de aniversariantes que já foi carregada
    for nome, idade, dt_nascimento, telefone in birthdays:
        # Pula o envio se o nome estiver na lista de exceções
        if nome in exception_list:
            print(f"PULANDO: {nome} está na lista de exceções.")
            continue

        # Pula o envio se o telefone for inválido
        if telefone == "Telefone não encontrado" or not is_valid_number(telefone):
            print(f"PULANDO: {nome} com telefone inválido ou não encontrado.")
            continue

        # Determina a saudação correta (Prezado/Prezada)
        saudacao = determinar_saudacao(nome)
        primeiro_nome = format_name(nome)

        # Monta a mensagem
        mensagem = (
            f"{saudacao} {primeiro_nome},\n\n"
            f"Nós lhe desejamos um feliz aniversário e parabeniza pelos seus {idade} anos! 🎉\n\n"
            "Para comemorar esta data tão especial, preparamos uma oferta exclusiva. "
            "Venha celebrar conosco e desfrute de uma experiência inesquecível.\n\n"
            "Acesse nosso site e aproveite a melhor tarifa disponível: https://www.google.com 🎁"
        )

        try:
            print(f"Enviando mensagem para {nome} no número {telefone}...")
            # Envia a mensagem via WhatsApp
            kit.sendwhatmsg_instantly(telefone, mensagem, wait_time=15, tab_close=True, close_time=3)
            contador_enviados += 1
            
            # Pausa de 15 segundos entre os envios
            # Apenas se não for o último da lista
            if contador_enviados < total_aniversariantes:
                time.sleep(15)

        except Exception as e:
            print(f"Ocorreu um erro ao enviar para {nome}: {e}")
            messagebox.showwarning("Erro de Envio", f"Não foi possível enviar a mensagem para {nome}.\nVerifique sua conexão ou se o WhatsApp Web está funcionando corretamente.")
            continue # Continua para o próximo da lista

    # Exibe a mensagem de conclusão
    messagebox.showinfo("Envio Concluído", f"Envio em massa finalizado!\n\n{contador_enviados} de {total_aniversariantes} mensagens foram enviadas com sucesso.")
    
    # Reativa o botão após o término
    send_button.config(state=tk.NORMAL)

# Função para aplicar o modo escuro/claro
def apply_dark_mode(root, dark_mode=True):
    global light_bg, dark_bg, light_fg, dark_fg, light_button_bg, dark_button_bg, light_button_fg, dark_button_fg, light_frame_bg, dark_frame_bg
    
    # Cores do modo escuro
    dark_bg = "#1c1c1c"
    dark_fg = "#ffffff"
    dark_button_bg = "#333333"
    dark_button_fg = "#ffffff"
    dark_frame_bg = "#2e2e2e"

    # Cores do modo claro
    light_bg = "#f0f0f0"
    light_fg = "#000000"
    light_button_bg = "#4caf50"
    light_button_fg = "#ffffff"
    light_frame_bg = "#ffffff"

    # Alternar cores com base no modo
    bg_color = dark_bg if dark_mode else light_bg
    fg_color = dark_fg if dark_mode else light_fg
    button_bg = dark_button_bg if dark_mode else light_button_bg
    button_fg = dark_button_fg if dark_mode else light_button_fg
    frame_bg = dark_frame_bg if dark_mode else light_frame_bg

    # Alterar o fundo da janela principal
    root.configure(bg=bg_color)

    # Estilo para o frame de aniversariantes
    for widget in root.winfo_children():
        if isinstance(widget, tk.Frame) or isinstance(widget, tk.Label):
            widget.configure(bg=bg_color, fg=fg_color)
    
    # Estilo dos botões
    for button in root.winfo_children():
        if isinstance(button, tk.Button):
            button.configure(bg="#333333" if dark_mode else "#f0f0f0", fg=fg_color)

# Função para reiniciar o programa
def restart_program():
    python = sys.executable
    os.execl(python, python, *sys.argv)

# Função para alternar o modo escuro/claro
def toggle_dark_mode():
    global is_dark_mode
    if is_dark_mode:  # Modo escuro está ativado, vamos para o claro
        result = messagebox.askyesno("Reiniciar", "Para mudar de modo escuro para modo claro, uma reinicialização no programa é necessária. Deseja executar essa ação agora?")
        if result:
            is_dark_mode = False
            preferences['dark_mode'] = False
            save_preferences(preferences)
            restart_program()
        else:
            messagebox.showinfo("Modo Claro", "O modo claro será ativado na próxima vez em que o programa for reiniciado.")
            is_dark_mode = False
            preferences['dark_mode'] = False
            save_preferences(preferences)
    else:  # Modo claro está ativado, vamos para o escuro
        result = messagebox.askyesno("Reiniciar", "Para mudar de modo claro para modo escuro, uma reinicialização no programa é necessária. Deseja executar essa ação agora?")
        if result:
            is_dark_mode = True
            preferences['dark_mode'] = True
            save_preferences(preferences)
            restart_program()
        else:
            messagebox.showinfo("Modo Noturno", "O modo noturno será ativado na próxima vez em que o programa for reiniciado.")
            is_dark_mode = True
            preferences['dark_mode'] = True
            save_preferences(preferences)

# Função para iniciar o envio de mensagens
def start_sending():
    """
    Esta função é chamada pelo botão. Ela confirma com o usuário
    e inicia a função de envio em uma nova thread.
    """
    if not birthdays:
        messagebox.showinfo("Aviso", "Não há aniversariantes na lista para enviar mensagens.")
        return

    resposta = messagebox.askyesno("Confirmar Envio", 
                                   f"Você está prestes a iniciar o envio de {len(birthdays)} mensagem(ns) de aniversário.\n\n"
                                   "Este processo será executado em segundo plano e levará alguns minutos. Deseja continuar?")

    if resposta:
        # Inicia a função processar_fila_de_envio em uma nova thread
        threading.Thread(target=processar_fila_de_envio, daemon=True).start()
        messagebox.showinfo("Envio Iniciado", "O envio em segundo plano foi iniciado! Você pode continuar usando o programa.")

# Função para mostrar um pop-up com animação fade in e fade out
def show_popup(widget, text, x, y):
    popup = Toplevel()
    popup.wm_overrideredirect(True)
    popup.geometry(f"+{x + 10}+{y + 10}")
    label = tk.Label(popup, text=text, bg="white", relief="solid", bd=1)
    label.pack()
    popup.attributes("-alpha", 0.0)
    fade_in(popup)
    widget.popup = popup

def fade_in(window, alpha=0.0, increment=0.1):
    if alpha < 1.0:
        alpha += increment
        window.attributes("-alpha", alpha)
        window.after(50, fade_in, window, alpha)

def fade_out(window, alpha=1.0, decrement=0.1):
    if window.winfo_exists():
        if alpha > 0.0:
            alpha -= decrement
            window.attributes("-alpha", alpha)
            window.after(50, fade_out, window, alpha)
        else:
            window.destroy()

# Função para exibir o pop-up com um delay de 1 segundo
def delayed_popup(widget, text, event):
    x, y = event.x_root, event.y_root
    widget.popup_after_id = widget.after(1000, lambda: show_popup(widget, text, x, y))

# Função para cancelar o pop-up se o mouse sair antes do delay
def cancel_popup(widget):
    if hasattr(widget, 'popup_after_id') and widget.popup_after_id is not None:
        widget.after_cancel(widget.popup_after_id)
        widget.popup_after_id = None
    if hasattr(widget, 'popup') and widget.popup is not None:
        fade_out(widget.popup)

def show_add_exception_popup(widget, text, x, y):
    popup = Toplevel()
    popup.wm_overrideredirect(True)
    popup.geometry(f"+{x + 10}+{y + 10}")
    label = tk.Label(popup, text=text, bg="white", relief="solid", bd=1)
    label.pack()
    popup.attributes("-alpha", 0.0)
    fade_in(popup)
    widget.popup = popup

def add_exception_popup(event):
    x, y = event.x_root, event.y_root
    text = """Não quer mandar mensagem de aniversário para algum hóspede
em específico ? Essa função lhe permite criar uma lista de
exceção de hóspedes para qual não será enviado a
mensagem de parabenização, clique, adicione os hóspedes
desejados, e clique em concluir"""
    show_add_exception_popup(add_exception_button, text, x, y)

# Função para enviar mensagem individual
def send_individual_message(name, phone, age):
    saudacao = determine_salutation(name)
    message = (
        f"{saudacao} {name},\n\n"
        f"Nós lhe desejamos um feliz aniversário e parabeniza pelos seus {age} anos! 🎉\n\n"
        "Para comemorar esta data tão especial, preparamos uma oferta exclusiva e imperdível. "
        "Venha celebrar conosco e desfrute de uma experiência inesquecível.\n\n"
        "Reserve diretamente em nosso site e aproveite a melhor tarifa disponível. "
        "Estamos prontos para tornar sua celebração única e memorável. 🎁\n\n"
        "Acesse: https:/www.google.com"
    )
    kit.sendwhatmsg_instantly(phone, message, wait_time=10)
        
# Função para adicionar o aniversariante à lista de exceção
def add_to_exception_list(name):
    if name not in exception_list:
        exception_list.append(name)
        save_exceptions_to_file()  # Salvar exceções no arquivo
    update_exception_list_message()
    for frame, phone, frame_name, _ in frames:
        if frame.winfo_exists() and frame_name == name:
            frame.config(bg="#ffcccc")
            for widget in frame.winfo_children():
                widget.config(bg="#ffcccc")
                
# Função para remover o aniversariante da lista de exceção
def remove_from_exception_list(name):
    if name in exception_list:
        exception_list.remove(name)
        save_exceptions_to_file()  # Salvar exceções no arquivo
        update_exception_list_message()
        for frame, phone, frame_name, _ in frames:
            if frame.winfo_exists() and frame_name == name:
                frame.config(bg="#ffffff")
                for widget in frame.winfo_children():
                    widget.config(bg="#ffffff")
                    
# Função para limpar a lista de exceções
def clear_exception_list():
    global frames
    if exception_list:
        response = messagebox.askyesno("Limpar exceções", "Deseja retirar todos os hóspedes da lista de exceção?")
        if response:
            exception_list.clear()
            clear_exceptions_file()  # Limpar o conteúdo do arquivo de exceções
            update_exception_list_message()
            for frame, _, _, _ in frames:
                if frame.winfo_exists():
                    frame.config(bg="#ffffff")
                    for widget in frame.winfo_children():
                        widget.config(bg="#ffffff")
                    
def clear_exceptions_file():
    with open("exceções.txt", "w") as f:
        f.write("")  # Escreve uma string vazia para limpar o arquivo

# Função para enviar mensagem personalizada
def custom_message(name, phone, age):
    # Implementar a lógica para enviar mensagem personalizada aqui
    pass

# Função para enviar mensagem sem idade
def send_without_age(name, phone):
    # Implementar a lógica para enviar mensagem sem idade aqui
    pass

# Função para exibir informações do hóspede
def show_info(name, age, phone):
    # --- LÓGICA DE BUSCA DE DADOS (CORRIGIDA) ---
    cpf, last_stay, gender = "Não encontrado", "Não encontrado", "Não informado"

    if hospedes_df is not None and not hospedes_df.empty:
        # Busca o hóspede na base de dados (ignorando maiúsculas/minúsculas)
        guest_row = hospedes_df[hospedes_df['NOME'].str.upper() == name.upper()]
        
        if not guest_row.empty:
            # Pega o primeiro resultado encontrado
            guest_data = guest_row.iloc[0]
            cpf = guest_data.get('CPF', 'Não encontrado')
            last_stay = guest_data.get('ÚLTIMA HOSP.', 'Não encontrado')
            gender = guest_data.get('SEXO', 'Não informado')
            # Traduz M/F para Masculino/Feminino
            if gender == 'M':
                gender = 'Masculino'
            elif gender == 'F':
                gender = 'Feminino'

    # --- CRIAÇÃO DA JANELA E LAYOUT (IDÊNTICO AO SEU DESIGN ORIGINAL) ---
    info_window = tk.Toplevel()
    info_window.title("Informações do Hóspede")
    info_window.overrideredirect(True) # Remove bordas e barra de título

    # Cores e fontes para estilização
    bg_color = "#e0f7e0"
    fg_color = "#004d00"
    font_bold = ("Helvetica", 12, "bold")
    font_normal = ("Helvetica", 12)

    info_window.configure(bg=bg_color)
    info_window.geometry("400x300") # Ajuste o tamanho se necessário
    
    # Configura a coluna principal para se expandir e centralizar o conteúdo
    info_window.grid_columnconfigure(0, weight=1)

    # Título da janela
    title_label = tk.Label(info_window, text="Detalhes do Hóspede", font=("Helvetica", 16, "bold"), bg=bg_color, fg=fg_color)
    title_label.grid(row=0, column=0, pady=(20, 10))

    # Frame principal para organizar as informações em grid
    content_frame = tk.Frame(info_window, bg=bg_color)
    content_frame.grid(row=1, column=0, pady=5, padx=20, sticky="ew")

    # Lista de informações para exibir
    info_list = [
        ("Nome:", name),
        ("Idade:", f"{age} anos"),
        ("Telefone:", phone),
        ("Gênero:", gender),
        ("CPF:", cpf),
        ("Última Hospedagem:", last_stay)
    ]

    # Cria e posiciona os labels na grade
    for i, (label_text, value_text) in enumerate(info_list):
        label = tk.Label(content_frame, text=label_text, font=font_bold, bg=bg_color, fg=fg_color)
        label.grid(row=i, column=0, sticky="w", padx=5, pady=2)

        value = tk.Label(content_frame, text=value_text, font=font_normal, bg=bg_color, fg=fg_color)
        value.grid(row=i, column=1, sticky="w", padx=5, pady=2)

    # Botão para fechar a janela
    close_button = tk.Button(info_window, text="Fechar", command=info_window.destroy,
    width=15,  # <-- ADICIONE ISSO
    bg=fg_color, fg="white", font=("Helvetica", 10, "bold"), relief="flat")
    close_button.grid(row=2, column=0, pady=20)

    # Centralizando e permitindo mover a janela
    info_window.update_idletasks()
    window_width = info_window.winfo_width()
    window_height = info_window.winfo_height()
    screen_width = info_window.winfo_screenwidth()
    screen_height = info_window.winfo_screenheight()
    x = (screen_width // 2) - (window_width // 2)
    y = (screen_height // 2) - (window_height // 2)
    info_window.geometry(f"+{x}+{y}")
    info_window.grab_set()

    def move_window(event):
        info_window.geometry(f"+{event.x_root - 10}+{event.y_root - 10}") # Pequeno ajuste para o cursor

    info_window.bind("<B1-Motion>", move_window)

# Função para salvar a lista de exceção em um arquivo
def save_exceptions_to_file():
    with open("exceções.txt", "w") as f:
        for name in exception_list:
            f.write(f"{name}\n")

# Função para atualizar a label da lista de exceções
def update_exception_list_label():
    exception_count = len(exception_list)
    exception_list_label.config(text=f"Lista de Exceções: {exception_count} hóspede(s) na lista de exceção")

# Função para mostrar o menu de contexto ao clicar com o botão direito
def show_context_menu(event, name, phone, age, root):
    context_menu = Menu(root, tearoff=0)

    # Verifica se o nome já está na lista de exceção
    if name in exception_list:
        context_menu.add_command(label="Remover da lista de exceções", command=lambda: remove_from_exception_list(name))
    else:
        context_menu.add_command(label="Adicionar à lista de exceção", command=lambda: add_to_exception_list(name))

    context_menu.add_command(label="Enviar mensagem", command=lambda: send_individual_message(name, phone, age))
    context_menu.add_command(label="Enviar mensagem personalizada", command=lambda: custom_message(name, phone, age))
    context_menu.add_command(label="Enviar mensagem sem idade", command=lambda: send_without_age(name, phone))
    context_menu.add_command(label="Informações", command=lambda: show_info(name, age, phone))

    context_menu.tk_popup(event.x_root, event.y_root)

# Função para destacar entradas sem telefone
def highlight_missing_phone_entries(birthdays):
    for frame, phone, _, _ in frames:
        if phone == "Telefone não encontrado":
            frame.config(bg="#ffcccc")

# Função para atualizar a mensagem de aniversariantes sem telefone
def update_missing_phone_message(birthdays):
    missing_phone_count = sum(1 for _, _, _, phone in birthdays if phone == "Telefone não encontrado")
    if missing_phone_count > 0:
        plural = "hóspedes" if missing_phone_count > 1 else "hóspede"
        message = f"{missing_phone_count} {plural} está/estão sem o telefone cadastrado"
        missing_phone_label.config(text=message)
    else:
        missing_phone_label.config(text="")
        
# Função para criar a imagem do ícone
def create_image():
    # Caminho para o ícone
    image_path = r'C:\\Users\Lucas Siqueira\\Downloads\\aniversario atualizado 18-10\\principal.ico'
    return Image.open(image_path)

# Função para definir as opções padrão
def set_default_options():
    allow_background_execution.set(True)
    enable_notifications.set(True)
    default_send_without_age.set(False)
    remove_intro.set(False)
    disallow_manual_mode.set(False)
    block_manual_mode.set(False)
    show_guest_info_on_double_click.set(False)

# Função para abrir as configurações
def open_settings():
    global allow_background_execution, enable_notifications, default_send_without_age
    global remove_intro, disallow_manual_mode, block_manual_mode
    global block_calendar, send_time_warning, show_guest_info_on_double_click, filter_by_date

    settings_window = tk.Toplevel(root)
    settings_window.title("Configurações")
    settings_window.geometry("1150x650")
    settings_window.configure(bg=light_bg if not is_dark_mode else dark_bg)
    
    # Frame para as opções do lado esquerdo
    left_frame = ttk.Frame(settings_window)
    left_frame.pack(side=tk.LEFT, padx=20, pady=20, anchor='n')
    
    # Frame para as opções do lado direito
    right_frame = ttk.Frame(settings_window)
    right_frame.pack(side=tk.RIGHT, padx=20, pady=20, anchor='n')

    allow_background_execution = tk.BooleanVar(value=preferences.get("allow_background_execution", True))
    enable_notifications = tk.BooleanVar(value=preferences.get("enable_notifications", True))
    default_send_without_age = tk.BooleanVar(value=preferences.get("default_send_without_age", False))
    remove_intro = tk.BooleanVar(value=preferences.get("remove_intro", False))
    disallow_manual_mode = tk.BooleanVar(value=preferences.get("disallow_manual_mode", False))
    block_manual_mode = tk.BooleanVar(value=preferences.get("block_manual_mode", False))
    block_calendar = tk.BooleanVar(value=preferences.get("block_calendar", False))
    send_time_warning = tk.BooleanVar(value=preferences.get("send_time_warning", False))
    password_to_start = tk.BooleanVar(value=preferences.get("password_to_start", False))
    start_system = tk.BooleanVar(value=preferences.get("start_system", False))
    show_guest_info_on_double_click = tk.BooleanVar(value=preferences.get("show_guest_info_on_double_click", False))
    filter_by_date = tk.BooleanVar(value=preferences.get("filter_by_date", False))
    
    # Checkbuttons para as opções no lado esquerdo
    tk.Checkbutton(left_frame, text="Permitir execução em segundo plano", variable=allow_background_execution).pack(anchor=tk.W, pady=5)
    tk.Checkbutton(left_frame, text="Ativar notificações", variable=enable_notifications).pack(anchor=tk.W, pady=5)
    tk.Checkbutton(left_frame, text="Definir envio sem idade por padrão", variable=default_send_without_age).pack(anchor=tk.W, pady=5)
    tk.Checkbutton(left_frame, text="Retirar introdução", variable=remove_intro).pack(anchor=tk.W, pady=5)
    tk.Checkbutton(left_frame, text="Não permitir modo manual", variable=disallow_manual_mode).pack(anchor=tk.W, pady=5)
    tk.Checkbutton(left_frame, text="Bloquear modo manual", variable=block_manual_mode).pack(anchor=tk.W, pady=5)
    tk.Checkbutton(left_frame, text="Mostrar informações do hóspede ao clicar 2 vezes", variable=show_guest_info_on_double_click).pack(anchor=tk.W, pady=5)
    
    # Checkbuttons para as opções no lado direito
    tk.Checkbutton(right_frame, text="Bloquear calendário", variable=block_calendar).pack(anchor=tk.W, pady=5)
    tk.Checkbutton(right_frame, text="Aviso de horário de envio de mensagens", variable=send_time_warning).pack(anchor=tk.W, pady=5)
    tk.Checkbutton(right_frame, text="Pedir senha ao iniciar fora de horário", variable=password_to_start).pack(anchor=tk.W, pady=5)
    tk.Checkbutton(right_frame, text="Iniciar com o sistema", variable=start_system).pack(anchor=tk.W, pady=5)

    # Opção "Filtrar por data de hospedagem"
    tk.Checkbutton(right_frame, text="Filtrar por data de hospedagem", variable=filter_by_date, command=lambda: toggle_date_filter(date_frame, settings_window)).pack(anchor=tk.W, pady=5)

    # Frame para os campos de data (inicialmente escondido)
    date_frame = tk.Frame(right_frame)
    date_frame.pack(anchor=tk.W, pady=5)
    date_frame.pack_forget()

    # Campos de entrada de data com calendário
    tk.Label(date_frame, text="De:").pack(side=tk.LEFT)
    start_date_entry = DateEntry(date_frame, width=12, background='green', foreground='white', borderwidth=2, year=2024, 
                                 locale='pt_BR', date_pattern='dd/MM/yyyy', selectbackground='green')
    start_date_entry.pack(side=tk.LEFT, padx=5)

    tk.Label(date_frame, text="Até:").pack(side=tk.LEFT)
    end_date_entry = DateEntry(date_frame, width=12, background='green', foreground='white', borderwidth=2, year=2024, 
                               locale='pt_BR', date_pattern='dd/MM/yyyy', selectbackground='green')
    end_date_entry.pack(side=tk.LEFT, padx=5)

    # Botão "Aplicar filtro"
    apply_filter_button = tk.Button(date_frame, text="Aplicar filtro", command=apply_date_filter, font=("Helvetica", 12), bg="#4caf50", fg="white")
    apply_filter_button.pack(anchor=tk.W, pady=10)

    # Botão de modo noturno
    mode_text = "Alterar para modo claro" if is_dark_mode else "Alterar para modo noturno"
    dark_mode_button = tk.Button(settings_window, text=mode_text, command=toggle_dark_mode)
    dark_mode_button.pack(pady=10)

    # Função para mostrar/ocultar o frame de datas e redimensionar a janela
    def toggle_date_filter(frame, window):
        if filter_by_date.get():
            frame.pack(anchor=tk.W, pady=5)
            window.geometry("1280x650")
        else:
            frame.pack_forget()
            window.geometry("1150x650")

    # Frame para os botões na parte inferior
    button_frame = ttk.Frame(settings_window)
    button_frame.pack(side=tk.BOTTOM, pady=20, fill='x')

    # Centralizar os botões no frame usando grid
    close_button = tk.Button(button_frame, text="Fechar", command=settings_window.destroy, font=("Helvetica", 14), bg="#f44336", fg="white", width=15)
    close_button.grid(row=0, column=0, padx=10)

    restore_button = tk.Button(button_frame, text="Restaurar padrões", command=set_default_options, font=("Helvetica", 14), bg="#0000ff", fg="white", width=15)
    restore_button.grid(row=0, column=1, padx=10)

    save_button = tk.Button(button_frame, text="Salvar", command=lambda: save_options(settings_window), font=("Helvetica", 14), bg="#4caf50", fg="white", width=15)
    save_button.grid(row=0, column=2, padx=10)

    button_frame.grid_columnconfigure(0, weight=1)
    button_frame.grid_columnconfigure(1, weight=1)
    button_frame.grid_columnconfigure(2, weight=1)

def save_options(window):
    # Salvar as opções no arquivo de configurações
    preferences['allow_background_execution'] = allow_background_execution.get()
    preferences['enable_notifications'] = enable_notifications.get()
    preferences['default_send_without_age'] = default_send_without_age.get()
    preferences['remove_intro'] = remove_intro.get()
    preferences['disallow_manual_mode'] = disallow_manual_mode.get()
    preferences['block_manual_mode'] = block_manual_mode.get()
    preferences['block_calendar'] = block_calendar.get()
    preferences['send_time_warning'] = send_time_warning.get()
    preferences['password_to_start'] = password_to_start.get()
    preferences['start_system'] = start_system.get()
    preferences['show_guest_info_on_double_click'] = show_guest_info_on_double_click.get()
    preferences['filter_by_date'] = filter_by_date.get()

    save_preferences(preferences)
    window.destroy()

def apply_date_filter():
    start_date = start_date_entry.get_date()
    end_date = end_date_entry.get_date()
    print(f"Aplicar filtro: De {start_date} até {end_date}")

# Função para encerrar o ícone da barra de tarefas e o aplicativo
def on_quit(icon, item):
    icon.visible = False  # Torna o ícone invisível antes de fechar
    root.destroy()  # Fecha a janela principal do Tkinter
    os._exit(0)  # Força o encerramento imediato do programa e de todos os processos

# Configuração inicial do ícone da barra de tarefas com as opções "Fechar" e "Configurações"
def setup(icon):
    icon.menu = pystray.Menu(
        item('Configurações', open_settings),  # Adiciona a opção "Configurações"
        item('Desativar notificações', on_quit),
        item('Iniciar envio agora', on_quit),
        item('Fechar', on_quit)  # Altera o texto da opção "Quit" para "Fechar"
    )
    icon.visible = True
    
def show_search_results(search_results, query):
    """
    Cria uma nova janela para exibir os resultados paginados de uma busca.
    Visualmente similar à tela de aniversariantes do calendário.
    """
    if not search_results:
        messagebox.showinfo("Busca", f"Nenhum resultado encontrado para '{query}'.")
        return

    # Criar uma nova janela para exibir os resultados
    results_window = tk.Toplevel(root)
    results_window.title("Resultados da Busca")
    results_window.configure(bg="#f0f0f0")
    results_window.geometry("600x650") # Um pouco maior para os resultados

    items_per_page = 10
    max_pages = (len(search_results) + items_per_page - 1) // items_per_page
    frames = []

    content_frame = tk.Frame(results_window, bg="#f0f0f0")
    content_frame.pack(pady=5, padx=10, fill='both', expand=True)

    def display_page(page):
        # Limpar o content_frame
        for widget in content_frame.winfo_children():
            widget.destroy()
            
        start_index = page * items_per_page
        end_index = min(start_index + items_per_page, len(search_results))

        for name, age, dt_nascimento, phone in search_results[start_index:end_index]:
            frame = tk.Frame(content_frame, bg="#ffffff", relief="flat", bd=0)
            frame.pack(pady=5, padx=10, fill='x')
            frames.append((frame, phone, name, dt_nascimento))

            name_label = tk.Label(frame, text=name, bg="#ffffff", font=("Helvetica", 14), bd=0, relief="flat")
            name_label.pack(side="left", padx=10)

            age_label = tk.Label(frame, text=f"{age} anos", bg="#ffffff", font=("Helvetica", 14), bd=0, relief="flat")
            age_label.pack(side="right", padx=10)

            # Adicionar menu de contexto ao clicar com o botão direito
            name_label.bind("<Button-3>", lambda e, n=name, p=phone, a=age: show_context_menu(e, n, p, a, results_window))

    def next_page():
        nonlocal current_page
        if current_page < max_pages - 1:
            current_page += 1
            display_page(current_page)
        page_label.config(text=f"Página {current_page + 1} de {max_pages}")

    def previous_page():
        nonlocal current_page
        if current_page > 0:
            current_page -= 1
            display_page(current_page)
        page_label.config(text=f"Página {current_page + 1} de {max_pages}")

    current_page = 0

    # Frame fixo para os botões de navegação
    nav_frame = tk.Frame(results_window, bg="#f0f0f0")
    nav_frame.pack(pady=5)
    
    prev_button = tk.Button(nav_frame, text="←", command=previous_page, font=("Helvetica", 16), bg="#4caf50", fg="white")
    prev_button.pack(side="left", padx=10)

    page_label = tk.Label(nav_frame, text=f"Página 1 de {max_pages}", bg="#f0f0f0", font=("Helvetica", 14))
    page_label.pack(side="left")

    next_button = tk.Button(nav_frame, text="→", command=next_page, font=("Helvetica", 16), bg="#4caf50", fg="white")
    next_button.pack(side="left", padx=10)

    display_page(current_page)

    # Texto informativo na parte inferior
    info_label_text = f"Mostrando {len(search_results)} resultado(s) para a busca."
    info_label = tk.Label(results_window, text=info_label_text, bg="#f0f0f0", fg="green", font=("Helvetica", 14))
    info_label.pack(side="bottom", pady=5)
    
def show_birthdays_for_selected_date(selected_birthdays, selected_date, root):
    # Criar uma nova janela para exibir os aniversariantes
    new_window = tk.Toplevel(root)
    new_window.title(f"Aniversariantes em {selected_date}")
    new_window.configure(bg="#f0f0f0")

    items_per_page = 10
    max_pages = (len(selected_birthdays) + items_per_page - 1) // items_per_page
    frames = []  # Armazena referências para os frames e informações relacionadas

    content_frame = tk.Frame(new_window, bg="#f0f0f0")
    content_frame.pack(pady=5, padx=10, fill='both', expand=True)

    def display_page(page):
        # Limpar o content_frame
        for widget in content_frame.winfo_children():
            widget.destroy()
        start_index = page * items_per_page
        end_index = min(start_index + items_per_page, len(selected_birthdays))
        for name, age, dt_nascimento, phone in selected_birthdays[start_index:end_index]:
            frame = tk.Frame(content_frame, bg="#ffffff" if phone != "Telefone não encontrado" else "#ffcccc", relief="flat", bd=0)
            frame.pack(pady=5, padx=10, fill='x')
            frames.append((frame, phone, name, dt_nascimento))

            name_label = tk.Label(frame, text=name, bg="#ffffff" if phone != "Telefone não encontrado" else "#ffcccc", font=("Helvetica", 14), bd=0, relief="flat")
            name_label.pack(side="left", padx=10)

            age_label = tk.Label(frame, text=f"{age} anos", bg="#ffffff" if phone != "Telefone não encontrado" else "#ffcccc", font=("Helvetica", 14), bd=0, relief="flat")
            age_label.pack(side="right", padx=10)

            # Adicionar menu de contexto ao nome
            name_label.bind("<Button-3>", lambda e, name=name, phone=phone, age=age, frame=frame: show_context_menu(e, name, phone, age, frame))

            # Atualizar cor de fundo se o nome estiver na lista de exceções
            if name in exception_list:
                frame.config(bg="#ffcccc")
                for widget in frame.winfo_children():
                    widget.config(bg="#ffcccc")

        page_label.config(text=f"Página {page + 1} de {max_pages}")

    def next_page():
        nonlocal current_page
        if current_page < max_pages - 1:
            current_page += 1
            display_page(current_page)

    def previous_page():
        nonlocal current_page
        if current_page > 0:
            current_page -= 1
            display_page(current_page)

    def show_context_menu(event, name, phone, age, frame):
        menu = tk.Menu(new_window, tearoff=0)
        
        # Adicionar à lista de exceção ou remover se já estiver na lista
        if name in exception_list:
            menu.add_command(label="Remover da lista de exceções", command=lambda: remove_from_exception_list(name, frame))
        else:
            menu.add_command(label="Adicionar à lista de exceção", command=lambda: add_to_exception_list(name, frame))
        
        # Enviar mensagem sem idade
        menu.add_command(label="Enviar mensagem sem idade", command=lambda: send_without_age(name, phone))
        
        # Definir pronome
        menu.add_command(label="Definir pronome", command=lambda: define_pronoun_from_menu(name, age, phone))
        
        # Ativar notificação
        menu.add_command(label="Notificar-me", command=lambda: mark_notification(name, age, phone))
        
        # Adicionar a VIP
        menu.add_command(label="Marcar como VIP", command=lambda: define_as_vip(name, age, phone))
        
        # Informações
        menu.add_command(label="Informações", command=lambda: show_info(name, age, phone))

        menu.tk_popup(event.x_root, event.y_root)
        
    def mark_notification(name, age, phone):
        print ('A função de notificação será implementada aqui')
        
    # Função para adicionar um hóspede à lista de VIPs e salvar no arquivo
    def define_as_vip(name, age, phone):
        # Criar uma nova janela de aviso
        vip_window = tk.Toplevel(root)
        vip_window.title("Confirmação de VIP")
        vip_window.configure(bg="#2c3e50", padx=20, pady=20)

        # Estilizando a janela e criando um efeito de fade-in
        vip_window.attributes("-alpha", 0.0)
        fade_in(vip_window)

        title_font = ("Helvetica", 14, "bold")
        label_font = ("Helvetica", 12)
        button_font = ("Helvetica", 12, "bold")

        # Configurar a largura da coluna para garantir que os botões estejam bem posicionados
        vip_window.grid_columnconfigure(0, weight=1)

        # Título da janela
        tk.Label(vip_window, text="Adicionar à lista VIP", bg="#2c3e50", fg="white", font=title_font).grid(row=0, column=0, padx=10, pady=(0, 10), sticky="w")

        # Mensagem de confirmação
        message = f"Deseja mesmo adicionar o hóspede {name} na lista VIP?"
        tk.Label(vip_window, text=message, bg="#2c3e50", fg="white", font=label_font, wraplength=300, justify="left").grid(row=1, column=0, padx=10, pady=(0, 20), sticky="w")

        # Frame para os botões
        button_frame = tk.Frame(vip_window, bg="#2c3e50")
        button_frame.grid(row=2, column=0, pady=10)

        yes_button = tk.Button(button_frame, text="Sim", command=lambda: add_vip(name, phone), font=button_font, bg="#27ae60", fg="white", relief="flat", padx=10, pady=5)
        yes_button.pack(side="left", padx=5)

        no_button = tk.Button(button_frame, text="Não", command=lambda: tv_off_animation(vip_window), font=button_font, bg="#c0392b", fg="white", relief="flat", padx=10, pady=5)
        no_button.pack(side="left", padx=5)

        vip_window.mainloop()

    # Função para aplicar o gradiente VIP e atualizar o menu contextual
    def add_vip(name, phone):
        with open("vips.txt", "a") as f:
            f.write(f"{name},{phone}\n")

        for frame, _, frame_name, _ in frames:
            if frame.winfo_exists() and frame_name == name:
                apply_vip_gradient(frame)

        update_vip_list_message()
        messagebox.showinfo("VIP", f"{name} foi adicionado à lista VIP!")

    # Animação de TV desligando mais rápida
    def tv_off_animation(window):
        def shrink():
            nonlocal scale
            scale -= 0.1  # Aumentar a velocidade da animação
            if scale > 0:
                window.geometry(f"{int(400 * scale)}x{int(300 * scale)}")
                window.attributes("-alpha", scale)
                window.after(30, shrink)  # Reduzir o intervalo para aumentar a velocidade
            else:
                window.destroy()

        scale = 1.0
        shrink()

    # Função para aplicar o gradiente VIP no frame
    def apply_vip_gradient(frame):
        frame_name = frame.winfo_children()[0].cget("text")

        # Limpar o frame atual e criar um Canvas para substituí-lo
        for widget in frame.winfo_children():
            widget.destroy()

        width = frame.winfo_width()
        height = frame.winfo_height()

        # Criar um Canvas para desenhar o gradiente
        vip_canvas = tk.Canvas(frame, width=width, height=height, highlightthickness=0)
        vip_canvas.pack(fill="both", expand=True)

        # Gradiente de laranja para roxo
        for i in range(width):
            r = 255 - int(i * 127 / width)
            g = 165 - int(i * 165 / width)
            b = 0 + int(i * 255 / width)
            color = f'#{r:02x}{g:02x}{b:02x}'
            vip_canvas.create_line(i, 0, i, height, fill=color)

        # Adicionar o nome do hóspede sobre o gradiente
        vip_canvas.create_text(width//2, height//2, text=frame_name, fill="white", font=("Helvetica", 14, "bold"))

    def define_pronoun_from_menu(name, phone, age):
        first_name = format_name(name)
        pronoun_window = tk.Toplevel(new_window)
        pronoun_window.title("Definir Pronome")
        pronoun_window.grab_set()  # Impede a interação com outras janelas até que essa seja fechada
        pronoun_window.configure(bg="#e0e0e0", padx=20, pady=20)

        title_font = ("Helvetica", 14, "bold")
        label_font = ("Helvetica", 12)
        button_font = ("Helvetica", 12, "bold")

        tk.Label(pronoun_window, text="Selecione o pronome", bg="#e0e0e0", font=title_font).grid(row=0, column=0, padx=10, pady=(0, 10), sticky="w")
        tk.Label(pronoun_window, text=f"Hóspede: {first_name}", bg="#e0e0e0", font=label_font).grid(row=1, column=0, padx=10, pady=(0, 10), sticky="w")

        pronoun_var = tk.StringVar(pronoun_window)

        if first_name in nomes_masculinos:
            pronoun_var.set("Masculino")
        elif first_name in nomes_femininos:
            pronoun_var.set("Feminino")
        else:
            pronoun_var.set("Não informado")

        pronoun_options = ["Masculino", "Feminino", "Não informado"]
        pronoun_menu = tk.OptionMenu(pronoun_window, pronoun_var, *pronoun_options)
        pronoun_menu.config(font=label_font, bg="#f7f7f7", fg="#333333", highlightthickness=1, highlightbackground="#b0b0b0")
        pronoun_menu.grid(row=2, column=0, padx=10, pady=10, sticky="ew")

        def save_pronoun():
            selected_pronoun = pronoun_var.get()
            save_name_to_file(first_name, selected_pronoun)
            pronoun_window.destroy()

        save_button = tk.Button(pronoun_window, text="Salvar", command=save_pronoun, font=button_font, bg="#4caf50", fg="white", relief="flat", padx=10, pady=5)
        save_button.grid(row=3, column=0, padx=10, pady=20, sticky="e")

        pronoun_window.geometry("300x200")
        pronoun_window.resizable(False, False)
                                
    def send_without_age(name, phone):
        print(f"Enviar sem idade: {name}, {phone}")
        # Lógica para enviar mensagem sem idade

    def show_info(name, age, phone):
        print(f"Mostrando informações para {name}.")
        # Lógica para mostrar informações adicionais

    current_page = 0

    # Frame fixo para os botões de navegação
    nav_frame = tk.Frame(new_window, bg="#f0f0f0")
    nav_frame.pack(pady=5)

    prev_button = tk.Button(nav_frame, text="←", command=previous_page, font=("Helvetica", 16), bg="#4caf50", fg="white")
    prev_button.pack(side="left", padx=10)

    page_label = tk.Label(nav_frame, text="", bg="#f0f0f0", font=("Helvetica", 14))
    page_label.pack(side="left")

    next_button = tk.Button(nav_frame, text="→", command=next_page, font=("Helvetica", 16), bg="#4caf50", fg="white")
    next_button.pack(side="left", padx=10)

    display_page(current_page)

    # Texto informativo na parte inferior central
    total_birthdays = len(selected_birthdays)
    info_label_text = f"Existem {total_birthdays} aniversariantes em {selected_date}!"
    info_label = tk.Label(new_window, text=info_label_text, bg="#f0f0f0", fg="green", font=("Helvetica", 14))
    info_label.pack(side="bottom", pady=5)
    
sending_completed = False
    
def show_main_screen():
    global missing_phone_label, frames, add_exception_button, is_adding_exception, exception_list_label, birthdays, page_label
    global current_page, max_pages, root  # Certifique-se de declarar root como global
    global hospedes_df
    global send_button
    global root
    
    # Verificar se o horário atual está dentro do intervalo de 08:00 às 17:00
    current_time = datetime.now().time()
    start_time = datetime.strptime("08:00", "%H:%M").time()
    end_time = datetime.strptime("17:00", "%H:%M").time()

    if not (start_time <= current_time <= end_time):
        proceed = messagebox.askyesno(
            "Aviso de Horário",
            "O horário de envio de mensagens de feliz aniversário para os hóspedes deve ser entre 08:00h e 17:00h. Você está inicializando o mensageiro fora desse horário. Deseja mesmo prosseguir?",
            icon='warning'  # Isso altera o ícone para o de alerta (exclamação)
        )
        if not proceed:
            return

    current_page = 0
    items_per_page = 10

    # Tela principal
    root = tk.Tk()
    root.title("Mensageiro de Aniversariantes")
    root.resizable(False, False)
    root.configure(bg="#f0f0f0")
    
    apply_dark_mode(root, dark_mode=is_dark_mode)

    root.iconbitmap('C:\\Users\\Lucas Siqueira\\Downloads\\aniversario atualizado 18-10\\principal.ico')
    
    # Carregar a última posição da janela
    load_window_position(root)

    # Configurar o evento de fechamento para salvar a posição da janela
    root.protocol("WM_DELETE_WINDOW", lambda: [save_window_position(root), root.destroy()])

    # Adiciona o ícone à barra de tarefas
    icon_image = create_image()
    menu = (item('Quit', on_quit),)
    icon = pystray.Icon("Mensageiro de Aniversariantes", icon_image, "Mensageiro de Aniversariantes", menu)
    threading.Thread(target=icon.run, args=(setup,)).start()

    # Inicializa a variável de exceções
    is_adding_exception = False

    # Carregar lista de exceções do arquivo
    global exception_list
    exception_list = load_exceptions_from_file()

    # Carrega o DataFrame de telefones
    hospedes_df = carregar_base_de_dados("hóspedes.csv")
    if hospedes_df is None:
        root.destroy() # Se não encontrar o arquivo, fecha o programa
        return

    # Exibe os aniversariantes
    birthdays = encontrar_aniversariantes(hospedes_df)
    frames = []
    max_pages = (len(birthdays) + items_per_page - 1) // items_per_page  # Calcular número máximo de páginas

    # Create a single frame for the page content
    content_frame = tk.Frame(root, bg="#f0f0f0")
    content_frame.pack(pady=5, padx=10, fill='both', expand=True)

    def display_page(page):
        # Clear the content frame
        for widget in content_frame.winfo_children():
            widget.destroy()
        start_index = page * items_per_page
        end_index = min(start_index + items_per_page, len(birthdays))
        for name, age, dt_nascimento, phone in birthdays[start_index:end_index]:
            frame = tk.Frame(content_frame, bg="#ffffff" if phone != "Telefone não encontrado" else "#ffcccc", relief="flat", bd=0)
            frame.pack(pady=5, padx=10, fill='x')
            frames.append((frame, phone, name, dt_nascimento))

            name_label = tk.Label(frame, text=name, bg="#ffffff" if phone != "Telefone não encontrado" else "#ffcccc", font=("Helvetica", 14), bd=0, relief="flat")
            name_label.pack(side="left", padx=10)
            name_label.bind("<Enter>", lambda e, name=name, phone=phone: delayed_popup(name_label, f"Telefone: {phone}", e))
            name_label.bind("<Leave>", lambda e: cancel_popup(name_label))
            name_label.bind("<Button-1>", lambda e, name=name: toggle_exception(e, name))

            # Adiciona o menu de contexto para o botão direito do mouse
            name_label.bind("<Button-3>", lambda e, name=name, phone=phone, age=age: show_context_menu(e, name, phone, age, root))

            age_label = tk.Label(frame, text=f"{age} anos", bg="#ffffff" if phone != "Telefone não encontrado" else "#ffcccc", font=("Helvetica", 14), bd=0, relief="flat")
            age_label.pack(side="right", padx=10)
            age_label.bind("<Enter> ", lambda e, age=age, dt_nascimento=dt_nascimento: delayed_popup(age_label, f"Data de Nascimento: {dt_nascimento}", e))
            age_label.bind("<Leave>", lambda e: cancel_popup(age_label))

            # Atualizar cor de fundo se o nome estiver na lista de exceções
            if name in exception_list:
                frame.config(bg="#ffcccc")
                for widget in frame.winfo_children():
                    widget.config(bg="#ffcccc")

        page_label.config(text=f"Página {current_page + 1} de {max_pages}")

    def next_page():
        global current_page
        if current_page < max_pages - 1:
            current_page += 1
            display_page(current_page)

    def previous_page():
        global current_page
        if current_page > 0:
            current_page -= 1
            display_page(current_page)
            
    # Definir calendar_window como uma variável global
    calendar_window = None

    def open_calendar():
        global calendar_window  # Garantir que estamos utilizando a variável global

        # Verifica se o calendário já está aberto
        try:
            if calendar_window is not None and calendar_window.winfo_exists():
                fade_out(calendar_window)
                return
        except NameError:
            calendar_window = None

        # Criar uma nova janela para o calendário sem bordas
        calendar_window = tk.Toplevel(root)
        calendar_window.overrideredirect(True)  # Remove a barra de título e bordas
        calendar_window.configure(bg="#f0f0f0")

        # Posicionar a janela do calendário à direita da janela principal
        x = root.winfo_x() + root.winfo_width()
        y = root.winfo_y()
        calendar_window.geometry(f"+{x}+{y}")

        # Criar o widget de calendário com a configuração de idioma para pt-BR
        cal = Calendar(calendar_window, selectmode='day', year=datetime.now().year, 
                       month=datetime.now().month, day=datetime.now().day, 
                       date_pattern='dd/mm/yyyy', background='#4caf50', 
                       foreground='white', selectbackground='#4caf50', 
                       font=("Helvetica", 12), locale='pt_BR')  # Definir o idioma como pt-BR
        cal.pack(pady=10, padx=10)

        # Efeito de fade-in no calendário
        calendar_window.attributes("-alpha", 0.0)
        fade_in(calendar_window)

        # Atualizar a função confirm_date para abrir a janela de aniversariantes ao confirmar a data
        def confirm_date():
            selected_date = cal.get_date()
            # Converter a data selecionada para o formato esperado (%d/%m)
            formatted_date = datetime.strptime(selected_date, "%d/%m/%Y").strftime("%d/%m")
            selected_birthdays = encontrar_aniversariantes(hospedes_df, specific_date=formatted_date)
            show_birthdays_for_selected_date(selected_birthdays, selected_date, root)
            fade_out(calendar_window)

        confirm_button = tk.Button(calendar_window, text="Confirmar", 
                                   command=confirm_date, font=("Helvetica", 12), 
                                   bg="#4caf50", fg="white")
        confirm_button.pack(pady=5)

    def fade_in(window, alpha=0.0, increment=0.05):
        alpha += increment
        if alpha <= 1.0:
            window.attributes("-alpha", alpha)
            window.after(30, fade_in, window, alpha)

    def fade_out(window, alpha=1.0, decrement=0.05):
        alpha -= decrement
        if alpha > 0.0:
            window.attributes("-alpha", alpha)
            window.after(30, fade_out, window, alpha)
        else:
            window.destroy()

    # Frame fixo para os botões de navegação
    nav_frame = tk.Frame(root, bg="#f0f0f0")
    nav_frame.pack(pady=5)

    prev_button = tk.Button(nav_frame, text="←", command=previous_page, font=("Helvetica", 16), bg="#4caf50", fg="white")
    prev_button.pack(side="left", padx=10)

    page_label = tk.Label(nav_frame, text="", bg="#f0f0f0", font=("Helvetica", 14))
    page_label.pack(side="left")

    next_button = tk.Button(nav_frame, text="→", command=next_page, font=("Helvetica", 16), bg="#4caf50", fg="white")
    next_button.pack(side="left", padx=10)

    display_page(current_page)

    # Frame fixo para os botões de ação
    action_frame = tk.Frame(root, bg="#f0f0f0")
    action_frame.pack(pady=5)

    # Botão para iniciar o envio de mensagens
    send_button = tk.Button(action_frame, text="Começar envio", command=start_sending, font=("Helvetica", 16), bg="#4caf50", fg="white")
    send_button.pack(side="left", padx=10)

    # Botão para adicionar exceção
    add_exception_button = tk.Button(action_frame, text="Adicionar exceção", command=toggle_exception_mode, font=("Helvetica", 16), bg="#f44336", fg="white")
    add_exception_button.pack(side="left", padx=10)

    # Novo botão "Modo manual"
    manual_mode_button = tk.Button(action_frame, text="Modo manual", command=manual_mode, font=("Helvetica", 16), bg="#ffa500", fg="white")
    manual_mode_button.pack(side="left", padx=10)

    # Frame adicional para colocar o botão "Pesquisar hóspede" e o botão "Calendário"
    search_calendar_frame = tk.Frame(root, bg="#f0f0f0")
    search_calendar_frame.pack(pady=5, anchor="w")

    # Botão "Pesquisar hóspede"
    search_button = tk.Button(search_calendar_frame, text="Pesquisar hóspede", command=search_guest, font=("Helvetica", 16), bg="#0000ff", fg="white")
    search_button.pack(side="left", padx=10)

    # Botão "Calendário"
    calendar_button = tk.Button(search_calendar_frame, text="Calendário", command=open_calendar, font=("Helvetica", 16), bg="#dda0dd", fg="white")
    calendar_button.pack(side="left", padx=10)
    
    # Botão "Configurações"
    settings_button = tk.Button(search_calendar_frame, text="Configurações", command=open_settings, font=("Helvetica", 16), bg="#808080", fg="white")
    settings_button.pack(side="left", padx=10)

    # Defina os bindings aqui
    add_exception_button.bind("<Enter>", lambda e: delayed_popup(add_exception_button, """Não quer mandar mensagem de aniversário para algum hóspede
em específico ? Essa função lhe permite criar uma lista de
exceção de hóspedes para qual não será enviado a
mensagem de parabenização, clique, adicione os hóspedes
desejados, e clique em concluir""", e))
    add_exception_button.bind("<Leave>", lambda e: cancel_popup(add_exception_button))
    add_exception_button.bind("<Button-1>", add_exception_popup)
    send_button.bind("<Enter>", lambda e: delayed_popup(send_button, "Começa o envio das mensagens de parabenização via WhatsApp Web", e))
    send_button.bind("<Leave>", lambda e: cancel_popup(send_button))
    prev_button.bind("<Enter>", lambda e: delayed_popup(prev_button, "Página anterior", e))
    prev_button.bind("<Leave>", lambda e: cancel_popup(prev_button))
    next_button.bind("<Enter>", lambda e: delayed_popup(next_button, "Próxima página", e))
    next_button.bind("<Leave>", lambda e: cancel_popup(next_button))

    # Label para mensagens de aniversarantes sem telefone
    missing_phone_label = tk.Label(root, text="", bg="#f0f0f0", fg="red", font=("Helvetica", 12))
    missing_phone_label.pack(pady=5)

    # Label para lista de exceções
    exception_list_label = tk.Label(root, text="", bg="#f0f0f0", fg="red", font=("Helvetica", 12))
    exception_list_label.pack(pady=5)
    exception_list_label.bind("<Button-1>", lambda event: clear_exception_list())

    update_missing_phone_message(birthdays)
    update_exception_list_message()

    # Texto informativo na parte inferior central
    total_birthdays = len(birthdays)
    info_label_text = f"Hoje temos {total_birthdays} hóspedes que fazem aniversário!!!"
    info_label = tk.Label(root, text=info_label_text, bg="#f0f0f0", fg="green", font=("Helvetica", 14))
    info_label.pack(side="bottom", pady=5)
    
    show_birthday_notification(total_birthdays)

    def blink():
        current_color = info_label.cget("fg")
        next_color = "green" if current_color == "#f0f0f0" else "#f0f0f0"
        info_label.config(fg=next_color)
        root.after(500, blink)  # Chama a função novamente após 500ms

    blink()  # Inicia o efeito de piscar

    # Carrega a imagem do ícone de ajuda
    help_icon = Image.open("ajuda.png")
    help_icon = help_icon.resize((50, 50), Image.Resampling.LANCZOS)
    help_icon_photo = ImageTk.PhotoImage(help_icon)

    # Função para abrir o arquivo help.html
    def open_help():
        webbrowser.open("help.html")

    # Adiciona o ícone de ajuda no canto inferior esquerdo
    left_help_button = tk.Button(root, image=help_icon_photo, command=open_help, bd=0, bg="#f0f0f0", activebackground="#f0f0f0")
    left_help_button.place(relx=0.01, rely=0.95, anchor="sw")

    # Adiciona o ícone de ajuda no canto inferior direito
    right_help_button = tk.Button(root, image=help_icon_photo, command=open_help, bd=0, bg="#f0f0f0", activebackground="#f0f0f0")
    right_help_button.place(relx=0.99, rely=0.95, anchor="se")
    
    # --- CÓDIGO PARA CENTRALIZAR A JANELA PRINCIPAL ---
    root.update_idletasks() # Garante que o tamanho da janela foi calculado
    window_width = root.winfo_width()
    window_height = root.winfo_height()
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    x = (screen_width // 2) - (window_width // 2)
    y = (screen_height // 2) - (window_height // 2)
    root.geometry(f'{window_width}x{window_height}+{x}+{y}')

    root.mainloop()

# Adicione o contexto menu
def show_context_menu(event, name, phone, age, root):
    menu = Menu(root, tearoff=0)
    menu.add_command(label="Enviar mensagem", command=lambda: send_individual_message(name, phone, age))
    # Verifica se o nome já está na lista de exceção
    if name in exception_list:
        menu.add_command(label="Remover da lista de exceções", command=lambda: remove_from_exception_list(name))
    else:
        menu.add_command(label="Adicionar à lista de exceção", command=lambda: add_to_exception_list(name))
    menu.add_command(label="Enviar mensagem personalizada", command=lambda: custom_message(name, phone, age))
    menu.add_command(label="Enviar sem idade", command=lambda: send_without_age(name, phone))
    menu.add_command(label="Informações", command=lambda: show_info(name, age, phone))
    menu.tk_popup(event.x_root, event.y_root)

def valida_cpf(cpf):
    pattern = re.compile(r'^\d{3}\.\d{3}\.\d{3}-\d{2}$')
    if not pattern.match(cpf):
        return False

    num1 = int(cpf[0])
    num2 = int(cpf[1])
    num3 = int(cpf[2])
    num4 = int(cpf[4])
    num5 = int(cpf[5])
    num6 = int(cpf[6])
    num7 = int(cpf[8])
    num8 = int(cpf[9])
    num9 = int(cpf[10])
    num10 = int(cpf[12])
    num11 = int(cpf[13])

    if (num1 == num2 == num3 == num4 == num5 == num6 == num7 == num8 == num9 == num10 == num11):
        return False

    soma1 = num1 * 10 + num2 * 9 + num3 * 8 + num4 * 7 + num5 * 6 + num6 * 5 + num7 * 4 + num8 * 3 + num9 * 2
    resto1 = (soma1 * 10) % 11
    if resto1 == 10:
        resto1 = 0

    soma2 = num1 * 11 + num2 * 10 + num3 * 9 + num4 * 8 + num5 * 7 + num6 * 6 + num7 * 5 + num8 * 4 + num9 * 3 + num10 * 2
    resto2 = (soma2 * 10) % 11
    if resto2 == 10:
        resto2 = 0

    return resto1 == num10 and resto2 == num11

def validate_telefone(telefone):
    pattern = re.compile(r'^\+55 \(\d{2}\) \d{5}-\d{4}$')
    return bool(pattern.match(telefone))

def validate_data_nascimento(dob):
    pattern = re.compile(r'^\d{2}/\d{2}/\d{4}$')
    if not pattern.match(dob):
        return False
    
    dia, mes, ano = map(int, dob.split('/'))
    
    if not (1900 <= ano <= 2100):
        return False
    
    if not (1 <= mes <= 12):
        return False
    
    dias_por_mes = [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]
    
    if (ano % 4 == 0 and ano % 100 != 0) or (ano % 400 == 0):
        dias_por_mes[1] = 29
    
    if not (1 <= dia <= dias_por_mes[mes - 1]):
        return False
    
    return True

def search_guest():
    search_window = tk.Toplevel()
    search_window.title("Pesquisar hóspede")
    search_window.geometry("400x250")
    search_window.configure(bg="#f0f0f0")

    tk.Label(search_window, text="Pesquisar por:", bg="#f0f0f0", font=("Helvetica", 12)).pack(pady=10)

    search_option = tk.StringVar(value="Nome")
    options = ["Nome", "CPF", "Telefone", "Data de nascimento"]
    search_menu = ttk.Combobox(search_window, textvariable=search_option, values=options, state="readonly", font=("Helvetica", 12))
    search_menu.pack(pady=5)

    search_entry = tk.Entry(search_window, font=("Helvetica", 12))

    cpf_frame = tk.Frame(search_window, bg="#f0f0f0")
    cpf_entries = [tk.Entry(cpf_frame, width=3, font=("Helvetica", 12), justify='center') for _ in range(4)]
    for i, entry in enumerate(cpf_entries):
        entry.pack(side="left")
        if i < 3:
            separator = tk.Label(cpf_frame, text="." if i < 2 else "-", bg="#f0f0f0", font=("Helvetica", 12))
            separator.pack(side="left")

    telefone_frame = tk.Frame(search_window, bg="#f0f0f0")

    # Gera a lista de DDI + Código ISO (ex: "+55 BR")
    iso_country_list = [f"+{ddi} {iso}" for name, iso, ddi in COUNTRIES_DATA]

    # Cria a Combobox (caixa de seleção)
    country_code_menu = ttk.Combobox(
        telefone_frame,
        values=iso_country_list,
        width=10,
        font=("Helvetica", 12),
        state="readonly"
    )

    # Criamos a variável e a "ancoramos" diretamente no widget do menu
    country_code_menu.variable = tk.StringVar()
    country_code_menu['textvariable'] = country_code_menu.variable

    # Definimos o valor padrão na variável que agora está segura
    country_code_menu.variable.set("+55 BR")

    country_code_menu.pack(side="left")

    # O resto do código para criar as caixas de entrada continua igual...
    # Recipiente para os campos de número
    number_fields_frame = tk.Frame(telefone_frame, bg="#f0f0f0")
    number_fields_frame.pack(side="left", padx=5)

    # Cria as 3 caixas de entrada para DDD e número
    telefone_entries = [
        tk.Entry(number_fields_frame, width=3, font=("Helvetica", 12), justify='center'), # DDD
        tk.Entry(number_fields_frame, width=6, font=("Helvetica", 12), justify='center'), # Parte 1
        tk.Entry(number_fields_frame, width=5, font=("Helvetica", 12), justify='center')  # Parte 2
    ]

    # Adiciona os campos e separadores
    tk.Label(number_fields_frame, text="(", bg="#f0f0f0", font=("Helvetica", 12)).pack(side="left")
    telefone_entries[0].pack(side="left")
    tk.Label(number_fields_frame, text=")", bg="#f0f0f0", font=("Helvetica", 12)).pack(side="left")
    telefone_entries[1].pack(side="left")
    tk.Label(number_fields_frame, text="-", bg="#f0f0f0", font=("Helvetica", 12)).pack(side="left")
    telefone_entries[2].pack(side="left")

    dob_frame = tk.Frame(search_window, bg="#f0f0f0")
    dob_entries = [
        tk.Entry(dob_frame, width=2, font=("Helvetica", 12), justify='center'),  # Dia
        tk.Entry(dob_frame, width=2, font=("Helvetica", 12), justify='center'),  # Mês
        tk.Entry(dob_frame, width=4, font=("Helvetica", 12), justify='center', validate="key", validatecommand=(search_window.register(lambda char: char.isdigit() and len(dob_entries[2].get()) < 4), '%S'))  # Ano
    ]
    dob_entries[0].pack(side="left")
    tk.Label(dob_frame, text="/", bg="#f0f0f0", font=("Helvetica", 12)).pack(side="left")
    dob_entries[1].pack(side="left")
    tk.Label(dob_frame, text="/", bg="#f0f0f0", font=("Helvetica", 12)).pack(side="left")
    dob_entries[2].pack(side="left")

    def on_cpf_entry_key(event, index):
        if index < 3 and len(cpf_entries[index].get()) == 3:
            cpf_entries[index + 1].focus_set()
        elif index == 3 and len(cpf_entries[index].get()) == 2:
            search_button.focus_set()

    def on_telefone_entry_key(event, index):
        if (index == 0 and len(telefone_entries[index].get()) == 2) or (index == 1 and len(telefone_entries[index].get()) == 5) or (index == 2 and len(telefone_entries[index].get()) == 4):
            if index < 2:
                telefone_entries[index + 1].focus_set()

    def on_dob_entry_key(event, index):
        if (index == 0 and len(dob_entries[index].get()) == 2) or (index == 1 and len(dob_entries[index].get()) == 2) or (index == 2 and len(dob_entries[index].get()) == 4):
            if index < 2:
                dob_entries[index + 1].focus_set()

    for i, entry in enumerate(cpf_entries):
        entry.bind('<KeyRelease>', lambda event, index=i: on_cpf_entry_key(event, index))
        entry.config(validate="key", validatecommand=(search_window.register(lambda char: char.isdigit()), '%S'))

    for i, entry in enumerate(telefone_entries):
        entry.bind('<KeyRelease>', lambda event, index=i: on_telefone_entry_key(event, index))
        entry.config(validate="key", validatecommand=(search_window.register(lambda char: char.isdigit()), '%S'))

    for i, entry in enumerate(dob_entries):
        entry.bind('<KeyRelease>', lambda event, index=i: on_dob_entry_key(event, index))
        entry.config(validate="key", validatecommand=(search_window.register(lambda char: char.isdigit()), '%S'))

    def execute_search():
        search_type = search_option.get()
        search_value = search_entry.get().strip()
        results_df = pd.DataFrame() # Cria um DataFrame vazio para os resultados

        # --- LÓGICA DE BUSCA COM PANDAS ---
        if search_type == "Nome":
            # O pulo do gato: .str.contains() busca partes do nome e case=False ignora maiúsculas/minúsculas
            results_df = hospedes_df[hospedes_df['NOME'].str.contains(search_value, case=False, na=False)]

        elif search_type == "CPF":
            # Remove qualquer formatação do input para buscar apenas números
            cpf_value = "".join(filter(str.isdigit, search_value))
            results_df = hospedes_df[hospedes_df['CPF'] == cpf_value]

        elif search_type == "Telefone":
            selected_code_full = country_code_menu.variable.get() # ex: "+55 BR"
            country_code = selected_code_full.split(' ')[0]
            if phone_value: # Apenas busca se houver algum número
                results_df = hospedes_df[hospedes_df['TELEFONE'].str.replace(r'\D', '', regex=True).str.contains(phone_value, na=False)]

        elif search_type == "Data de Nascimento":
            # Busca pela data exata no formato DD/MM/YYYY
            results_df = hospedes_df[hospedes_df['DT. NASCIMENTO'] == search_value]
        
        # --- PREPARA OS DADOS PARA EXIBIÇÃO ---
        results_list = []
        if not results_df.empty:
            for index, row in results_df.iterrows():
                nome = row['NOME']
                dt_nascimento = row['DT. NASCIMENTO']
                telefone = row.get('TELEFONE', 'Não encontrado')
                idade = calculate_age(dt_nascimento)
                
                # Formata o telefone
                if pd.notna(telefone) and str(telefone).strip():
                     telefone = str(telefone).replace(" ", "").replace("-", "").replace("(", "").replace(")", "")
                     if not telefone.startswith("+"):
                         telefone = "+55" + telefone
                else:
                    telefone = "Não encontrado"

                results_list.append((nome, idade, dt_nascimento, telefone))

        # Fecha a janela de busca e abre a de resultados
        search_window.destroy()
        show_search_results(results_list, search_value)

    search_button = tk.Button(search_window, text="Pesquisar", command=execute_search, font=("Helvetica", 12), bg="#0000ff", fg="white")
    search_button.pack(side="bottom", pady=10)

    def update_entry(event=None):
        search_type = search_option.get()
        if search_type == "CPF":
            search_entry.pack_forget()
            telefone_frame.pack_forget()
            dob_frame.pack_forget()
            cpf_frame.pack(pady=10)
        elif search_type == "Telefone":
            search_entry.pack_forget()
            cpf_frame.pack_forget()
            dob_frame.pack_forget()
            telefone_frame.pack(pady=10)
        elif search_type == "Data de nascimento":
            search_entry.pack_forget()
            cpf_frame.pack_forget()
            telefone_frame.pack_forget()
            dob_frame.pack(pady=10)
        else:
            cpf_frame.pack_forget()
            telefone_frame.pack_forget()
            dob_frame.pack_forget()
            search_entry.pack(pady=10)

    search_menu.bind("<<ComboboxSelected>>", update_entry)
    update_entry()

def manual_mode():
    manual_window = tk.Toplevel()
    manual_window.title("Modo Manual")
    manual_window.configure(bg="#f0f0f0")

    # Função para permitir apenas letras e espaços no campo Nome
    def only_letters(char):
        return char.isalpha() or char == " "

    # Função para permitir apenas números no campo Idade com limite de 2 caracteres
    def only_two_digit_numbers(char, entry_value):
        return char.isdigit() and len(entry_value) < 3

    # Campo Nome
    tk.Label(manual_window, text="Nome:", bg="#f0f0f0", font=("Helvetica", 12)).grid(row=0, column=0, padx=10, pady=5, sticky="w")
    name_entry = tk.Entry(manual_window, font=("Helvetica", 12))
    name_entry.grid(row=0, column=1, padx=10, pady=5)
    name_entry.config(validate="key", validatecommand=(manual_window.register(only_letters), "%S"))

    # Campo Telefone
    tk.Label(manual_window, text="Número de telefone:", bg="#f0f0f0", font=("Helvetica", 12)).grid(row=1, column=0, padx=10, pady=5, sticky="w")
    phone_entry_frame = tk.Frame(manual_window, bg="#f0f0f0")
    phone_entry_frame.grid(row=1, column=1, padx=10, pady=5, sticky="w")
    phone_code_label = tk.Label(phone_entry_frame, text="+55", bg="#f0f0f0", font=("Helvetica", 12))
    phone_code_label.pack(side="left")
    phone_number_entry = tk.Entry(phone_entry_frame, font=("Helvetica", 12))
    phone_number_entry.pack(side="left")

    def only_numbers(char):
        return char.isdigit()

    phone_number_entry.config(validate="key", validatecommand=(manual_window.register(only_numbers), "%S"))

    # Campo Idade
    tk.Label(manual_window, text="Idade:", bg="#f0f0f0", font=("Helvetica", 12)).grid(row=2, column=0, padx=10, pady=5, sticky="w")
    age_entry = tk.Entry(manual_window, font=("Helvetica", 12))
    age_entry.grid(row=2, column=1, padx=10, pady=5)
    age_entry.config(validate="key", validatecommand=(manual_window.register(only_two_digit_numbers), "%S", "%P"))

    def send_individual_message(name, phone, age, without_age=False):
        salutation = determine_salutation(name)

        if without_age or not age:
            message = (f"{salutation},\n\n"
                       "Nós te desejamos um feliz aniversário! 🎉\n\n"
                       "Para celebrar esta ocasião especial, temos o prazer de oferecer uma tarifa exclusiva e imperdível. "
                       "Venha comemorar conosco e viva uma experiência única.\n\n"
                       "Reserve agora em nosso site e garanta a melhor tarifa. Estamos à sua espera para tornar esta data inesquecível. 🎁🌴\n\n"
                       "Acesse: https://www.google.com")
        else:
            message = (f"{salutation},\n\n"
                       f"Nós te desejamos um feliz aniversário e parabeniza pelos seus {age} anos! 🎉\n\n"
                       "Para comemorar esta data tão especial, preparamos uma oferta exclusiva e imperdível. "
                       "Venha celebrar conosco e desfrute de uma experiência inesquecível.\n\n"
                       "Reserve diretamente em nosso site e aproveite a melhor tarifa disponível. "
                       "Estamos prontos para tornar sua celebração única e memorável. 🎁\n\n"
                       "Acesse: https://www.google.com")
        
        kit.sendwhatmsg_instantly(phone, message, wait_time=10)

    def add_to_list():
        list_window = tk.Toplevel(manual_window)
        list_window.title("Adicionar à Lista")
        list_window.configure(bg="#f0f0f0")

        tk.Label(list_window, text="Nome:", bg="#f0f0f0", font=("Helvetica", 12)).grid(row=0, column=0, padx=10, pady=5, sticky="w")
        name_list_entry = tk.Entry(list_window, font=("Helvetica", 12))
        name_list_entry.grid(row=0, column=1, padx=10, pady=5)
        name_list_entry.config(validate="key", validatecommand=(list_window.register(only_letters), "%S"))

        tk.Label(list_window, text="Número de telefone:", bg="#f0f0f0", font=("Helvetica", 12)).grid(row=1, column=0, padx=10, pady=5, sticky="w")
        phone_list_entry_frame = tk.Frame(list_window, bg="#f0f0f0")
        phone_list_entry_frame.grid(row=1, column=1, padx=10, pady=5, sticky="w")
        phone_list_code_label = tk.Label(phone_list_entry_frame, text="+55", bg="#f0f0f0", font=("Helvetica", 12))
        phone_list_code_label.pack(side="left")
        phone_list_number_entry = tk.Entry(phone_list_entry_frame, font=("Helvetica", 12))
        phone_list_number_entry.pack(side="left")
        phone_list_number_entry.config(validate="key", validatecommand=(list_window.register(only_numbers), "%S"))

        tk.Label(list_window, text="Idade:", bg="#f0f0f0", font=("Helvetica", 12)).grid(row=2, column=0, padx=10, pady=5, sticky="w")
        age_list_entry = tk.Entry(list_window, font=("Helvetica", 12))
        age_list_entry.grid(row=2, column=1, padx=10, pady=5)
        age_list_entry.config(validate="key", validatecommand=(list_window.register(only_two_digit_numbers), "%S", "%P"))

        without_age_dict = {}

        def toggle_without_age(item_id):
            age = treeview.item(item_id, "values")[2]
            if age == "":
                messagebox.showerror("Erro", "Por motivos óbvios, não é possível desmarcar o envio sem idade de um hóspede no qual não teve sua idade informada")
                return
            without_age_dict[item_id] = not without_age_dict.get(item_id, False)
            update_context_menu(item_id)
            
        def check_for_undefined_pronouns_and_ask_to_define(treeview):
            undefined_items = []

            for item in treeview.get_children():
                values = treeview.item(item, "values")
                name, phone, age = values
                first_name = format_name(name)

                if first_name not in nomes_masculinos and first_name not in nomes_femininos:
                    undefined_items.append(item)

            if undefined_items:
                show_custom_dialog(
                    "Pronomes indefinidos",
                    "Existem hóspedes com pronomes indefinidos.\nGostaria de definir agora?",
                    confirm_action=lambda: ask_to_define_pronouns(undefined_items),
                    cancel_action=lambda: send_messages()
                )
            else:
                send_messages()
                
        def ask_to_define_pronouns(undefined_items):
            def show_next_pronoun_window():
                if undefined_items:
                    selected_item = undefined_items.pop(0)
                    define_pronoun(treeview, selected_item, show_next_pronoun_window)
                else:
                    confirm_sending_dialog()

            def confirm_sending_dialog():
                dialog = tk.Toplevel()
                dialog.title("Todos os pronomes definidos")
                dialog.geometry("300x150")
                dialog.grab_set()

                tk.Label(dialog, text="Todos os hóspedes tiveram seus pronomes definidos.\nGostaria de começar o envio agora?",
                         wraplength=280, justify="center").pack(pady=20)

                response = tk.StringVar(value="no")

                def on_yes():
                    response.set("yes")
                    dialog.destroy()

                def on_no():
                    response.set("no")
                    dialog.destroy()

                tk.Button(dialog, text="Sim", command=on_yes, width=15).pack(side="left", padx=10, pady=10)
                tk.Button(dialog, text="Não", command=on_no, width=15).pack(side="right", padx=10, pady=10)

                dialog.wait_window()

                if response.get() == "yes":
                    send_messages()

            show_next_pronoun_window()

        def start_sending():
            check_for_undefined_pronouns_and_ask_to_define(treeview)
            undefined_items = []

            # Verifica na treeview se há algum hóspede com o pronome "Não informado"
            for item in treeview.get_children():
                values = treeview.item(item, "values")
                name, phone, age = values
                current_salutation = determinar_saudacao(name)

                if current_salutation == "Prezado(a)":
                    undefined_items.append(item)

            if undefined_items:
                ask_to_define_pronouns(undefined_items)
            else:
                send_messages()

        def send_messages():
            global sending_completed
            if sending_completed:
                return  # Se o envio já foi concluído, não faz nada

            # Função que executa o envio das mensagens
            for i, item in enumerate(treeview.get_children()):
                values = treeview.item(item, "values")
                name, phone, age = values
                without_age = without_age_dict.get(item, False)

                # Verifica se o número é válido antes de enviar a mensagem
                if not is_valid_number(phone):
                    print(f"Pulando o envio para {name}, número inválido: {phone}")
                    continue  # Pula para o próximo item na lista

                send_individual_message(name, phone, age, without_age)

                # Aguarda um intervalo aleatório antes do próximo envio
                if i < len(treeview.get_children()) - 1:
                    time.sleep(random.randint(20, 60))  # Espera aleatória de 20 a 60 segundos antes de enviar a próxima mensagem

            # Marcar como concluído após o envio
            sending_completed = True

            # Exibe uma mensagem de envio concluído
            messagebox.showinfo("Envio Concluído", "Mensagens enviadas com sucesso!!!")

            # Após exibir a mensagem, desativa o botão "Começar" para evitar reenvio acidental
            start_button.config(state=tk.DISABLED)

        def reset_sending_state():
            global sending_completed
            sending_completed = False
            # Reativar o botão "Começar" caso você queira permitir novos envios após redefinir o estado
            start_button.config(state=tk.NORMAL)

        # Dentro da função add_to_list_action, quando uma nova pessoa for adicionada à lista:
        def add_to_list_action():
            name = name_list_entry.get().strip()
            phone = "+55" + phone_list_number_entry.get().strip()
            age = age_list_entry.get().strip()
            without_age = not age  # Marca como sem idade se a idade não for informada

            # Carrega o pronome salvo ou define a saudação baseada nas listas
            salutation = determine_salutation(name)

            # Inserção na Treeview
            item_id = treeview.insert("", "end", values=(format_name(name), phone, age))
            without_age_dict[item_id] = without_age  # Inicialmente desativado ou ativado com base na idade informada

            # Limpa os campos após adicionar
            name_list_entry.delete(0, tk.END)
            phone_list_number_entry.delete(0, tk.END)
            age_list_entry.delete(0, tk.END)
            update_start_button_state()

        add_list_button = tk.Button(list_window, text="Adicionar", command=add_to_list_action, font=("Helvetica", 12), bg="orange", fg="white")
        add_list_button.grid(row=3, column=0, padx=10, pady=10, sticky="ew")

        columns = ("Nome", "Telefone", "Idade")
        treeview = ttk.Treeview(list_window, columns=columns, show="headings", height=5)
        treeview.heading("Nome", text="Nome")
        treeview.heading("Telefone", text="Telefone")
        treeview.heading("Idade", text="Idade")
        treeview.column("Nome", width=150)
        treeview.column("Telefone", width=100)
        treeview.column("Idade", width=50)
        treeview.grid(row=4, column=0, columnspan=2, padx=10, pady=10)

        full_menu = tk.Menu(list_window, tearoff=0)

        def update_context_menu(item_id):
            full_menu.delete(0, "end")
            full_menu.add_command(label="Envio individual", command=lambda: send_selected_message(treeview))
            if without_age_dict.get(item_id, False):
                full_menu.add_command(label="Enviar sem idade ✓", command=lambda: toggle_without_age(item_id))
            else:
                full_menu.add_command(label="Enviar sem idade", command=lambda: toggle_without_age(item_id))
            full_menu.add_command(label="Definir pronome", command=lambda: define_pronoun_from_menu(treeview, item_id))
            full_menu.add_command(label="Remover", command=lambda: remove_selected(treeview))

        def clear_list(treeview):
            for item in treeview.get_children():
                treeview.delete(item)
            update_start_button_state()

        empty_menu = tk.Menu(list_window, tearoff=0)
        empty_menu.add_command(label="Limpar lista", command=lambda: clear_list(treeview))

        def show_context_menu(event):
            selected_item = treeview.identify_row(event.y)
            if selected_item:
                treeview.selection_set(selected_item)
                update_context_menu(selected_item)
                full_menu.post(event.x_root, event.y_root)
            else:
                empty_menu.post(event.x_root, event.y_root)

        treeview.bind("<Button-3>", show_context_menu)

        def send_selected_message(treeview):
            selected_item = treeview.selection()[0]
            values = treeview.item(selected_item, "values")
            name, phone, age = values
            without_age = without_age_dict.get(selected_item, False)
            send_individual_message(name, phone, age, without_age)

        def remove_selected(treeview):
            selected_item = treeview.selection()[0]
            treeview.delete(selected_item)
            update_start_button_state()

        def update_start_button_state():
            if len(treeview.get_children()) > 0:
                start_button.grid(row=3, column=1, padx=10, pady=10, sticky="ew")
            else:
                start_button.grid_remove()

        start_button = tk.Button(list_window, text="Começar", command=start_sending, font=("Helvetica", 12), bg="#4caf50", fg="white")
        start_button.grid(row=3, column=1, padx=10, pady=10, sticky="ew")
        start_button.grid_remove()

    send_button = tk.Button(manual_window, text="Envio individual", command=lambda: send_individual_message(name_entry.get(), "+55" + phone_number_entry.get().strip(), age_entry.get().strip()), font=("Helvetica", 12), bg="#4caf50", fg="white")
    send_button.grid(row=3, column=0, padx=10, pady=10, sticky="ew")

    add_button = tk.Button(manual_window, text="Adicionar a lista", command=add_to_list, font=("Helvetica", 12), bg="orange", fg="white")
    add_button.grid(row=3, column=1, padx=10, pady=10, sticky="ew")

# Função para alternar o modo de adição de exceções
def toggle_exception_mode():
    global is_adding_exception
    is_adding_exception = not is_adding_exception
    add_exception_button.config(text="Concluir" if is_adding_exception else "Adicionar exceção")
    if is_adding_exception:
        for frame, _, name, _ in frames:
            if frame.winfo_exists():
                frame.after(500, blink_name, frame.winfo_children()[0], True)
    else:
        confirm_exception_changes()
        save_exceptions_to_file()  # Salvar exceções no arquivo
    update_exception_list_message()

# Função para fazer os nomes piscarem
def blink_name(label, visible):
    if is_adding_exception and label.winfo_exists():
        label.config(fg="white" if visible else "black")
        label.after(500, blink_name, label, not visible)
        
# Função para marcar/desmarcar exceções
def toggle_exception(event, name):
    if is_adding_exception:
        for frame, phone, frame_name, _ in frames:
            if frame.winfo_exists() and frame_name == name:
                current_color = frame.cget("bg")
                new_color = "#ffcccc" if current_color == "#ffffff" else "#ffffff"
                frame.config(bg=new_color)
                for widget in frame.winfo_children():
                    widget.config(bg=new_color, fg="black")

# Função para confirmar mudanças de exceção
def confirm_exception_changes():
    global exception_list
    new_exceptions = [name for frame, _, name, _ in frames if frame.winfo_exists() and frame.cget("bg") == "#ffcccc"]
    exception_list.extend([name for name in new_exceptions if name not in exception_list])
    if exception_list:
        message = f"Você está incluindo {len(exception_list)} hóspede(s) na exceção de envio de mensagens de parabenização. Deseja continuar?"
        response = messagebox.askyesno("Confirmação de exceção", message)
        if not response:
            exception_list = [name for name in exception_list if name not in new_exceptions]
        update_exception_list_message()
            
# Função para resetar as cores dos frames dos aniversariantes quando confirmado exceção
def reset_frame_colors():
    for frame, _, _, _ in frames:
        if frame.winfo_exists():
            frame.config(bg="#ffffff")
            for widget in frame.winfo_children():
                widget.config(bg="#ffffff")

# Função para atualizar a mensagem de aniversariantes sem telefone
def update_missing_phone_message(birthdays):
    missing_phone_count = sum(1 for _, _, _, phone in birthdays if phone == "Telefone não encontrado")
    if missing_phone_count > 0:
        plural = "hóspedes" if missing_phone_count > 1 else "hóspede"
        message = f"{missing_phone_count} {plural} está/estão sem o telefone cadastrado"
        missing_phone_label.config(text=message)
    else:
        missing_phone_label.config(text="")

# Função para destacar entradas sem telefone
def highlight_missing_phone_entries(birthdays):
    for frame, phone, _, _ in frames:
        if phone == "Telefone não encontrado":
            frame.config(bg="#ffcccc")

# Função para mostrar um pop-up com os nomes da lista de exceção
def show_exception_list_popup(widget, text, x, y):
    popup = Toplevel()
    popup.wm_overrideredirect(True)
    popup.geometry(f"+{x + 10}+{y + 10}")
    label = tk.Label(popup, text=text, bg="white", relief="solid", bd=1, justify="left")
    label.pack()
    popup.attributes("-alpha", 0.0)
    fade_in(popup)
    widget.popup = popup

def delayed_exception_list_popup(widget, text, event):
    x, y = event.x_root, event.y_root
    widget.popup_after_id = widget.after(1000, lambda: show_exception_list_popup(widget, text, x, y))

def update_exception_list_message():
    exception_list = load_exceptions_from_file()
    if exception_list:
        plural = "hóspedes" if len(exception_list) > 1 else "hóspede"
        message = f"{len(exception_list)} {plural} incluso(s) na lista de exceção"
        exception_list_label.config(text=message)
        exception_list_label.bind("<Enter>", lambda e: delayed_exception_list_popup(exception_list_label, "\n".join(exception_list), e))
        exception_list_label.bind("<Leave>", lambda e: cancel_popup(exception_list_label))
    else:
        exception_list_label.config(text="")
        exception_list_label.unbind("<Enter>")
        exception_list_label.unbind("<Leave>")

# Função para criar o efeito de fade-in
def fade_in_intro(window, interval=50, increment=0.05):
    alpha = 0
    while alpha < 1:
        window.attributes("-alpha", alpha)
        alpha += increment
        window.update()
        time.sleep(interval / 5000.0)  # Convert milliseconds to seconds

# Tela de carregamento
loading_screen = tk.Tk()
loading_screen.title("Mensageiro de Aniversário")
loading_screen.configure(bg="#f0f0f0")

# Fix the path for iconbitmap
loading_screen.iconbitmap('loading.ico')  # Ensure the path is correct or use a relative path

# Carrega a imagem da logo
logo_image = Image.open("logo.png")
logo_image = logo_image.resize((300, 300), Image.LANCZOS)  # Correct attribute
logo_photo = ImageTk.PhotoImage(logo_image)

# Exibe a logo e o texto de carregamento
logo_label = tk.Label(loading_screen, image=logo_photo, bg="#f0f0f0")
logo_label.pack(pady=20)
loading_label = tk.Label(loading_screen, text="Carregando...", font=("Helvetica", 16), bg="#f0f0f0")
loading_label.pack()
credit_label = tk.Label(loading_screen, text="Criado por Lucas Siqueira", font=("Helvetica", 10), bg="#f0f0f0")
credit_label.pack()

# Inicialmente a janela está oculta
loading_screen.attributes("-alpha", 0)

# Inicia o efeito de fade-in
fade_in_intro(loading_screen)

# Fecha a tela de carregamento após 1.5 segundos e abre a tela principal
loading_screen.after(2000, lambda: (loading_screen.destroy(), show_main_screen()))

loading_screen.mainloop()