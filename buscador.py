import re
import pandas as pd
import fitz  # PyMuPDF
import tkinter as tk
from tkinter import ttk, messagebox
import os
import datetime

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

# Função para ler o arquivo Excel de telefones
def read_phones_excel(excel_path):
    try:
        return pd.read_excel(excel_path)
    except FileNotFoundError:
        response = messagebox.askyesno("Banco de dados de telefone dos hóspedes não encontrado",
                                       "Banco de dados de telefone dos hóspedes não encontrado, deseja continuar assim mesmo? "
                                       "Caso sim, o programa não irá funcionar corretamente, e o envio de mensagens não poderá ser iniciado.")
        if response:
            return None  # Return None to indicate que o arquivo não foi encontrado
        else:
            root.destroy()

# Função para encontrar o telefone pelo nome
def find_phone_by_name(name, phones_df):
    phone_row = phones_df[phones_df['NOME'] == name]
    if not phone_row.empty:
        phone = phone_row.iloc[0]['TELEFONE']
        if pd.isna(phone):
            return "Telefone não encontrado"
        else:
            phone = phone.replace(" ", "").replace("-", "")
            if not phone.startswith("+55"):
                phone = "+55" + phone
            return phone
    return "Telefone não encontrado"

# Função para extrair aniversariantes do PDF
def extract_birthdays_from_pdf(pdf_path, phones_df):
    birthdays = []
    today = datetime.datetime.now().strftime("%d/%m")
    
    if not os.path.exists(pdf_path):
        messagebox.showerror("Erro", "Banco de dados de hóspedes não encontrado")
        root.destroy()  # Fechar a tela principal
        return []

    with fitz.open(pdf_path) as doc:
        for page in doc:
            text = page.get_text()
            lines = text.split('\n')
            for i in range(len(lines)):
                if lines[i].strip().isdigit():  # Identificar linhas que começam com código numérico
                    try:
                        nome = lines[i + 1].strip()
                        dt_nascimento = lines[i + 5].strip()
                        dt_nascimento_fmt = dt_nascimento[:5]
                        if dt_nascimento_fmt == today:
                            idade = calculate_age(dt_nascimento)
                            telefone = find_phone_by_name(nome, phones_df)
                            if telefone == "Telefone não encontrado":
                                nome = lines[i + 1].strip()  # Correção: usar o nome correto
                            birthdays.append((nome, idade, dt_nascimento, telefone))
                    except IndexError as e:
                        print(f"Erro ao processar linha {i}: {e}")
    return birthdays

def read_excel(file_path):
    try:
        df = pd.read_excel(file_path)
        df.columns = df.columns.str.lower()  # Convertendo os nomes das colunas para minúsculas
        print(f"Excel columns: {df.columns.tolist()}")  # Debug: Print the column names
        if 'nome' not in df.columns or 'telefone' not in df.columns:
            raise ValueError("Excel file must contain 'nome' and 'telefone' columns.")
        guests = []
        for _, row in df.iterrows():
            guests.append({
                "name": row["nome"],
                "phone": row["telefone"],
            })
        return guests
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return []

def combine_data(pdf_data, excel_data):
    combined_data = []

    print(f"PDF data: {pdf_data}")  # Debug: Print PDF data
    print(f"Excel data: {excel_data}")  # Debug: Print Excel data

    for pdf_guest in pdf_data:
        for excel_guest in excel_data:
            if pdf_guest["name"].lower() in excel_guest["name"].lower():
                combined_data.append({
                    "name": pdf_guest["name"],
                    "cpf": pdf_guest["cpf"],
                    "dob": pdf_guest["dob"],
                    "phone": excel_guest["phone"]
                })

    return combined_data

def search_guests(criteria, value, guests_data):
    results = []
    if criteria == "Nome":
        results = [guest for guest in guests_data if value.lower() in guest["name"].lower()]
    elif criteria == "CPF":
        results = [guest for guest in guests_data if value in guest["cpf"]]
    elif criteria == "Telefone":
        results = [guest for guest in guests_data if value in guest["phone"]]
    elif criteria == "Data de nascimento":
        results = [guest for guest in guests_data if value in guest["dob"]]
    return results

def search_guest():
    search_window = tk.Toplevel()
    search_window.title("Pesquisar hóspede")
    search_window.geometry("500x400")
    search_window.configure(bg="#f0f0f0")

    tk.Label(search_window, text="Pesquisar por:", bg="#f0f0f0", font=("Helvetica", 12)).pack(pady=10)

    search_option = tk.StringVar(value="Nome")
    options = ["Nome", "CPF", "Telefone", "Data de nascimento"]
    search_menu = ttk.Combobox(search_window, textvariable=search_option, values=options, state="readonly", font=("Helvetica", 12))
    search_menu.pack(pady=5)

    search_entry = tk.Entry(search_window, font=("Helvetica", 12))
    search_entry.pack(pady=10)

    result_frame = tk.Frame(search_window, bg="#f0f0f0")
    result_frame.pack(pady=10)
    
    cpf_frame = tk.Frame(search_window, bg="#f0f0f0")
    cpf_entries = [tk.Entry(cpf_frame, width=3, font=("Helvetica", 12), justify='center') for _ in range(4)]
    for i, entry in enumerate(cpf_entries):
        entry.pack(side="left")
        if i < 3:
            separator = tk.Label(cpf_frame, text="." if i < 2 else "-", bg="#f0f0f0", font=("Helvetica", 12))
            separator.pack(side="left")

    telefone_frame = tk.Frame(search_window, bg="#f0f0f0")
    telefone_entries = [
        tk.Entry(telefone_frame, width=2, font=("Helvetica", 12), justify='center'),  # DDD
        tk.Entry(telefone_frame, width=5, font=("Helvetica", 12), justify='center'),  # Primeira parte do número
        tk.Entry(telefone_frame, width=4, font=("Helvetica", 12), justify='center')   # Segunda parte do número
    ]
    tk.Label(telefone_frame, text="+55 (", bg="#f0f0f0", font=("Helvetica", 12)).pack(side="left")
    telefone_entries[0].pack(side="left")
    tk.Label(telefone_frame, text=") ", bg="#f0f0f0", font=("Helvetica", 12)).pack(side="left")
    telefone_entries[1].pack(side="left")
    tk.Label(telefone_frame, text="-", bg="#f0f0f0", font=("Helvetica", 12)).pack(side="left")
    telefone_entries[2].pack(side="left")

    def format_cpf():
        cpf = "".join(entry.get() for entry in cpf_entries)
        return f"{cpf[:3]}.{cpf[3:6]}.{cpf[6:9]}-{cpf[9:11]}"

    def format_telefone():
        ddd = telefone_entries[0].get()
        first_part = telefone_entries[1].get()
        second_part = telefone_entries[2].get()
        return f"+55 ({ddd}) {first_part}-{second_part}"

    def clear_results():
        for widget in result_frame.winfo_children():
            widget.destroy()

    def display_results(results):
        clear_results()
        if not results:
            tk.Label(result_frame, text="Nenhum hóspede encontrado.", bg="#f0f0f0", font=("Helvetica", 12)).pack()
            return
        for guest in results:
            result_text = f"Nome: {guest['name']}\nCPF: {guest['cpf']}\nData de nascimento: {guest['dob']}\nTelefone: {guest['phone']}\n"
            tk.Label(result_frame, text=result_text, bg="#f0f0f0", font=("Helvetica", 12), justify="left").pack(anchor="w", pady=5)

    def perform_search():
        criteria = search_option.get()
        value = search_entry.get()

        if criteria == "CPF":
            value = format_cpf()
        elif criteria == "Telefone":
            value = format_telefone()

        print(f"Searching for {criteria} with value {value}")  # Debug: Print the search criteria and value

        guests = combine_data(extract_birthdays_from_pdf("data.pdf", read_phones_excel("telefones.xlsx")), read_excel("telefones.xlsx"))
        results = search_guests(criteria, value, guests)
        display_results(results)

    search_button = tk.Button(search_window, text="Pesquisar", command=perform_search, font=("Helvetica", 12))
    search_button.pack(pady=10)

    def update_search_input(*args):
        if search_option.get() == "CPF":
            search_entry.pack_forget()
            telefone_frame.pack_forget()
            cpf_frame.pack(pady=10)
        elif search_option.get() == "Telefone":
            search_entry.pack_forget()
            cpf_frame.pack_forget()
            telefone_frame.pack(pady=10)
        else:
            cpf_frame.pack_forget()
            telefone_frame.pack_forget()
            search_entry.pack(pady=10)

    search_option.trace("w", update_search_input)

    update_search_input()

    search_window.mainloop()

# Inicializando a janela principal do Tkinter
root = tk.Tk()
root.title("Sistema de Gestão de Hóspedes")
root.geometry("300x200")
root.configure(bg="#f0f0f0")

tk.Button(root, text="Pesquisar Hóspede", command=search_guest, font=("Helvetica", 12)).pack(pady=20)

root.mainloop()