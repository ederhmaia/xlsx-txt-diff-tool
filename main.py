import xlsxwriter
import pandas as pd
import os
from typing import List
from colorama import init, Fore

init(autoreset=True)

def banner() -> None:
    print(f'''{Fore.CYAN}
    {Fore.CYAN}            ,    _             
    {Fore.CYAN}           /|   | |         {Fore.WHITE}   -.. .. .- -. .-
    {Fore.CYAN}         _/_\_  >_<         {Fore.MAGENTA}   diff xlsx - txt    
    {Fore.CYAN}        .-\-/.   |          {Fore.MAGENTA}    by @ederhmaia
    {Fore.CYAN}       /  | | \_ |          {Fore.WHITE}   -.. .. .- -. .-
    {Fore.CYAN}       \ \| |\__(/                  
    {Fore.CYAN}       /(`---')  |          {Fore.RED}    the responsability
    {Fore.CYAN}      / /     \  |          {Fore.RED}    of using this tool
    {Fore.CYAN}   _.'  \\'-'  /  |          {Fore.RED}    is yours, not mine.
    {Fore.CYAN}   `----'`=-='   '    
    ''')

def clear():
    os.system('cls' if os.name == 'nt' else 'clear')
    banner()

def process_files(txt_filename: str, xlsx_filename: str, xlsx_output: str) -> None:
    df = pd.read_excel(xlsx_filename)

    df_cpfs_list = [str(cpf) for cpf in df['CPF'].tolist()]

    matched_cpfs = []

    with open(txt_filename, 'r', encoding='utf-8') as TXT_FILE:
        for line in TXT_FILE:
            cpf = line.strip()
            if cpf in df_cpfs_list:
                matched_cpfs.append(cpf)

    if matched_cpfs:
        clear()
        print(f'{Fore.GREEN}✔ {Fore.WHITE}A semelhança é de {Fore.GREEN}{len(matched_cpfs)}{Fore.WHITE} linhas. {Fore.GREEN}✔')

        workbook = xlsxwriter.Workbook(xlsx_output, {'nan_inf_to_errors': True})
        worksheet = workbook.add_worksheet()

        headers = list(df.columns)

        for col, header in enumerate(headers):
            worksheet.write(0, col, header)

        matched_indices = [df_cpfs_list.index(cpf) for cpf in matched_cpfs]

   
        row = 1 
        for idx in matched_indices:
            for col, value in enumerate(df.iloc[idx]):
                worksheet.write(row, col, value)
            row += 1

        workbook.close()
    else:
        print(f'{Fore.RED}✘ {Fore.WHITE}Nenhuma semelhança foi encontra {Fore.RED}✘')


def main():
    clear()

    print(f"{Fore.GREEN}[↓] {Fore.WHITE}Digite o nome do XLSX a ser filtrado")
    while True:
        try:
            xlsx_filename: str = input(f"{Fore.GREEN}→{Fore.WHITE} ")


            if not xlsx_filename.endswith(".xlsx"):
                xlsx_filename += ".xlsx"

            if os.path.exists(xlsx_filename):
                break
            else:
                raise ValueError()

        except ValueError:
            print(f"{Fore.RED}✘ {Fore.WHITE}Arquivo {Fore.RED}{xlsx_filename}{Fore.WHITE} não existente {Fore.RED}✘")
            print()
            continue


    clear()

    print(f"{Fore.GREEN}[↓] {Fore.WHITE}Digite o nome do TXT para ser usado como filtro")
    while True:
        try:
            txt_filename: str = input(f"{Fore.GREEN}→{Fore.WHITE} ")

            if not txt_filename.endswith(".txt"):
                txt_filename += ".txt"

            if os.path.exists(txt_filename):
                break
            else:
                raise ValueError()

        except ValueError:
            print(f"{Fore.RED}✘ {Fore.WHITE}Arquivo {Fore.RED}{txt_filename}{Fore.WHITE} não existente {Fore.RED}✘")
            print()
            continue

    clear()
    print(f"{Fore.GREEN}[↓] {Fore.WHITE}Digite o nome do XLSX filtrado (saída)")
    while True:
        try:
            xlsx_output: str = input(f"{Fore.GREEN}→{Fore.WHITE} ")

            if not xlsx_output.endswith(".xlsx"):
                xlsx_output += ".xlsx"

            if os.path.exists(xlsx_output):
                raise ValueError()

            break

        except ValueError:
            print(f"{Fore.RED}✘ {Fore.WHITE}Arquivo {Fore.RED}{xlsx_output}{Fore.WHITE} já existe {Fore.RED}✘")
            print()
            continue

    process_files(txt_filename, xlsx_filename, xlsx_output)
    clear()
    print(f"{Fore.GREEN}✔ {Fore.WHITE}Processo finalizado {Fore.GREEN}✔")


if __name__ == '__main__':
    try:
        main()
    except KeyboardInterrupt:
        print(f"{Fore.RED}✘ {Fore.WHITE}Goodbye! {Fore.RED}✘")
