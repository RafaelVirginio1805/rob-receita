import pandas as pd
import requests
import time
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
import threading
import os 
from datetime import datetime 

TIME_BETWEEN = 20  # segundos entre consultas

def consultar_cnpj(cnpj):
    url = f"https://www.receitaws.com.br/v1/cnpj/{cnpj}"
    try:
        resp = requests.get(url)
        if resp.status_code == 429:
            time.sleep(60)
            return consultar_cnpj(cnpj)
        if resp.status_code == 200:
            return resp.json()
        else:
            return {"status": str(resp.status_code)}
    except Exception as e:
        return {"status": "Erro"}

def rodar_consulta(caminho_entrada, caminho_saida, barra, botao, label_status):
    botao.config(state="disabled")
    barra.start(10)

    try:
        df = pd.read_excel(caminho_entrada)
        cnpjs = df.iloc[:, 0].dropna().astype(str).str.zfill(14)
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao ler o Excel:\n{e}")
        barra.stop()
        botao.config(state="normal")
        return

    resultados = []
    total = len(cnpjs)
    sucesso = 0
    erro = 0
    os.makedirs("backups", exist_ok=True)

    for i, cnpj in enumerate(cnpjs, 1):
        status_txt = f"🔍 Consultando ({i}/{total}): {cnpj}"
        label_status.config(text=status_txt)
        label_status.update()
        

        d = consultar_cnpj(cnpj)

        if d and d.get("status") == "OK":
            res = {
                "CNPJ": cnpj,
                "Nome Empresarial Cartão CNPJ": d.get("nome"),
                "Nome de Fantasia Cartão CNPJ": d.get("fantasia"),
                "Logradouro do Cartão CNPJ": d.get("logradouro"),
                "nº Cartão CNPJ": d.get("numero"),
                "Complemento Cartão CNPJ": d.get("complemento"),
                "CEP do Cartão CNPJ": d.get("cep"),
                "Bairro do Cartão CNPJ": d.get("bairro"),
                "Município do Cartão do CNPJ": d.get("municipio"),
                "UF Cartão CNPJ": d.get("uf"),
                "Telefone Cartão CNPJ": d.get("telefone"),
                "E-mail Cartão CNPJ": d.get("email"),
                "Situação Cadastral": d.get("situacao"),
                "Data da Situação Cadastral": d.get("data_situacao"),
            }
            sucesso += 1
        else:
            res = {
                "CNPJ": cnpj,
                "Nome Empresarial Cartão CNPJ": None,
                "Nome de Fantasia Cartão CNPJ": None,
                "Logradouro do Cartão CNPJ": None,
                "nº Cartão CNPJ": None,
                "Complemento Cartão CNPJ": None,
                "CEP do Cartão CNPJ": None,
                "Bairro do Cartão CNPJ": None,
                "Município do Cartão do CNPJ": None,
                "UF Cartão CNPJ": None,
                "Telefone Cartão CNPJ": None,
                "E-mail Cartão CNPJ": None,
                "Situação Cadastral": d.get("status") if d else "Erro",
                "Data da Situação Cadastral": None,
            }
            erro += 1

        resultados.append(res)

        if i % 50 == 0:
            backup = pd.DataFrame(resultados)
            nome_backup = os.path.join("backups", f"backup_parcial_{i}.xlsx")
            backup.to_excel(nome_backup, index=False)
            print(f"📦 Backup salvo: {nome_backup}")

        time.sleep(TIME_BETWEEN)

    df_out = pd.DataFrame(resultados)
    df_out.to_excel(caminho_saida, index=False)

    barra.stop()
    botao.config(state="normal")
    label_status.config(text="✅ Consulta finalizada.")

    resumo = (
        f"🔎 Consulta Finalizada!\n\n"
        f"📄 Total: {total}\n"
        f"✅ Sucesso: {sucesso}\n"
        f"❌ Erros: {erro}\n\n"
        f"📁 Arquivo salvo em:\n{caminho_saida}"
    )
    messagebox.showinfo("Resumo da Consulta", resumo)

def iniciar_consulta(botao, barra, label_status):
    entrada = filedialog.askopenfilename(title="Selecione o Excel com os CNPJs", filetypes=[("Arquivos Excel", "*.xlsx")])
    if not entrada:
        return

    saida = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Arquivos Excel", "*.xlsx")])
    if not saida:
        return

    thread = threading.Thread(target=rodar_consulta, args=(entrada, saida, barra, botao, label_status))
    thread.start()

def criar_interface():
    janela = tk.Tk()
    janela.title("Consulta de CNPJ Automática")
    janela.geometry("420x240")

    tk.Label(janela, text="Consulta de Cartão CNPJ\nvia ReceitaWS", font=("Arial", 12, "bold")).pack(pady=10)

    barra = ttk.Progressbar(janela, mode="indeterminate")
    barra.pack(pady=5, fill="x", padx=20)

    label_status = tk.Label(janela, text="", font=("Arial", 10))
    label_status.pack(pady=5)

    botao = tk.Button(janela, text="Selecionar arquivo e iniciar", width=30,
                      command=lambda: iniciar_consulta(botao, barra, label_status))
    botao.pack(pady=10)

    tk.Label(janela, text="Feito para uso no trabalho 💼", font=("Arial", 8)).pack(side="bottom", pady=5)

    return janela

if __name__ == "__main__":
    app = criar_interface()
    app.mainloop()
