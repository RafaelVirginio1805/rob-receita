import pandas as pd
import requests
import time
import tkinter as tk
from tkinter import filedialog, messagebox

TIME_BETWEEN = 20  # segundos entre consultas

def consultar_cnpj(cnpj):
    url = f"https://www.receitaws.com.br/v1/cnpj/{cnpj}"
    try:
        resp = requests.get(url)
        if resp.status_code == 429:
            print("⚠️ Limite atingido. Aguardando 60s...")
            time.sleep(60)
            return consultar_cnpj(cnpj)
        if resp.status_code == 200:
            return resp.json()
        else:
            print(f"Erro HTTP {resp.status_code} no CNPJ {cnpj}")
            return {"status": str(resp.status_code)}
    except Exception as e:
        print(f"Erro geral no CNPJ {cnpj}: {e}")
        return {"status": "Erro"}

def rodar_consulta(caminho_entrada, caminho_saida):
    df = pd.read_excel(caminho_entrada)
    cnpjs = df.iloc[:, 0].dropna().astype(str).str.zfill(14)
    resultados = []

    total = len(cnpjs)
    sucesso = 0
    erro = 0

    for i, cnpj in enumerate(cnpjs, 1):
        print(f"🔍 Consultando {cnpj} ({i}/{total})")
        d = consultar_cnpj(cnpj)

        if d and d.get("status") == "OK":
            res = {
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

        # Backup a cada 50 CNPJs
        if i % 50 == 0:
            backup = pd.DataFrame(resultados)
            backup.to_excel(f"backup_{i}.xlsx", index=False)
            print(f"📁 Backup salvo: backup_{i}.xlsx")

        time.sleep(TIME_BETWEEN)

    df_out = pd.DataFrame(resultados)
    df_out.to_excel(caminho_saida, index=False)
    print("✅ Consulta finalizada.")

    # Mostra resumo
    resumo = (
        f"🔎 Consulta Finalizada!\n\n"
        f"📄 Total: {total}\n"
        f"✅ Sucesso: {sucesso}\n"
        f"❌ Erros: {erro}\n\n"
        f"📁 Arquivo salvo em:\n{caminho_saida}"
    )
    messagebox.showinfo("Resumo da Consulta", resumo)

def selecionar_arquivo():
    return filedialog.askopenfilename(title="Selecione o Excel com os CNPJs", filetypes=[("Arquivos Excel", "*.xlsx")])

def salvar_como():
    return filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Arquivos Excel", "*.xlsx")])

def iniciar():
    entrada = selecionar_arquivo()
    if not entrada:
        return

    saida = salvar_como()
    if not saida:
        return

    rodar_consulta(entrada, saida)

# GUI
janela = tk.Tk()
janela.title("Consulta de CNPJ Automática")
janela.geometry("360x160")

tk.Label(janela, text="Consulta de Cartão CNPJ\nvia ReceitaWS", font=("Arial", 12, "bold")).pack(pady=10)
tk.Button(janela, text="Selecionar arquivo e iniciar", command=iniciar, width=30).pack(pady=10)
tk.Label(janela, text="Feito para uso no trabalho 💼", font=("Arial", 8)).pack(side="bottom", pady=5)

janela.mainloop()
