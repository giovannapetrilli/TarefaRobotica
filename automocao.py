import pyautogui
import pandas as pd
import time
from openpyxl import Workbook
from datetime import datetime


posicoes = {
    "navegador_icone": (100, 200)  
}

relatorio = []

def executar_tarefa(tarefa, tipo, dado):
    status = "Sucesso"
    inicio = datetime.now()

    try:
        if tipo == "click":
            pos = posicoes.get(dado)
            if pos:
                pyautogui.click(pos)
            else:
                raise ValueError(f"Posição não encontrada: {dado}")

        elif tipo == "texto":
            pyautogui.write(dado)

        elif tipo == "tecla":
            pyautogui.press(dado)

        elif tipo == "espera":
            time.sleep(int(dado))

        else:
            raise ValueError(f"Tipo de tarefa desconhecido: {tipo}")

    except Exception as e:
        status = f"Erro: {e}"

    fim = datetime.now()
    duracao = (fim - inicio).total_seconds()

    relatorio.append({
        "Tarefa": tarefa,
        "Tipo": tipo,
        "Dado": dado,
        "Status": status,
        "Tempo (s)": duracao
    })

def gerar_relatorio():
    df = pd.DataFrame(relatorio)
    df.to_excel("relatorio_tarefas.xlsx", index=False)

def main():
    tarefas = pd.read_csv("tarefas.csv")

    for index, linha in tarefas.iterrows():
        executar_tarefa(linha["Tarefa"], linha["Tipo"], linha["Dado"])

    gerar_relatorio()

if __name__ == "__main__":
    print("Iniciando assistente virtual...")
    time.sleep(3)  
    main()