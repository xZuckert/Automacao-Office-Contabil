from time import sleep
from pywinauto import Application
import pandas as pd
import os
import json
import ctypes
from datetime import datetime


PARAR_AUTOMACAO = False
DELAY = 0.2
PASTA_LOG = None
CHECKPOINT_FILE = None
CONFIG_LANCAMENTOS = {
    "saida": {
        "provisao": {"deb": 5, "cred": 104},
        "pagamento": {"deb": 1, "cred": 5}
    },
    "servico": {
        "provisao": {"deb": 5, "cred": 377},
        "pagamento": {"deb": 1, "cred": 5}
    }
}


# CONEXÃO
def conectarOffice():
    try:
        app = Application(backend="uia").connect(title_re="^IOB Office Contábil")
        janela = app.window(title_re="^IOB Office Contábil")
        return janela
    except Exception:
        raise Exception("Office Contábil não está aberto.")


# LOG
def registrarLog(msg):
    global PASTA_LOG
    if not PASTA_LOG:
        return

    caminho_log = os.path.join(PASTA_LOG, "log_automacao.txt")

    with open(caminho_log, "a", encoding="utf-8") as log:
        log.write(f"{datetime.now()} - {msg}\n")


# BLOQUEIO INPUT
def bloquearInput():
    ctypes.windll.user32.BlockInput(True)

def desbloquearInput():
    ctypes.windll.user32.BlockInput(False)


# PREPARAÇÃO AMBIENTE
def prepararAmbiente(janela, caminhoExcel):

    pasta = os.path.dirname(caminhoExcel)
    codigoEmpresa = os.path.basename(pasta)

    nomeArquivo = os.path.basename(caminhoExcel)
    base = os.path.splitext(nomeArquivo)[0]

    mes = base[-6:-4]
    ano = base[-4:]

    janela.set_focus()

    # Sai de qualquer tela
    for i in range(5):
        janela.type_keys("{ESC}")
        sleep(0.2)

    janela.type_keys("E")  # Ativar empresa
    sleep(0.5)

    janela.type_keys(codigoEmpresa)
    janela.type_keys("{ENTER}")

    janela.type_keys(f"{mes}{ano}")
    janela.type_keys("{ENTER}")

    janela.type_keys("{TAB}")  # pular digitador
    janela.type_keys("{ENTER}")

    sleep(0.5)

    janela.type_keys("D")  # Digitação
    sleep(0.5)


# LANÇAMENTOS
def lancamento(janela, dia, valor, numero, debito, credito):

    janela.type_keys("N")
    sleep(DELAY)

    janela.type_keys("{TAB}")
    janela.type_keys(str(dia))
    janela.type_keys("{ENTER}")

    janela.type_keys("{TAB}")

    janela.type_keys(str(debito))
    janela.type_keys("{ENTER}")

    janela.type_keys(str(credito))
    janela.type_keys("{ENTER}")

    janela.type_keys(str(valor))
    janela.type_keys("{ENTER}")

    janela.type_keys(str(numero))
    janela.type_keys("{PGDN}")

    sleep(DELAY)

# EXECUÇÃO PRINCIPAL
def executarAutomacao(caminhoExcel, tipo, progressCallback=None):

    global PARAR_AUTOMACAO, PASTA_LOG, CHECKPOINT_FILE
    PARAR_AUTOMACAO = False
    PASTA_LOG = os.path.dirname(caminhoExcel)
    CHECKPOINT_FILE = os.path.join(PASTA_LOG, "checkpoint.json")

    try:
        df = pd.read_excel(caminhoExcel)

        config = CONFIG_LANCAMENTOS[tipo]

        debProv = config["provisao"]["deb"]
        credProv = config["provisao"]["cred"]

        debPag = config["pagamento"]["deb"]
        credPag = config["pagamento"]["cred"]
        df = df[df["Dia"].notna()]

        janela = conectarOffice()
        prepararAmbiente(janela, caminhoExcel)

        total = len(df)

        # CHECKPOINT
        indiceInicial = 0

        if os.path.exists(CHECKPOINT_FILE):
            with open(CHECKPOINT_FILE, "r") as f:
                indiceInicial = json.load(f).get("indice", 0)

        bloquearInput()

        for i in range(indiceInicial, total):

            if PARAR_AUTOMACAO:
                registrarLog("Interrompido pelo usuário")
                break

            row = df.iloc[i]

            dia = int(row["Dia"])
            valor = float(row["valor contabil (R$)"])
            numero = row["número"]

            registrarLog(f"Lançando dia {dia}")

            lancamento(janela, dia, valor, numero, debProv, credProv) # Provisão
            lancamento(janela, dia, valor, numero, debPag, credPag) # Pagamento

            # Atualiza checkpoint
            with open(CHECKPOINT_FILE, "w") as f:
                json.dump({"indice": i + 1}, f)

            if progressCallback:
                progressCallback(i + 1, total)

        # Se finalizou sem parar
        if not PARAR_AUTOMACAO:
            registrarLog("Automação finalizada com sucesso")
            if os.path.exists(CHECKPOINT_FILE):
                os.remove(CHECKPOINT_FILE)

    except Exception as e:
        registrarLog(f"Erro: {e}")
        print("Erro na automação:", e)

    finally:
        desbloquearInput()

# PARAR
def pararAutomacao():
    global PARAR_AUTOMACAO
    PARAR_AUTOMACAO = True
