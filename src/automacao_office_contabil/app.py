import os
import tkinter as tk
import threading
from tkinter import filedialog, messagebox, ttk
from .reader import processarSaida, processarServico, detectarTipoArquivo
from .xlGenerator import gerarExcelSaida, gerarExcelServico
from .automation import executarAutomacao, pararAutomacao


class RelatorioApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Gerador de Relatorio Contábil")
        self.root.geometry("500x300")
        self.root.resizable(False, False)

        # Mantém sempre na frente
        self.root.attributes("-topmost", True)

        self.caminhoTXT = ""
        self.threadAutomacao = None

        # Título
        titulo = tk.Label(root, text="Gerador de Relatório Contábil", font=("Arial", 16, "bold"))
        titulo.pack(pady=10)

        # Botão selecionar
        btnSelec = tk.Button(root, text="Selecionar Arquivo TXT", command=self.selecArq, width=30)
        btnSelec.pack(pady=10)

        # Label caminho
        self.LabelCaminho = tk.Label(root, text="Nenhum arquivo selecionado", wraplength=450)
        self.LabelCaminho.pack(pady=5)

        # Botão gerar
        btnIniciar = tk.Button(root, text="Executar Automação", command=self.iniciarAutomacao, width=30, bg="#1F4E79", fg="white")
        btnIniciar.pack(pady=20)

        # Botão Parar
        btnParar = tk.Button(root, text="Parar Automação", command=self.pararAutom, width=30, bg="red", fg="white")
        btnParar.pack(pady=5)

        # Progresso
        self.progress = ttk.Progressbar(root, length=400, mode="determinate")
        self.progress.pack(pady=10)

        # Status
        self.label_status = tk.Label(root, text="")
        self.label_status.pack(pady=5)

    # Seleciona o txt
    def selecArq(self):
        caminho = filedialog.askopenfilename(
            title="Selecione o arquivo",
            filetypes=[("Arquivos TXT", "*.txt")]
        )

        if caminho:
            self.caminhoTXT = caminho
            self.LabelCaminho.config(text=caminho)

    # Iniciar o programa
    def iniciarAutomacao(self):
        if not self.caminhoTXT:
            messagebox.showwarning("Aviso", "Selecione um arquivo primeiro.")
            return

        try:
            # Processa TXT
            pasta = os.path.dirname(self.caminhoTXT)
            nomeBase = os.path.splitext(os.path.basename(self.caminhoTXT))[0]
            caminhoExcel = os.path.join(pasta, f"{nomeBase}.xlsx")

            # Detecta tipo automaticamente
            tipo = detectarTipoArquivo(self.caminhoTXT)

            # Processa e gera Excel conforme tipo
            if tipo == "saida":
                dados = processarSaida(self.caminhoTXT)
                gerarExcelSaida(dados, caminhoExcel)
            else:
                dados = processarServico(self.caminhoTXT)
                gerarExcelServico(dados, caminhoExcel)

            self.label_status.config(text="Executando...", fg="green")

            # Inicia automação em thread
            self.threadAutomacao = threading.Thread(
                target=self.executarComTratamento,
                args=(caminhoExcel, tipo),
                daemon=True
            )
            self.threadAutomacao.start()

        except Exception as e:
            messagebox.showerror("Erro", str(e))

    # Wrapper de segurança para rodar a automação em um thread separada
    def executarComTratamento(self, caminhoExcel, tipo):
        try: # Executa a automação
            executarAutomacao(caminhoExcel, tipo, self.atualizarProgresso)
        except Exception as e: # Captura erros
            self.root.after(0, lambda:
                messagebox.showerror("Erro na Automação", str(e)))
        finally: # Atualiza a interface
            # Quando terminar (normal ou parada)
            self.root.after(0, self.finalizarUI)

    # Remove a mensagem de pausa ao ocorrer a pausa
    def finalizarUI(self):
        self.label_status.config(text="")
        self.progress["value"] = 0

    #  Solicita a pausa da automação
    def pararAutom(self):
        pararAutomacao()
        self.label_status.config(text="Solicitação de parada enviada",
                                 fg="orange")

    # Atualiza a barra de progresso com o deccorrer dos lancamentos
    def atualizarProgresso(self, atual, total):
        self.progress["maximum"] = total
        self.progress["value"] = atual
        self.root.update_idletasks()

# Executa o programa
if __name__ == "__main__":
    root = tk.Tk()
    app = RelatorioApp(root)
    root.mainloop()
