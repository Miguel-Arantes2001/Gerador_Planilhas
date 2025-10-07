import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
from pathlib import Path
from openpyxl import load_workbook
import json

# Função para ajustar largura das colunas
def ajustar_colunas(arquivo):
    wb = load_workbook(arquivo)
    ws = wb.active
    ws.column_dimensions['A'].width = 20   # Telefone
    ws.column_dimensions['B'].width = 50   # Estabelecimento
    wb.save(arquivo)

# Função para carregar progresso do controle.json
def carregar_progresso(caminho_original):
    controle_path = Path(caminho_original).with_suffix(".json")
    if controle_path.exists():
        with open(controle_path, "r", encoding="utf-8") as f:
            return json.load(f)
    return {"ultima_posicao": 0, "total_planilhas_geradas": 0}

# Função para salvar progresso no controle.json
def salvar_progresso(caminho_original, ultima_posicao, total_planilhas_geradas):
    controle_path = Path(caminho_original).with_suffix(".json")
    with open(controle_path, "w", encoding="utf-8") as f:
        json.dump({"ultima_posicao": ultima_posicao, "total_planilhas_geradas": total_planilhas_geradas}, f)

# Função para corrigir nomes das colunas
def corrigir_colunas(df, caminho):
    def salvar_colunas(df_local):
        nova_col_telefone = entrada_telefone.get().strip()
        nova_col_estabelecimento = entrada_estabelecimento.get().strip()
        if not nova_col_telefone or not nova_col_estabelecimento:
            messagebox.showerror("Erro", "Ambos os campos devem ser preenchidos!")
            return

        # Verifica se os nomes das colunas existem no DataFrame
        colunas_existentes = df_local.columns.tolist()
        if nova_col_telefone not in colunas_existentes or nova_col_estabelecimento not in colunas_existentes:
            messagebox.showerror("Erro", "Um ou ambos os nomes de colunas não existem na planilha! Verifique os nomes e tente novamente.")
            return

        janela_corrigir.destroy()  # Fecha a janela pop-up
        # Renomeia as colunas com base no que o usuário digitou
        df_local.rename(columns={nova_col_telefone: 'telefone'}, inplace=True)
        df_local.rename(columns={nova_col_estabelecimento: '{{estabelecimento}}'}, inplace=True)
        # Define as colunas na ordem desejada
        df_local = df_local[['telefone', '{{estabelecimento}}']]
        # Continua o processamento
        gerar_planilhas_continuar(df_local, caminho)

    janela_corrigir = tk.Toplevel()
    janela_corrigir.title("Corrigir Colunas")
    janela_corrigir.geometry("500x400")
    janela_corrigir.configure(bg='#32CD32')  # Fundo lime escuro

    # Labels e Entrys
    ttk.Label(janela_corrigir, text="Nome da coluna telefones :", style='Custom.TLabel').pack(pady=10)
    entrada_telefone = ttk.Entry(janela_corrigir, style='Custom.TEntry', width=20)
    entrada_telefone.pack(pady=5)

    ttk.Label(janela_corrigir, text="Nome da coluna estabelecimentos :", style='Custom.TLabel').pack(pady=10)
    entrada_estabelecimento = ttk.Entry(janela_corrigir, style='Custom.TEntry', width=20)
    entrada_estabelecimento.pack(pady=5)

    # Botão para salvar
    ttk.Button(
        janela_corrigir,
        text="Salvar e Continuar",
        command=lambda: salvar_colunas(df),
        style='Custom.TButton'
    ).pack(pady=20)

    janela_corrigir.transient(janela)  # Faz a janela pop-up ficar sobre a janela principal
    janela_corrigir.grab_set()  # Bloqueia a interação com a janela principal até fechar
    janela_corrigir.wait_window()  # Pausa até a janela ser fechada

# Função auxiliar para continuar o processamento após correção
def gerar_planilhas_continuar(df, caminho):
    try:
        # Carregar progresso anterior
        progresso = carregar_progresso(caminho)
        ultima_posicao = progresso["ultima_posicao"]
        total_planilhas_geradas = progresso["total_planilhas_geradas"]

        # Perguntar parâmetros ao usuário
        num_planilhas = int(num_planilhas_entry.get())
        contatos_por_arquivo = int(contatos_por_entry.get())

        pasta_saida = Path(caminho).parent / "planilhas_disparo"
        pasta_saida.mkdir(exist_ok=True)

        total_gerados = 0
        num_geradas = 0  # Contador de planilhas geradas nesta execução
        for i in range(num_planilhas):
            inicio = ultima_posicao + (i * contatos_por_arquivo)
            fim = min(inicio + contatos_por_arquivo, len(df))
            if inicio >= len(df):
                break
            disparo = df.iloc[inicio:fim]

            if disparo.empty:
                break

            arquivo_saida = pasta_saida / f"disparo_limao_{total_planilhas_geradas + num_geradas + 1}.xlsx"  # Nome único
            # Salva com colunas na ordem desejada
            disparo = disparo[['telefone', '{{estabelecimento}}']]
            disparo.to_excel(arquivo_saida, index=False)
            ajustar_colunas(arquivo_saida)
            total_gerados += len(disparo)
            num_geradas += 1  # Incrementa apenas se gerou a planilha
            print(f"Gerada planilha {total_planilhas_geradas + num_geradas}: linhas {inicio} a {fim-1}, contatos {len(disparo)}")  # Depuração

        # Atualizar progresso
        novo_total_planilhas_geradas = total_planilhas_geradas + num_geradas
        salvar_progresso(caminho, ultima_posicao + total_gerados, novo_total_planilhas_geradas)

        messagebox.showinfo(
            "Sucesso",
            f"{num_geradas} planilhas geradas.\n"
            f"{total_gerados} contatos usados.\n"
            f"Progresso salvo em {Path(caminho).with_suffix('.json').name}"
        )

    except Exception as e:
        messagebox.showerror("Erro", f"Ocorreu um erro: {str(e)}")

# Função principal
def gerar_planilhas():
    caminho = filedialog.askopenfilename(
        title="Selecione a planilha original",
        filetypes=[("Arquivos Excel", "*.xlsx *.xls")]
    )
    if not caminho:
        return

    try:
        # Carregar dados
        df = pd.read_excel(caminho, dtype=str)

        # Normalizar colunas para garantir consistência
        colunas_necessarias = {
            'Telefone 1': 'telefone',
            'Nomes corrigidos': '{{estabelecimento}}'
        }
        df = df.rename(columns=colunas_necessarias)

        if not {'telefone', '{{estabelecimento}}'}.issubset(df.columns):
            # Chama a janela de correção em vez de apenas exibir erro
            corrigir_colunas(df, caminho)
        else:
            gerar_planilhas_continuar(df, caminho)

    except Exception as e:
        messagebox.showerror("Erro", f"Ocorreu um erro: {str(e)}")

# ---------------- INTERFACE ----------------
janela = tk.Tk()
janela.title("Gerador de Planilhas de Disparo")
janela.geometry('600x600')

# Definir cor de fundo da janela (lime escuro)
janela.configure(bg='#32CD32')

# Criar estilo ttk
style = ttk.Style()
style.theme_use('clam')

# Estilizar Label
style.configure('Custom.TLabel',
                background='#32CD32',
                foreground='#000000',
                font=('Arial', 14, 'bold'))

# Estilizar Entry
style.configure('Custom.TEntry',
                fieldbackground='#FFFFFF',
                foreground='#000000',
                font=('Arial', 14))

# Estilizar Botão
style.configure('Custom.TButton',
                background='#006400',
                foreground='white',
                font=('Arial', 16, 'bold'),
                padding=10)
style.map('Custom.TButton',
          background=[('active', '#004d00')],
          foreground=[('active', 'white')])

# Frame para centralizar
frame = ttk.Frame(janela, style='Custom.TFrame')
style.configure('Custom.TFrame', background='#32CD32')
frame.pack(expand=True, padx=50, pady=50)

# Título
ttk.Label(
    frame,
    text="Gerador de Planilhas",
    style='Custom.TLabel',
    font=('Arial', 18, 'bold')
).pack(pady=20)

# Label e Entry 1
ttk.Label(
    frame,
    text="Número de planilhas a gerar:",
    style='Custom.TLabel'
).pack(pady=15)
num_planilhas_entry = ttk.Entry(
    frame,
    style='Custom.TEntry',
    width=12
)
num_planilhas_entry.pack(pady=10)
num_planilhas_entry.insert(0, "5")

# Label e Entry 2
ttk.Label(
    frame,
    text="Contatos por planilha:",
    style='Custom.TLabel'
).pack(pady=15)
contatos_por_entry = ttk.Entry(
    frame,
    style='Custom.TEntry',
    width=12
)
contatos_por_entry.pack(pady=10)
contatos_por_entry.insert(0, "250")

# Botão
ttk.Button(
    frame,
    text="Selecionar Planilha e Gerar",
    command=gerar_planilhas,
    style='Custom.TButton'
).pack(pady=30)

janela.mainloop()
