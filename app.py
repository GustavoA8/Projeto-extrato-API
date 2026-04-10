import customtkinter as ctk
from tkinter import filedialog, messagebox
from PIL import Image, ImageTk
import os

# ===== IMPORTA SEU CÓDIGO =====
from core.reader import (
    read_with_fallback
)


from services.processador import processar_arquivo

# Configurações de tema
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

class BankReportApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        
        self.arquivo_selecionado = None
        
        # Configurações da janela
        self.title("Gerador de Relatórios Bancários")
        self.geometry("700x600")
        self.resizable(False, False)
        
        # Centralizar janela
        self.center_window()
        
        # Criar interface
        self.create_widgets()
    
    def center_window(self):
        """Centraliza a janela na tela"""
        self.update_idletasks()
        width = self.winfo_width()
        height = self.winfo_height()
        x = (self.winfo_screenwidth() // 2) - (width // 2)
        y = (self.winfo_screenheight() // 2) - (height // 2)
        self.geometry(f'{width}x{height}+{x}+{y}')
    
    def create_widgets(self):
        """Cria todos os widgets da interface"""
        
        # ===== CONTAINER PRINCIPAL =====
        main_container = ctk.CTkFrame(self, fg_color="transparent")
        main_container.pack(fill="both", expand=True, padx=30, pady=30)
        
        # ===== CABEÇALHO =====
        header_frame = ctk.CTkFrame(main_container, fg_color="#1a1a2e", corner_radius=15)
        header_frame.pack(fill="x", pady=(0, 25))
        
        title_label = ctk.CTkLabel(
            header_frame,
            text="💼 Gerador de Relatórios Bancários",
            font=ctk.CTkFont(family="Segoe UI", size=28, weight="bold"),
            text_color="#00d9ff"
        )
        title_label.pack(pady=20)
        
        subtitle_label = ctk.CTkLabel(
            header_frame,
            text="Gere relatórios profissionais de forma simples e rápida",
            font=ctk.CTkFont(family="Segoe UI", size=13),
            text_color="#a0a0a0"
        )
        subtitle_label.pack(pady=(0, 15))
        
        # ===== ÁREA DE SELEÇÃO DE BANCO =====
        bank_frame = ctk.CTkFrame(main_container, fg_color="#252547", corner_radius=12)
        bank_frame.pack(fill="x", pady=(0, 20))
        
        bank_label = ctk.CTkLabel(
            bank_frame,
            text="Selecione o Banco",
            font=ctk.CTkFont(family="Segoe UI", size=16, weight="bold"),
            text_color="#ffffff"
        )
        bank_label.pack(pady=(20, 10), padx=20, anchor="w")
        
        # Dropdown de bancos
        self.banco_var = ctk.StringVar(value="Selecione um banco")
        
        self.banco_dropdown = ctk.CTkOptionMenu(
            bank_frame,
            variable=self.banco_var,
            values=["Itau", "Santander", "Bradesco"],
            command=self.on_banco_change,
            font=ctk.CTkFont(family="Segoe UI", size=14),
            dropdown_font=ctk.CTkFont(family="Segoe UI", size=13),
            width=300,
            height=40,
            corner_radius=8,
            fg_color="#3d3d6b",
            button_color="#00d9ff",
            button_hover_color="#00b8d4",
            dropdown_fg_color="#2d2d4f",
            dropdown_hover_color="#3d3d6b"
        )
        self.banco_dropdown.pack(pady=(0, 20), padx=20)

        # ===== CAMPO CONDOMÍNIO (começa oculto) =====

        self.label_condominio = ctk.CTkLabel(
            bank_frame,
           text="Nome do Condomínio",
           font=ctk.CTkFont(family="Segoe UI", size=14),
          text_color="#ffffff"
        )

        self.entry_condominio = ctk.CTkEntry(
            bank_frame,
           width=300,
           height=38,
           corner_radius=8,
          placeholder_text="Digite o nome do condomínio..."
        )

        # Não damos pack → começa invisível

        
        # ===== ÁREA DE SELEÇÃO DE ARQUIVO =====
        file_frame = ctk.CTkFrame(main_container, fg_color="#252547", corner_radius=12)
        file_frame.pack(fill="x", pady=(0, 20))
        
        file_label = ctk.CTkLabel(
            file_frame,
            text="Arquivo do Extrato",
            font=ctk.CTkFont(family="Segoe UI", size=16, weight="bold"),
            text_color="#ffffff"
        )
        file_label.pack(pady=(20, 10), padx=20, anchor="w")
        
        # Container do arquivo
        file_container = ctk.CTkFrame(file_frame, fg_color="transparent")
        file_container.pack(fill="x", padx=20, pady=(0, 20))
        
        self.lbl_arquivo = ctk.CTkLabel(
            file_container,
            text="📄 Nenhum arquivo selecionado",
            font=ctk.CTkFont(family="Segoe UI", size=13),
            text_color="#a0a0a0",
            anchor="w"
        )
        self.lbl_arquivo.pack(side="left", fill="x", expand=True, padx=(0, 10))
        
        btn_selecionar = ctk.CTkButton(
            file_container,
            text="Procurar",
            command=self.escolher_arquivo,
            width=120,
            height=35,
            corner_radius=8,
            font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"),
            fg_color="#4a4a7a",
            hover_color="#5a5a8a"
        )
        btn_selecionar.pack(side="right")
        
        # ===== BOTÃO GERAR RELATÓRIO =====
        btn_gerar = ctk.CTkButton(
            main_container,
            text="🚀 GERAR RELATÓRIO",
            command=self.gerar,
            width=300,
            height=50,
            corner_radius=10,
            font=ctk.CTkFont(family="Segoe UI", size=16, weight="bold"),
            fg_color="#00d9ff",
            hover_color="#00b8d4",
            text_color="#000000"
        )
        btn_gerar.pack(pady=(10, 0))
        # ===== BOTÃO APLICAÇÃO =====
        btn_aplicacao = ctk.CTkButton(
            main_container,
            text="📂 Gerenciar Aplicações",
            command=self.abrir_aplicacoes,
            width=300,
            height=40,
            corner_radius=10,
            font=ctk.CTkFont(family="Segoe UI", size=14, weight="bold"),
            fg_color="#4a4a7a",
            hover_color="#5a5a8a"
        )
        btn_aplicacao.pack(pady=(15, 0))
        
        # ===== RODAPÉ =====
        footer_label = ctk.CTkLabel(
            main_container,
            text="v1.0 | Sistema de Gestão Bancária",
            font=ctk.CTkFont(family="Segoe UI", size=11),
            text_color="#666666"
        )
        footer_label.pack(side="bottom", pady=(20, 0))
    
    def on_banco_change(self, escolha):
        """Mostra ou esconde campo de condomínio e ajusta tamanho da janela"""

        if escolha in ["Santander", "Bradesco"]:
         # Mostra campos
            self.label_condominio.pack(pady=(5, 0), padx=20, anchor="w")
            self.entry_condominio.pack(pady=(0, 15), padx=20)

         # AUMENTA JANELA
            self.geometry("700x670")

        else:
             # Esconde campos
            self.label_condominio.pack_forget()
            self.entry_condominio.pack_forget()

            # VOLTA TAMANHO PADRÃO
            self.geometry("700x600")

    def escolher_arquivo(self):
        """Abre diálogo para escolher arquivo Excel"""
        arquivo = filedialog.askopenfilename(
            title="Selecione o extrato bancário",
            filetypes=[
                ("Arquivos Excel", "*.xls *.xlsx"),
                ("Todos os arquivos", "*.*")
            ]
        )
        
        if arquivo:
            self.arquivo_selecionado = arquivo
            # Exibe apenas o nome do arquivo
            nome_arquivo = os.path.basename(arquivo)
            self.lbl_arquivo.configure(
                text=f"📄 {nome_arquivo}",
                text_color="#00d9ff"
            )
    
    def gerar(self):
        """Processa e gera o relatório"""
        print("Iniciando geração do relatório...")
        print("Caminho: ", self.arquivo_selecionado)
        
        # Validações
        if not self.arquivo_selecionado:
            messagebox.showwarning(
                "Atenção",
                "Por favor, selecione um arquivo de extrato primeiro!"
            )
            return
        
        banco = self.banco_var.get()
        
        if banco == "Selecione um banco":
            messagebox.showwarning(
                "Atenção",
                "Por favor, selecione um banco!"
            )
            return
        
        condominio = self.entry_condominio.get().strip()

        if banco in ["Santander", "Bradesco"] and not condominio:
          messagebox.showwarning(
             "Atenção",
                "Informe o nome do condomínio!"
            )
          return 
        
        try:
            # Desabilita botão durante processamento
            self.disable_generate_button()
            
            # Lê o arquivo
            df = read_with_fallback(self.arquivo_selecionado)
            df.attrs["arquivo_origem"] = self.arquivo_selecionado
            
            processar_arquivo(
                df=df,
                banco=banco.lower(),
                condominio=condominio
            )           
            mensagem = "Relatório gerado com sucesso! ✅"
            
            # Sucesso
            messagebox.showinfo("Sucesso", mensagem)
            
            # Reseta interface
            self.reset_interface()
            
        except Exception as e:
            messagebox.showerror(
                "Erro",
                f"Erro ao processar o arquivo:\n\n{str(e)}\n\nVerifique se o arquivo está no formato correto."
            )
        
        finally:
            # Reabilita botão
            self.enable_generate_button()
    
    def disable_generate_button(self):
        """Desabilita o botão de gerar durante processamento"""
        for widget in self.winfo_children():
            if isinstance(widget, ctk.CTkFrame):
                for child in widget.winfo_children():
                    if isinstance(child, ctk.CTkButton) and "GERAR" in child.cget("text"):
                        child.configure(state="disabled", text="⏳ Processando...")
    
    def enable_generate_button(self):
        """Reabilita o botão de gerar"""
        for widget in self.winfo_children():
            if isinstance(widget, ctk.CTkFrame):
                for child in widget.winfo_children():
                    if isinstance(child, ctk.CTkButton) and "Processando" in child.cget("text"):
                        child.configure(state="normal", text="🚀 GERAR RELATÓRIO")
    
    def reset_interface(self):
        """Reseta a interface após gerar relatório"""
        self.arquivo_selecionado = None
        self.lbl_arquivo.configure(
            text="📄 Nenhum arquivo selecionado",
            text_color="#a0a0a0"
        )
        self.banco_var.set("Selecione um banco")
    
    def abrir_aplicacoes(self):
        import pandas as pd
        import sys
        import os
        from tkinter import ttk
    
        janela = ctk.CTkToplevel(self)
        janela.title("Gerenciar Aplicações")
        janela.geometry("800x600")
        janela.grab_set()

        if getattr(sys, 'frozen', False):
            base_path = os.path.dirname(sys.executable)
        else:
            base_path = os.path.dirname(os.path.abspath(__file__))

        caminho = os.path.join(base_path, "grupo-aplic.xlsx")
        print("CAMINHO REAL:", caminho)

        # ===== TÍTULO =====
        titulo = ctk.CTkLabel(
            janela,
            text="Gerenciar Grupos de Aplicação",
            font=ctk.CTkFont(size=22, weight="bold"),
            text_color="#00d9ff"
        )
        titulo.pack(pady=20)

        # ===== DROPDOWN BANCO =====
        banco_var = ctk.StringVar(value="Itau")

        dropdown = ctk.CTkOptionMenu(
            janela,
            variable=banco_var,
            values=["Itau", "Bradesco", "Santander"],
            width=200
        )
        dropdown.pack(pady=10)

        # ===== TABELA =====
        frame_tabela = ctk.CTkFrame(janela)
        frame_tabela.pack(fill="both", expand=True, padx=20, pady=20)

        colunas = ("Descrição", "Grupo")
        tree = ttk.Treeview(frame_tabela, columns=colunas, show="headings")

        for col in colunas:
            tree.heading(col, text=col)
            tree.column(col, width=300)

        tree.pack(fill="both", expand=True)

        # ===== FUNÇÃO CARREGAR =====
        def carregar():
            tree.delete(*tree.get_children())
            banco = banco_var.get()
            df = pd.read_excel(caminho, sheet_name=banco)

            for _, row in df.iterrows():
                tree.insert("", "end", values=(row["Descrição"], row["Grupo"]))

        dropdown.configure(command=lambda e: carregar())

        # ===== ÁREA ADICIONAR =====
        frame_add = ctk.CTkFrame(janela)
        frame_add.pack(pady=10)

        entry_desc = ctk.CTkEntry(frame_add, placeholder_text="Descrição", width=250)
        entry_desc.grid(row=0, column=0, padx=5)

        entry_grupo = ctk.CTkEntry(frame_add, placeholder_text="Grupo", width=200)
        entry_grupo.grid(row=0, column=1, padx=5)

        def adicionar():
            desc = entry_desc.get().strip().upper()
            grupo = entry_grupo.get().strip()

            if not desc or not grupo:
                messagebox.showwarning("Aviso", "Preencha todos os campos.")
                return

            banco = banco_var.get()
            df = pd.read_excel(caminho, sheet_name=banco)

            if desc in df["Descrição"].values:
                messagebox.showwarning("Aviso", "Descrição já cadastrada.")
                return

            novo = pd.DataFrame([{"Descrição": desc, "Grupo": grupo}])
            df = pd.concat([df, novo], ignore_index=True)

            with pd.ExcelWriter(caminho, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
                df.to_excel(writer, sheet_name=banco, index=False)

            carregar()
            entry_desc.delete(0, "end")
            entry_grupo.delete(0, "end")

        btn_add = ctk.CTkButton(
            frame_add,
            text="Adicionar",
            command=adicionar,
            fg_color="green"
        )
        btn_add.grid(row=0, column=2, padx=5)

        # ===== EXCLUIR =====
        def excluir():
            selecionado = tree.selection()
            if not selecionado:
                messagebox.showwarning("Aviso", "Selecione um item.")
                return

            valores = tree.item(selecionado)["values"]
            desc = valores[0]

            banco = banco_var.get()
            df = pd.read_excel(caminho, sheet_name=banco)

            df = df[df["Descrição"] != desc]

            with pd.ExcelWriter(caminho, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
                df.to_excel(writer, sheet_name=banco, index=False)

            carregar()

        btn_del = ctk.CTkButton(
            janela,
            text="Excluir Selecionado",
            command=excluir,
            fg_color="red"
        )
        btn_del.pack(pady=10)

        carregar()


def main():
    """Função principal para executar a aplicação"""
    app = BankReportApp()
    app.mainloop()


if __name__ == "__main__":
    main()