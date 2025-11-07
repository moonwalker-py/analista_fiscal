import tkinter
from tkinter import filedialog, ttk
import customtkinter as ctk
import xml.etree.ElementTree as ET
import zipfile
import rarfile
import os
from collections import defaultdict

# Define a aparência padrão do CustomTkinter
ctk.set_appearance_mode("system")
ctk.set_default_color_theme("blue")

# Variável global para armazenar detalhes dos itens (Produtos/Serviços)
# Estrutura: {chave_doc: [{nItem: 1, xProd: '...', NCM: '...', CFOP: '...'}, ...]}
GLOBAL_ITEM_DETAILS = defaultdict(list)

class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        # VERSÃO ATUALIZADA
        self.title("Analisador Fiscal de Documentos Eletrônicos (NF-e / CT-e / NFSE) v4.0.0")
        self.geometry("1400x980")
        
        # Variáveis de Estado
        self.documentos_processados = {}
        self.arquivos_contados = 0
        self.cfop_totals = defaultdict(float) 
        self.nfse_item_totals = defaultdict(float) 
        self.partner_totals = defaultdict(float) 
        self.cnpj_empresa_var = ctk.StringVar(value="") # Armazena o CNPJ da empresa (Input do usuário)
        # Limpa o global ao iniciar
        GLOBAL_ITEM_DETAILS.clear() 

        # Configuração do Grid Principal da Janela: 
        # Linha 0: Cabeçalho Fixo
        # Linha 1: Resultados Roláveis
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=0) # Linha 0 (Top Frame) não se expande
        self.grid_rowconfigure(1, weight=1) # Linha 1 (Scrollable Frame) se expande

        # --- Frame Superior (Controles e Resumos) - FIXO ---
        self.top_frame = ctk.CTkFrame(self, height=180) 
        self.top_frame.grid(row=0, column=0, padx=10, pady=(10, 0), sticky="ew")
        self.top_frame.grid_columnconfigure((0, 1, 2, 3), weight=1)

        # Entrada de CNPJ da Empresa (Fila 0)
        self.label_cnpj_empresa = ctk.CTkLabel(self.top_frame, text="CNPJ da Empresa (Apenas números):", font=ctk.CTkFont(size=13, weight="bold"))
        self.label_cnpj_empresa.grid(row=0, column=2, padx=(10, 5), pady=10, sticky="e")
        self.cnpj_entry = ctk.CTkEntry(self.top_frame, textvariable=self.cnpj_empresa_var, placeholder_text="Ex: 12345678000190", width=180)
        self.cnpj_entry.grid(row=0, column=3, padx=(5, 10), pady=10, sticky="w")


        # Botões (Fila 1)
        self.select_button = ctk.CTkButton(self.top_frame, text="Selecionar Arquivos ou Pasta", command=self.selecionar_e_processar_arquivos)
        self.select_button.grid(row=1, column=0, padx=10, pady=(0, 10), sticky="ew")
        self.clear_button = ctk.CTkButton(self.top_frame, text="Limpar Resultados", command=self.limpar_resultados)
        self.clear_button.grid(row=1, column=1, padx=10, pady=(0, 10), sticky="ew")
        
        # Labels de Totais (Fila 2 e 3)
        self.label_valor_total = ctk.CTkLabel(self.top_frame, text="Valor Total (Autorizados): R$ 0,00", font=ctk.CTkFont(size=14, weight="bold"))
        self.label_valor_total.grid(row=2, column=0, padx=10, pady=5, columnspan=2, sticky="w")
        
        self.label_total_autorizadas = ctk.CTkLabel(self.top_frame, text="Docs Autorizados (NF-e/CT-e/NFSE): 0 / 0 / 0", font=ctk.CTkFont(size=14))
        self.label_total_autorizadas.grid(row=3, column=0, padx=10, pady=5, sticky="w")
        self.label_total_canceladas = ctk.CTkLabel(self.top_frame, text="Docs Cancelados (NF-e/CT-e/NFSE): 0 / 0 / 0", font=ctk.CTkFont(size=14))
        self.label_total_canceladas.grid(row=3, column=1, padx=10, pady=5, sticky="w")
        self.label_arquivos_processados = ctk.CTkLabel(self.top_frame, text="Arquivos Físicos Processados: 0", font=ctk.CTkFont(size=14))
        self.label_arquivos_processados.grid(row=3, column=2, padx=10, pady=5, columnspan=2, sticky="e")


        # --- Frame Rolável (Resultados) ---
        self.results_scrollable_frame = ctk.CTkScrollableFrame(self)
        self.results_scrollable_frame.grid(row=1, column=0, padx=10, pady=10, sticky="nsew")
        self.results_scrollable_frame.grid_columnconfigure(0, weight=1)
        
        # Estilos para Treeview (mantido)
        style = ttk.Style()
        style.theme_use("default")
        style.configure("Treeview", background="#2b2b2b", foreground="white", rowheight=25, fieldbackground="#2b2b2b", bordercolor="#333333", borderwidth=0)
        style.map('Treeview', background=[('selected', '#1f6aa5')])
        style.configure("Treeview.Heading", background="#565b5e", foreground="white", relief="flat")
        style.map("Treeview.Heading", background=[('active', '#3484D2')])


        # --- Frame Tabela Principal (Filho do Scrollable Frame) ---
        self.main_table_frame = ctk.CTkFrame(self.results_scrollable_frame)
        self.main_table_frame.grid(row=0, column=0, padx=0, pady=(0, 10), sticky="ew") # row 0 dentro do Scrollable Frame
        self.main_table_frame.grid_rowconfigure(0, weight=1)
        self.main_table_frame.grid_columnconfigure(0, weight=1)

        # Tabela Principal
        self.tree = ttk.Treeview(self.main_table_frame, columns=("Chave", "Tipo", "Nº Doc", "Valor", "Status", "Fluxo", "Parceiro", "CNPJ/CPF", "Inconsistência", "Arquivo"), show="headings", height=10) 
        self.tree.heading("Chave", text="Chave do Documento")
        self.tree.heading("Tipo", text="Tipo")
        self.tree.heading("Nº Doc", text="Nº Doc") # NOVO
        self.tree.heading("Valor", text="Valor (R$)")
        self.tree.heading("Status", text="Status")
        self.tree.heading("Fluxo", text="Fluxo") 
        self.tree.heading("Parceiro", text="Parceiro (Dest/Tomador)")
        self.tree.heading("CNPJ/CPF", text="CNPJ/CPF")
        self.tree.heading("Inconsistência", text="Inconsistência Fiscal")
        self.tree.heading("Arquivo", text="Origem")
        self.tree.column("Chave", width=180) 
        self.tree.column("Tipo", width=50, anchor='center')
        self.tree.column("Nº Doc", width=70, anchor='center') # NOVO
        self.tree.column("Valor", width=100, anchor='e')
        self.tree.column("Status", width=80, anchor='center')
        self.tree.column("Fluxo", width=130, anchor='center')
        self.tree.column("Parceiro", width=200)
        self.tree.column("CNPJ/CPF", width=120, anchor='center')
        self.tree.column("Inconsistência", width=120, anchor='center')
        self.tree.column("Arquivo", width=180)
        self.tree.grid(row=0, column=0, sticky="ew")
        
        scrollbar = ctk.CTkScrollbar(self.main_table_frame, command=self.tree.yview)
        scrollbar.grid(row=0, column=1, sticky='ns')
        self.tree.configure(yscrollcommand=scrollbar.set)
        
        
        # --- Frame Produtos/Itens (Filho do Scrollable Frame) - NOVO ---
        self.itens_frame = ctk.CTkFrame(self.results_scrollable_frame)
        self.itens_frame.grid(row=1, column=0, padx=0, pady=(10, 10), sticky="ew") 
        self.itens_frame.grid_columnconfigure(0, weight=1)

        self.label_itens_title = ctk.CTkLabel(self.itens_frame, text="Detalhes de Itens (Produtos/Serviços) de Documentos Autorizados", font=ctk.CTkFont(size=14, weight="bold"))
        self.label_itens_title.grid(row=0, column=0, padx=10, pady=(10, 5), sticky="w")

        self.tree_itens = ttk.Treeview(self.itens_frame, columns=("Nº Doc", "Produto", "NCM", "CFOP", "Valor", "CST", "vICMS"), show="headings", height=10)
        self.tree_itens.heading("Nº Doc", text="Nº Doc")
        self.tree_itens.heading("Produto", text="Produto / Serviço")
        self.tree_itens.heading("NCM", text="NCM")
        self.tree_itens.heading("CFOP", text="CFOP")
        self.tree_itens.heading("Valor", text="Valor Item (R$)")
        self.tree_itens.heading("CST", text="CST/CSOSN")
        self.tree_itens.heading("vICMS", text="vICMS (R$)")
        self.tree_itens.column("Nº Doc", width=70, anchor='center')
        self.tree_itens.column("Produto", width=300)
        self.tree_itens.column("NCM", width=100, anchor='center')
        self.tree_itens.column("CFOP", width=70, anchor='center')
        self.tree_itens.column("Valor", width=100, anchor='e')
        self.tree_itens.column("CST", width=80, anchor='center')
        self.tree_itens.column("vICMS", width=90, anchor='e')
        self.tree_itens.grid(row=1, column=0, padx=10, pady=(0, 10), sticky="ew")
        
        itens_scrollbar = ctk.CTkScrollbar(self.itens_frame, command=self.tree_itens.yview)
        itens_scrollbar.grid(row=1, column=1, sticky='ns')
        self.tree_itens.configure(yscrollcommand=itens_scrollbar.set)
        
        # --- Frame CFOP (Filho do Scrollable Frame) ---
        self.cfop_frame = ctk.CTkFrame(self.results_scrollable_frame)
        self.cfop_frame.grid(row=2, column=0, padx=0, pady=(0, 10), sticky="ew") # row 2 dentro do Scrollable Frame
        self.cfop_frame.grid_columnconfigure(0, weight=1)

        self.label_cfop_title = ctk.CTkLabel(self.cfop_frame, text="Totalização de Valores por CFOP (Apenas NF-e Autorizadas)", font=ctk.CTkFont(size=14, weight="bold"))
        self.label_cfop_title.grid(row=0, column=0, padx=10, pady=(10, 5), sticky="w")

        self.tree_cfop = ttk.Treeview(self.cfop_frame, columns=("CFOP", "Valor"), show="headings", height=5)
        self.tree_cfop.heading("CFOP", text="CFOP")
        self.tree_cfop.heading("Valor", text="Valor Total (R$)")
        self.tree_cfop.column("CFOP", width=100, anchor='center')
        self.tree_cfop.column("Valor", width=150, anchor='e')
        self.tree_cfop.grid(row=1, column=0, padx=10, pady=(0, 10), sticky="ew")
        
        cfop_scrollbar = ctk.CTkScrollbar(self.cfop_frame, command=self.tree_cfop.yview)
        cfop_scrollbar.grid(row=1, column=1, sticky='ns')
        self.tree_cfop.configure(yscrollcommand=cfop_scrollbar.set)
        
        # --- Frame NFSE (Filho do Scrollable Frame) ---
        self.nfse_frame = ctk.CTkFrame(self.results_scrollable_frame)
        self.nfse_frame.grid(row=3, column=0, padx=0, pady=(0, 10), sticky="ew") # row 3 dentro do Scrollable Frame
        self.nfse_frame.grid_columnconfigure(0, weight=1)

        self.label_nfse_title = ctk.CTkLabel(self.nfse_frame, text="Totalização de Valores por Item de Serviço (Apenas NFSE Autorizadas)", font=ctk.CTkFont(size=14, weight="bold"))
        self.label_nfse_title.grid(row=0, column=0, padx=10, pady=(10, 5), sticky="w")

        self.tree_nfse = ttk.Treeview(self.nfse_frame, columns=("ItemLista", "Valor"), show="headings", height=5)
        self.tree_nfse.heading("ItemLista", text="Item Lista Serviço")
        self.tree_nfse.heading("Valor", text="Valor Total (R$)")
        self.tree_nfse.column("ItemLista", width=150, anchor='center')
        self.tree_nfse.column("Valor", width=150, anchor='e')
        self.tree_nfse.grid(row=1, column=0, padx=10, pady=(0, 10), sticky="ew")
        
        nfse_scrollbar = ctk.CTkScrollbar(self.nfse_frame, command=self.tree_nfse.yview)
        nfse_scrollbar.grid(row=1, column=1, sticky='ns')
        self.tree_nfse.configure(yscrollcommand=nfse_scrollbar.set)
        
        # --- Frame Parceiro (Filho do Scrollable Frame) ---
        self.partner_frame = ctk.CTkFrame(self.results_scrollable_frame)
        self.partner_frame.grid(row=4, column=0, padx=0, pady=(0, 10), sticky="ew") # row 4 dentro do Scrollable Frame
        self.partner_frame.grid_columnconfigure(0, weight=1)

        self.label_partner_title = ctk.CTkLabel(self.partner_frame, text="Totalização de Valores por Parceiro (Autorizados)", font=ctk.CTkFont(size=14, weight="bold"))
        self.label_partner_title.grid(row=0, column=0, padx=10, pady=(10, 5), sticky="w")

        self.tree_partner = ttk.Treeview(self.partner_frame, columns=("Parceiro", "CNPJ/CPF", "Valor"), show="headings", height=5)
        self.tree_partner.heading("Parceiro", text="Nome do Parceiro")
        self.tree_partner.heading("CNPJ/CPF", text="CNPJ/CPF")
        self.tree_partner.heading("Valor", text="Valor Total (R$)")
        self.tree_partner.column("Parceiro", width=300)
        self.tree_partner.column("CNPJ/CPF", width=150, anchor='center')
        self.tree_partner.column("Valor", width=150, anchor='e')
        self.tree_partner.grid(row=1, column=0, padx=10, pady=(0, 10), sticky="ew")
        
        partner_scrollbar = ctk.CTkScrollbar(self.partner_frame, command=self.tree_partner.yview)
        partner_scrollbar.grid(row=1, column=1, sticky='ns')
        self.tree_partner.configure(yscrollcommand=partner_scrollbar.set)


    def selecionar_e_processar_arquivos(self):
        caminhos = filedialog.askopenfilenames(
            title="Selecione os arquivos XML, ZIP ou RAR",
            filetypes=[("Todos os arquivos suportados", ".xml .zip .rar"), ("XML", "*.xml"), ("ZIP", "*.zip"), ("RAR", "*.rar")]
        )
        if not caminhos:
            caminho_pasta = filedialog.askdirectory(title="Ou selecione uma pasta para processar")
            if caminho_pasta:
                caminhos = [os.path.join(caminho_pasta, f) 
                            for f in os.listdir(caminho_pasta) 
                            if os.path.splitext(f)[1].lower() in ['.xml', '.zip', '.rar']]
                
        if not caminhos: return
        
        # Limpa os resultados antes de iniciar novo processamento
        self.limpar_resultados(clear_files=False) 
        
        initial_count = self.arquivos_contados
        
        for caminho in caminhos:
            self.arquivos_contados += 1
            ext = os.path.splitext(caminho)[1].lower()
            try:
                if ext == '.xml':
                    with open(caminho, 'rb') as f:
                        self.processar_conteudo_xml(f.read(), os.path.basename(caminho))
                elif ext == '.zip':
                    with zipfile.ZipFile(caminho, 'r') as zf:
                        for nome in zf.namelist():
                            if nome.lower().endswith('.xml'):
                                self.processar_conteudo_xml(zf.read(nome), f"{os.path.basename(caminho)}/{nome}")
                elif ext == '.rar':
                    with rarfile.RarFile(caminho, 'r') as rf:
                        for nome in rf.namelist():
                            if nome.lower().endswith('.xml'):
                                self.processar_conteudo_xml(rf.read(nome), f"{os.path.basename(caminho)}/{nome}")
            except rarfile.MissingExternalError:
                 print(f"Erro: O arquivo RAR requer o comando 'unrar' instalado no sistema. Arquivo: {caminho}")
            except Exception as e:
                print(f"Erro ao processar o arquivo {caminho}: {e}")
                
        if self.arquivos_contados > initial_count:
            self.atualizar_interface()

    def processar_conteudo_xml(self, conteudo_xml, nome_arquivo):
        try:
            # Captura e normaliza o CNPJ da empresa (usuário)
            cnpj_empresa = self.cnpj_empresa_var.get().replace('.', '').replace('/', '').replace('-', '').strip()
            
            # 1. Definições de Namespace
            ns_nfe = {'nfe': 'http://www.portalfiscal.inf.br/nfe'}
            ns_cte = {'cte': 'http://www.portalfiscal.inf.br/cte'}
            ns_nfse = {'nfse': 'http://www.ginfes.com.br/tipos'} 
            
            try:
                # Tenta analisar o XML
                root = ET.fromstring(conteudo_xml)
            except ET.ParseError as e:
                print(f"Arquivo {nome_arquivo} não é um XML válido: {e}")
                return

            chave, valor, tipo, status, nNF = None, 0.0, 'N/A', 'autorizada', 'N/A'
            cfop_breakdown = []
            nfse_item_breakdown = None 
            doc_ns = None 
            prefix = None 
            
            # Variáveis para dados detalhados
            cnpj_parceiro, nome_parceiro = 'N/A', 'N/A'
            cnpj_emitente = 'N/A' 
            inconsistencia_fiscal = 'OK' 
            tpNF = None 
            fluxo = 'N/A' 
            
            # Lista para armazenar detalhes dos itens (para a nova aba)
            itens_lista = []

            # 2. IDENTIFICAR O TIPO DE DOCUMENTO E SUA CHAVE
            infNFe_tag = root.find('.//nfe:infNFe', ns_nfe)
            infCte_tag = root.find('.//cte:infCte', ns_cte)
            identificacao_nfse_tag = root.find('.//nfse:IdentificacaoNfse', ns_nfse) 

            # --- PROCESSAMENTO NFE ---
            if infNFe_tag is not None:
                chave = infNFe_tag.get('Id').replace('NFe', '')
                tipo = 'NF-e'
                doc_ns = ns_nfe
                prefix = 'nfe'
                
                # Extração de Nº Doc
                nNF_tag = root.find('.//nfe:ide/nfe:nNF', ns_nfe)
                nNF = nNF_tag.text if nNF_tag is not None else 'N/A'

                # Extração de CNPJ do Emitente
                emit_tag = root.find('.//nfe:emit', ns_nfe)
                if emit_tag is not None: 
                    cnpj_emitente_tag = emit_tag.find('nfe:CNPJ', ns_nfe)
                    cnpj_emitente = cnpj_emitente_tag.text if cnpj_emitente_tag is not None else 'N/A'
                
                # Extração de Valores e Parceiro
                tpNF_tag = root.find('.//nfe:ide/nfe:tpNF', ns_nfe)
                tpNF = tpNF_tag.text if tpNF_tag is not None else None
                valor_tag = root.find('.//nfe:ICMSTot/nfe:vNF', ns_nfe)
                valor = float(valor_tag.text) if valor_tag is not None and valor_tag.text else 0.0
                
                dest_tag = root.find('.//nfe:dest', ns_nfe)
                if dest_tag is not None:
                    cnpj_parceiro = dest_tag.find('nfe:CNPJ', ns_nfe).text if dest_tag.find('nfe:CNPJ', ns_nfe) is not None else (dest_tag.find('nfe:CPF', ns_nfe).text if dest_tag.find('nfe:CPF', ns_nfe) is not None else 'N/A')
                    nome_parceiro = dest_tag.find('nfe:xNome', ns_nfe).text if dest_tag.find('nfe:xNome', ns_nfe) is not None else 'N/A'
                
                # Extração dos Detalhes dos Itens (PRODUTOS)
                for det_tag in root.findall('.//nfe:det', ns_nfe):
                    prod_tag = det_tag.find('nfe:prod', ns_nfe)
                    imposto_tag = det_tag.find('nfe:imposto', ns_nfe)
                    
                    if prod_tag is not None:
                        item_detail = {
                            'nNF': nNF,
                            'nItem': det_tag.get('nItem'),
                            'xProd': prod_tag.find('nfe:xProd', ns_nfe).text if prod_tag.find('nfe:xProd', ns_nfe) is not None else 'N/A',
                            'NCM': prod_tag.find('nfe:NCM', ns_nfe).text if prod_tag.find('nfe:NCM', ns_nfe) is not None else 'N/A',
                            'CFOP': prod_tag.find('nfe:CFOP', ns_nfe).text if prod_tag.find('nfe:CFOP', ns_nfe) is not None else 'N/A',
                            'vProd': float(prod_tag.find('nfe:vProd', ns_nfe).text) if prod_tag.find('nfe:vProd', ns_nfe) is not None and prod_tag.find('nfe:vProd', ns_nfe).text else 0.0,
                            'CST_ICMS': 'N/A',
                            'vICMS': 0.0
                        }
                        
                        # Extração de ICMS
                        if imposto_tag is not None:
                            icms_tags = imposto_tag.find('nfe:ICMS', ns_nfe)
                            if icms_tags is not None:
                                for icms_child in icms_tags:
                                    cst_tag = icms_child.find(f'nfe:CST', ns_nfe) # Busca CST para ICMS00, ICMS10, etc.
                                    if cst_tag is None:
                                        cst_tag = icms_child.find(f'nfe:CSOSN', ns_nfe) # Busca CSOSN
                                        
                                    if cst_tag is not None:
                                        item_detail['CST_ICMS'] = cst_tag.text
                                        
                                    v_icms_tag = icms_child.find(f'nfe:vICMS', ns_nfe)
                                    if v_icms_tag is not None:
                                        try:
                                            item_detail['vICMS'] = float(v_icms_tag.text)
                                        except:
                                            pass
                                    break # Pega apenas o primeiro ICMS (que deve ser o relevante)

                        itens_lista.append(item_detail)

            # --- PROCESSAMENTO CTE ---
            elif infCte_tag is not None:
                chave = infCte_tag.get('Id').replace('CTe', '')
                tipo = 'CT-e'
                doc_ns = ns_cte
                prefix = 'cte'
                
                # Extração de Nº Doc
                nNF_tag = root.find('.//cte:ide/cte:nCT', ns_cte)
                nNF = nNF_tag.text if nNF_tag is not None else 'N/A'
                
                # Extração de CNPJ do Emitente
                emit_tag = root.find('.//cte:emit', ns_cte)
                if emit_tag is not None: 
                    cnpj_emitente_tag = emit_tag.find('cte:CNPJ', ns_cte)
                    cnpj_emitente = cnpj_emitente_tag.text if cnpj_emitente_tag is not None else 'N/A'

                # Extração de Valores e Parceiro
                valor_tag = root.find('.//cte:vPrest/cte:vTPrest', ns_cte)
                valor = float(valor_tag.text) if valor_tag is not None and valor_tag.text else 0.0
                dest_tag = root.find('.//cte:dest', ns_cte)
                if dest_tag is not None:
                    cnpj_parceiro = dest_tag.find('cte:CNPJ', ns_cte).text if dest_tag.find('cte:CNPJ', ns_cte) is not None else (dest_tag.find('cte:CPF', ns_cte).text if dest_tag.find('cte:CPF', ns_cte) is not None else 'N/A')
                    nome_parceiro = dest_tag.find('cte:xNome', ns_cte).text if dest_tag.find('cte:xNome', ns_cte) is not None else 'N/A'
                
                # Adiciona item de serviço genérico para o CT-e
                itens_lista.append({
                    'nNF': nNF,
                    'nItem': 1,
                    'xProd': 'Prestação de Serviço de Transporte',
                    'NCM': 'N/A',
                    'CFOP': 'N/A',
                    'vProd': valor,
                    'CST_ICMS': 'N/A', # CT-e usa tributação diferente
                    'vICMS': 0.0
                })
                

            # --- PROCESSAMENTO NFSE ---
            elif identificacao_nfse_tag is not None:
                chave_tag = identificacao_nfse_tag.find('nfse:Numero', ns_nfse)
                chave = chave_tag.text if chave_tag is not None and chave_tag.text else None
                tipo = 'NFSE'
                doc_ns = ns_nfse
                prefix = 'nfse'
                
                # Extração de Nº Doc
                nNF = chave # Para NFSE, o número é a chave
                
                # Extração de CNPJ do Prestador (Emitente)
                prestador_tag = root.find('.//nfse:PrestadorServico', ns_nfse)
                if prestador_tag is not None:
                    cnpj_emitente_tag = prestador_tag.find('nfse:IdentificacaoPrestador/nfse:CpfCnpj/nfse:Cnpj', ns_nfse)
                    cnpj_emitente = cnpj_emitente_tag.text if cnpj_emitente_tag is not None else 'N/A'
                
                # Extração de Valores e Parceiro
                valor_tag = root.find('.//nfse:ValorLiquidoNfse', ns_nfse)
                valor = float(valor_tag.text) if valor_tag is not None and valor_tag.text else 0.0
                item_lista_tag = root.find('.//nfse:Servico/nfse:ItemListaServico', ns_nfse)
                nfse_item_breakdown = item_lista_tag.text if item_lista_tag is not None and item_lista_tag.text else None
                
                tomador_tag = root.find('.//nfse:TomadorServico', ns_nfse)
                if tomador_tag is not None:
                    identificacao_tag = tomador_tag.find('nfse:IdentificacaoTomador', ns_nfse)
                    if identificacao_tag is not None:
                        cpfcnpj_tag = identificacao_tag.find('nfse:CpfCnpj', ns_nfse)
                        if cpfcnpj_tag is not None:
                            cnpj_parceiro = cpfcnpj_tag.find('nfse:Cnpj', ns_nfse).text if cpfcnpj_tag.find('nfse:Cnpj', ns_nfse) is not None and cpfcnpj_tag.find('nfse:Cnpj', ns_nfse).text else (cpfcnpj_tag.find('nfse:Cpf', ns_nfse).text if cpfcnpj_tag.find('nfse:Cpf', ns_nfse) is not None and cpfcnpj_tag.find('nfse:Cpf', ns_nfse).text else 'N/A')
                        razao_social_tag = tomador_tag.find('nfse:RazaoSocial', ns_nfse)
                        if razao_social_tag is not None:
                            nome_parceiro = razao_social_tag.text
                
                # Adiciona item de serviço específico para NFSE
                itens_lista.append({
                    'nNF': nNF,
                    'nItem': 1,
                    'xProd': f'Serviço: Item {nfse_item_breakdown}' if nfse_item_breakdown else 'Serviço Prestado',
                    'NCM': 'N/A',
                    'CFOP': 'N/A',
                    'vProd': valor,
                    'CST_ICMS': 'N/A',
                    'vICMS': 0.0
                })
            
            # --- Tenta encontrar a chave em XML de Evento (se não for documento principal) ---
            chNFe_tag = root.find('.//nfe:chNFe', ns_nfe)
            chCTe_tag = root.find('.//cte:chCTe', ns_cte)
            
            if not chave and chNFe_tag is not None:
                chave = chNFe_tag.text
                tipo = 'NF-e'
                doc_ns = ns_nfe
                prefix = 'nfe'
            elif not chave and chCTe_tag is not None:
                chave = chCTe_tag.text
                tipo = 'CT-e'
                doc_ns = ns_cte
                prefix = 'cte'

            # Se a chave foi encontrada (seja NF-e, CT-e, NFSE ou Evento), prossegue
            if chave:
                
                # 3. VERIFICAR CANCELAMENTO (Lógica Aprimorada)
                is_cancellation_event = False
                
                # 3a. Busca Evento de Cancelamento no XML atual (principalmente para XML 'avulso' ou 'procEventoNFe')
                if prefix in ['nfe', 'cte']:
                    # Procura por retEvento 110111 (Cancelamento) ou procEvento com 101/135 (Cancelamento)
                    cancel_ret_evento_tag = root.find(f'.//{prefix}:retEvento/{prefix}:infEvento', doc_ns)
                    if cancel_ret_evento_tag is not None:
                        cStat_evento_tag = cancel_ret_evento_tag.find(f'{prefix}:cStat', doc_ns)
                        tpEvento_tag = cancel_ret_evento_tag.find(f'{prefix}:tpEvento', doc_ns)
                        if cStat_evento_tag is not None and cStat_evento_tag.text == '135' and tpEvento_tag is not None and tpEvento_tag.text == '110111':
                            status = 'cancelada'
                            is_cancellation_event = True
                            
                    # Procura por evento (para o caso do evento estar na raiz procEventoNFe)
                    if not is_cancellation_event and (root.tag.endswith('procEventoNFe') or root.tag.endswith('procEventoCTe')):
                        tpEvento_tag = root.find(f'.//{prefix}:evento/{prefix}:infEvento/{prefix}:tpEvento', doc_ns)
                        if tpEvento_tag is not None and tpEvento_tag.text == '110111':
                             status = 'cancelada'
                             is_cancellation_event = True
                             
                    # Procura por infProt 101 (Cancelamento)
                    if root.tag.endswith('Proc') and prefix in ['nfe', 'cte']:
                        cStat_prot_tag = root.find(f'.//{prefix}:prot{tipo[-3:]}/{prefix}:infProt/{prefix}:cStat', doc_ns)
                        if cStat_prot_tag is not None and cStat_prot_tag.text == '101': 
                            status = 'cancelada'
                            is_cancellation_event = True # Define como evento de cancelamento para sobrescrita

                # 4. LÓGICA DE FLUXO (Entrada/Saída)
                if cnpj_empresa:
                    # Normaliza CNPJs para comparação
                    cnpj_emitente_norm = cnpj_emitente.replace('.', '').replace('/', '').replace('-', '').strip()
                    cnpj_parceiro_norm = cnpj_parceiro.replace('.', '').replace('/', '').replace('-', '').strip()

                    if tipo == 'NF-e':
                        if cnpj_emitente_norm == cnpj_empresa:
                            fluxo = 'Saída Própria' if tpNF == '1' else ('Entrada Própria (Dev/Canc)' if tpNF == '0' else 'Saída Própria (tpNF Desconhecido)')
                        elif cnpj_parceiro_norm == cnpj_empresa:
                            fluxo = 'Entrada de Terceiros'
                        else:
                            fluxo = 'Terceiros'
                    
                    elif tipo in ['CT-e', 'NFSE']:
                        if cnpj_parceiro_norm == cnpj_empresa:
                            fluxo = f'Entrada ({tipo})'
                        elif cnpj_emitente_norm == cnpj_empresa:
                            fluxo = f'Saída ({tipo})'
                        else:
                            fluxo = 'Terceiros'
                
                # Se estiver cancelado, ajusta o fluxo
                if status == 'cancelada':
                    fluxo = 'CANCELADO'
                
                # 5. REGISTRO DO DOCUMENTO
                novo_doc = {
                    'nNF': nNF, # NOVO CAMPO
                    'valor': valor, 
                    'status': status, 
                    'tipo': tipo, 
                    'arquivo': nome_arquivo, 
                    'cfop_breakdown': cfop_breakdown,
                    'nfse_item_breakdown': nfse_item_breakdown, 
                    'cnpj_parceiro': cnpj_parceiro,          
                    'nome_parceiro': nome_parceiro,          
                    'inconsistencia_fiscal': inconsistencia_fiscal, 
                    'fluxo': fluxo, 
                    'cnpj_emitente': cnpj_emitente,
                }

                # Lógica de Sobrescrita de Cancelamento
                if is_cancellation_event:
                    # Se for um evento de cancelamento (XML avulso ou aninhado)
                    if chave in self.documentos_processados:
                        # Se já existe a NF-e, atualiza o status
                        doc_existente = self.documentos_processados[chave]
                        doc_existente['status'] = 'cancelada'
                        doc_existente['fluxo'] = 'CANCELADO'
                        doc_existente['inconsistencia_fiscal'] = 'CANCELADA'
                        doc_existente['arquivo'] += f" / Evento: {nome_arquivo}"
                        # Limpa itens se a nota for cancelada
                        GLOBAL_ITEM_DETAILS.pop(chave, None)
                    else:
                        # Se o evento de cancelamento veio antes da NF-e, registra o placeholder
                        self.documentos_processados[chave] = {
                            'nNF': 'N/A',
                            'valor': 0.0, 
                            'status': 'cancelada', 
                            'tipo': tipo,
                            'arquivo': f"Evento: {nome_arquivo}",
                            'cnpj_parceiro': 'N/A', 
                            'nome_parceiro': f"Verificar Chave: {chave}",
                            'inconsistencia_fiscal': 'CANCELADA',
                            'fluxo': 'CANCELADO',
                            'cfop_breakdown': [], 'nfse_item_breakdown': None,
                            'cnpj_emitente': 'N/A',
                        }
                
                # Se for uma NF-e/CT-e/NFSE principal
                elif chave not in self.documentos_processados or self.documentos_processados[chave]['status'] != 'cancelada':
                    self.documentos_processados[chave] = novo_doc
                    GLOBAL_ITEM_DETAILS[chave] = itens_lista # Adiciona detalhes dos itens
                    
                # 6. EXTRAÇÃO DE CFOP E INCONSISTÊNCIA (Apenas para NF-e Autorizadas)
                if tipo == 'NF-e' and novo_doc['status'] == 'autorizada' and 'cfop_breakdown' in novo_doc:
                    for item in itens_lista:
                        if item.get('CFOP'):
                           novo_doc['cfop_breakdown'].append({'cfop': item['CFOP'], 'vProd': item['vProd']})
                        
                        # Lógica de Inconsistência Fiscal (CFOP vs. tpNF)
                        cfop = item.get('CFOP')
                        if cfop and tpNF in ['0', '1']:
                            # CFOP de SAÍDA (5, 6, 7) em NF de ENTRADA (tpNF=0)
                            if tpNF == '0' and cfop[0] in ['5', '6', '7']: 
                                novo_doc['inconsistencia_fiscal'] = 'CFOP SAÍDA em NF ENTRADA'
                                break
                            # CFOP de ENTRADA (1, 2, 3) em NF de SAÍDA (tpNF=1)
                            elif tpNF == '1' and cfop[0] in ['1', '2', '3']:
                                novo_doc['inconsistencia_fiscal'] = 'CFOP ENTRADA em NF SAÍDA'
                                break
                                
                    # Se inconsistência foi detectada, atualiza o status do documento para o fluxo principal
                    if novo_doc['inconsistencia_fiscal'] != 'OK':
                        self.documentos_processados[chave]['inconsistencia_fiscal'] = novo_doc['inconsistencia_fiscal']
                        
        except Exception as e:
            print(f"Erro inesperado ao processar {nome_arquivo}: {e}")

    def atualizar_interface(self):
        # 1. Limpar Treeviews
        for i in self.tree.get_children():
            self.tree.delete(i)
        for i in self.tree_cfop.get_children(): 
            self.tree_cfop.delete(i)
        for i in self.tree_nfse.get_children(): 
            self.tree_nfse.delete(i)
        for i in self.tree_partner.get_children(): 
            self.tree_partner.delete(i)
        for i in self.tree_itens.get_children(): # NOVO
            self.tree_itens.delete(i) 

        # 2. Re-calcular totais (Documento, CFOP, NFSE e Parceiro)
        total_valor = 0.0
        counts = defaultdict(lambda: defaultdict(int))
        self.cfop_totals.clear() 
        self.nfse_item_totals.clear() 
        self.partner_totals.clear() 

        for chave, doc in sorted(self.documentos_processados.items()):
            # Adicionar na Treeview principal
            valor_formatado = f"{doc['valor']:.2f}".replace('.', ',')
            
            # Formatação de cor para melhor visualização
            tag = "cancelada" if doc['status'] == 'cancelada' else ("inconsistencia" if doc.get('inconsistencia_fiscal') != 'OK' else "")
            
            self.tree.insert("", "end", values=(
                chave, 
                doc.get('tipo', 'N/A'), 
                doc.get('nNF', 'N/A'), # NOVO CAMPO
                valor_formatado, 
                doc['status'].title(), 
                doc.get('fluxo', 'N/A'), 
                doc.get('nome_parceiro', 'N/A'), 
                doc.get('cnpj_parceiro', 'N/A'), 
                doc.get('inconsistencia_fiscal', 'N/A'), 
                doc['arquivo']
            ), tags=(tag,))
            
            # Estilos de cor para tags
            self.tree.tag_configure('cancelada', foreground='red')
            self.tree.tag_configure('inconsistencia', foreground='#FFD700')
            
            # Contagem de status/tipo
            counts[doc['status']][doc.get('tipo', 'N/A')] += 1
            
            # Totalizar apenas se o status for AUTORIZADA
            if doc['status'] == 'autorizada':
                total_valor += doc['valor']
                
                # Agregação de totais (CFOP, NFSE, Parceiro)
                if doc.get('tipo') == 'NF-e' and 'cfop_breakdown' in doc:
                    for item in doc['cfop_breakdown']:
                        self.cfop_totals[item['cfop']] += item['vProd']
                        
                if doc.get('tipo') == 'NFSE' and doc.get('nfse_item_breakdown'):
                    item_lista = doc['nfse_item_breakdown']
                    self.nfse_item_totals[item_lista] += doc['valor']
                    
                partner_key = (doc.get('cnpj_parceiro', 'N/A'), doc.get('nome_parceiro', 'N/A'))
                if partner_key[0] != 'N/A':
                    self.partner_totals[partner_key] += doc['valor']
        
        # 3. Atualizar Labels de Totais de Documento
        total_valor_str = f"{total_valor:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
        self.label_valor_total.configure(text=f"Valor Total (Autorizados): R$ {total_valor_str}")
        
        auth_nfe = counts['autorizada']['NF-e']
        auth_cte = counts['autorizada']['CT-e']
        auth_nfse = counts['autorizada']['NFSE']
        self.label_total_autorizadas.configure(text=f"Docs Autorizados (NF-e/CT-e/NFSE): {auth_nfe} / {auth_cte} / {auth_nfse}")
        
        canc_nfe = counts['cancelada']['NF-e']
        canc_cte = counts['cancelada']['CT-e']
        canc_nfse = counts['cancelada']['NFSE']
        self.label_total_canceladas.configure(text=f"Docs Cancelados (NF-e/CT-e/NFSE): {canc_nfe} / {canc_cte} / {canc_nfse}")
        self.label_arquivos_processados.configure(text=f"Arquivos Físicos Processados: {self.arquivos_contados}")

        # 4. Atualizar Treeview de Itens (NOVO)
        for chave, itens in GLOBAL_ITEM_DETAILS.items():
            doc = self.documentos_processados.get(chave)
            if doc and doc['status'] == 'autorizada':
                for item in itens:
                    valor_item_formatado = f"{item['vProd']:.2f}".replace('.', ',')
                    valor_icms_formatado = f"{item['vICMS']:.2f}".replace('.', ',')
                    
                    self.tree_itens.insert("", "end", values=(
                        item['nNF'],
                        item['xProd'],
                        item['NCM'],
                        item['CFOP'],
                        valor_item_formatado,
                        item['CST_ICMS'],
                        valor_icms_formatado
                    ))


        # 5. Atualizar Treeview de CFOP
        for cfop, valor in sorted(self.cfop_totals.items()):
            valor_cfop_formatado = f"{valor:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
            self.tree_cfop.insert("", "end", values=(cfop, valor_cfop_formatado))
            
        # 6. Atualizar Treeview de NFSE
        for item_lista, valor in sorted(self.nfse_item_totals.items()):
            valor_nfse_formatado = f"{valor:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
            self.tree_nfse.insert("", "end", values=(item_lista, valor_nfse_formatado))
            
        # 7. Atualizar Treeview de Parceiro
        for partner_key, valor in sorted(self.partner_totals.items(), key=lambda x: x[1], reverse=True):
            cnpj_parceiro, nome_parceiro = partner_key
            valor_parceiro_formatado = f"{valor:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
            self.tree_partner.insert("", "end", values=(nome_parceiro, cnpj_parceiro, valor_parceiro_formatado))


    def limpar_resultados(self, clear_files=True):
        self.documentos_processados.clear()
        self.cfop_totals.clear() 
        self.nfse_item_totals.clear() 
        self.partner_totals.clear() 
        GLOBAL_ITEM_DETAILS.clear() # NOVO: Limpa os detalhes dos itens
        if clear_files:
             self.arquivos_contados = 0
        self.atualizar_interface()

if __name__ == "__main__":
    try:
        # Tenta configurar rarfile e zipfile (necessário em alguns ambientes)
        rarfile.PATH_SEP = os.path.sep
        zipfile.Path = os.path.join
    except AttributeError:
        pass 

    app = App()
    app.mainloop()
