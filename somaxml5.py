# somaxml5_pro.py
import tkinter
from tkinter import filedialog, ttk, messagebox
import customtkinter as ctk
import xml.etree.ElementTree as ET
import zipfile
import rarfile
import os
import csv
import logging
from collections import defaultdict

# --- Config logging ---
LOG_FILENAME = "erros.log"
logging.basicConfig(
    filename=LOG_FILENAME,
    level=logging.INFO,
    format="%(asctime)s %(levelname)s: %(message)s",
)

# --- Aparência CustomTkinter ---
ctk.set_appearance_mode("system")
ctk.set_default_color_theme("blue")

# GLOBAL para detalhes de itens por documento
GLOBAL_ITEM_DETAILS = defaultdict(list)

class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Analisador Fiscal de Documentos Eletrônicos (NF-e / CT-e / NFSe) v5.0.0")
        self.geometry("1400x980")

        # Estado
        self.documentos_processados = {}
        self.arquivos_contados = 0
        self.cfop_totals = defaultdict(float)
        # nfse_item_totals agrupa por (ItemListaServico, CNAE, CodigoTributacaoMunicipio)
        self.nfse_item_totals = defaultdict(float)
        self.partner_totals = defaultdict(float)
        self.cnpj_empresa_var = ctk.StringVar(value="")
        GLOBAL_ITEM_DETAILS.clear()

        # Novos contadores / diagnósticos
        self.erros_detectados = 0
        self.quebras_sequencia_alerts = []  # lista de strings de alerta

        # Layout
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=0)
        self.grid_rowconfigure(1, weight=1)

        # Top frame (controles)
        self.top_frame = ctk.CTkFrame(self, height=210)
        self.top_frame.grid(row=0, column=0, padx=10, pady=(10,0), sticky="ew")
        self.top_frame.grid_columnconfigure((0,1,2,3,4,5), weight=1)

        # CNPJ empresa
        self.label_cnpj_empresa = ctk.CTkLabel(self.top_frame, text="CNPJ da Empresa (Apenas números):", font=ctk.CTkFont(size=13, weight="bold"))
        self.label_cnpj_empresa.grid(row=0, column=2, padx=(10,5), pady=8, sticky="e")
        self.cnpj_entry = ctk.CTkEntry(self.top_frame, textvariable=self.cnpj_empresa_var, placeholder_text="Ex: 12345678000190", width=180)
        self.cnpj_entry.grid(row=0, column=3, padx=(5,10), pady=8, sticky="w")

        # Botões principais
        self.select_button = ctk.CTkButton(self.top_frame, text="Selecionar Arquivos / Pasta", command=self.selecionar_e_processar_arquivos)
        self.select_button.grid(row=1, column=0, padx=6, pady=(4,6), sticky="ew")
        self.clear_button = ctk.CTkButton(self.top_frame, text="Limpar Resultados", command=self.limpar_resultados)
        self.clear_button.grid(row=1, column=1, padx=6, pady=(4,6), sticky="ew")
        self.export_button = ctk.CTkButton(self.top_frame, text="Exportar CSV", command=self.exportar_csv)
        self.export_button.grid(row=1, column=2, padx=6, pady=(4,6), sticky="ew")

        # Diagnóstico rápido
        self.label_valor_total = ctk.CTkLabel(self.top_frame, text="Valor Total (Autorizados): R$ 0,00", font=ctk.CTkFont(size=14, weight="bold"))
        self.label_valor_total.grid(row=2, column=0, padx=10, pady=6, columnspan=2, sticky="w")

        self.label_total_autorizadas = ctk.CTkLabel(self.top_frame, text="Docs Autorizados (NF-e/CT-e/NFSE): 0 / 0 / 0", font=ctk.CTkFont(size=13))
        self.label_total_autorizadas.grid(row=3, column=0, padx=10, pady=2, sticky="w")
        self.label_total_canceladas = ctk.CTkLabel(self.top_frame, text="Docs Cancelados (NF-e/CT-e/NFSE): 0 / 0 / 0", font=ctk.CTkFont(size=13))
        self.label_total_canceladas.grid(row=3, column=1, padx=10, pady=2, sticky="w")

        self.label_arquivos_processados = ctk.CTkLabel(self.top_frame, text="Arquivos Processados: 0", font=ctk.CTkFont(size=13))
        self.label_arquivos_processados.grid(row=3, column=2, padx=10, pady=2, sticky="w")

        # Novo: contador de erros e quebras
        self.label_erros = ctk.CTkLabel(self.top_frame, text=f"Erros: {self.erros_detectados}", text_color="red", font=ctk.CTkFont(size=13, weight="bold"))
        self.label_erros.grid(row=2, column=4, padx=10, pady=6, sticky="e")
        self.label_quebras_count = ctk.CTkLabel(self.top_frame, text="Quebras sequência: 0", font=ctk.CTkFont(size=13))
        self.label_quebras_count.grid(row=2, column=5, padx=10, pady=6, sticky="e")
        self.btn_ver_quebras = ctk.CTkButton(self.top_frame, text="Ver Quebras", command=self.mostrar_quebras_popup)
        self.btn_ver_quebras.grid(row=1, column=5, padx=6, pady=(4,6), sticky="ew")

        # Scrollable results
        self.results_scrollable_frame = ctk.CTkScrollableFrame(self)
        self.results_scrollable_frame.grid(row=1, column=0, padx=10, pady=10, sticky="nsew")
        self.results_scrollable_frame.grid_columnconfigure(0, weight=1)

        style = ttk.Style()
        style.theme_use("default")
        style.configure("Treeview", background="#2b2b2b", foreground="white", rowheight=25, fieldbackground="#2b2b2b", bordercolor="#333333", borderwidth=0)
        style.map('Treeview', background=[('selected', '#1f6aa5')])
        style.configure("Treeview.Heading", background="#565b5e", foreground="white", relief="flat")

        # Main table
        self.main_table_frame = ctk.CTkFrame(self.results_scrollable_frame)
        self.main_table_frame.grid(row=0, column=0, padx=0, pady=(0,10), sticky="ew")
        self.main_table_frame.grid_rowconfigure(0, weight=1)
        self.main_table_frame.grid_columnconfigure(0, weight=1)

        self.tree = ttk.Treeview(self.main_table_frame, columns=("Chave", "Tipo", "Nº Doc", "Valor", "Status", "Fluxo", "Parceiro", "CNPJ/CPF", "Inconsistência", "Arquivo"), show="headings", height=10)
        self.tree.heading("Chave", text="Chave do Documento")
        self.tree.heading("Tipo", text="Tipo")
        self.tree.heading("Nº Doc", text="Nº Doc")
        self.tree.heading("Valor", text="Valor (R$)")
        self.tree.heading("Status", text="Status")
        self.tree.heading("Fluxo", text="Fluxo")
        self.tree.heading("Parceiro", text="Parceiro (Dest/Tomador)")
        self.tree.heading("CNPJ/CPF", text="CNPJ/CPF")
        self.tree.heading("Inconsistência", text="Inconsistência Fiscal")
        self.tree.heading("Arquivo", text="Origem")
        self.tree.column("Chave", width=180)
        self.tree.column("Tipo", width=50, anchor='center')
        self.tree.column("Nº Doc", width=70, anchor='center')
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

        # Itens
        self.itens_frame = ctk.CTkFrame(self.results_scrollable_frame)
        self.itens_frame.grid(row=1, column=0, padx=0, pady=(10,10), sticky="ew")
        self.itens_frame.grid_columnconfigure(0, weight=1)

        self.label_itens_title = ctk.CTkLabel(self.itens_frame, text="Detalhes de Itens (Produtos/Serviços) de Documentos Autorizados", font=ctk.CTkFont(size=14, weight="bold"))
        self.label_itens_title.grid(row=0, column=0, padx=10, pady=(10,5), sticky="w")

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
        self.tree_itens.grid(row=1, column=0, padx=10, pady=(0,10), sticky="ew")

        itens_scrollbar = ctk.CTkScrollbar(self.itens_frame, command=self.tree_itens.yview)
        itens_scrollbar.grid(row=1, column=1, sticky='ns')
        self.tree_itens.configure(yscrollcommand=itens_scrollbar.set)

        # CFOP
        self.cfop_frame = ctk.CTkFrame(self.results_scrollable_frame)
        self.cfop_frame.grid(row=2, column=0, padx=0, pady=(0,10), sticky="ew")
        self.cfop_frame.grid_columnconfigure(0, weight=1)

        self.label_cfop_title = ctk.CTkLabel(self.cfop_frame, text="Totalização de Valores por CFOP (Apenas NF-e Autorizadas)", font=ctk.CTkFont(size=14, weight="bold"))
        self.label_cfop_title.grid(row=0, column=0, padx=10, pady=(10,5), sticky="w")

        self.tree_cfop = ttk.Treeview(self.cfop_frame, columns=("CFOP", "Valor"), show="headings", height=5)
        self.tree_cfop.heading("CFOP", text="CFOP")
        self.tree_cfop.heading("Valor", text="Valor Total (R$)")
        self.tree_cfop.column("CFOP", width=100, anchor='center')
        self.tree_cfop.column("Valor", width=150, anchor='e')
        self.tree_cfop.grid(row=1, column=0, padx=10, pady=(0,10), sticky="ew")

        cfop_scrollbar = ctk.CTkScrollbar(self.cfop_frame, command=self.tree_cfop.yview)
        cfop_scrollbar.grid(row=1, column=1, sticky='ns')
        self.tree_cfop.configure(yscrollcommand=cfop_scrollbar.set)

        # NFSe totals (agregado por ItemListaServico / CNAE / CodigoTributacaoMunicipio)
        self.nfse_frame = ctk.CTkFrame(self.results_scrollable_frame)
        self.nfse_frame.grid(row=3, column=0, padx=0, pady=(0,10), sticky="ew")
        self.nfse_frame.grid_columnconfigure(0, weight=1)

        self.label_nfse_title = ctk.CTkLabel(self.nfse_frame, text="Totalização de NFSe por (ItemListaServico / CNAE / Código Trib. Município)", font=ctk.CTkFont(size=14, weight="bold"))
        self.label_nfse_title.grid(row=0, column=0, padx=10, pady=(10,5), sticky="w")

        self.tree_nfse = ttk.Treeview(self.nfse_frame, columns=("ItemLista", "CNAE", "CodTribMun", "Valor"), show="headings", height=6)
        self.tree_nfse.heading("ItemLista", text="Item Lista Serviço")
        self.tree_nfse.heading("CNAE", text="CNAE")
        self.tree_nfse.heading("CodTribMun", text="Código Trib. Município")
        self.tree_nfse.heading("Valor", text="Valor Total (R$)")
        self.tree_nfse.column("ItemLista", width=150, anchor='center')
        self.tree_nfse.column("CNAE", width=120, anchor='center')
        self.tree_nfse.column("CodTribMun", width=150, anchor='center')
        self.tree_nfse.column("Valor", width=150, anchor='e')
        self.tree_nfse.grid(row=1, column=0, padx=10, pady=(0,10), sticky="ew")

        nfse_scrollbar = ctk.CTkScrollbar(self.nfse_frame, command=self.tree_nfse.yview)
        nfse_scrollbar.grid(row=1, column=1, sticky='ns')
        self.tree_nfse.configure(yscrollcommand=nfse_scrollbar.set)

        # Parceiro
        self.partner_frame = ctk.CTkFrame(self.results_scrollable_frame)
        self.partner_frame.grid(row=4, column=0, padx=0, pady=(0,10), sticky="ew")
        self.partner_frame.grid_columnconfigure(0, weight=1)

        self.label_partner_title = ctk.CTkLabel(self.partner_frame, text="Totalização de Valores por Parceiro (Autorizados)", font=ctk.CTkFont(size=14, weight="bold"))
        self.label_partner_title.grid(row=0, column=0, padx=10, pady=(10,5), sticky="w")

        self.tree_partner = ttk.Treeview(self.partner_frame, columns=("Parceiro", "CNPJ/CPF", "Valor"), show="headings", height=5)
        self.tree_partner.heading("Parceiro", text="Nome do Parceiro")
        self.tree_partner.heading("CNPJ/CPF", text="CNPJ/CPF")
        self.tree_partner.heading("Valor", text="Valor Total (R$)")
        self.tree_partner.column("Parceiro", width=300)
        self.tree_partner.column("CNPJ/CPF", width=150, anchor='center')
        self.tree_partner.column("Valor", width=150, anchor='e')
        self.tree_partner.grid(row=1, column=0, padx=10, pady=(0,10), sticky="ew")

        partner_scrollbar = ctk.CTkScrollbar(self.partner_frame, command=self.tree_partner.yview)
        partner_scrollbar.grid(row=1, column=1, sticky='ns')
        self.tree_partner.configure(yscrollcommand=partner_scrollbar.set)

    # -----------------------
    # Helpers de extração
    # -----------------------
    def _find_text_any(self, root, candidates):
        """
        Tenta várias paths (relativas) e retorna o texto do primeiro match.
        Usa wildcard '{*}' para ignorar namespace.
        candidates: lista de caminhos como './/{*}Tag/...'
        """
        for p in candidates:
            node = root.find(p)
            if node is not None and node.text and node.text.strip() != '':
                return node.text.strip()
        return None

    def _find_node_any(self, root, candidates):
        for p in candidates:
            node = root.find(p)
            if node is not None:
                return node
        return None

    # -----------------------
    # Seleção e processamento
    # -----------------------
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
        if not caminhos:
            return

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
                logging.error(f"Arquivo RAR requer 'unrar': {caminho}")
                self._registrar_erro(f"RAR missing unrar: {caminho}")
            except Exception as e:
                logging.exception(f"Erro ao processar arquivo container {caminho}: {e}")
                self._registrar_erro(f"Erro ao processar container {caminho}: {e}")

        if self.arquivos_contados > initial_count:
            self.atualizar_interface()

    # -----------------------
    # Parser principal
    # -----------------------
    def processar_conteudo_xml(self, conteudo_xml, nome_arquivo):
        try:
            cnpj_empresa = self.cnpj_empresa_var.get().replace('.', '').replace('/', '').replace('-', '').strip()

            ns_nfe = {'nfe': 'http://www.portalfiscal.inf.br/nfe'}
            ns_cte = {'cte': 'http://www.portalfiscal.inf.br/cte'}

            try:
                root = ET.fromstring(conteudo_xml)
            except ET.ParseError as e:
                logging.warning(f"Arquivo {nome_arquivo} não é XML válido: {e}")
                self._registrar_erro(f"XML inválido {nome_arquivo}: {e}")
                return

            chave, valor, tipo, status, nNF = None, 0.0, 'N/A', 'autorizada', 'N/A'
            cfop_breakdown = []
            nfse_item_breakdown = None
            doc_ns = None
            prefix = None

            cnpj_parceiro, nome_parceiro = 'N/A', 'N/A'
            cnpj_emitente = 'N/A'
            inconsistencia_fiscal = 'OK'
            tpNF = None
            fluxo = 'N/A'
            itens_lista = []

            # Detectar NF-e / CT-e via namespaces conhecidos
            infNFe_tag = root.find('.//{http://www.portalfiscal.inf.br/nfe}infNFe')
            infCte_tag = root.find('.//{http://www.portalfiscal.inf.br/cte}infCte')

            # --- NF-e ---
            if infNFe_tag is not None:
                chave = infNFe_tag.get('Id').replace('NFe', '') if infNFe_tag.get('Id') else None
                tipo = 'NF-e'
                doc_ns = ns_nfe
                prefix = 'nfe'

                nNF_tag = root.find('.//{http://www.portalfiscal.inf.br/nfe}ide/{http://www.portalfiscal.inf.br/nfe}nNF')
                nNF = nNF_tag.text if nNF_tag is not None else 'N/A'

                emit_tag = root.find('.//{http://www.portalfiscal.inf.br/nfe}emit')
                if emit_tag is not None:
                    cnpj_emitente_tag = emit_tag.find('{http://www.portalfiscal.inf.br/nfe}CNPJ')
                    cnpj_emitente = cnpj_emitente_tag.text if cnpj_emitente_tag is not None else 'N/A'

                tpNF_tag = root.find('.//{http://www.portalfiscal.inf.br/nfe}ide/{http://www.portalfiscal.inf.br/nfe}tpNF')
                tpNF = tpNF_tag.text if tpNF_tag is not None else None

                valor_tag = root.find('.//{http://www.portalfiscal.inf.br/nfe}ICMSTot/{http://www.portalfiscal.inf.br/nfe}vNF')
                try:
                    valor = float(valor_tag.text) if valor_tag is not None and valor_tag.text else 0.0
                except:
                    valor = 0.0

                dest_tag = root.find('.//{http://www.portalfiscal.inf.br/nfe}dest')
                if dest_tag is not None:
                    cnpj_parceiro_tag = dest_tag.find('{http://www.portalfiscal.inf.br/nfe}CNPJ') or dest_tag.find('{http://www.portalfiscal.inf.br/nfe}CPF')
                    cnpj_parceiro = cnpj_parceiro_tag.text if cnpj_parceiro_tag is not None else 'N/A'
                    nome_parceiro_tag = dest_tag.find('{http://www.portalfiscal.inf.br/nfe}xNome')
                    nome_parceiro = nome_parceiro_tag.text if nome_parceiro_tag is not None else 'N/A'

                for det_tag in root.findall('.//{http://www.portalfiscal.inf.br/nfe}det'):
                    prod_tag = det_tag.find('{http://www.portalfiscal.inf.br/nfe}prod')
                    imposto_tag = det_tag.find('{http://www.portalfiscal.inf.br/nfe}imposto')
                    if prod_tag is not None:
                        item_detail = {
                            'nNF': nNF,
                            'nItem': det_tag.get('nItem'),
                            'xProd': prod_tag.find('{http://www.portalfiscal.inf.br/nfe}xProd').text if prod_tag.find('{http://www.portalfiscal.inf.br/nfe}xProd') is not None else 'N/A',
                            'NCM': prod_tag.find('{http://www.portalfiscal.inf.br/nfe}NCM').text if prod_tag.find('{http://www.portalfiscal.inf.br/nfe}NCM') is not None else 'N/A',
                            'CFOP': prod_tag.find('{http://www.portalfiscal.inf.br/nfe}CFOP').text if prod_tag.find('{http://www.portalfiscal.inf.br/nfe}CFOP') is not None else 'N/A',
                            'vProd': float(prod_tag.find('{http://www.portalfiscal.inf.br/nfe}vProd').text) if prod_tag.find('{http://www.portalfiscal.inf.br/nfe}vProd') is not None and prod_tag.find('{http://www.portalfiscal.inf.br/nfe}vProd').text else 0.0,
                            'CST_ICMS': 'N/A',
                            'vICMS': 0.0
                        }
                        if imposto_tag is not None:
                            icms_tags = imposto_tag.find('{http://www.portalfiscal.inf.br/nfe}ICMS')
                            if icms_tags is not None:
                                for icms_child in icms_tags:
                                    cst_tag = icms_child.find('{http://www.portalfiscal.inf.br/nfe}CST') or icms_child.find('{http://www.portalfiscal.inf.br/nfe}CSOSN')
                                    if cst_tag is not None:
                                        item_detail['CST_ICMS'] = cst_tag.text
                                    v_icms_tag = icms_child.find('{http://www.portalfiscal.inf.br/nfe}vICMS')
                                    if v_icms_tag is not None:
                                        try:
                                            item_detail['vICMS'] = float(v_icms_tag.text)
                                        except:
                                            pass
                                    break
                        itens_lista.append(item_detail)

            # --- CT-e ---
            elif infCte_tag is not None:
                chave = infCte_tag.get('Id').replace('CTe', '') if infCte_tag.get('Id') else None
                tipo = 'CT-e'
                doc_ns = ns_cte
                prefix = 'cte'

                nNF_tag = root.find('.//{http://www.portalfiscal.inf.br/cte}ide/{http://www.portalfiscal.inf.br/cte}nCT')
                nNF = nNF_tag.text if nNF_tag is not None else 'N/A'

                emit_tag = root.find('.//{http://www.portalfiscal.inf.br/cte}emit')
                if emit_tag is not None:
                    cnpj_emitente_tag = emit_tag.find('{http://www.portalfiscal.inf.br/cte}CNPJ')
                    cnpj_emitente = cnpj_emitente_tag.text if cnpj_emitente_tag is not None else 'N/A'

                valor_tag = root.find('.//{http://www.portalfiscal.inf.br/cte}vPrest/{http://www.portalfiscal.inf.br/cte}vTPrest')
                try:
                    valor = float(valor_tag.text) if valor_tag is not None and valor_tag.text else 0.0
                except:
                    valor = 0.0

                dest_tag = root.find('.//{http://www.portalfiscal.inf.br/cte}dest')
                if dest_tag is not None:
                    cnpj_parceiro_tag = dest_tag.find('{http://www.portalfiscal.inf.br/cte}CNPJ') or dest_tag.find('{http://www.portalfiscal.inf.br/cte}CPF')
                    cnpj_parceiro = cnpj_parceiro_tag.text if cnpj_parceiro_tag is not None else 'N/A'
                    nome_parceiro_tag = dest_tag.find('{http://www.portalfiscal.inf.br/cte}xNome')
                    nome_parceiro = nome_parceiro_tag.text if nome_parceiro_tag is not None else 'N/A'

                itens_lista.append({
                    'nNF': nNF,
                    'nItem': 1,
                    'xProd': 'Prestação de Serviço de Transporte',
                    'NCM': 'N/A',
                    'CFOP': 'N/A',
                    'vProd': valor,
                    'CST_ICMS': 'N/A',
                    'vICMS': 0.0
                })

            # --- NFSe (GENÉRICO) ---
            else:
                # Localiza bloco InfNfse/InfDeclaracaoPrestacaoServico/Nfse
                infnfse = self._find_node_any(root, ['.//{*}InfNfse', './/{*}InfDeclaracaoPrestacaoServico', './/{*}Nfse/{*}InfNfse', './/{*}Nfse'])
                if infnfse is not None:
                    tipo = 'NFSE'

                    # Número
                    numero = self._find_text_any(infnfse, ['.//{*}Numero', './/{*}NumeroNfse', './/{*}NumeroNfSe'])
                    if not numero:
                        numero = self._find_text_any(root, ['.//{*}Numero', './/{*}NumeroNfse'])
                    nNF = numero if numero else 'N/A'
                    chave = nNF

                    # Verifica cancelamento específico (NfseCancelamento dentro de CompNfse)
                    nfse_cancelamento = root.find('.//{*}NfseCancelamento')
                    if nfse_cancelamento is not None:
                        # marca cancelada
                        status = 'cancelada'
                        # opcional: zera valor em canceladas
                        valor = 0.0
                        logging.info(f"NFSe {nNF} identificada como CANCELADA no arquivo {nome_arquivo}")
                    else:
                        # tenta buscar valor normalmente (ValorLiquidoNfse, ValorServicos)
                        valor_text = self._find_text_any(infnfse, ['.//{*}ValoresNfse/{*}ValorLiquidoNfse', './/{*}Valores/{*}ValorServicos', './/{*}Valores/{*}ValorLiquidoNfse'])
                        if not valor_text:
                            valor_text = self._find_text_any(root, ['.//{*}ValorLiquidoNfse', './/{*}ValorServicos'])
                        try:
                            valor = float(valor_text) if valor_text is not None else 0.0
                        except:
                            valor = 0.0

                    # ItemListaServico / CNAE / CodigoTributacaoMunicipio
                    item_lista = self._find_text_any(infnfse, ['.//{*}ItemListaServico', './/{*}Servico/{*}ItemListaServico', './/{*}Servico/{*}Valores/{*}ItemListaServico'])
                    cnae = self._find_text_any(infnfse, ['.//{*}Servico/{*}CodigoCnae', './/{*}CodigoCnae', './/{*}Cnae'])
                    codtrib = self._find_text_any(infnfse, ['.//{*}CodigoTributacaoMunicipio', './/{*}CodigoTributacao', './/{*}CodigoTribMunicipio', './/{*}CodigoTributacaoMunicipio'])

                    # Prestador e Tomador
                    prestador_cnpj = self._find_text_any(infnfse, ['.//{*}PrestadorServico/{*}CpfCnpj/{*}Cnpj', './/{*}Prestador/{*}CpfCnpj/{*}Cnpj', './/{*}PrestadorServico/{*}Cnpj'])
                    prestador_nome = self._find_text_any(infnfse, ['.//{*}PrestadorServico/{*}RazaoSocial', './/{*}Prestador/{*}RazaoSocial', './/{*}PrestadorServico/{*}NomeFantasia'])

                    tomador_cnpj = self._find_text_any(infnfse, ['.//{*}TomadorServico/{*}IdentificacaoTomador/{*}CpfCnpj/{*}Cnpj', './/{*}TomadorServico/{*}CpfCnpj/{*}Cnpj'])
                    tomador_nome = self._find_text_any(infnfse, ['.//{*}TomadorServico/{*}RazaoSocial', './/{*}TomadorServico/{*}RazaoSocial'])

                    if prestador_cnpj:
                        cnpj_emitente = prestador_cnpj

                    if tomador_cnpj:
                        cnpj_parceiro = tomador_cnpj
                    if tomador_nome:
                        nome_parceiro = tomador_nome
                    elif prestador_nome:
                        nome_parceiro = prestador_nome

                    itens_lista.append({
                        'nNF': nNF,
                        'nItem': 1,
                        'xProd': f'Serviço: Item {item_lista}' if item_lista else 'Serviço Prestado',
                        'NCM': 'N/A',
                        'CFOP': 'N/A',
                        'vProd': valor,
                        'CST_ICMS': 'N/A',
                        'vICMS': 0.0,
                        'nfse_item': item_lista if item_lista else 'N/A',
                        'cnae': cnae if cnae else 'N/A',
                        'codtrib': codtrib if codtrib else 'N/A'
                    })

                    nfse_item_breakdown = (item_lista if item_lista else 'N/A', cnae if cnae else 'N/A', codtrib if codtrib else 'N/A')

            # --- procurar eventos de cancelamento (ProcEventoNFe / infEvento / retEvento) ---
            is_cancellation_event = False
            try:
                if prefix in ['nfe', 'cte']:
                    cancel_ret_evento_tag = root.find(f'.//{prefix}:retEvento/{prefix}:infEvento', {'nfe': 'http://www.portalfiscal.inf.br/nfe', 'cte': 'http://www.portalfiscal.inf.br/cte'})
                    if cancel_ret_evento_tag is not None:
                        cStat_evento_tag = cancel_ret_evento_tag.find(f'{prefix}:cStat', {'nfe': 'http://www.portalfiscal.inf.br/nfe', 'cte': 'http://www.portalfiscal.inf.br/cte'})
                        tpEvento_tag = cancel_ret_evento_tag.find(f'{prefix}:tpEvento', {'nfe': 'http://www.portalfiscal.inf.br/nfe', 'cte': 'http://www.portalfiscal.inf.br/cte'})
                        if cStat_evento_tag is not None and cStat_evento_tag.text == '135' and tpEvento_tag is not None and tpEvento_tag.text == '110111':
                            status = 'cancelada'
                            is_cancellation_event = True

                    if not is_cancellation_event and (root.tag.endswith('procEventoNFe') or root.tag.endswith('procEventoCTe')):
                        tpEvento_tag = root.find(f'.//{prefix}:evento/{prefix}:infEvento/{prefix}:tpEvento', {'nfe': 'http://www.portalfiscal.inf.br/nfe', 'cte': 'http://www.portalfiscal.inf.br/cte'})
                        if tpEvento_tag is not None and tpEvento_tag.text == '110111':
                            status = 'cancelada'
                            is_cancellation_event = True

                    if root.tag.endswith('Proc') and prefix in ['nfe', 'cte']:
                        cStat_prot_tag = root.find(f'.//{prefix}:prot{tipo[-3:]}/{prefix}:infProt/{prefix}:cStat', {'nfe': 'http://www.portalfiscal.inf.br/nfe', 'cte': 'http://www.portalfiscal.inf.br/cte'})
                        if cStat_prot_tag is not None and cStat_prot_tag.text == '101':
                            status = 'cancelada'
                            is_cancellation_event = True
            except Exception:
                pass

            # Lógica de fluxo
            if cnpj_empresa:
                cnpj_emitente_norm = cnpj_emitente.replace('.', '').replace('/', '').replace('-', '').strip() if cnpj_emitente else ''
                cnpj_parceiro_norm = cnpj_parceiro.replace('.', '').replace('/', '').replace('-', '').strip() if cnpj_parceiro else ''

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

            if status == 'cancelada':
                fluxo = 'CANCELADO'

            novo_doc = {
                'nNF': nNF,
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

            # Sobrescrita de cancelamento
            if is_cancellation_event:
                if chave in self.documentos_processados:
                    doc_existente = self.documentos_processados[chave]
                    doc_existente['status'] = 'cancelada'
                    doc_existente['fluxo'] = 'CANCELADO'
                    doc_existente['inconsistencia_fiscal'] = 'CANCELADA'
                    doc_existente['arquivo'] += f" / Evento: {nome_arquivo}"
                    GLOBAL_ITEM_DETAILS.pop(chave, None)
                else:
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
            else:
                # se documento principal, armazena (somente sobrescreve se não estiver cancelada)
                if chave not in self.documentos_processados or self.documentos_processados[chave]['status'] != 'cancelada':
                    self.documentos_processados[chave] = novo_doc
                    # Salva itens (se houver)
                    GLOBAL_ITEM_DETAILS[chave] = itens_lista

                # se NF-e autorizada com produtos, calcule cfop_breakdown
                if tipo == 'NF-e' and novo_doc['status'] == 'autorizada' and itens_lista:
                    for item in itens_lista:
                        cf = item.get('CFOP')
                        if cf:
                            novo_doc['cfop_breakdown'].append({'cfop': cf, 'vProd': item.get('vProd', 0.0)})
                            # inconsistência simples
                            if cf and tpNF in ['0', '1']:
                                if tpNF == '0' and cf[0] in ['5','6','7']:
                                    novo_doc['inconsistencia_fiscal'] = 'CFOP SAÍDA em NF ENTRADA'
                                    break
                                elif tpNF == '1' and cf[0] in ['1','2','3']:
                                    novo_doc['inconsistencia_fiscal'] = 'CFOP ENTRADA em NF SAÍDA'
                                    break

                # para NFSe, atualiza agregação por item/cnae/codtrib (se autorizada)
                if tipo == 'NFSE' and novo_doc['status'] == 'autorizada' and itens_lista:
                    for item in itens_lista:
                        key = (item.get('nfse_item', 'N/A'), item.get('cnae', 'N/A'), item.get('codtrib', 'N/A'))
                        self.nfse_item_totals[key] += item.get('vProd', 0.0)

        except Exception as e:
            logging.exception(f"Erro inesperado ao processar {nome_arquivo}: {e}")
            self._registrar_erro(f"Erro inesperado {nome_arquivo}: {e}")

    # -----------------------
    # Detectar quebras de sequência
    # -----------------------
    def detectar_quebra_sequencia(self):
        """
        Para cada emitente (cnpj_emitente), verifica seq. numérica de nNF para NF-e e NFSe (não canceladas).
        Produz alertas se houver saltos ou duplicatas.
        """
        sequencias = defaultdict(list)
        for chave, doc in self.documentos_processados.items():
            try:
                if doc["tipo"] in ["NF-e", "NFSE"] and doc["status"] != "cancelada":
                    num = int(str(doc.get("nNF")).strip())
                    cnpj = doc.get("cnpj_emitente", "N/A")
                    sequencias[cnpj].append(num)
            except Exception:
                # ignora nNF não numérico
                continue

        alerts = []
        for cnpj, nums in sequencias.items():
            nums_ordenados = sorted(set(nums))
            if not nums_ordenados:
                continue
            # checar duplicatas (antes de set)
            duplicates = [n for n in nums if nums.count(n) > 1]
            if duplicates:
                for d in sorted(set(duplicates)):
                    alerts.append(f"CNPJ {cnpj}: duplicata número {d}")

            for i in range(1, len(nums_ordenados)):
                prev_n = nums_ordenados[i-1]
                cur_n = nums_ordenados[i]
                if cur_n != prev_n + 1:
                    alerts.append(f"CNPJ {cnpj}: quebra entre {prev_n} e {cur_n} (faltam {cur_n - prev_n - 1} números)")
        return alerts

    # -----------------------
    # Atualizar UI
    # -----------------------
    def atualizar_interface(self):
        # limpar treeviews
        for i in self.tree.get_children():
            self.tree.delete(i)
        for i in self.tree_cfop.get_children():
            self.tree_cfop.delete(i)
        for i in self.tree_nfse.get_children():
            self.tree_nfse.delete(i)
        for i in self.tree_partner.get_children():
            self.tree_partner.delete(i)
        for i in self.tree_itens.get_children():
            self.tree_itens.delete(i)

        total_valor = 0.0
        counts = defaultdict(lambda: defaultdict(int))
        self.cfop_totals.clear()
        partner_acc = defaultdict(float)

        for chave, doc in sorted(self.documentos_processados.items()):
            valor_formatado = f"{doc['valor']:.2f}".replace('.', ',')
            tag = "cancelada" if doc['status'] == 'cancelada' else ("inconsistencia" if doc.get('inconsistencia_fiscal') != 'OK' else "")
            self.tree.insert("", "end", values=(
                chave,
                doc.get('tipo', 'N/A'),
                doc.get('nNF', 'N/A'),
                valor_formatado,
                doc['status'].title(),
                doc.get('fluxo', 'N/A'),
                doc.get('nome_parceiro', 'N/A'),
                doc.get('cnpj_parceiro', 'N/A'),
                doc.get('inconsistencia_fiscal', 'N/A'),
                doc['arquivo']
            ), tags=(tag,))

            self.tree.tag_configure('cancelada', foreground='red')
            self.tree.tag_configure('inconsistencia', foreground='#FFD700')

            counts[doc['status']][doc.get('tipo', 'N/A')] += 1

            if doc['status'] == 'autorizada':
                total_valor += doc['valor']

                if doc.get('tipo') == 'NF-e' and 'cfop_breakdown' in doc:
                    for item in doc['cfop_breakdown']:
                        self.cfop_totals[item['cfop']] += item['vProd']

                partner_key = (doc.get('cnpj_parceiro', 'N/A'), doc.get('nome_parceiro', 'N/A'))
                if partner_key[0] != 'N/A':
                    partner_acc[partner_key] += doc['valor']

        # atualizar labels
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
        self.label_arquivos_processados.configure(text=f"Arquivos Processados: {self.arquivos_contados}")

        # atualizar contador de erros
        self.label_erros.configure(text=f"Erros: {self.erros_detectados}")

        # atualizar itens (GLOBAL_ITEM_DETAILS)
        for chave, itens in GLOBAL_ITEM_DETAILS.items():
            doc = self.documentos_processados.get(chave)
            if doc and doc['status'] == 'autorizada':
                for item in itens:
                    valor_item_formatado = f"{item['vProd']:.2f}".replace('.', ',')
                    valor_icms_formatado = f"{item.get('vICMS', 0.0):.2f}".replace('.', ',')
                    self.tree_itens.insert("", "end", values=(
                        item.get('nNF', 'N/A'),
                        item.get('xProd', 'N/A'),
                        item.get('NCM', 'N/A'),
                        item.get('CFOP', 'N/A'),
                        valor_item_formatado,
                        item.get('CST_ICMS', 'N/A'),
                        valor_icms_formatado
                    ))

        # atualizar CFOP
        for cfop, valor in sorted(self.cfop_totals.items()):
            valor_cfop_formatado = f"{valor:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
            self.tree_cfop.insert("", "end", values=(cfop, valor_cfop_formatado))

        # atualizar NFSe agregada (ItemLista / CNAE / CodTrib)
        for key, valor in sorted(self.nfse_item_totals.items(), key=lambda x: (x[0][0] or '', x[0][1] or '', x[0][2] or '')):
            item_lista, cnae, codtrib = key
            valor_nfse_formatado = f"{valor:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
            self.tree_nfse.insert("", "end", values=(item_lista, cnae, codtrib, valor_nfse_formatado))

        # atualizar parceiro
        for partner_key, valor in sorted(partner_acc.items(), key=lambda x: x[1], reverse=True):
            cnpj_parceiro, nome_parceiro = partner_key
            valor_parceiro_formatado = f"{valor:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
            self.tree_partner.insert("", "end", values=(nome_parceiro, cnpj_parceiro, valor_parceiro_formatado))

        # recalcula e exibe quebras de sequência
        self.quebras_sequencia_alerts = self.detectar_quebra_sequencia()
        self.label_quebras_count.configure(text=f"Quebras sequência: {len(self.quebras_sequencia_alerts)}")

    # -----------------------
    # Registrar erro
    # -----------------------
    def _registrar_erro(self, mensagem):
        """Incrementa contador, loga e atualiza label"""
        self.erros_detectados += 1
        logging.error(mensagem)
        # atualiza label imediata (se UI já criada)
        try:
            self.label_erros.configure(text=f"Erros: {self.erros_detectados}")
        except Exception:
            pass

    # -----------------------
    # Mostrar popup com quebras
    # -----------------------
    def mostrar_quebras_popup(self):
        if not self.quebras_sequencia_alerts:
            messagebox.showinfo("Quebras de Sequência", "Nenhuma quebra detectada.")
            return
        top = ctk.CTkToplevel(self)
        top.title("Quebras de Sequência Detectadas")
        top.geometry("700x400")
        txt = tkinter.Text(top, wrap="word")
        txt.pack(fill="both", expand=True, padx=6, pady=6)
        for a in self.quebras_sequencia_alerts:
            txt.insert("end", a + "\n")
        txt.configure(state="disabled")

    # -----------------------
    # Exportar CSV
    # -----------------------
    def exportar_csv(self):
        if not self.documentos_processados:
            messagebox.showwarning("Exportar CSV", "Nenhum documento processado para exportar.")
            return
        caminho = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV", "*.csv")], title="Salvar resumo como")
        if not caminho:
            return
        try:
            with open(caminho, 'w', newline='', encoding='utf-8') as csvfile:
                writer = csv.writer(csvfile)
                writer.writerow(["Chave", "Tipo", "Nº Doc", "Valor", "Status", "Fluxo", "Parceiro", "CNPJ/CPF", "Inconsistência", "Arquivo"])
                for chave, doc in sorted(self.documentos_processados.items()):
                    writer.writerow([
                        chave,
                        doc.get('tipo'),
                        doc.get('nNF'),
                        f"{doc.get('valor', 0.0):.2f}",
                        doc.get('status'),
                        doc.get('fluxo'),
                        doc.get('nome_parceiro'),
                        doc.get('cnpj_parceiro'),
                        doc.get('inconsistencia_fiscal'),
                        doc.get('arquivo')
                    ])
            messagebox.showinfo("Exportar CSV", f"Resumo exportado para:\n{caminho}")
        except Exception as e:
            logging.exception(f"Erro ao exportar CSV: {e}")
            self._registrar_erro(f"Erro exportar CSV: {e}")
            messagebox.showerror("Exportar CSV", f"Erro ao exportar CSV: {e}")

    # -----------------------
    # Limpar
    # -----------------------
    def limpar_resultados(self, clear_files=True):
        self.documentos_processados.clear()
        self.cfop_totals.clear()
        self.nfse_item_totals.clear()
        self.partner_totals.clear()
        GLOBAL_ITEM_DETAILS.clear()
        self.quebras_sequencia_alerts.clear()
        if clear_files:
            self.arquivos_contados = 0
            self.erros_detectados = 0
        self.atualizar_interface()

if __name__ == "__main__":
    try:
        rarfile.PATH_SEP = os.path.sep
        zipfile.Path = os.path.join
    except AttributeError:
        pass

    app = App()
    app.mainloop()
