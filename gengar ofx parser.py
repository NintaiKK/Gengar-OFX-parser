import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pdfplumber
from openpyxl import Workbook, load_workbook
import ofxparse
import os

class ParserApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Conversor PDF/XLSX/OFX")
        self.root.geometry("600x400")
        
        # Variáveis de controle
        self.pdf_path = tk.StringVar()
        self.xlsx_path = tk.StringVar()
        self.ofx_path = tk.StringVar()
        
        # Criar interface
        self.create_widgets()
        
    def create_widgets(self):
        # Frame principal
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        
        # Abas
        notebook = ttk.Notebook(main_frame)
        notebook.pack(fill=tk.BOTH, expand=True)
        
        # Aba PDF para XLSX
        tab_pdf_xlsx = ttk.Frame(notebook)
        notebook.add(tab_pdf_xlsx, text="PDF para XLSX")
        
        # Widgets PDF para XLSX
        ttk.Label(tab_pdf_xlsx, text="Arquivo PDF:").grid(row=0, column=0, sticky=tk.W, pady=5)
        ttk.Entry(tab_pdf_xlsx, textvariable=self.pdf_path, width=50).grid(row=0, column=1, pady=5)
        ttk.Button(tab_pdf_xlsx, text="Procurar", command=self.browse_pdf).grid(row=0, column=2, pady=5)
        
        ttk.Label(tab_pdf_xlsx, text="Arquivo XLSX de saída:").grid(row=1, column=0, sticky=tk.W, pady=5)
        ttk.Entry(tab_pdf_xlsx, textvariable=self.xlsx_path, width=50).grid(row=1, column=1, pady=5)
        ttk.Button(tab_pdf_xlsx, text="Procurar", command=self.browse_xlsx_output).grid(row=1, column=2, pady=5)
        
        ttk.Button(tab_pdf_xlsx, text="Converter PDF para XLSX", 
                  command=self.convert_pdf_to_xlsx).grid(row=2, column=1, pady=20)
        
        # Aba XLSX para OFX
        tab_xlsx_ofx = ttk.Frame(notebook)
        notebook.add(tab_xlsx_ofx, text="XLSX para OFX")
        
        # Widgets XLSX para OFX
        ttk.Label(tab_xlsx_ofx, text="Arquivo XLSX:").grid(row=0, column=0, sticky=tk.W, pady=5)
        ttk.Entry(tab_xlsx_ofx, textvariable=self.xlsx_path, width=50).grid(row=0, column=1, pady=5)
        ttk.Button(tab_xlsx_ofx, text="Procurar", command=self.browse_xlsx).grid(row=0, column=2, pady=5)
        
        ttk.Label(tab_xlsx_ofx, text="Arquivo OFX de saída:").grid(row=1, column=0, sticky=tk.W, pady=5)
        ttk.Entry(tab_xlsx_ofx, textvariable=self.ofx_path, width=50).grid(row=1, column=1, pady=5)
        ttk.Button(tab_xlsx_ofx, text="Procurar", command=self.browse_ofx_output).grid(row=1, column=2, pady=5)
        
        ttk.Button(tab_xlsx_ofx, text="Converter XLSX para OFX", 
                  command=self.convert_xlsx_to_ofx).grid(row=2, column=1, pady=20)
        
        # Barra de status
        self.status = ttk.Label(main_frame, text="Pronto", relief=tk.SUNKEN)
        self.status.pack(fill=tk.X, pady=10)
    
    def browse_pdf(self):
        filename = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
        if filename:
            self.pdf_path.set(filename)
            # Sugerir nome do arquivo XLSX
            base = os.path.splitext(filename)[0]
            self.xlsx_path.set(base + ".xlsx")
    
    def browse_xlsx(self):
        filename = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        if filename:
            self.xlsx_path.set(filename)
            # Sugerir nome do arquivo OFX
            base = os.path.splitext(filename)[0]
            self.ofx_path.set(base + ".ofx")
    
    def browse_xlsx_output(self):
        filename = filedialog.asksaveasfilename(defaultextension=".xlsx", 
                                              filetypes=[("Excel Files", "*.xlsx")])
        if filename:
            self.xlsx_path.set(filename)
    
    def browse_ofx_output(self):
        filename = filedialog.asksaveasfilename(defaultextension=".ofx", 
                                              filetypes=[("OFX Files", "*.ofx")])
        if filename:
            self.ofx_path.set(filename)
    
    def convert_pdf_to_xlsx(self):
        pdf_file = self.pdf_path.get()
        xlsx_file = self.xlsx_path.get()
        
        if not pdf_file or not xlsx_file:
            messagebox.showerror("Erro", "Por favor, especifique os arquivos de entrada e saída")
            return
        
        try:
            self.status.config(text="Convertendo PDF para XLSX...")
            self.root.update()
            
            # Criar um novo workbook Excel
            wb = Workbook()
            ws = wb.active
            ws.title = "Dados do PDF"
            
            # Extrair texto do PDF
            with pdfplumber.open(pdf_file) as pdf:
                for page in pdf.pages:
                    text = page.extract_text()
                    if text:
                        for line in text.split('\n'):
                            # Adicionar cada linha como uma linha no Excel
                            ws.append([line])
            
            # Salvar o arquivo Excel
            wb.save(xlsx_file)
            
            messagebox.showinfo("Sucesso", f"Arquivo convertido com sucesso para:\n{xlsx_file}")
            self.status.config(text="Pronto")
            
        except Exception as e:
            messagebox.showerror("Erro", f"Ocorreu um erro durante a conversão:\n{str(e)}")
            self.status.config(text="Erro")
    
    def convert_xlsx_to_ofx(self):
        xlsx_file = self.xlsx_path.get()
        ofx_file = self.ofx_path.get()
        
        if not xlsx_file or not ofx_file:
            messagebox.showerror("Erro", "Por favor, especifique os arquivos de entrada e saída")
            return
        
        try:
            self.status.config(text="Convertendo XLSX para OFX...")
            self.root.update()
            
            # Carregar o arquivo Excel
            wb = load_workbook(xlsx_file)
            ws = wb.active
            
            # Criar um arquivo OFX básico
            # NOTA: Esta é uma implementação simplificada. O formato OFX é complexo.
            # Em uma aplicação real, você precisaria mapear os dados do Excel para
            # a estrutura OFX adequada, incluindo cabeçalhos, contas, transações, etc.
            
            ofx_content = """OFXHEADER:100
DATA:OFXSGML
VERSION:102
SECURITY:NONE
ENCODING:USASCII
CHARSET:1252
COMPRESSION:NONE
OLDFILEUID:NONE
NEWFILEUID:NONE

<OFX>
    <SIGNONMSGSRSV1>
        <SONRS>
            <STATUS>
                <CODE>0
                <SEVERITY>INFO
            </STATUS>
            <DTSERVER>20230514
            <LANGUAGE>POR
        </SONRS>
    </SIGNONMSGSRSV1>
    <BANKMSGSRSV1>
        <STMTTRNRS>
            <TRNUID>1
            <STATUS>
                <CODE>0
                <SEVERITY>INFO
            </STATUS>
            <STMTRS>
                <CURDEF>BRL
                <BANKACCTFROM>
                    <BANKID>123
                    <ACCTID>456789
                    <ACCTTYPE>CHECKING
                </BANKACCTFROM>
                <BANKTRANLIST>
                    <DTSTART>20230501
                    <DTEND>20230514
"""
            
            # Adicionar transações (simplificado)
            for row in ws.iter_rows(values_only=True):
                if row and row[0]:  # Se a célula não estiver vazia
                    # Aqui você precisaria mapear os dados do Excel para o formato OFX
                    # Esta é uma implementação muito básica
                    ofx_content += f"                    <STMTTRN>\n"
                    ofx_content += f"                        <TRNTYPE>DEBIT\n"
                    ofx_content += f"                        <DTPOSTED>20230514\n"
                    ofx_content += f"                        <TRNAMT>-100.00\n"
                    ofx_content += f"                        <FITID>123456789\n"
                    ofx_content += f"                        <MEMO>{row[0]}\n"
                    ofx_content += f"                    </STMTTRN>\n"
            
            ofx_content += """                </BANKTRANLIST>
                <LEDGERBAL>
                    <BALAMT>1000.00
                    <DTASOF>20230514
                </LEDGERBAL>
            </STMTRS>
        </STMTTRNRS>
    </BANKMSGSRSV1>
</OFX>"""
            
            # Salvar o arquivo OFX
            with open(ofx_file, 'w') as f:
                f.write(ofx_content)
            
            messagebox.showinfo("Sucesso", f"Arquivo convertido com sucesso para:\n{ofx_file}")
            self.status.config(text="Pronto")
            
        except Exception as e:
            messagebox.showerror("Erro", f"Ocorreu um erro durante a conversão:\n{str(e)}")
            self.status.config(text="Erro")

if __name__ == "__main__":
    root = tk.Tk()
    app = ParserApp(root)
    root.mainloop()
