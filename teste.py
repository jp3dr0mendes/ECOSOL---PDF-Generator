from tkinter import *
from tkinter import ttk
import webbrowser
import base64
from PIL import ImageTk, Image
from docx import Document
from docx.shared import Cm,Pt,Inches
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx2pdf import convert
import aspose.words as aw
from reportlab.pdfgen import canvas
# import fitz
import os
import io
import PyPDF2 as ppdf

#class File_Functions(File):
class File_Functions():
    def mm2p(mm: int):
        return mm/0.352777
    
    def base64_images_decode(self):
        print('qlq coisa')

    def document_set_margins(self,document):
        sections = document.sections

        for section in sections:
            section.top_margin = Cm(3)
            section.right_margin = Cm(2)
            section.bottom_margin = Cm(2)
            section.left_margin = Cm(3)

#Creating a class for documents notations
class File(File_Functions):
    def __init__(self):
        self.document = Document()
        self.cnv = None
        self.style = self.document.styles['Normal']
        self.font = self.style.font
        self.font.name = 'Arial'
        # self.pages=[]
        # self.contents=[]
        # self.client=''
        # self.name=''
        # self.table_clients=''
        self.pg_one()  # webbrowser.open('teste1.pdf')
    
    def pg_one(self):
        for i in range(100):
            self.document.add_paragraph('teste')
        paragraph = self.document.add_paragraph('Teste1')
        paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

    def create_file(self):
        # # print(self.document)
        # for paragraph in self.document.paragraphs:
        #     # print('paragraph:')
        #     # print(paragraph)
        #     if '[Cliente]' in paragraph.text:
        #         paragraph.text =  paragraph.text.replace('[Cliente]',self.client)
        #     if '[Tabela Cliente]' in paragraph.text:
        #         print('espaço de tabela achado')
        #         self.table_clients_write(paragraph)
        #         # self.document.add_table(rows=1,cols=2)

        print(self.name)
        print(self.client)
        print('passou aqui')
        self.document.save(self.name+'.docx')
        convert(self.name+'.docx',self.name+'.pdf')
        os.remove(self.name+'.docx')
        
        webbrowser.open(self.name+'.pdf')
    #starting development to put images on file
    def file_reader(self):
        self.file=open(self.dname+'.pdf','rb')
        self.cnv = canvas.Canvas(self.name+'.pdf')
        self.file_content = ppdf.PdfFileReader(self.file)
        
        print('pdf content readed')

        for i in range(self.file_content.numPages):
            self.pages.append(self.file_content.getPage[i])
        
        print('pdf pages getted')

        for i in range(self.file_content.numPages):
            self.contents.append(self.pages[i].extractText())

        print(self.contents)
        print('pdf pages contents getted')
    #Function to make a client data table
    def table_clients_write(self,place):
        self.client_data = (
            ('Cliente',self.client),
            ('Nome Fantasia',self.client_name_f),
            ('CNPJ',self.client_cnpj),
            ('N°s do Cliente',self.client_num),
            ('Modalidade da Tarifa',self.modalidade_tarifa),
            ('Classe/Subclasse',self.client_classe),
            ('Endereço',self.client_address),
            ('Cidade',self.client_city),
            ('Estado',self.client_state),
            ('Telefone',self.client_phone),
            ('E-mail',self.client_mail),
            ('Contato',self.client_contact),
            ('Ramo da Atividade',self.client_activity)
        )

        table = self.document.add_table(1,2)

        for data,info in self.client_data:
            row = table.add_row().cells
            row[0].text = str(data)
            row[1].text = info
        print('Tabela de clientes criada')
        print(type(table))
        

#Creating a main application functions
class Application_Functions():
    def base64_images(self):
        self.ecosol_logo= "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"
    def file_generate(self):
        self.gets_entry_values()
        self.pdf.create_file()

    def gets_entry_values(self):
        self.pdf.name= self.document_name_entry.get()
        self.pdf.client = self.document_client_entry.get()
        self.pdf.client = self.client_name_entry.get()
        self.pdf.client_name_f = self.client_name_f_entry.get()
        self.pdf.client_cnpj = self.client_cnpj_entry.get()
        self.pdf.client_num = self.client_num_entry.get()
        self.pdf.modalidade_tarifa = self.client_mod_entry.get()
        self.pdf.client_classe = self.client_class_entry.get()
        self.pdf.client_address = self.client_address_entry.get()
        self.pdf.client_city = self.client_city_entry.get()
        self.pdf.client_state = self.client_state_entry.get()
        self.pdf.client_phone = self.client_phone_entry.get()
        self.pdf.client_mail = self.client_mail_entry.get()
        self.pdf.client_contact = self.client_contact_entry.get()
        self.pdf.client_activity = self.client_activity_entry.get()

        print(self.pdf.name)
        print(self.pdf.client)

#building the application
class Application(Application_Functions):
    def __init__(self):
        self.root=Tk()
        self.pdf = File()
        self.base64_images()
        self.screen()
        self.window_buttons()
        self.frame_logo()
        self.main_frame()
        self.pages()
        self.widgets_page1()
        self.widgets_page2()
        self.labels_frame_logo()
        self.root.mainloop()

    def screen(self):
        self.root.title('ECOSOL - PDF AUTO GENERATE')
        self.root.geometry("1000x700")
        self.root.resizable(True,True)
        #self.root.maxsize(width=,height)
        #self.root.minsize(width=,height=)
    #set first frame definitions
    def frame_logo(self):
        self.fr_logo=Frame(self.root,bd=4,)
        self.fr_logo.place(relx=0.3,rely=0.01,relwidth=0.4,relheight=0.2)
    #setting main frame definitions
    def main_frame(self):
        self.main_fr=Frame(self.root,bd=4,highlightbackground='black',highlightthickness=0.5)
        self.main_fr.place(relx=0.01,rely=0.2,relwidth=0.98,relheight=0.725)
    #setting notebook in main frame
    def pages(self):
        self.notebook = ttk.Notebook(self.main_fr)
        self.page1 = Frame(self.notebook)
        self.page2 = Frame(self.notebook)

        self.page1.configure(background='lightgray')
        self.page2.configure(background='lightgray')

        self.notebook.add(self.page1,text='Document Settings')
        self.notebook.add(self.page2, text='Client Information')

        self.notebook.place(relx=0,rely=0,relwidth=1,relheight=1)
    #setting logo
    def labels_frame_logo(self):
        self.ecosol_img = PhotoImage(data=base64.b64decode(self.ecosol_logo))
        # self.ecosol_img = self.resize_image(self.ecosol_logo,2,2)
        # self.ecosol_img_aux = PhotoImage(self.ecosol_img)
        self.ecosol_img.subsample(2,2)
        self.ecosol_label_img = Label(self.fr_logo, image=self.ecosol_img)
        self.ecosol_label_img.place(relx=0.0,rely=0.0,relheight=1,relwidth=1)
    def widgets_page1(self):
        #creating page1 labels
        self.document_name_label = Label(self.page1,text="Document Name:",bg = 'lightgray',)
        self.document_client_label = Label(self.page1,text="Client Name:",bg = "lightgray")
        
        #locating page1 labels
        self.document_name_label.place(relx=0,rely=0,relwidth=0.12,relheight=0.1)
        self.document_client_label.place(relx=0,rely=0.11,relwidth=0.1,relheight=0.1)
        
        #creating page1 entrys
        self.document_name_entry = Entry(self.page1)
        self.document_name_entry.place(relx=0.01,rely=0.071,relwidth=0.15,relheight=0.06)
        
        self.document_client_entry = Entry(self.page1)
        self.document_client_entry.place(relx=0.01,rely=0.181,relwidth=0.15,relheight=0.06)

    def widgets_page2(self):
        self.client_name_label = Label(self.page2,text='Nome do Cliente',bg='lightgray')
        self.client_name_label.place(relx=0.01,rely=0.0,relwidth=0.10,relheight=0.1)

        self.client_name_entry = Entry(self.page2)
        self.client_name_entry.place(relx=0.01,rely=0.08,relwidth=0.15,relheight=0.06)

        self.client_name_f_label = Label(self.page2,text='Nome Fantasia',bg='lightgray')
        self.client_name_f_label.place(relx=0.23,rely=0.0,relwidth=0.15,relheight=0.1)

        self.client_name_f_entry = Entry(self.page2)
        self.client_name_f_entry.place(relx=0.25,rely=0.08,relwidth=0.15,relheight=0.06)
        
        self.client_cnpj_label = Label(self.page2,text='CNPJ',bg='lightgray')
        self.client_cnpj_label.place(relx=0.45,rely=0.0,relwidth=0.1,relheight=0.1)

        self.client_cnpj_entry = Entry(self.page2)
        self.client_cnpj_entry.place(relx=0.48,rely=0.08,relwidth=0.15,relheight=0.06)
        
        self.client_num_label = Label(self.page2,text='N°s do Cliente',bg='lightgray')
        self.client_num_label.place(relx=0.7,rely=0.0,relwidth=0.15,relheight=0.1)

        self.client_num_entry = Entry(self.page2)
        self.client_num_entry.place(relx=0.73,rely=0.08,relwidth=0.15,relheight=0.06)

        self.client_mod_label = Label(self.page2,text='Modalidade Tarifária',bg='lightgray')
        self.client_mod_label.place(relx=0.015,rely=0.15,relwidth=0.115,relheight=0.1)

        self.client_mod_entry = Entry(self.page2)
        self.client_mod_entry.place(relx=0.01,rely=0.23,relwidth=0.15,relheight=0.06)
        
        self.client_class_label = Label(self.page2,text='Classe/Subclasse',bg='lightgray')
        self.client_class_label.place(relx=0.25,rely=0.15,relwidth=0.11,relheight=0.1)

        self.client_class_entry = Entry(self.page2)
        self.client_class_entry.place(relx=0.25,rely=0.23,relwidth=0.15,relheight=0.06)

        self.client_address_label = Label(self.page2,text='Endedreço',bg='lightgray')
        self.client_address_label.place(relx=0.463,rely=0.15,relwidth=0.11,relheight=0.1)

        self.client_address_entry = Entry(self.page2)
        self.client_address_entry.place(relx=0.48,rely=0.23,relwidth=0.15,relheight=0.06)

        self.client_city_label = Label(self.page2,text='Cidade',bg='lightgray')
        self.client_city_label.place(relx=0.705,rely=0.15,relwidth=0.11,relheight=0.1)

        self.client_city_entry = Entry(self.page2)
        self.client_city_entry.place(relx=0.73,rely=0.23,relwidth=0.15,relheight=0.06)

        self.client_state_label = Label(self.page2,text='Estado',bg='lightgray')
        self.client_state_label.place(relx=0,rely=0.3,relwidth=0.08,relheight=0.1)

        self.client_state_entry = Entry(self.page2)
        self.client_state_entry.place(relx=0.015,rely=0.38,relwidth=0.15,relheight=0.06)

        self.client_phone_label = Label(self.page2,text='Telefone',bg='lightgray')
        self.client_phone_label.place(relx=0.22,rely=0.3,relwidth=0.12,relheight=0.1)

        self.client_phone_entry = Entry(self.page2)
        self.client_phone_entry.place(relx=0.25,rely=0.38,relwidth=0.15,relheight=0.06)

        self.client_mail_label = Label(self.page2,text='E-mail',bg='lightgray')
        self.client_mail_label.place(relx=0.47,rely=0.3,relwidth=0.08,relheight=0.1)

        self.client_mail_entry = Entry(self.page2)
        self.client_mail_entry.place(relx=0.48,rely=0.38,relwidth=0.15,relheight=0.06)

        self.client_contact_label = Label(self.page2,text='Contato',bg='lightgray')
        self.client_contact_label.place(relx=0.72,rely=0.3,relwidth=0.08,relheight=0.1)

        self.client_contact_entry = Entry(self.page2)
        self.client_contact_entry.place(relx=0.73,rely=0.38,relwidth=0.15,relheight=0.06)

        self.client_activity_label = Label(self.page2,text='Ramo de Atividade',bg='lightgray')
        self.client_activity_label.place(relx=0.01,rely=0.475,relwidth=0.13,relheight=0.1)

        self.client_activity_entry = Entry(self.page2)
        self.client_activity_entry.place(relx=0.015,rely=0.57,relwidth=0.15,relheight=0.06)
        
    def window_buttons(self):
        self.generate_pdf_button = Button(self.root,text = "Gerar PDF",command = self.file_generate)
        self.generate_pdf_button.place(relx=0.45,rely=0.94,relwidth=0.1,relheight=0.04)

        self.quit_application_button = Button(self.root,text='Quit',command = self.root.destroy)
        self.quit_application_button.place(relx=0.85,rely=0.94,relwidth=0.1,relheight=0.04) 

Application()