import tempfile
import pandas as pd
import tkinter as tk
from tkinter import PhotoImage, ttk
from tkinter.messagebox import showinfo
from tkinter import filedialog
from PIL import ImageTk, Image
import os
from fpdf import FPDF
import pyperclip

root = tk.Tk()
root.title("Sistema ERF (Emissor de Relatório Fotográfico)")
window_width = 1050
window_height = 690

# get the screen dimension
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()

# find the center point
center_x = int(screen_width/2 - window_width / 2)
center_y = int(screen_height/2 - window_height / 2)
root.geometry(f'{window_width}x{window_height}+{center_x}+{center_y}')
# root.iconbitmap('./assets/CDE.ico')
im = Image.open('./assets/CDE.ico')
photo = ImageTk.PhotoImage(im)
root.wm_iconphoto(True, photo)
root.config(background='#eeeeee')


# Adiciona a foto
if "nt" == os.name:
    photoGerar = PhotoImage(file=r"assets\PDF.png")
else:
    photoGerar = PhotoImage(file="./assets/PDF.png")
    
# Cria o botão de gerar relatório
gerbtn = ttk.Button(root, text="Gerar Relatório (PDF)",
                   image=photoGerar, compound='left', command=lambda: genreport())
gerbtn.place(x=450, y=650)
outDir = ''
excelPlan = ''
photosdirask = ''


columnsSel = ('Caminhofoto')

treeSel = ttk.Treeview(root, columns=columnsSel, show='headings')

treeSel.heading('Caminhofoto', text='Caminho da foto')


treeSel.bind('<<TreeviewSelect>>')

treeSel.grid(row=0, column=0, sticky='nsew')

treeSel.place(x=25, y=290, width=700)

columnsName = ('Nomefoto')


def copysel():
    sel_item = treeName.focus()
    sel_items = treeName.selection()

    det = treeName.item(sel_item)

    if len(sel_items) == 1:
        pyperclip.copy(det['values'][0])
    else:
        temp_var = ''
        for item in sel_items:
            cur_item = treeName.item(item)
            cur_text = cur_item['values']
            temp_var += str(cur_text[0]) + '\n'
        pyperclip.copy(temp_var)


treeName = ttk.Treeview(root, columns=columnsName,
                        show='headings', selectmode=tk.EXTENDED)

treeName.heading('Nomefoto', text='Nome da foto')


treeName.bind('<<TreeviewSelect>>')

treeName.grid(row=0, column=0, sticky='nsew')

treeName.place(x=735, y=290, width=290)

btnCopy = ttk.Button(root, text='Copiar seleção', command=copysel)
btnCopy.place(x=755, y=520)


def ask(q):
    if q == 'out':
        outDir = filedialog.askdirectory(
            title='Selecione a pasta de saída', initialdir='.')
        labelOut.config(text=outDir)
    elif q == 'plan':
        filetypes = (
            ('Planilha excel', '*.xls'),
            ('Planilha excel', '*.xlsx')
        )
        excelPlan = filedialog.askopenfilename(
            title='Selecione a planilha', filetypes=filetypes, initialdir='.')
        labelExc.config(text=excelPlan)
    elif q == 'pho':
        photosdirask = filedialog.askdirectory(
            title='Selecione a pasta das fotos', initialdir='.')
        labelPho.config(text=photosdirask)
        files = os.listdir(photosdirask)
        treeSel.delete(*treeSel.get_children())
        treeName.delete(*treeName.get_children())
        filetypes = ('.png', '.jpg', '.jpeg', '.jfif')
        for file in files:
            if file.endswith(filetypes):
                fileName = file.split('.')[0]
                treeSel.insert('', tk.END, values=f'{photosdirask}/{file}')
                treeName.insert('', tk.END, values=f'{fileName}')


def genreport():
    planilhaPath = labelExc.cget('text')
    outPath = labelOut.cget('text')
    photosPath = labelPho.cget('text')
    if len(planilhaPath) != 0 and len(outPath) != 0 and len(photosPath) != 0:

        df = pd.read_excel(planilhaPath)

        pdf = FPDF()
        pdf.set_auto_page_break(auto=False)

        def titleCabec():
            pdf.set_xy(0.0, 0.0)
            pdf.set_font('Helvetica', 'B', 11)
            pdf.set_text_color(0, 0, 0)
            if entryNomeEmp.get() != 0:
                pdf.cell(w=210.0, h=20.0, align='C',
                        txt=entryNomeEmp.get(), border=0)
            else:
                print("Erro", "Coloque o Nome do Condominio")

            pdf.set_xy(0.0, 0.0)
            pdf.set_font('Helvetica', '', 11)
            pdf.set_text_color(0, 0, 0)
            if len(entryEnderEmp.get()) != 0:
                listPalavra = entryEnderEmp.get().split('\\n')
                pdf.cell(w=210.0, h=34.0, align='C',
                        txt=listPalavra[0], border=0)
                pdf.set_xy(0.0, 0.0)
                pdf.set_font('Helvetica', '', 11)
                pdf.set_text_color(0, 0, 0)
                pdf.cell(w=210.0, h=45.0, align='C',
                        txt=listPalavra[1], border=0)
                pdf.set_xy(0.0, 0.0)
                pdf.set_font('Helvetica', '', 11)
                pdf.set_text_color(0, 0, 0)
                pdf.cell(w=210.0, h=56.0, align='C',
                        txt=listPalavra[2], border=0)

            

        def logo():
            pdf.set_xy(6.0, 0.0)
            logo = 'assets/CDE.png'
            pdf.image(logo, link='', type='', w=250/6, h=250/6)

        def barrinha(ambiente):

            pdf.set_fill_color(47, 11, 97)
            pdf.rect(8.0, 35.0, 193.0, 7.00, 'F')
            pdf.set_xy(9.0, 35.4)
            pdf.set_font('Helvetica', 'B', 14)
            pdf.set_text_color(255, 255, 255)
            pdf.cell(txt=str(ambiente), border=0, align='C', w=193.0, h=7.0)
            # writeText(ambiente, 98.0, 36, 16, 'B', 255, 255, 255)

        def addNumeroPaginaFirst():
            writeText(str(numPag), 196.0, 292.0, 10, '')

        def add_first_page(gut: bool = False):
            pdf.add_page()
            titleCabec()
            logo()
            if gut == False:
                barrinha(df.iloc[0]['GRUPO'])
            else:
                barrinha(df.iloc[0]['SISTEMA PRINCIPAL'])
            addNumeroPaginaFirst()

        def myAddPage(ambiente):
            pdf.add_page()
            titleCabec()
            logo()
            barrinha(ambiente)
            addNumeroPagina(numPag)

        def addNumeroPagina(n):
            # pdf.set_xy(100.0, 272.0)
            # pdf.set_font('Helvetica', '', 10)
            # pdf.set_text_color(0, 0, 0)
            # pdf.cell(txt=str(n), border=0)
            writeText(str(n), 196.0, 292.0, 10, '')

        def writeText(text, x: float, y: float, fontsize: int, bold: str = '', r: int = 0, g: int = 0, b: int = 0):
            pdf.set_xy(x, y)
            pdf.set_font('Helvetica', bold, fontsize)
            pdf.set_text_color(r, g, b)
            pdf.cell(txt=str(text), border=0, w=10,)

        def writeTextParecer(text, x: float, y: float, fontsize: int, bold: str = '', r: int = 0, g: int = 0, b: int = 0):
            pdf.set_xy(x, y)
            pdf.set_font('Helvetica', bold, fontsize)
            pdf.set_text_color(r, g, b)
            pdf.multi_cell(txt=str(text), border=0, h=3.4, w=93, align='J')

        def writeTextParecerRet(text, x: float, y: float, fontsize: int, bold: str = '', r: int = 0, g: int = 0, b: int = 0):
            pdf.set_xy(x, y)
            pdf.set_font('Helvetica', bold, fontsize)
            pdf.set_text_color(r, g, b)
            pdf.multi_cell(txt=str(text), border=0, h=3.4, w=181, align='J')
        
        def writeTextDescricaoGut(text, y: float, fontsize: int, bold: str = '', r: int = 0, g: int = 0, b: int = 0):
            
            pdf.set_xy(103, y)
            pdf.set_font('Helvetica', 'B', fontsize)
            pdf.set_text_color(47, 11, 97)
            pdf.multi_cell(txt='Descrição:', border=0, h=3.4, w=100, align='J')
            pdf.set_text_color(0, 0, 0)
            pdf.set_font('Helvetica', '', fontsize)
            pdf.set_xy(103, y+4.7)
            pdf.multi_cell(txt=str(text), border=0, h=3.8, w=98.4, align='J')

        def writeTextSistemaConstrutivoGut(text: any, y: float, fontsize: int, bold: str = '', r: int = 0, g: int = 0, b: int = 0):
            pdf.set_xy(103, y)
            pdf.set_font('Helvetica', 'B', fontsize)
            pdf.set_text_color(47, 11, 97)
            pdf.multi_cell(txt='Sistema Construtivo:', border=0, h=3.4, w=100, align='J')

            pdf.set_xy(103, y+3.9)
            pdf.set_font('Helvetica', bold, fontsize)
            pdf.set_text_color(r, g, b)
            pdf.multi_cell(txt=str(text), border=0, h=3.4, w=100, align='J')

        def writeTextOrigemGut(text: any, y: float, fontsize: int, bold: str = '', r: int = 0, g: int = 0, b: int = 0):
            pdf.set_xy(103, y)
            pdf.set_font('Helvetica', 'B', fontsize)
            pdf.set_text_color(47, 11, 97)
            pdf.multi_cell(txt='Origem:', border=0, h=3.4, w=100, align='J')

            pdf.set_xy(103, y+3.9)
            pdf.set_font('Helvetica', bold, fontsize)
            pdf.set_text_color(r, g, b)
            pdf.multi_cell(txt=str(text), border=0, h=3.4, w=100, align='J')

        def writeTextCriterioAceitacao(text: any, y: float, fontsize: int, bold: str = '', r: int = 0, g: int = 0, b: int = 0):
            pdf.set_xy(103, y)
            pdf.set_font('Helvetica', 'B', fontsize)
            pdf.set_text_color(47, 11, 97)
            pdf.multi_cell(txt='Critério de aceitação:', border=0, h=3.4, w=100, align='J')

            pdf.set_xy(103, y+3.9)
            pdf.set_font('Helvetica', bold, fontsize)
            pdf.set_text_color(r, g, b)
            pdf.multi_cell(txt=str(text), border=0, h=3.4, w=100, align='J')

        def writeAmbienteGut(text: any, x: float, y: float, fontsize: int, bold: str = '', r: int = 0, g: int = 0, b: int = 0):
            pdf.set_xy(x, y)
            pdf.set_font('Helvetica', 'B', fontsize)
            pdf.set_text_color(47, 11, 97)
            pdf.multi_cell(txt='Ambiente:', border=0, h=3.4, w=100, align='J')

            pdf.set_xy(x+18.2, y)
            pdf.set_font('Helvetica', bold, fontsize)
            pdf.set_text_color(r, g, b)
            pdf.multi_cell(txt=str(text), border=0, h=3.4, w=100, align='J')
        
        '''
            g: Gravidade
            u: Urgência
            t: Tendência
            numfoto: Se é a primeira ou segunda foto
        '''
        def writeGUT(g: int, u: int, t: int, numfoto: int, fontsize: int):
            match g:
                case 1:
                    gravidade = 'Nenhuma'
                case 2:
                    gravidade = 'Baixa'
                case 3:
                    gravidade = 'Média'
                case 4:
                    gravidade = 'Alta'
                case 5:
                    gravidade = 'Alta'
            match u:
                case 1:
                    urgencia = 'Nenhuma'
                case 2:
                    urgencia = 'Baixa'
                case 3:
                    urgencia = 'Media'
                case 4:
                    urgencia = 'Alta'
                case 5:
                    urgencia = 'Alta'
            match t:
                case 1:
                    tendencia = 'Nenhuma'
                case 2:
                    tendencia = 'Baixa'
                case 3:
                    tendencia = 'Média'
                case 4:
                    tendencia = 'Alta'
                case 5:
                    tendencia = 'Alta'
            '''
            r 47, g 11, b 97
            '''
            print(gravidade, urgencia, tendencia, sep='\n')
            if numfoto == 1:
                writeText(f'GRAVIDADE:    {gravidade}', x=103, y=131.6, r=47, g=11, b=97,fontsize=fontsize)
                writeText(f'URGÊNCIA:      {urgencia}', x=103, y=136.3, r=47, g=11, b=97,fontsize=fontsize)
                writeText(f'TENDÊNCIA:    {tendencia}', x=103, y=141, r=47, g=11, b=97,fontsize=fontsize)
            elif numfoto == 2:
                writeText(f'GRAVIDADE:    {gravidade}', x=103, y=150, r=47, g=11, b=97,fontsize=fontsize)
                writeText(f'URGÊNCIA:      {urgencia}', x=103, y=150, r=47, g=11, b=97,fontsize=fontsize)
                writeText(f'TENDÊNCIA:    {tendencia}', x=103, y=150, r=47, g=11, b=97,fontsize=fontsize)


        numPag = int(current_value.get())

        listaDeFotos = df['IMAGEM'].to_list()

        podecriar = True
        listaDeLinhas = []
        primeirox = 8.0
        primeiroy = 50.0
        iters = 1
        iImagem = 1
        numFoto = 1
        iIloc = 0
        
        lenListaFotos = len(listaDeFotos)
        if str(alignment_var.get()) == "Laudo de Inspeção (quadrado)":
            ambienteAtual = df.iloc[iIloc]['GRUPO']
            add_first_page(gut=False)
            for imagem in listaDeFotos:
                if len(df.iloc[iIloc]['DESCRIÇÃO']) > 210:
                    print(df.iloc[iIloc]['DESCRIÇÃO'])
                    numerodarow = df[df['DESCRIÇÃO'] == df.iloc[iIloc]['DESCRIÇÃO']].index[0]
                    listaDeLinhas.append(numerodarow)
                    podecriar = False
                    

                ambienteAtual = df.iloc[iIloc]['GRUPO']
                numFotoreal = f'%3s' % numFoto
                # print("imagem: ", imagem, "iImagem: ", iImagem, "iIloc:", iIloc, "numFoto:",
                #     numFoto, "iters: ", iters, "numPag: ", numPag-1, "ambienteAtual: ", ambienteAtual, "len: ", lenListaFotos)

                if iImagem == 1:
                    if iters > 1:
                        if ambienteAtual == df.iloc[iIloc-1]['GRUPO']:
                            ambienteAtual = df.iloc[iIloc]['GRUPO']
                            pdf.image(f'{photosPath}/{imagem}.jpg', link='', type='',
                                    w=93, h=93, x=primeirox, y=primeiroy)
                            pdf.set_draw_color(0,0,0)
                            pdf.rect(primeirox, primeiroy, w=93, h=93)
                            writeText(f'Foto {numFotoreal.replace(" ", "0")}', primeirox +
                                    40, primeiroy+97, 10, '')
                            writeTextParecer(df.iloc[iIloc]['DESCRIÇÃO'],
                                    primeirox, primeiroy+100, 10)

                        else:
                            pdf.image(f'{photosPath}/{imagem}.jpg', link='', type='',
                                    w=93, h=93, x=primeirox, y=primeiroy)
                            pdf.set_draw_color(0,0,0)
                            pdf.rect(primeirox, primeiroy, w=93, h=93)
                            writeText(f'Foto {numFotoreal.replace(" ", "0")}', primeirox +
                                    40, primeiroy+97, 10, '')
                            writeTextParecer(df.iloc[iIloc]['DESCRIÇÃO'],
                                    primeirox, primeiroy+100, 10)
                            iters += 1
                            iImagem = 1

                    else:
                        pdf.image(f'{photosPath}/{imagem}.jpg', link='', type='',
                                w=93, h=93, x=primeirox, y=primeiroy)
                        pdf.set_draw_color(0,0,0)
                        pdf.rect(primeirox, primeiroy, w=93, h=93)
                        writeText(f'Foto {numFotoreal.replace(" ", "0")}', primeirox +
                                40, primeiroy+97, 10, '')
                        writeTextParecer(df.iloc[iIloc]['DESCRIÇÃO'],
                                primeirox, primeiroy+100, 10)
                    numFoto += 1
                    iImagem += 1
                    iIloc += 1

                elif iImagem == 2:
                    if ambienteAtual == df.iloc[iIloc-1]['GRUPO']:
                        ambienteAtual = df.iloc[iIloc]['GRUPO']
                        pdf.image(f'{photosPath}/{imagem}.jpg', link='', type='',
                                w=93, h=93, x=primeirox+100, y=primeiroy)
                        pdf.set_draw_color(0,0,0)
                        pdf.rect(primeirox+100, primeiroy, w=93, h=93)
                        writeText(f'Foto {numFotoreal.replace(" ", "0")}', primeirox +
                                140, primeiroy+97, 10, '')
                        writeTextParecer(df.iloc[iIloc]['DESCRIÇÃO'],
                                primeirox+100, primeiroy+100, 10)

                    else:

                        myAddPage(ambienteAtual)
                        pdf.image(f'{photosPath}/{imagem}.jpg', link='', type='',
                                w=93, h=93, x=primeirox, y=primeiroy)
                        pdf.set_draw_color(0,0,0)
                        pdf.rect(primeirox, primeiroy, w=93, h=93)
                        writeText(f'Foto {numFotoreal.replace(" ", "0")}', primeirox +
                                40, primeiroy+97, 10, '')
                        writeTextParecer(df.iloc[iIloc]['DESCRIÇÃO'],
                                primeirox, primeiroy+100, 10)
                        iters += 1
                        iImagem = 1
                        numPag += 1

                    numFoto += 1
                    iImagem += 1
                    iIloc += 1

                elif iImagem == 3:
                    if ambienteAtual == df.iloc[iIloc-1]['GRUPO']:
                        ambienteAtual = df.iloc[iIloc]['GRUPO']
                        pdf.image(f'{photosPath}/{imagem}.jpg', link='', type='',
                                w=93, h=93, x=primeirox, y=primeiroy+120)
                        pdf.set_draw_color(0,0,0)
                        pdf.rect(primeirox, primeiroy+120, w=93, h=93)
                        writeText(f'Foto {numFotoreal.replace(" ", "0")}', primeirox +
                                40, primeiroy+217, 10, '')
                        writeTextParecer(df.iloc[iIloc]['DESCRIÇÃO'],
                                primeirox, primeiroy+220, 10)
                    else:

                        myAddPage(ambienteAtual)
                        pdf.image(f'{photosPath}/{imagem}.jpg', link='', type='',
                                w=93, h=93, x=primeirox, y=primeiroy)
                        pdf.set_draw_color(0,0,0)
                        pdf.rect(primeirox, primeiroy, w=93, h=93)
                        writeText(f'Foto {numFotoreal.replace(" ", "0")}', primeirox +
                                40, primeiroy+97, 10, '')
                        writeTextParecer(df.iloc[iIloc]['DESCRIÇÃO'],
                                primeirox, primeiroy+100, 10)
                        iters += 1
                        iImagem = 1
                        numPag += 1

                    numFoto += 1
                    iImagem += 1
                    iIloc += 1
                elif iImagem == 4:
                    if ambienteAtual == df.iloc[iIloc-1]['GRUPO']:
                        ambienteAtual = df.iloc[iIloc]['GRUPO']
                        pdf.image(f'{photosPath}/{imagem}.jpg', link='', type='',
                                w=93, h=93, x=primeirox+100, y=primeiroy+120)
                        pdf.set_draw_color(0,0,0)
                        pdf.rect(primeirox+100, primeiroy+120, w=93, h=93)
                        writeText(f'Foto {numFotoreal.replace(" ", "0")}', primeirox +
                                140, primeiroy+217, 10, '')
                        writeTextParecer(df.iloc[iIloc]['DESCRIÇÃO'],
                                primeirox+100, primeiroy+220, 10)
                        if not numFoto == lenListaFotos:
                            myAddPage(ambienteAtual)
                        iImagem = 1

                    else:

                        myAddPage(ambienteAtual)
                        pdf.image(f'{photosPath}/{imagem}.jpg', link='', type='',
                                w=93, h=93, x=primeirox, y=primeiroy)
                        pdf.set_draw_color(0,0,0)
                        pdf.rect(primeirox, primeiroy, w=93, h=93)
                        writeText(f'Foto {numFotoreal.replace(" ", "0")}', primeirox +
                                40, primeiroy+97, 10, '')
                        writeTextParecer(df.iloc[iIloc]['DESCRIÇÃO'],
                                primeirox, primeiroy+100, 10)
                        iImagem = 2

                    iIloc += 1
                    numFoto += 1
                    numPag += 1

                    iters += 1
        
        elif str(alignment_var.get()) == "Laudo de Inspeção (retangular)":
            ambienteAtual = df.iloc[iIloc]['GRUPO']
            add_first_page(gut=False)            
            for imagem in listaDeFotos:
                if len(df.iloc[iIloc]['DESCRIÇÃO']) > 400:
                    print(df.iloc[iIloc]['DESCRIÇÃO'])
                    numerodarow = df[df['DESCRIÇÃO'] == df.iloc[iIloc]['DESCRIÇÃO']].index[0]
                    listaDeLinhas.append(numerodarow)
                    podecriar = False
                
                ambienteAtual = df.iloc[iIloc]['GRUPO']
                numFotoreal = f'%3s' % numFoto

                if iImagem == 1:
                    if iters > 1:
                        if ambienteAtual == df.iloc[iIloc-1]['GRUPO']:
                            ambienteAtual = df.iloc[iIloc]['GRUPO']
                            pdf.image(f'{photosPath}/{imagem}.jpg', link='', type='',
                                    w=181, h=87, x=primeirox+6, y=primeiroy)
                            pdf.set_draw_color(0,0,0)
                            pdf.rect(primeirox+6, primeiroy, w=181, h=87)
                            writeText(f'Foto {numFotoreal.replace(" ", "0")}', (pdf.w/2)-9, primeiroy+90.2, 10, 'B')
                            writeTextParecerRet(df.iloc[iIloc]['DESCRIÇÃO'],
                                    primeirox+6, primeiroy+93.2, 10)

                        else:
                            pdf.image(f'{photosPath}/{imagem}.jpg', link='', type='',
                                    w=181, h=87, x=primeirox+6, y=primeiroy)
                            pdf.set_draw_color(0,0,0)
                            pdf.rect(primeirox+6, primeiroy, w=181, h=87)
                            writeText(f'Foto {numFotoreal.replace(" ", "0")}', (pdf.w/2)-9, primeiroy+90.2, 10, 'B')
                            writeTextParecerRet(df.iloc[iIloc]['DESCRIÇÃO'],
                                    primeirox+6, primeiroy+93.2, 10)
                            iters += 1
                            iImagem = 1

                    else:
                        pdf.image(f'{photosPath}/{imagem}.jpg', link='', type='',
                                w=181, h=87, x=primeirox+6, y=primeiroy)
                        pdf.set_draw_color(0,0,0)
                        pdf.rect(primeirox+6, primeiroy, w=181, h=87)
                        writeText(f'Foto {numFotoreal.replace(" ", "0")}', (pdf.w/2)-9, primeiroy+90.2, 10, 'B')
                        writeTextParecerRet(df.iloc[iIloc]['DESCRIÇÃO'],
                                primeirox+6, primeiroy+93.2, 10)
                    numFoto += 1
                    iImagem += 1
                    iIloc += 1

                elif iImagem == 2:
                    if ambienteAtual == df.iloc[iIloc-1]['GRUPO']:
                        ambienteAtual = df.iloc[iIloc]['GRUPO']
                        pdf.image(f'{photosPath}/{imagem}.jpg', link='', type='',
                                w=181, h=87, x=primeirox+6, y=primeiroy+120)
                        pdf.set_draw_color(0,0,0)
                        pdf.rect(primeirox+6, primeiroy+120, w=181, h=87)
                        writeText(f'Foto {numFotoreal.replace(" ", "0")}', (pdf.w/2)-10, primeiroy+210.2, 10, 'B')
                        writeTextParecerRet(df.iloc[iIloc]['DESCRIÇÃO'],
                                primeirox+6, primeiroy+213.2, 10)
                        if not numFoto == lenListaFotos:
                            myAddPage(ambienteAtual)
                        iImagem = 1
                        iters += 1

                    else:

                        myAddPage(ambienteAtual)
                        pdf.image(f'{photosPath}/{imagem}.jpg', link='', type='',
                                w=181, h=87, x=primeirox+6, y=primeiroy)
                        pdf.set_draw_color(0,0,0)
                        pdf.rect(primeirox+6, primeiroy, w=181, h=87)
                        writeText(f'Foto {numFotoreal.replace(" ", "0")}', (pdf.w/2)-10, primeiroy+90.2, 10, 'B')
                        writeTextParecerRet(df.iloc[iIloc]['DESCRIÇÃO'],
                                primeirox+6, primeiroy+93.2, 10)
                        iters += 1
                        iImagem = 1

                    numFoto += 1
                    iImagem = 1
                    iIloc += 1
                    iters += 1
                    numPag += 1

        elif str(alignment_var.get()) == "Laudo de Inspeção com GUT":
            ambienteAtual = df.iloc[iIloc]['AMBIENTE']

            add_first_page(gut=True)
            for imagem in listaDeFotos:
                # if len(df.iloc[iIloc]['DESCRIÇÃO']) > 400:
                #     print(df.iloc[iIloc]['DESCRIÇÃO'])
                #     numerodarow = df[df['DESCRIÇÃO'] == df.iloc[iIloc]['DESCRIÇÃO']].index[0]
                #     listaDeLinhas.append(numerodarow)
                #     podecriar = False
                
                ambienteAtual = df.iloc[iIloc]['AMBIENTE']
                numFotoreal = f'%3s' % numFoto
                sistema = df.iloc[0]['SISTEMA PRINCIPAL']
                sisConstru = df.iloc[iIloc]['SISTEMA CONSTRUTIVO']
                g, u, t = df.iloc[iIloc]['GRAVIDADE'], df.iloc[iIloc]['URGÊNCIA'], df.iloc[iIloc]['TENDÊNCIA']
                descricao = df.iloc[iIloc]['DESCRIÇÃO']
                origem = df.iloc[iIloc]['ORIGEM']
                criterio = df.iloc[iIloc]['CRITÉRIO DE ACEITAÇÃO']
                # print(g, u ,t, sep='\n')
                if len(descricao) > 230:
                    podecriar = False
                    showinfo("Erro", "Alguma descrição está muito longa.")
                if iImagem == 1:
                    if iters > 1:
                        if ambienteAtual == df.iloc[iIloc-1]['AMBIENTE']:
                            ambienteAtual = df.iloc[iIloc]['AMBIENTE']
                            pdf.image(f'{photosPath}/{imagem}.jpg', link='', type='',
                                    w=93, h=93, x=primeirox, y=primeiroy)
                            pdf.set_draw_color(0,0,0)
                            pdf.rect(primeirox, primeiroy, w=93, h=93)
                            writeText(f'Foto {numFotoreal.replace(" ", "0")}', primeirox, primeiroy+97, 10, 'B')
                            writeTextSistemaConstrutivoGut(sisConstru, y=primeiroy,fontsize=10)
                            writeGUT(g, u, t, 1,fontsize=10)
                            writeTextDescricaoGut(descricao, y=62,fontsize=10)
                            writeTextOrigemGut(origem, y=85.3, fontsize=10)
                            writeTextCriterioAceitacao(criterio, y=88.3, fontsize=10)
                            writeAmbienteGut(ambienteAtual, primeirox, primeiroy+101, 10)
                            iImagem += 1
                        else:
                            pdf.image(f'{photosPath}/{imagem}.jpg', link='', type='',
                                    w=93, h=93, x=primeirox, y=primeiroy)
                            pdf.set_draw_color(0,0,0)
                            pdf.rect(primeirox, primeiroy, w=93, h=93)
                            writeText(f'Foto {numFotoreal.replace(" ", "0")}', primeirox, primeiroy+97, 10, 'B')
                            writeTextSistemaConstrutivoGut(sisConstru, y=primeiroy, fontsize=10)
                            writeGUT(g, u, t, 1,fontsize=10)
                            writeTextDescricaoGut(descricao, y=62,fontsize=10)
                            writeTextOrigemGut(origem, y=85.3, fontsize=10)
                            writeTextCriterioAceitacao(criterio, y=88.3, fontsize=10)
                            writeAmbienteGut(ambienteAtual, primeirox, primeiroy+101, 10)
                            iters += 1
                            iImagem = 1

                    else:
                        pdf.image(f'{photosPath}/{imagem}.jpg', link='', type='',
                                w=93, h=93, x=primeirox, y=primeiroy)
                        pdf.set_draw_color(0,0,0)
                        pdf.rect(primeirox, primeiroy, w=93, h=93)
                        writeText(f'Foto {numFotoreal.replace(" ", "0")}', primeirox, primeiroy+97, 10, 'B')
                        writeTextSistemaConstrutivoGut(sisConstru, y=primeiroy,fontsize=10)
                        print(g, u ,t)
                        writeGUT(g, u, t, 1,fontsize=10)
                        writeTextDescricaoGut(descricao, y=62,fontsize=10)
                        writeTextOrigemGut(origem, y=85.5, fontsize=10)
                        writeTextCriterioAceitacao(criterio, y=97.2, fontsize=10)
                        writeAmbienteGut(ambienteAtual, primeirox, primeiroy+101, 10)
                        iImagem += 1
                    
                    numFoto += 1
                    
                    iIloc += 1

                elif iImagem == 2:
                    if ambienteAtual == df.iloc[iIloc-1]['AMBIENTE']:
                    #     ambienteAtual = df.iloc[iIloc]['AMBIENTE']
                    #     pdf.image(f'{photosPath}/{imagem}.jpg', link='', type='',
                    #             w=93, h=93, x=primeirox, y=primeiroy+120)
                    #     pdf.set_draw_color(0,0,0)
                    #     pdf.rect(primeirox, primeiroy+120, w=93, h=93)
                    #     writeText(f'Foto {numFotoreal.replace(" ", "0")}', primeirox, primeiroy+210.2, 10, 'B')
                    #     writeTextParecerRet(df.iloc[iIloc]['DESCRIÇÃO'],
                    #             primeirox, primeiroy+213.2, 10)
                    #     if not numFoto == lenListaFotos:
                    #         myAddPage(sistema)
                    #     iImagem = 1
                    #     iters += 1

                    # else:

                    #     myAddPage(sistema)
                    #     pdf.image(f'{photosPath}/{imagem}.jpg', link='', type='',
                    #             w=93, h=93, x=primeirox, y=primeiroy)
                    #     pdf.set_draw_color(0,0,0)
                    #     pdf.rect(primeirox, primeiroy, w=93, h=93)
                    #     writeText(f'Foto {numFotoreal.replace(" ", "0")}', primeirox, primeiroy+90.2, 10, 'B')
                    #     writeTextParecerRet(df.iloc[iIloc]['DESCRIÇÃO'],
                    #             primeirox, primeiroy+93.2, 10)
                    #     iters += 1
                    #     iImagem = 1

                    numFoto += 1
                    iImagem = 1
                    iIloc += 1
                    iters += 1
                    numPag += 1
        else:
            podecriar = False
            showinfo("Erro", "Selecione algum tipo de documento")

        if len(listaDeLinhas) != 0:
            showinfo("Erro", f"A(s) linha(s) {listaDeLinhas} tem mais de 210 caracteres")
            
        elif len(entryNomeEmp.get()) != 0:
            if podecriar:
                pdf.output(f'{outPath}/ERF (Laudo de Inspeção).pdf')
                showinfo("Sucesso", "Relatório gerado com sucesso")
                


    elif len(planilhaPath) == 0 and len(outPath) != 0 and len(photosPath) != 0:
        showinfo("Erro", "Seleciona a planilha")
    elif len(planilhaPath) == 0 and len(outPath) == 0 and len(photosPath) != 0:
        showinfo("Erro", "Seleciona a planilha e a pasta de saída")
    elif len(planilhaPath) == 0 and len(outPath) == 0 and len(photosPath) == 0:
        showinfo("Erro", "Seleciona a planilha, pasta de saída e pasta de fotos")
    elif len(planilhaPath) != 0 and len(outPath) == 0 and len(photosPath) != 0:
        showinfo("Erro", "Seleciona pasta de saída")
    elif len(planilhaPath) != 0 and len(outPath) != 0 and len(photosPath) == 0:
        showinfo("Erro", "Seleciona a pasta de fotos")
    elif len(planilhaPath) != 0 and len(outPath) == 0 and len(photosPath) == 0:
        showinfo("Erro", "Seleciona a pasta de saída e pasta de fotos")


# Cria o botão de selecionar a saida
if "nt" == os.name:
    photoOut = PhotoImage(file=r"assets\pasta.png")
else:
    photoOut = PhotoImage(file=r"./assets/pasta.png")

outDirBtn = ttk.Button(root, text="Selecionar pasta de saída",
                       image=photoOut, compound='left', command=lambda: ask('out'))
outDirBtn.place(x=25, y=150)


labelOut = ttk.Label(root, text=outDir, background='#FFFFFF')
labelOut.place(x=195, y=154, width=840)

# Cria o botão de selecionar a planilha
if "nt" == os.name:
    photoPlan = PhotoImage(file=r"assets\planilha.png")
else:
    photoPlan = PhotoImage(file=r"./assets/planilha.png")
    
excelPlanilhaBtn = ttk.Button(root, text="Selecionar planilha Excel ",
                              image=photoPlan, compound='left', command=lambda: ask('plan'))
excelPlanilhaBtn.place(x=25, y=190)

labelExc = ttk.Label(root, text=excelPlan, background='#FFFFFF')
labelExc.place(x=195, y=194, width=840)

# Cria o botão de selecionar as fotos
if "nt" == os.name:
    photoDir = PhotoImage(file=r"assets\foto.png")
else:
        photoDir = PhotoImage(file=r"./assets/foto.png")
photosDirBtn = ttk.Button(root, text="Selecionar pasta de fotos",
                          image=photoDir, compound='left', command=lambda: ask('pho'))
photosDirBtn.place(x=25, y=230)

labelPho = ttk.Label(root, text=photosdirask, background='#FFFFFF')
labelPho.place(x=195, y=234, width=840)

# Escolha de tipo de doc
lf = ttk.LabelFrame(root, text='Tipo de documento')
lf.grid(column=0, row=0, padx=25, pady=20)
alignment_var = tk.StringVar()
alignments = ('Laudo de Inspeção (quadrado)', 'Laudo de Inspeção (retangular)', 'Laudo de Inspeção com GUT')

grid_column = 0
for alignment in alignments:
    # create a radio button
    radio = ttk.Radiobutton(
        lf, text=alignment, value=alignment, variable=alignment_var)
    if alignment != 'Laudo de Inspeção com GUT':
        radio.grid(column=grid_column, row=0, ipadx=10, ipady=10)
    else:
        radio.grid(column=0, row=1, ipadx=10, ipady=10)

    # grid column
    grid_column += 1

lf.place(x=270, y=30)
# Adiciona os endereços das fotos


labelPagIni = ttk.Label(root, text="Página inicial:")
labelPagIni.place(x=25, y=120)
current_value = tk.IntVar(value=1)
spin_box = ttk.Spinbox(
    root,
    from_=1,
    to=999999,
    textvariable=current_value,
    wrap=True)

spin_box.place(y=120, x=120, width=65)

entryNomeEmp = ttk.Entry(root)
entryNomeEmp.place(x=800, y=50, width=235)

labelNomeEmp = ttk.Label(root, text="Nome:")
labelNomeEmp.place(x=720, y=50)

entryEnderEmp = ttk.Entry(root)
entryEnderEmp.place(x=800, y=80, width=235, height=55)

labelEnderEmp = ttk.Label(root, text="Endereço:")
labelEnderEmp.place(x=720, y=90)

labelCabec = ttk.Label(
    root, text='Cabeçalho (empreendimento)', foreground='#fc0328')
labelCabec.place(x=720, y=25)

img = Image.open('assets/CDE.png')
img = img.resize((100, 100), Image.Resampling.LANCZOS)
img = ImageTk.PhotoImage(img)
panelLogo = ttk.Label(root, image=img)
panelLogo.place(x=50, y=10, width=100, height=100, )

root.mainloop()
