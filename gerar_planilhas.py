import os
import pandas as pd
import tkinter as tk
from time import sleep
from tkinter import filedialog, messagebox
from reportlab.lib.pagesizes import landscape, A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Table, TableStyle, Image, Spacer, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

estilos = getSampleStyleSheet()

# Estilo personalizado para quebrar padronizar a quebra de texto grande
estilo_quebra_texto = ParagraphStyle(name='EstiloQuebraTexto', wordWrap='CJK', leading=15, spaceAfter=10, fontSize=8, maxLineHeight=36, alignment=1)

nome_secretario = 'Alexandro Olegario dos Santos'
digitador = 'Daniel Cordeiro da Costa'

def carregar_arquivo():
    nome_arquivo = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
    if nome_arquivo:
        messagebox.showinfo("Aviso", "Aguarde enquanto as planilhas estão sendo geradas...")
        try:
            nome_pasta = os.path.splitext(nome_arquivo)[0]
            dados = pd.read_csv(f'{nome_arquivo}').fillna('')
            if not os.path.exists(f'{nome_pasta}'):
                os.makedirs(f'{nome_pasta}')
            for (saida, carro), grupo in dados.groupby(['Saída', 'Carro']):
                grupo = grupo.reset_index(drop=True)
                for i in range(0, len(grupo), 4):
                    subset = grupo.iloc[i:i+4]
                    nome_arquivo = f'{nome_pasta}/{saida}_{carro}_{i//4+1}.pdf'
                    criar_pdf(subset, nome_arquivo)
        except Exception as e:
            messagebox.showerror("Erro", str(e))
        else:
            messagebox.showinfo("Sucesso", "Planilhas geradas com sucesso!")
            sleep(2)
            root.destroy()

def adicionar_cabecalho(story):
    # Cabeçalho do PDF com as logos da prefeitura e da secretaria
    cabecalho = [
        [Image('logo-pma.png', width=180, height=50), '', '', Image('logo-transporte.png', width=180, height=50)],
        ['Prefeitura de Araçagi', '', '', 'Secretarias dos Transportes']
    ]

    tabela_cabecalho = Table(cabecalho, colWidths=[20, 200, 200, 20])
    tabela_cabecalho.setStyle(TableStyle([
        ('SPAN', (0, 0), (0, 1)),  # Mescla a célula do logotipo
        ('SPAN', (3, 0), (3, 1)),  # Mescla a célula do logotipo
        ('ALIGN', (0, 0), (0, 0), 'CENTER'),
        ('ALIGN', (3, 0), (3, 0), 'CENTER'),
        ('ALIGN', (1, 0), (2, 0), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('FONTSIZE', (1, 0), (2, 0), 20),
        ('BOTTOMPADDING', (1, 0), (0, 0), 0),
        ('TOPPADDING', (0, 0), (-10, -10), -10),  # Remove o espaçamento superior
    ]))

    story.append(tabela_cabecalho)
    story.append(Spacer(1, 10)) 
    
    # Titutlo informarivo do PDF
def adicionar_titulo_principal(story):
    titulo_principal = "DECLARAMOS PARA OS DEVIDOS FINS DE DIREITO JUNTO AOS ÓRGÃOS DE FISCALIZAÇÃO, QUE FOI DISPONIBILIZADO UMA VIAGEM NO VEÍCULO DE PLACA: _______ SAINDO DE ARAÇAGI/PB PARA TRATAMENTO DE SAÚDE NA CIDADE DE JOÃO PESSOA."
    story.append(Paragraph(titulo_principal, estilos['Normal']))
    story.append(Spacer(0, 7))

    # Funcao para criar o PDF
def criar_pdf(grupo, nome_arquivo):
    doc = SimpleDocTemplate(nome_arquivo, pagesize=landscape(A4))
    story = []
    adicionar_cabecalho(story)
    adicionar_titulo_principal(story)
    dados_tabela = [['Data', 'Nome do Paciente', 'CPF', 'Telefone', 'Saída', 'Marcado', 'Nome do Hospital', 'Assinatura do Paciente/Responsável']]
    for index, row in grupo.iterrows():
        dados_tabela.append([
            Paragraph(row['Data'], estilo_quebra_texto),
            Paragraph(row['Nome do Paciente'], estilo_quebra_texto),
            Paragraph(row['CPF'], estilo_quebra_texto),
            Paragraph(row['Telefone'], estilo_quebra_texto),
            Paragraph(row['Saída'], estilo_quebra_texto),
            Paragraph(row['Marcado'], estilo_quebra_texto),
            Paragraph(row['Nome do Hospital'], estilo_quebra_texto),
            ' '
        ])
    
    # Verificando se a quantidade de linhas é menor que 5, para completar com linhas vazias
    while len(dados_tabela) < 5:
        dados_tabela.append(['', '', '', '', '', '', '', ' '])
    
    # Criando a tabela com os dados
    tabela = Table(dados_tabela, colWidths=[50, 90, 70, 70, 40, 40, 170, 220])
    tabela.setStyle(TableStyle([
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
        ('FONTSIZE', (0, 0), (-1, -1), 10),
        ('FONTSIZE', (0, 0), (-1, 0), 10),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.black),
        ('LEFTPADDING', (0, 0), (-1, -1), 5),
        ('RIGHTPADDING', (0, 0), (-1, -1), 5),
        ('TOPPADDING', (0, 0), (-1, -1), 5),  
        ('BOTTOMPADDING', (0, 0), (-1, -1), 10),
    ]))
    story.append(tabela)
    story.append(Spacer(1, 10))
    
    # Rodapé final do PDF
    rodape = [
        [f'Digitado por: {digitador}'],
        [''],
        [''],
        ['__________________________________________'],
        [f'{nome_secretario}'],
        [' Secretário de Transportes de Araçagi/PB']
    ]
    
    tabela_rodape = Table(rodape, colWidths=[500])
    tabela_rodape.setStyle(TableStyle([
        ('ALIGN', (0, 0), (0, 0), 'CENTER'),
        ('ALIGN', (0, 1), (0, 1), 'CENTER'),
        ('ALIGN', (0, 2), (0, 2), 'CENTER'),
        ('ALIGN', (0, 3), (0, 3), 'CENTER'),
        ('ALIGN', (0, 4), (0, 4), 'CENTER'),
        ('ALIGN', (0, 5), (0, 5), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('FONTNAME', (0, 0), (0, 4), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (0, 2), 10),
        ('BOTTOMPADDING', (0, 0), (0, 2), 0)
    ]))
    story.append(tabela_rodape)
    story.append(PageBreak())
    
    doc.build(story)

root = tk.Tk()
root.geometry("300x100")
root.resizable(False, False)
root.title("Gerador de relatórios PDF")
root.eval('tk::PlaceWindow . center')

frame = tk.Frame(root)
frame.pack(padx=10, pady=10)

botao_carregar = tk.Button(frame, text="Carregar arquivo CSV", command=carregar_arquivo)
botao_carregar.pack(fill=tk.X)

root.mainloop()