import aspose.words as aw
import glob
import os

#Para o uso do conversor, envie arquivos rtf para o diretorio da pasta do projeto.

ListaArquivos = []

#Informe o diretorio de arquivos abaixc.
diretorio_rtf = r"C:\Users\gustavoledo\PycharmProjects\pythonProject"

for arquivo in glob.glob(os.path.join(diretorio_rtf, '*.rtf')):
    ListaArquivos.append(arquivo)

def ConvertePdf(EntradaRtf, SaidaPdf):
    try:
        doc = aw.Document(EntradaRtf)
        doc.save(SaidaPdf, aw.SaveFormat.PDF)
        print(f"Conversão PDF concluída! Arquivo gerado em {SaidaPdf}")
    except Exception as e:
        print(f"Erro: {str(e)}")

for arquivo_rtf in ListaArquivos:
    nome_base = os.path.splitext(os.path.basename(arquivo_rtf))[0]
    arquivo_saida_pdf = os.path.join(diretorio_rtf, nome_base + '.pdf')
    ConvertePdf(arquivo_rtf, arquivo_saida_pdf)


def ConverteTxt(EntradaRtf, SaidaTxt):
    try:
        doc = aw.Document(EntradaRtf)
        doc.save(SaidaTxt, aw.SaveFormat.TEXT)
        print(f"Conversão TEXT concluída! Arquivo gerado em {SaidaTxt}")
    except Exception as e:
        print(f"Erro: {str(e)}")

for arquivo_rtf in ListaArquivos:
    nome_base = os.path.splitext(os.path.basename(arquivo_rtf))[0]
    arquivo_saida_txt = os.path.join(diretorio_rtf, nome_base + '.txt')
    ConverteTxt(arquivo_rtf, arquivo_saida_txt)


def ConverteOdt(EntradaRtf, SaidaOdt):
    try:
        doc = aw.Document(EntradaRtf)
        doc.save(SaidaOdt, aw.SaveFormat.ODT)
        print(f"Conversão ODT concluída! Arquivo gerado em {SaidaOdt}")
    except Exception as e:
        print(f"Erro: {str(e)}")

for arquivo_rtf in ListaArquivos:
    nome_base = os.path.splitext(os.path.basename(arquivo_rtf))[0]
    arquivo_saida_odt = os.path.join(diretorio_rtf, nome_base + '.odt')
    ConverteOdt(arquivo_rtf, arquivo_saida_odt)