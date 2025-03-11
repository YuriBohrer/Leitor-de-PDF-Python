from pandas as pd # type: ignore
from pypdf import PdfReader # type: ignore
import time

#Funções de extração e tratamento do texto extraido do pdf
solicitacao = 1

while solicitacao <= 7:
    #importação dos Dados da tabela
    mes_compet = time.strftime("%m.%Y", time.localtime())
    mes_compet_cont = time.strftime("%Y-%m", time.localtime())

    tabela = pd.read_excel(f'C:\\Users\\YURI\\Dropbox\\ROBOS\\ROBÔ BLIG\\situação fiscal\\{mes_compet}\\Planilha\\geral.xlsx', dtype= "str")

    cnptabela = []
    #rztabela = []
    tamtabela = len(tabela["CNPJ"])
    #tamtabela = len(tabela["RZ"])

    for i in tabela["CNPJ"]:
        cnptabela.append(i)
        #for p in tabela["RZ"]:
            #rztabela.append(p)


    #display(tabela)

    #Dicionario com as listas de empresas por parcelamento

    rel_geral = {}

    for emp in range(0, tamtabela):
        rel_geral[str(cnptabela[emp])]= []
        #rel_geral[f'{cnptabela[emp]};{rztabela[emp]}'] = []

    
    #Funções de extração e tratamento do texto extraido do pdf
    def consulta(n):
        if n == 1:
            par_con = ["SIMPLES NACIONAL - EM PARCELAMENTO",
                    "SIMPLES NACIONAL - PERT - EM PARCELAMENTO",
                    "SIMPLES NACIONAL - RELP - EM PARCELAMENTO",
                    "PARCELAMENTO SIMPLIFICADO",
                    "(SISPAR)",
                    "PENDÊNCIA - INSCRIÇÃO (SIDA)"]
            return par_con
        elif n == 2:
            fis_con = ["OMISSÃO DE PGDAS-D",
                    "PGDAS-D - MULTA",
                    "OMISSÃO DE ECF",
                    "OMISSÃO DE EFD-CONTRIB",
                    "OMISSÃO DE DCTF",
                    "DCTF - MULTA ATR",
                    "OMISSÃO DE DASN SIMEI"]
            return fis_con
        elif n == 3:
            pes_con = ["OMISSÃO DE DCTFWEB",
                    "OMISSÃO DE GFIP"]
            return pes_con
        elif n == 4:
            cont_con = ["OMISSÃO DE DEFIS",
                        "OMISSÃO DE DIRF"]
            return cont_con
        else:
            ger_con = ["OMISSÃO DE PGDAS-D",
                    "PGDAS-D - MULTA",
                    "OMISSÃO DE ECF",
                    "OMISSÃO DE EFD-CONTRIB",
                    "OMISSÃO DE DCTF",
                    "DCTF - MULTA ATR",
                    "OMISSÃO DE DASN SIMEI",
                    "OMISSÃO DE DCTFWEB",
                    "OMISSÃO DE GFIP",
                    "OMISSÃO DE DEFIS",
                    "OMISSÃO DE DIRF"]
            return ger_con


    def extracao_texto(numtb, pagpdf):
        
        for cn in consulta(solicitacao):
            consult_inter = cn
            
            teste = ""
            taman = 0
            recorrencia = 0

            for i in pagpdf:
                if i == consult_inter[taman]:
                    teste += i
                    taman += 1

                    if taman >= len(consult_inter):
                        recorrencia += 1
                        teste = ""
                        taman = 0
                else:
                    taman = 0
                    teste = ""

            if recorrencia >= 1:
                rel_geral[str(cnptabela[numtb])].append(consult_inter)
                #rel_geral[f'{cnptabela[numtb]};{rztabela[numtb]}'].append(consult_inter)
            else:
                rel_geral[str(cnptabela[numtb])].append("NÃO")
                #rel_geral[f'{cnptabela[numtb]};{rztabela[numtb]}'].append("NÃO")


    def tratamento_dados(number):

        if number == 2:#CRIAÇÃO DE PENDENCIA FISCAL 
            df = pd.DataFrame.from_dict(rel_geral, orient='index',columns= ["PGDAS-D - DECLARAÇÃO","PGDAS-D - MULTA","ECF","EFD-CONTRIB","DCTF - DECLARAÇÃO","DCTF - MULTA","DASN SIMEI"])
            df.to_excel(f"C:\\Users\\YURI\\Dropbox\\ROBOS\\ROBÔ BLIG\\situação fiscal\\{mes_compet}\\Planilha\\PENDENCIA FISCAL.xlsx")
        elif number == 4:#CRIAÇÃO DE PLANILHA DE PENDENCIA CONTABIL       
            df = pd.DataFrame.from_dict(rel_geral, orient='index',columns= ["DIRF","DEFIS"])
            df.to_excel(f"C:\\Users\\YURI\\Dropbox\\ROBOS\\ROBÔ BLIG\\situação fiscal\\{mes_compet}\\Planilha\\PENDENCIA CONTABIL.xlsx")
        elif number == 3:#CRIAÇÃO DE PLANILHA DE PENDENCIA PESSOAL
            df = pd.DataFrame.from_dict(rel_geral, orient='index',columns= ["DCTFWEB","GFIP"])
            df.to_excel(f"C:\\Users\\YURI\\Dropbox\\ROBOS\\ROBÔ BLIG\\situação fiscal\\{mes_compet}\\Planilha\\PENDENCIA PESSOAL.xlsx")
        elif number == 1:#CRIAÇÃO DE PLANILHA DE PARCELAMENTO
            df = pd.DataFrame.from_dict(rel_geral, orient='index',columns= ["SIMPLES NACIONAL","PERT","RELP","SIMPLIFICADO","(SISPAR)","PENDENCIA SIDA"])
            df.to_excel(f"C:\\Users\\YURI\\Dropbox\\ROBOS\\ROBÔ BLIG\\situação fiscal\\{mes_compet}\\Planilha\\PLANILHA MENSAL PARCELAMENTO.xlsx")
        else:#CRIAÇÃO DE PLANILHA DE PENDENCIA GERAL
            df = pd.DataFrame.from_dict(rel_geral, orient='index',columns= ["PGDAS-D - DECLARAÇÃO","PGDAS-D - MULTA","ECF","EFD-CONTRIB","DCTF - DECLARAÇÃO","DCTF - MULTA","DASN SIMEI","DCTFWEB","GFIP","DIRF","DEFIS"])
            df.to_excel(f"C:\\Users\\YURI\\Dropbox\\ROBOS\\ROBÔ BLIG\\situação fiscal\\{mes_compet}\\Planilha\\PLANILHA GERAL DE PENDENCIAS.xlsx")


    for tb in range(0, tamtabela):
        dia = "05"
        indice_table = tb
        #leitura_pdf = PdfReader(f'C:\\Users\\YURI\\Dropbox\\ROBOS\\ROBÔ BLIG\\situação fiscal\\07.2024\\Geral\\situacao_fiscal--{cnptabela[tb]}-{rztabela[tb]}')
        
        while True:
            try:
                leitura_pdf = PdfReader(f'C:\\Users\\YURI\\Dropbox\\ROBOS\\ROBÔ BLIG\\situação fiscal\\{mes_compet}\\Geral\\situacao-fiscal-{mes_compet_cont}-{dia}-{cnptabela[tb]}.pdf')
                break
            except:
                if dia < "09":
                    corre = int(dia) + 1
                    dia = f"0{corre}"
                    print("EU TO AQUI")
                    print(f'C:\\Users\\YURI\\Dropbox\\ROBOS\\ROBÔ BLIG\\situação fiscal\\{mes_compet}\\Geral\\situacao-fiscal-{mes_compet_cont}-{dia}-{cnptabela[tb]}.pdf')
                    continue
                else:
                    corre = int(dia) + 1
                    dia = f"{corre}"
                    print("EU TO AQUI")
                    print(f'C:\\Users\\YURI\\Dropbox\\ROBOS\\ROBÔ BLIG\\situação fiscal\\{mes_compet}\\Geral\\situacao-fiscal-{mes_compet_cont}-{dia}-{cnptabela[tb]}.pdf')
                    continue
             
        
        total_paginas = leitura_pdf.get_num_pages()
        conteudo_pdf = ""

        for num_pg in range(0, total_paginas):
            pagina_individual = leitura_pdf.pages[num_pg]
            conteudo_pdf += pagina_individual.extract_text().upper()        

    
        extracao_texto(indice_table, conteudo_pdf)


    tratamento_dados(solicitacao)
    solicitacao += 1
    rel_geral = {}
        
