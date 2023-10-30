import os

import pandas as pd
import xmltodict as xd


def pegar_infos(nome_arquivo, valores):
    with open(f"notas/{nome_arquivo}", "rb") as arquivo_xml:
        dic_arquivo = xd.parse(arquivo_xml)
        infos_nf = dic_arquivo["nfeProc"]["NFe"]["infNFe"]
        numero_nota = infos_nf["@Id"]
        empresa_emissora = infos_nf["emit"]["xNome"]
        nome_cliente = infos_nf["dest"]["xNome"]
        endereco_completo = infos_nf["dest"]["enderDest"]
        rua = endereco_completo["xLgr"]
        numero_casa = endereco_completo["nro"]
        bairro = endereco_completo["xBairro"]
        municipio = endereco_completo["xMun"]
        cep = endereco_completo["CEP"]
        peso = infos_nf["transp"]["vol"]["pesoB"]
        endereco_desejado = {"Rua":rua, "Número da Casa":numero_casa, "Bairro":bairro, "Cidade":municipio, "Estado":uf, "CEP":cep}
        valores.append([numero_nota, empresa_emissora, nome_cliente, endereco_desejado, peso])


arquivos = os.listdir("notas")
colunas = ["numero_nota", "empresa_emissora", "nome_cliente", "endereco", "peso"]
valores = []
for arquivo in arquivos:
    pegar_infos(arquivo, valores)
tabela = pd.DataFrame(columns=colunas, data=valores)
tabela.to_excel("NotasFiscais.xlsx", index=False)
