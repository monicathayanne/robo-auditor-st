import os
import xml.etree.ElementTree as ET
import pandas as pd
from tkinter import filedialog, Tk, messagebox

def escolher_pasta():
    root = Tk()
    root.withdraw()
    root.attributes('-topmost', True)
    pasta = filedialog.askdirectory(title='Selecione a pasta dos XMLs do SIEG')
    root.destroy()
    return pasta

def processar_xml(caminho_xml):
    try:
        tree = ET.parse(caminho_xml)
        root = tree.getroot()
        ns = {'nfe': 'http://www.portalfiscal.inf.br/nfe'}
        
        infNFe = root.find('.//nfe:infNFe', ns)
        if infNFe is None: return None
        
        # --- INICIALIZAÇÃO DAS VARIÁVEIS DE SOMA ---
        valor_total_qualificado = 0
        valor_st_total = 0
        encontrou_criterio = False
        csts_encontrados = set()
        cfops_encontrados = set()
        # ------------------------------------------

        # Dados do Emitente
        emit = root.find('.//nfe:emit', ns)
        cnpj_emitente = emit.find('nfe:CNPJ', ns).text
        
        # Dados do Destinatário (com trava para CPF/CNPJ)
        dest = root.find('.//nfe:dest', ns)
        cnpj_destinatario = "N/A"
        if dest is not None:
            id_dest = dest.find('nfe:CNPJ', ns) if dest.find('nfe:CNPJ', ns) is not None else dest.find('nfe:CPF', ns)
            cnpj_destinatario = id_dest.text if id_dest is not None else "Não identificado"

        # Dados da Nota
        ide = root.find('.//nfe:ide', ns)
        numero_nota = ide.find('nfe:nNF', ns).text
        total = root.find('.//nfe:total/nfe:ICMSTot', ns)
        valor_total_nota = float(total.find('nfe:vNF', ns).text) if total is not None else 0
        chave = infNFe.attrib['Id'].replace('NFe', '')
        dhEmi = ide.find('nfe:dhEmi', ns).text
        ano, mes = dhEmi[:4], dhEmi[5:7]
        
        criterios_cfop = ['5101', '5102', '5103', '5104', '5105', '5106']
        criterios_cst = ['000', '020', '520', '120', '0102', '1102']
        vBCST_total = 0
        vST_total = 0
        
        # Varredura dos Itens
        for det in root.findall('.//nfe:det', ns):
            prod = det.find('nfe:prod', ns)
            imposto = det.find('nfe:imposto', ns)
            cfop = prod.find('nfe:CFOP', ns).text
            
            cst_item = ""
            icms = imposto.find('.//nfe:ICMS', ns)
            for grupo in icms:
                vBCST_tag = grupo.find('nfe:vBCST', ns)
                vST_tag = grupo.find('nfe:vST', ns)

                if vBCST_tag is not None:
                    vBCST_total += float(vBCST_tag.text)
                if vST_tag is not None:
                    vST_total += float(vST_tag.text)

                tag_cst = grupo.find('nfe:CST', ns)
                tag_csosn = grupo.find('nfe:CSOSN', ns)
                if tag_cst is not None:
                    cst_item = tag_cst.text.zfill(3)
                elif tag_csosn is not None:
                    cst_item = tag_csosn.text.zfill(4)

            # --- LÓGICA DE FILTRO E CÁLCULO ---
            if cfop in criterios_cfop and cst_item in criterios_cst:
                encontrou_criterio = True
                csts_encontrados.add(cst_item)
                cfops_encontrados.add(cfop)
                
                vProd = float(prod.find('nfe:vProd', ns).text)
                valor_total_qualificado += vProd
                
                # Aplicação das Alíquotas que você definiu
                if cst_item in ['020', '520', '120']:
                    aliquota = 0.0264
                elif cst_item == '000':
                    aliquota = 0.04
                elif cst_item in ['0102', '1102']:
                    aliquota = 0.07  # 4% + 3%
                else:
                    aliquota = 0
                
                valor_st_total += (vProd * aliquota)

        if encontrou_criterio:
            return {
                "CNPJ Emitente": cnpj_emitente,
                "CNPJ Destinatário": cnpj_destinatario,
                "Ano": ano,
                "Mês": mes,
                "Número NF": numero_nota,
                "Chave": chave,
                "Valor Total da Nota": round(valor_total_nota, 2),
                "CFOPs na Nota": ", ".join(sorted(cfops_encontrados)),
                "CSTs na Nota": ", ".join(sorted(csts_encontrados)),
                "Base de Cálculo": valor_total_qualificado,
                "ICMS ST Calculado": round(valor_st_total, 2),
                "BC ST Destacada (XML)": round(vBCST_total, 2), 
                "Valor ST Destacado (XML)": round(vST_total, 2) 
            }
    except Exception as e:
        print(f"Erro no arquivo {os.path.basename(caminho_xml)}: {e}")
    return None

def iniciar_sistema():
    pasta_selecionada = escolher_pasta()
    if not pasta_selecionada: return

    resultados = []
    for root, dirs, files in os.walk(pasta_selecionada):
        for file in files:
            if file.lower().endswith(".xml"):
                dados = processar_xml(os.path.join(root, file))
                if dados:
                    resultados.append(dados)

    if resultados:
        df = pd.DataFrame(resultados)
        nome_arquivo = "relatorio_apuracao_st_1104.xlsx"
        df.to_excel(nome_arquivo, index=False)
        # Janela de Sucesso
        messagebox.showinfo("Sucesso!", f"Processamento concluído!\n{len(resultados)} notas encontradas.\nO arquivo '{nome_arquivo}' foi gerado.")
    else:
        # Janela de Aviso caso não encontre nada
        messagebox.showwarning("Atenção", "Nenhuma nota foi encontrada com os critérios (CFOPs e CSTs informados) nesta pasta.")

if __name__ == "__main__":
    iniciar_sistema()