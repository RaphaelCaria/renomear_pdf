import os
import shutil
from pathlib import Path
import re
import subprocess
import fitz  # PyMuPDF
import traceback

# Lista de valores a serem renomeados
valores_para_renomear = [
    "_KIT - ","_ASO - ","_FC-","_ASO- ", "_TERMO DE ENCAMINHAMENTO -",
    "_ASO DEMISSIONAL0 ","_ASO DEMISSIONAL1 ","_ASO DEMISSIONAL2 ","_ASO DEMISSIONAL3 ","_ASO DEMISSIONAL4 ","_ASO DEMISSIONAL5 ","_ASO DEMISSIONAL6 ","_ASO DEMISSIONAL7 ",
    "_ASO ADMISSIONAL0 ","_ASO ADMISSIONAL1 ","_ASO ADMISSIONAL2 ","_ASO ADMISSIONAL3 ","_ASO ADMISSIONAL4 ","_ASO ADMISSIONAL5 ","_ASO ADMISSIONAL6 ","_ASO ADMISSIONAL7 ",
    "_ASO PERIODICO0 ","_ASO PERIODICO1 ","_ASO PERIODICO2 ","_ASO PERIODICO3 ","_ASO PERIODICO4 ","_ASO PERIODICO5 ","_ASO PERIODICO6 ","_ASO PERIODICO7 ",
    "_ASO RETORNO AO TRABALHO0 ","_ASO RETORNO AO TRABALHO1 ","_ASO RETORNO AO TRABALHO2 ","_ASO RETORNO AO TRABALHO3 ","_ASO RETORNO AO TRABALHO4 ","_ASO RETORNO AO TRABALHO5 ","_ASO RETORNO AO TRABALHO6 ",
    "_ASO_RETORNO_AO_TRABALHO1","_ASO_RETORNO_AO_TRABALHO2","_ASO_RETORNO_AO_TRABALHO3","_ASO_RETORNO_AO_TRABALHO4","_ASO_RETORNO_AO_TRABALHO5",
    "_ASO_RETORNO_AO_TRABALHO6","_ASO_RETORNO_AO_TRABALHO7","_ASO_RETORNO_AO_TRABALHO8","_ASO_RETORNO_AO_TRABALHO9",  
    "_ACUIDADE_VISUAL1", "_ACUIDADE_VISUAL2", "_ACUIDADE_VISUAL3", "_ACUIDADE_VISUAL4", "_ACUIDADE_VISUAL5",
    "_ACUIDADE_VISUAL6", "_ACUIDADE_VISUAL7", "_ACUIDADE_VISUAL8", "_ACUIDADE_VISUAL9", "_ACUIDADE_VISUAL0",
    "_ASO_ADMISSIONAL1", "_ASO_ADMISSIONAL2", "_ASO_ADMISSIONAL3", "_ASO_ADMISSIONAL4", "_ASO_ADMISSIONAL5",
    "_ASO_ADMISSIONAL6", "_ASO_ADMISSIONAL7", "_ASO_ADMISSIONAL8", "_ASO_ADMISSIONAL9", "_ASO_ADMISSIONAL0",
    "_ASO_DEMISSIONAL1", "_ASO_DEMISSIONAL2", "_ASO_DEMISSIONAL3", "_ASO_DEMISSIONAL4", "_ASO_DEMISSIONAL5",
    "_ASO_DEMISSIONAL6", "_ASO_DEMISSIONAL7", "_ASO_DEMISSIONAL8", "_ASO_DEMISSIONAL9", "_ASO_DEMISSIONAL0",
    "_ASO_PERIODICO1", "_ASO_PERIODICO2", "_ASO_PERIODICO3", "_ASO_PERIODICO4", "_ASO_PERIODICO5",
    "_ASO_PERIODICO6", "_ASO_PERIODICO7", "_ASO_PERIODICO8", "_ASO_PERIODICO9", "_ASO_PERIODICO0",
    "_ASO_PERIODICO_1", "_ASO_PERIODICO_2", "_ASO_PERIODICO_3", "_ASO_PERIODICO_4", "_ASO_PERIODICO_5",
    "_ASO_PERIODICO_6", "_ASO_PERIODICO_7", "_ASO_PERIODICO_8", "_ASO_PERIODICO_9", "_ASO_PERIODICO_0",
    "_ASO PERIODICO0 ","_ASO PERIODICO1 ","_ASO PERIODICO2 ","_ASO PERIODICO3 ","_ASO PERIODICO4 ","_ASO PERIODICO5 ","_ASO PERIODICO6 ","_ASO PERIODICO7 ","_ASO PERIODICO8 ","_ASO PERIODICO9 ","_ASO PERIODICO10 ","_ASO PERIODICO11 ",
    "_AUDIOMETRIA1_AUDIO", "_AUDIOMETRIA2_AUDIO", "_AUDIOMETRIA3_AUDIO", "_AUDIOMETRIA4_AUDIO", "_AUDIOMETRIA5_AUDIO", "_AUDIOMETRIA6_AUDIO",
    "_AUDIOMETRIA7_AUDIO", "_AUDIOMETRIA8_AUDIO", "_AUDIOMETRIA9_AUDIO",
    "_AUDIOMETRIA_TONAL1", "_AUDIOMETRIA_TONAL2", "_AUDIOMETRIA_TONAL3", "_AUDIOMETRIA_TONAL4", "_AUDIOMETRIA_TONAL5",
    "_AUDIOMETRIA_TONAL6", "_AUDIOMETRIA_TONAL7", "_AUDIOMETRIA_TONAL8", "_AUDIOMETRIA_TONAL9", "_AUDIOMETRIA_TONAL0",
    "_AUDIOMETRIA1", "_AUDIOMETRIA2", "_AUDIOMETRIA3", "_AUDIOMETRIA4", "_AUDIOMETRIA5", "_AUDIOMETRIA6",
    "_AUDIOMETRIA7", "_AUDIOMETRIA8", "_AUDIOMETRIA9", "_AUDIOMETRIA0",
    "_AVALIAÇÃO_PSICOSSOCIAL1", "_AVALIAÇÃO_PSICOSSOCIAL2", "_AVALIAÇÃO_PSICOSSOCIAL3", "_AVALIAÇÃO_PSICOSSOCIAL4",
    "_AVALIAÇÃO_PSICOSSOCIAL5", "_AVALIAÇÃO_PSICOSSOCIAL6", "_AVALIAÇÃO_PSICOSSOCIAL7", "_AVALIAÇÃO_PSICOSSOCIAL8",
    "_AVALIAÇÃO_PSICOSSOCIAL9", "_AVALIAÇÃO_PSICOSSOCIAL0",
    "_ELETROCARDIOGRAMA1", "_ELETROCARDIOGRAMA2", "_ELETROCARDIOGRAMA3", "_ELETROCARDIOGRAMA4", "_ELETROCARDIOGRAMA5",
    "_ELETROCARDIOGRAMA6", "_ELETROCARDIOGRAMA7", "_ELETROCARDIOGRAMA8", "_ELETROCARDIOGRAMA9", "_ELETROCARDIOGRAMA0",
    "_ELETROENCEFALOGRAMA1", "_ELETROENCEFALOGRAMA2", "_ELETROENCEFALOGRAMA3", "_ELETROENCEFALOGRAMA4",
    "_ELETROENCEFALOGRAMA5", "_ELETROENCEFALOGRAMA6", "_ELETROENCEFALOGRAMA7", "_ELETROENCEFALOGRAMA8",
    "_ELETROENCEFALOGRAMA9", "_ELETROENCEFALOGRAMA0",
    "_ESPIROMETRIA1", "_ESPIROMETRIA2", "_ESPIROMETRIA3", "_ESPIROMETRIA4", "_ESPIROMETRIA5", "_ESPIROMETRIA6",
    "_ESPIROMETRIA7", "_ESPIROMETRIA8", "_ESPIROMETRIA9", "_ESPIROMETRIA0",
    "_EXAME_CLÍNICO1", "_EXAME_CLÍNICO2", "_EXAME_CLÍNICO3", "_EXAME_CLÍNICO4", "_EXAME_CLÍNICO5",
    "_EXAME_CLÍNICO6", "_EXAME_CLÍNICO7", "_EXAME_CLÍNICO8", "_EXAME_CLÍNICO9", "_EXAME_CLÍNICO0",
    "_EXAME CLÍNICO0 ","_EXAME CLÍNICO1 ","_EXAME CLÍNICO2 ","_EXAME CLÍNICO3 ","_EXAME CLÍNICO4 ","_EXAME CLÍNICO5 ","_EXAME CLÍNICO6 ","_EXAME CLÍNICO7 ","_EXAME CLÍNICO8 ","_EXAME CLÍNICO9 ",
    "_GLICOSE1", "_GLICOSE2", "_GLICOSE3", "_GLICOSE4", "_GLICOSE5", "_GLICOSE6", "_GLICOSE7", "_GLICOSE8", "_GLICOSE9", "_GLICOSE0",
    "_AVALIAÇÃO_MÉDICA1_AV", "_AVALIAÇÃO_MÉDICA2_AV", "_AVALIAÇÃO_MÉDICA3_AV", "_AVALIAÇÃO_MÉDICA4_AV", "_AVALIAÇÃO_MÉDICA5_AV", "_AVALIAÇÃO_MÉDICA6_AV",
    "_AVALIAÇÃO_MÉDICA7_AV", "_AVALIAÇÃO_MÉDICA8_AV", "_AVALIAÇÃO_MÉDICA9_AV", "_AVALIAÇÃO_MÉDICA0_AV",
    "_ASO_MUDANÇA_DE_RISCO1",
    "_LAUDO_PNE1_LAUDO_PNE","_LAUDO_PNE2_LAUDO_PNE","_LAUDO_PNE3_LAUDO_PNE","_LAUDO_PNE4_LAUDO_PNE","_LAUDO_PNE5_LAUDO_PNE","_LAUDO_PNE6_LAUDO_PNE",
    "_LAUDO_PNE7_LAUDO_PNE","_LAUDO_PNE8_LAUDO_PNE","_LAUDO_PNE9_LAUDO_PNE","_LAUDO_PNE0_LAUDO_PNE",
    "_RETORNO_AO_TRABALHO1","_RETORNO_AO_TRABALHO2","_RETORNO_AO_TRABALHO3","_RETORNO_AO_TRABALHO4","_RETORNO_AO_TRABALHO5","_RETORNO_AO_TRABALHO6",
    "_RETORNO_AO_TRABALHO7","_RETORNO_AO_TRABALHO8","_RETORNO_AO_TRABALHO9","_RETORNO_AO_TRABALHO0",
    "_ASO ADMISSIONAL0 ","_ASO ADMISSIONAL1 ","_ASO ADMISSIONAL2 ","_ASO ADMISSIONAL3 ","_ASO ADMISSIONAL4 ",
    "_EXAME CLÍNICO0 ","_EXAME CLÍNICO1 ","_EXAME CLÍNICO2 ","_EXAME CLÍNICO3 ","_EXAME CLÍNICO4 ",
    "_ DEMISSIONAL0 ","_ DEMISSIONAL1 ","_ DEMISSIONAL2 ","_ DEMISSIONAL3 ","_ DEMISSIONAL4 ",
    "_ PERIODICO0 ","_ PERIODICO1 ","_ PERIODICO2 ","_ PERIODICO3 ","_ PERIODICO4 ","_ PERIODICO5 ","_ PERIODICO6 ","_ PERIODICO7 ",
    "_AUDIOMETRIA TONAL0","_AUDIOMETRIA TONAL1","_AUDIOMETRIA TONAL2","_AUDIOMETRIA TONAL3","_AUDIOMETRIA TONAL4","_AUDIOMETRIA TONAL5","_AUDIOMETRIA TONAL6","_AUDIOMETRIA TONAL7","_AUDIOMETRIA TONAL8","_AUDIOMETRIA TONAL9",
    "_PESQUISA DE PLASMODIUM0","_PESQUISA DE PLASMODIUM1","_PESQUISA DE PLASMODIUM2","_PESQUISA DE PLASMODIUM3","_PESQUISA DE PLASMODIUM4","_PESQUISA DE PLASMODIUM5","_PESQUISA DE PLASMODIUM6","_PESQUISA DE PLASMODIUM7","_PESQUISA DE PLASMODIUM8","_PESQUISA DE PLASMODIUM9",
    "_RX TORAX PA COM LAUDO OIT0","_RX TORAX PA COM LAUDO OIT1","_RX TORAX PA COM LAUDO OIT2","_RX TORAX PA COM LAUDO OIT3","_RX TORAX PA COM LAUDO OIT4","_RX TORAX PA COM LAUDO OIT5","_RX TORAX PA COM LAUDO OIT6","_RX TORAX PA COM LAUDO OIT7","_RX TORAX PA COM LAUDO OIT8","_RX TORAX PA COM LAUDO OIT9",
    "_ACUIDADE VISUAL0","_ACUIDADE VISUAL1","_ACUIDADE VISUAL2","_ACUIDADE VISUAL3","_ACUIDADE VISUAL4","_ACUIDADE VISUAL5","_ACUIDADE VISUAL6","_ACUIDADE VISUAL7","_ACUIDADE VISUAL8","_ACUIDADE VISUAL9",
    "_AVALIAÇÃO PSICOLÓGICA0","_AVALIAÇÃO PSICOLÓGICA1","_AVALIAÇÃO PSICOLÓGICA2","_AVALIAÇÃO PSICOLÓGICA3","_AVALIAÇÃO PSICOLÓGICA4","_AVALIAÇÃO PSICOLÓGICA5","_AVALIAÇÃO PSICOLÓGICA6","_AVALIAÇÃO PSICOLÓGICA7","_AVALIAÇÃO PSICOLÓGICA8","_AVALIAÇÃO PSICOLÓGICA9",
    "_HEMOGRAMA COMPLETO E PLAQUETAS0","_HEMOGRAMA COMPLETO E PLAQUETAS1","_HEMOGRAMA COMPLETO E PLAQUETAS2","_HEMOGRAMA COMPLETO E PLAQUETAS3","_HEMOGRAMA COMPLETO E PLAQUETAS4","_HEMOGRAMA COMPLETO E PLAQUETAS5","_HEMOGRAMA COMPLETO E PLAQUETAS6","_HEMOGRAMA COMPLETO E PLAQUETAS7","_HEMOGRAMA COMPLETO E PLAQUETAS8","_HEMOGRAMA COMPLETO E PLAQUETAS9",
    "_AUDIOMETRIA VOCAL E TONAL0","_AUDIOMETRIA VOCAL E TONAL1","_AUDIOMETRIA VOCAL E TONAL2","_AUDIOMETRIA VOCAL E TONAL3","_AUDIOMETRIA VOCAL E TONAL4","_AUDIOMETRIA VOCAL E TONAL5","_AUDIOMETRIA VOCAL E TONAL6","_AUDIOMETRIA VOCAL E TONAL7","_AUDIOMETRIA VOCAL E TONAL8","_AUDIOMETRIA VOCAL E TONAL9",
    "_ AUDIO","_AUDIOMETRIA -","_EXAMES E FICHA CLINICA","_RH VIDA ASO ","_RH MED -","_PRONT_","_PRONTUARIO"
]

# Lista de valores para mover para o final do nome do arquivo
valores_para_mover_final = [
"_ELEKTRO - ",
"_VIVARA - ",
"_RH Med - ",    
"_ASO_RHMED_-",
"_ASO_-_EXAMES_-",
"_RHMED_-_ASO_-",
"_RHMED_-_EXAMES_-",
"_RHMED_-_FICHA_-",    
"_Registro_1_-",
"_ASO_+_PRONTUARIO_-_",
"__ADMISSIONAL",
"__CLÍNICO",
"__DEMISSIONAL",
"__PERIODICO",
"_FICHA", 
"_-_NORMAL_-", 
"_RHMED_-", 
"_RH_Med_-",
"_Pront.",
"_P-",
"_P_-",
"_KIT_-",
"_ASO",
"_RH_VIDA",
"_ASO_-",
"_kit_-",
"_RHMED",
"_ASO-",
"_ASO_",
"_FC-_",
"_FC-__",
"_-",
"_EXAMES_",
"_EXAMES -",
"_EXAME -"
]

# Função para OCR
def ocr_pdf(input_pdf, output_pdf, language='por'):
    try:
        subprocess.run([
            'ocrmypdf',
            '-l', language,
            '--optimize', '1',
            '--skip-text',
            input_pdf,
            output_pdf
        ], check=True)
        print(f'OCR processado com sucesso: {output_pdf}')
    except subprocess.CalledProcessError as e:
        print(f'Erro ao processar OCR: {e}')
    except FileNotFoundError:
        print("O comando ocrmypdf não foi encontrado. Certifique-se de que está instalado e no PATH.")
    except Exception as e:
        print(f"Ocorreu um erro inesperado: {e}")

# Função para extrair texto usando PyMuPDF
def extrair_texto(pdf_path):
    texto_completo = ""
    with fitz.open(pdf_path) as doc:
        for page in doc:
            texto_completo += page.get_text()
    return texto_completo

# Função para encontrar o nome do colaborador
def encontrar_nome(texto):
    match = re.search(r"\s(funcionário|ifuncionário):?[\s\S]*?\d+\s*-\s*([A-ZÇÁÉÍÓÚÃÕÂÊÎÔÛÀÈÌÒÙÄËÏÖÜ ]+) ", texto, re.IGNORECASE)
    if match:
        return match.group(2).strip()

    match = re.search(r"\s(nome|inome)\s*:[\W]*\s*([A-ZÇÁÉÍÓÚÃÕÂÊÎÔÛÀÈÌÒÙÄËÏÖÜ ]*) ", texto, re.IGNORECASE)
    if match:
        return match.group(2).strip()

    match = re.search(r"(?i)(nome|inome|funcionário|ifuncionário)\s*:\s*([A-ZÇÁÉÍÓÚÃÕÂÊÎÔÛÀÈÌÒÙÄËÏÖÜ ]{2,})", texto)
    if match:
        return match.group(2).strip()
    else:
        return "NOME_NAO_ENCONTRADO"

# Função para renomear arquivos
def renomear_arquivos(diretorio_entrada, diretorio_saida):
    for arquivo in Path(diretorio_entrada).glob("*.pdf"):
        nome_arquivo = arquivo.stem
        extensao = arquivo.suffix
        novo_nome = None

        # Verifica e renomeia com base na primeira lista de valores
        for valor in valores_para_renomear:
            if valor in nome_arquivo:
                partes = nome_arquivo.split(valor)
                if len(partes) == 2:
                    novo_nome = f"{partes[0].rstrip('_')}_{partes[1].lstrip('_')}_{valor.strip('_')}{extensao}"
                    break
                elif re.search(rf"{valor}\d{6}", nome_arquivo):
                    match = re.search(rf"({valor}\d{{6}})", nome_arquivo)
                    if match:
                        valor_com_numero = match.group(1)
                        partes = nome_arquivo.split(valor_com_numero)
                        novo_nome = f"{partes[0].rstrip('_')}_{partes[1].lstrip('_')}_{valor_com_numero.strip('_')}{extensao}"
                        break    
        
        # Se o arquivo contiver "ENCAMINHAMENTO" e não foi renomeado
        if "ENCAMINHAMENTO" in nome_arquivo and novo_nome is None:
            ocr_pdf(arquivo, arquivo)
            texto = extrair_texto(arquivo)
            nome_colaborador = encontrar_nome(texto)
            partes = nome_arquivo.split("_ENCAMINHAMENTO")
            if len(partes) == 2:
                novo_nome = f"{partes[0].rstrip('_')}_{nome_colaborador}_{partes[1].strip('_')}{extensao}"
        
        # Se o arquivo contiver "scan_color_inferior" e não foi renomeado
        if "scan_color_inferior" in nome_arquivo and novo_nome is None:
            ocr_pdf(arquivo, arquivo)
            texto = extrair_texto(arquivo)
            nome_colaborador = encontrar_nome(texto)
            partes = nome_arquivo.split("_scan_color_inferior")
            if len(partes) == 2:
                novo_nome = f"{partes[0].rstrip('_')}_{nome_colaborador}_{partes[1].strip('_')}{extensao}"

        # Se o arquivo contiver "scan_color_inferior" e não foi renomeado
        if "scan_color_superior" in nome_arquivo and novo_nome is None:
            ocr_pdf(arquivo, arquivo)
            texto = extrair_texto(arquivo)
            nome_colaborador = encontrar_nome(texto)
            partes = nome_arquivo.split("scan_color_superior")
            if len(partes) == 2:
                novo_nome = f"{partes[0].rstrip('_')}_{nome_colaborador}_{partes[1].strip('_')}{extensao}"
        # Se não foi possível renomear, usa o nome original
        if novo_nome is None:
            novo_nome = arquivo.name
        
        novo_caminho = Path(diretorio_saida) / novo_nome
        if not novo_caminho.exists():  # Evitar duplicação
            shutil.copy(arquivo, novo_caminho)
            print(f"Arquivo copiado: {arquivo.name} -> {novo_nome}")

# Função para mover valores específicos para o final do nome do arquivo
def mover_valores_final(diretorio_saida):
    for arquivo in Path(diretorio_saida).glob("*.pdf"):
        nome_arquivo = arquivo.stem
        extensao = arquivo.suffix

        for valor in valores_para_mover_final:
            if valor in nome_arquivo:
                partes = nome_arquivo.split(valor)
                if len(partes) == 2:
                    novo_nome = f"{partes[0].rstrip('_')}_{partes[1].lstrip('_')}_{valor.strip('_')}{extensao}"
                    novo_caminho = Path(diretorio_saida) / novo_nome
                    if not novo_caminho.exists() and len(str(novo_caminho)) <= 260:  # Evitar duplicação e verificar comprimento do caminho
                        try:
                            arquivo.rename(novo_caminho)
                            print(f"Arquivo renomeado: {arquivo.name} -> {novo_nome}")
                            break
                        except FileNotFoundError as e:
                            print(f"Erro ao renomear {arquivo.name}: {traceback.format_exc()}")

def mover_numeros_final(diretorio_saida):
    for arquivo in Path(diretorio_saida).glob("*.pdf"):
        nome_arquivo = arquivo.stem
        extensao = arquivo.suffix    

        # Verifica e move números após DDMMAAAA_HHM para o final
        match = re.search(r'(\d{8}\_\d{4})(\_\d+)(.*)', nome_arquivo)
        if match:
            data_hora = match.group(1)
            numeros = match.group(2)
            resto_nome = match.group(3)
            novo_nome = f"{data_hora}{resto_nome}{numeros}{extensao}"
            novo_caminho = Path(diretorio_saida) / novo_nome
            if not novo_caminho.exists() and len(str(novo_caminho)) <= 260:  # Evitar duplicação e verificar comprimento do caminho
                try:
                    arquivo.rename(novo_caminho)
                    print(f"Arquivo renomeado: {arquivo.name} -> {novo_nome}")
                except FileNotFoundError as e:
                    print(f"Erro ao renomear {arquivo.name}: {e}")

# Função para analisar palavras-chave nos arquivos PDF
def analisar_empresas(diretorio_saida):
    palavras_chave = input("Insira as palavras-chave (nomes das empresas) separadas por vírgula: ").split(',')
    palavras_chave = [palavra.strip().lower() for palavra in palavras_chave]

    for arquivo in Path(diretorio_saida).glob("*.pdf"):
        ocr_pdf(arquivo, arquivo)  # Executa o OCR em todos os arquivos antes da análise
        texto = extrair_texto(arquivo).lower()  # Converte o texto extraído para minúsculas
        for palavra in palavras_chave:
            if re.search(rf'{palavra}\s', texto, re.IGNORECASE):
                novo_nome = f"{arquivo.stem}_ENCONTRADO{arquivo.suffix}"
                novo_caminho = arquivo.parent / novo_nome
                if not novo_caminho.exists():  # Evitar duplicação
                    arquivo.rename(novo_caminho)
                    print(f"Palavra '{palavra}' encontrada. Arquivo renomeado: {arquivo.name} -> {novo_nome}")
                break

# Função para mover arquivos "ENCONTRADO" para uma nova pasta
def mover_arquivos_encontrados(diretorio_saida):
    pasta_encontrado = Path(diretorio_saida) / "ENCONTRADO"
    if not pasta_encontrado.exists():
        criar_pasta = input("Deseja criar a pasta 'ENCONTRADO' e mover os arquivos? (s/n): ").strip().lower()
        if criar_pasta == 's':
            pasta_encontrado.mkdir()
            for arquivo in Path(diretorio_saida).glob("*ENCONTRADO*.pdf"):
                novo_caminho = pasta_encontrado / arquivo.name
                shutil.move(str(arquivo), str(novo_caminho))
                print(f"Arquivo movido para 'ENCONTRADO': {arquivo.name}")

# Função principal
def main():
    diretorio_entrada = input("Insira o caminho do diretório de entrada: ")
    diretorio_saida = input("Insira o caminho do diretório de saída: ")

    renomear_arquivos(diretorio_entrada, diretorio_saida)
    mover_valores_final(diretorio_saida)
    mover_numeros_final(diretorio_saida)

    analisar = input("Deseja analisar o nome das empresas nos arquivos? (s/n): ").strip().lower()
    if analisar == 's':
        analisar_empresas(diretorio_saida)

    mover_arquivos_encontrados(diretorio_saida)

if __name__ == "__main__":
    main()
