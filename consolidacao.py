import pandas as pd
import os
import pyautogui

def pedir_diretorio_raiz():
    return pyautogui.prompt('Digite o caminho do diretório raiz onde estão as planilhas com as informações a serem consolidadas:')

def processar_arquivos(diretorio_raiz):
    caminho_consolidacao = os.path.join(diretorio_raiz, 'consolidacao.xlsx')
    
    if os.path.isfile(caminho_consolidacao):
        df_consolidado = pd.read_excel(caminho_consolidacao, engine='openpyxl')
    else:
        df_consolidado = pd.DataFrame(columns=['Nome da Planilha'])
    
    for pasta in os.listdir(diretorio_raiz):
        caminho_pasta = os.path.join(diretorio_raiz, pasta)
        if os.path.isdir(caminho_pasta):
            for arquivo in os.listdir(caminho_pasta):
                if arquivo.endswith(('.xls', '.xlsx')):
                    df_arquivo = ler_e_filtrar_arquivo(caminho_pasta, arquivo)
                    if df_arquivo is not None:
                        df_arquivo.insert(0, 'Nome da Planilha', arquivo)
                        df_consolidado = pd.concat([df_consolidado, df_arquivo], ignore_index=True)
    
    df_consolidado.to_excel(caminho_consolidacao, index=False, engine='openpyxl')
    print('Planilha de consolidação atualizada com sucesso.')

def ler_e_filtrar_arquivo(caminho_pasta, arquivo):
    caminho_arquivo = os.path.join(caminho_pasta, arquivo)
    try:
        df_arquivo = pd.read_excel(caminho_arquivo)
        return df_arquivo[df_arquivo.iloc[:, 5].notna()].reset_index(drop=True)
    except Exception as e:
        print(f'Erro ao processar o arquivo {caminho_arquivo}: {e}')
        return None

if __name__ == '__main__':
    diretorio_raiz = pedir_diretorio_raiz()
    if diretorio_raiz:
        processar_arquivos(diretorio_raiz)