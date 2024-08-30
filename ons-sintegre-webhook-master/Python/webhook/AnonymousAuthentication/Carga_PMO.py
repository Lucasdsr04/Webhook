import zipfile
import pandas as pd
import glob
import locale
from datetime import datetime
from dateutil.relativedelta import relativedelta
import re

def carga_pmo(file_name,path_arquivo):
    #pip install xlsxwriter
    # Obtendo o mês vigente / data vigente
    locale.setlocale(locale.LC_TIME, 'pt_BR.UTF-8')
    dat = datetime.now()
    mes_ano = dat.strftime('%B%y').capitalize()
    data_atual = dat.strftime('%d-%B-%y')

    # Obter o próximo mês
    proximo_mes = dat + relativedelta(months=1)

    # Formatar o próximo mês e ano
    proximo_mes_ano = proximo_mes.strftime('%B%y').capitalize()

    # Caminho do arquivo zip e diretório de extração
    zip_path = f'Solicitações/{file_name}/{path_arquivo}'
    extract_dir = 'Retornos/Carga-PMO'

    # Extrair arquivos do zip
    with zipfile.ZipFile(zip_path, 'r') as zip_ref:
        zip_ref.extractall(extract_dir)

    #Extrair o número da revisão de cada nome de arquivo
    try:
        file_pattern = f'Retornos/Carga-PMO/*{proximo_mes_ano}(Rev *.xlsx'
        file_list = glob.glob(file_pattern)
        Rvs = [int(re.search(r'Rev (\d+)', path).group(1)) for path in file_list]
        Rvs = max(Rvs)
        mes_atual = proximo_mes_ano
    except:
        file_pattern = f'Retornos/Carga-PMO/*{mes_ano}(Rev *.xlsx'
        file_list = glob.glob(file_pattern)
        Rvs = [int(re.search(r'Rev (\d+)', path).group(1)) for path in file_list]
        Rvs = max(Rvs)
        mes_atual = mes_ano
        
    def pmo(Rvs):
        # Define o padrão do nome do arquivo
        file_pattern = f'Retornos/Carga-PMO/Carga_PMO_{mes_atual}*.xlsx'

        # Encontra o arquivo que corresponde ao padrão
        file_list = glob.glob(file_pattern)
        #print(file_list) #Pra ver a lista dos arquivos que ta lendo 
        #O numero em file_list[numero] corresponde as RVs.
        try:
            df = pd.read_excel(file_list[Rvs])
            #limpa NaN no geral
            df.dropna()
            df = df.fillna('')
            df = df.replace('nan', '')

            # Renomear colunas 'Unnamed'
            df.columns = [f'Coluna_{i}' if 'Unnamed' in col else col for i, col in enumerate(df.columns)]

            # Escolher colunas (que estão vazias) para dropar
            colunas_para_dropar = ['Coluna_0', 'Coluna_1', 'Coluna_3', 'Coluna_4', 'Coluna_5']  # Substitua pelos nomes das colunas que deseja remover
            data = df.drop(columns=colunas_para_dropar)

            # Dropar linhas específicas (que estão vazias)
            linhas_para_dropar = [0,1,2,3,4,5,6,7,10,11] 
            data = data.drop(index=linhas_para_dropar)

            # Transformar os dados em um dicionário
            data_dict = data.to_dict(orient='list')

            # Extrair o primeiro item do dicionário
            primeiro_item = {k: v[0] for k, v in data_dict.items()}

            # Criar um novo dicionário com as chaves renomeadas (para ser as colunas posteriormente)
            novo_dicionario = {primeiro_item[k]: v for k, v in data_dict.items()}

            #Tratamento pra se existir 6ª Semana
            if novo_dicionario.get('6ª Semana')[2] == '':
                del novo_dicionario['6ª Semana']

            # Converter o dicionário limpo para um DataFrame
            clean_data_df = pd.DataFrame(dict([(k, pd.Series(v)) for k, v in novo_dicionario.items()]))

            # Limitar o número de casas decimais para 2
            def formatar_numeros(valor):
                if isinstance(valor, (int, float)):
                    return f'{valor:.0f}'
                return valor

            colunas = ['Mensal', '1ª Semana', '2ª Semana', '3ª Semana', '4ª Semana', '5ª Semana']

            for coluna in colunas:
                clean_data_df[coluna] = clean_data_df[coluna].apply(formatar_numeros)

            #Tratamento pra se existir 6ª Semana
            if '6ª Semana' in clean_data_df.columns:
                clean_data_df['6ª Semana'] = clean_data_df['6ª Semana'].apply(formatar_numeros)

            # formatando só os dados que quer (usar print em clean_data_df pra ver os dados que não estao sendo utilizados se quiser)
            data_final = clean_data_df.drop(index=[0,7,8])

            #formatando ordem dos dados
            data_final.loc[[2, 4]] = data_final.loc[[4, 2]].values
            data_final.loc[[3, 5]] = data_final.loc[[5, 3]].values
            
            # Convertendo as linhas específicas das colunas para tipos numéricos
            linhas_para_converter = range(2, 7)
            for coluna in colunas:
                data_final.loc[linhas_para_converter, coluna] = pd.to_numeric(data_final.loc[linhas_para_converter, coluna])
                
                #Tratamento pra se existir 6ª Semana
                if '6ª Semana' in data_final.columns:
                    data_final.loc[linhas_para_converter, '6ª Semana'] = pd.to_numeric(data_final.loc[linhas_para_converter, '6ª Semana'])
            
            #Se quiser ver o DataFrame tirar o # abaixo
            #print(data_final) 

            data_final.to_excel(f'Retornos/Carga-PMO/RV{Rvs} {mes_atual}.xlsx',index=False)

        except:
            print(f'Não existe uma RV{Rvs}')
        
    def comparacao(Rvs):
        # Carregar os arquivos Excel em DataFrames
        df1 = pd.read_excel(f'Retornos/Carga-PMO/RV{Rvs} {mes_atual}.xlsx')
        if Rvs > 0:
            df2 = pd.read_excel(f'Retornos/Carga-PMO/RV{Rvs - 1} {mes_atual}.xlsx')

            # Calculando a diferença entre os números de df2 e df1
            df3 = df2.copy()
            for col in df1.columns[1:]:
                df3[col] = pd.to_numeric(df1[col], errors='coerce') - pd.to_numeric(df2[col], errors='coerce')
            df2 = df3

            with pd.ExcelWriter(f'Retornos/Carga-PMO/Comparações/{data_atual} Comparação_PMO_{mes_atual}.xlsx', engine='xlsxwriter') as writer:
                # Escreve df1 no arquivo Excel
                df1.to_excel(writer, sheet_name='Comparação PMO', startrow=1, index=False)
                worksheet = writer.sheets['Comparação PMO']
                worksheet.write(0, 0, f"Valores Carga PMO RV{Rvs} {mes_ano}")
                
                # Escreve df2 no arquivo Excel com um espaço de 2 linhas abaixo de df1
                df2.to_excel(writer, sheet_name='Comparação PMO', startrow=len(df1) + 4, index=False)
                startrow = len(df1) + 5
                worksheet.write(startrow - 2, 0, f"Comparação com RV{Rvs - 1}")

            # Adicionando formatação condicional para destacar as diferenças
                format_blue = writer.book.add_format({'bg_color': '#DDEBF7', 'font_color': '#2F75B5'})
                format_orange = writer.book.add_format({'bg_color': '#FCE4D6', 'font_color': '#C65911'})
                
                for col in range(1, len(df2.columns) + 1):
                    worksheet.conditional_format(startrow + 1, col, startrow + len(df2), col, 
                                                {'type': 'cell', 'criteria': '>', 'value': 0, 'format': format_blue})
                    worksheet.conditional_format(startrow + 1, col, startrow + len(df2), col, 
                                                {'type': 'cell', 'criteria': '<', 'value': 0, 'format': format_orange})
        else:    
            with pd.ExcelWriter(f'Retornos/Carga-PMO/Comparações/{data_atual} Comparação_PMO_{mes_atual}.xlsx', engine='xlsxwriter') as writer:
                # Escreve df1 no arquivo Excel
                df1.to_excel(writer, sheet_name='Comparação PMO', startrow=1, index=False)
                worksheet = writer.sheets['Comparação PMO']
                worksheet.write(0, 0, f"Valores Carga PMO RV{Rvs} {proximo_mes_ano}")

    #Numero da RV como parametro. RV0 = 0; RV1=1 ...
    pmo(Rvs)
    comparacao(Rvs)

