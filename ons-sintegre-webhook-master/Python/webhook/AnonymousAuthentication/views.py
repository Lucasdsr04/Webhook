from django.http import JsonResponse
from django.views.decorators.csrf import csrf_exempt
from django.views.decorators.http import require_POST
import json
import requests
from webhook.AnonymousAuthentication import Carga_PMO
# use 'requests' em vez de 'urllib'

@csrf_exempt
@require_POST
def webhook(request):
    try:
        # Processa o corpo da solicitação
        body = json.loads(request.body)
        data = body['dataProduto'].replace('/', '-')
        global file_name
        file_name = body['nome']
        url = body['url']
        print(file_name)
        #print(body) #se quiser ver o conteudo do corpo
        #print(url) #se quiser ver a url 
        
        response = requests.get(url) 

        #tratamento da extensão do arquivo (sem especificar = .zip)
        #colocar nome dos arquivos nas listas 
        
        lista_xlsx = ["RDH"] 
        lista_xls = ["Acomph"]

        if  file_name in lista_xlsx :
            extension = 'xlsx'
        elif file_name in lista_xls:
            extension = 'xls'
        else:
            extension = 'zip'

        # Salva o conteúdo do arquivo        
        with open(f'C:/Users/8102081/OneDrive - Light/Área de Trabalho/Webhook/Solicitações/{file_name}/{file_name} {data}.{extension}', 'wb') as f:            
            f.write(response.content)       
        # substitua o diretório até antes de {file_name}
        
        print(f)
        print("Sucesso ao recuperar o arquivo.")

        global path_arquivo
        path_arquivo = f'{file_name} {data}.{extension}'
        
        if file_name == 'Carga por patamar - DECOMP':
            Carga_PMO.carga_pmo(file_name,path_arquivo)

    except Exception as e:
        print(f"Erro: {e}")
        return JsonResponse({"error": "Erro ao recuperar arquivo"}, status=500)
    
    return JsonResponse({"status": "ok"}, status=200)



