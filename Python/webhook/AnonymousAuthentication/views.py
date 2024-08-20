from django.http import JsonResponse
from django.views.decorators.csrf import csrf_exempt
from django.views.decorators.http import require_POST
import json
import requests  # use 'requests' em vez de 'urllib'

@csrf_exempt
@require_POST
def webhook(request):
    try:
        # Processa o corpo da solicitação
        body = json.loads(request.body)
        data = body['dataProduto'].replace('/', '-')
        file_name = body['nome']
        url = body['url']
        print(file_name)
        #print(body) #se quiser ver o conteudo do corpo
        #print(url) #se quiser ver a url 
        
        response = requests.get(url)          
                    
        # Salva o conteúdo do arquivo        
        with open(f'C:/Users/8102081/OneDrive - Light/Área de Trabalho/Webhook/Solicitações/{file_name} {data}.zip', 'wb') as f:            
            f.write(response.content)       
        # substitua o diretório até antes de {file_name} e se desejar mude a extensão do arquivo apos {data}
        
        print("Sucesso ao recuperar o arquivo.")

    except Exception as e:
        print(f"Erro: {e}")
        return JsonResponse({"error": "Erro ao recuperar arquivo"}, status=500)
    
    return JsonResponse({"status": "ok"}, status=200)


