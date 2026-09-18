import requests

API_KEY = 'Your Open Router API KEY'

response = requests.get('https://openrouter.ai/api/v1/models', headers={'Authorization': f'Bearer {API_KEY}'})
models = response.json().get('data', [])

free_models = [m['id'] for m in models if ':free' in m['id']]
print("Available free models:")
for m in free_models:
    print(m)
