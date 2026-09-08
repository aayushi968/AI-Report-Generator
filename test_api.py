import requests

API_KEY = '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'

response = requests.get('https://openrouter.ai/api/v1/models', headers={'Authorization': f'Bearer {API_KEY}'})
models = response.json().get('data', [])

free_models = [m['id'] for m in models if ':free' in m['id']]
print("Available free models:")
for m in free_models:
    print(m)
