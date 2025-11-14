import requests
import json
from openpyxl import Workbook, load_workbook
from api_key import headers

city = 'Los Angeles County'
state = 'CA'
year = '2025'

workbook = Workbook()
sheet = workbook.active

url = f'https://api.gsa.gov/travel/perdiem/v2/rates/city/{city}/state/{state}/year/{year}'

try:
    response = requests.get(url, headers=headers)
    response.raise_for_status()

    data = response.json()
    final = json.dumps(data, indent=4)

    #print(final)
    print(f'Meals: {data['rates'][0]['rate'][0]['meals']}\nCounty: {data['rates'][0]['rate'][0]['county']}\nCity {data['rates'][0]['rate'][0]['city']}\nStandard Rate: {data['rates'][0]['rate'][0]['standardRate']}\nState: {data['rates'][0]['state']}')

    sheet['A1'] = 'Meals'
    sheet['B1'] = data['rates'][0]['rate'][0]['meals']
    
    workbook.save(filename="test.xlsx")


except Exception as e:
    print(f'Error: {e}')
finally:
    print('done')
