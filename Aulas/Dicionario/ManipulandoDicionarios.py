locadora = []

filme1 = {'titulo':'star wars',
       'ano':'1990',
       'diretor':'george'
       }
filme1['nota'] = '10'

filme2 = {'titulo':'Vingadores',
       'ano':'2020',
       'diretor':'scepmpea'
       }

locadora.append(filme1)
locadora.append(filme2)

print(locadora)
print('____________________________')
print(filme1.values())
print(filme1.keys())
print(filme1.items())
print('____________________________')
for k, v in filme1.items():
    print(f'o {k} É {v}')





