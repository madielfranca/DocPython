numero = list()
while True:
    n = int(input('Digite um valor: '))
    if n not in numero:
        numero.append(n)
    else:
        print('valor duplicado')
    r = str(input('Quer continuar? [S/N]  ' ))
    if r in 'Nn':
        break

print('=-' * 30)
numero.sort()
print(f'voce digitou os valores {numero}')
numero.sort(reverse=True)
print(f'voce digitou os valores {numero}')
