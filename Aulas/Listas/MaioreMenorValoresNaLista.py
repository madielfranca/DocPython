listanum = []
maoir = 0
menor = 0

for c in range(0, 5):
    listanum.append(int(input(f'Digite um valor para a posição {c}: ')))
    if c == 0:
        maior = menor = listanum[c]
    else:
        if listanum[c] > maoir:
            maoir = listanum[c]
        if listanum[c] < menor:
            menor = listanum[c]


    print('=-' * 30)
    print(f'você digitou os valores {listanum}')
    print(f'O maior valor digitado foi {maior} nas posicões ', end='')
    for i, v in enumerate(listanum):
        if v == maior:
            print(f'{i}... ', end='')
    print()
    print(f'O menor valor digitado foi {menor} nas posicões ', end='')
    for i, v in enumerate(listanum):
        if v == menor:
            print(f'{i}... ', end='')
    print()

