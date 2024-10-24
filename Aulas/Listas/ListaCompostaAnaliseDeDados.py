pessoas = list()
while True:
    pessoas.append([input('Nome: '), float(input('Peso: '))])
    escolha = input('Quer continuar? [S/N] ').lower()[0]
    if escolha == 'n':
        break
print('-='*30)
print(f'Ao todo, você cadastrou {len(pessoas)} pessoas.')
Max = ['', 0]
for pessoa in pessoas:
    if Max[len(Max)-1] == 0 or pessoa[1] > Max[len(Max)-1]:
        Max = [pessoa[0], pessoa[1]]
    elif pessoa[1] == Max[len(Max)-1]:
        Max.insert(0, pessoa[0])
print(f'O maior peso foi de {Max[len(Max)-1]}Kg. Peso de ', end='')
for c in range(len(Max)-1):
    if c != len(Max)-2:
        print(Max[c], end=', ')
    else:
        print(Max[c])
Min = ['', 0]
for pessoa in pessoas:
    if Min[len(Min)-1] == 0 or pessoa[1] < Min[len(Min)-1]:
        Min = [pessoa[0], pessoa[1]]
    elif pessoa[1] == Min[len(Min)-1]:
        Min.insert(0, pessoa[0])
print(f'O menor peso foi de {Min[len(Min)-1]}Kg. Peso de ', end='')
for c in range(len(Min)-1):
    if c != len(Min)-2:
        print(Min[c], end=', ')
    else:
        print(Min[c])
