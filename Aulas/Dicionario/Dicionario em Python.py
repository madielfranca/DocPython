aluno = dict()
aluno['Nome'] = str(input('Digite o nome do(a) aluno(a): '))
aluno['Média'] = float(input('Digite a média agora: '))

if aluno['Média'] >= 7:
    aluno['Situação'] = 'Aprovado'
    print(f'O aluno {aluno["Nome"]} está aprovado!')
elif aluno['Média'] >= 5 and aluno['Média'] < 7:
    aluno['Situação'] = 'Na recuperação'
    print(f'O aluno {aluno["Nome"]} está na recuperação!')
else:
    aluno['Situação'] = 'Reprovado'
    print(f'O aluno {aluno["Nome"]} está reprovado!')

print('-+' * 30)
print('Dados: ')
print(aluno)
for k, v in aluno.items():
    print(f'{k} é {v}')
print('-+' * 30)


