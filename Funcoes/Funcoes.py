def linhas(titulo):
    print('-'*30)
    print(titulo)
    print('-'*30)


linhas('notas')
linhas('alunos')


def contador(i,f,p):
    c = i
    while c <= f:
        print(c, end='\n')
        print()
        c +=p
    print('fim!')

contador(2,10,2)
print('tese')
