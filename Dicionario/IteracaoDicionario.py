import random
perguntas = { '1':{
                'questao':'quanto e 2+2?',
                'resposta' : {'A':'1', 'B': '2', 'C':'4', 'D': '6'},
                'gabarito': 'C'},
                '2':{
                'questao':'quanto e 2+1?',
                'resposta' : {'A':'3', 'B': '2', 'C':'4', 'D': '6'},
                'gabarito': 'A'},
                }

for perguntaKey, perguntaValue in perguntas.items():
    print(f'{perguntaKey}: {perguntaValue["questao"]}')
    
    print('Respostas')
    for respostaKey, respostaValue in perguntaValue["resposta"].items():
        print(f'[{respostaKey}]: {respostaValue}')
    
    print()