try:
    a = int(4)
    b = int()

    r = a / b
except (ValueError, TypeError, NameError):
    print('Digite um numero')
except (ZeroDivisionError):
    print('Não é possivel dividir por zero')
else:
    print(f'O RESULTADO é {r:.1f}')
finally:
    print('FIM')
