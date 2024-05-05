import pyautogui

resp = 'N'

while resp not in 'S': 
    input('Posicione o Mouse e em seguida tecle ENTER: ')
     #pega o retorno da posicao atual de x e y do mouse e passa o valor da tupla para as duas
    x, y = pyautogui.position() 
    print( "Posicao atual do mouse:")
    print("x = "+str(x)+" y = "+str(y))
    resp = input('Deseja sair do programa ?:')[0].upper()
    if resp == 'S':
        break
