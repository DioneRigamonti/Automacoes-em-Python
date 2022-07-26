from time import sleep
import pyautogui as pag
import pyperclip as ppc

#Projeto de automação para envio de um e-mail no outlook do browser.
# a posição x , y devem ser encontradas com o comando pag.position(x= , y= )
#ATENÇÃO: as posições x , y apresentadas nesse projeto funcionam com base no tamanho do meu monitor , podendo não funcionar o código corretamente na sua máquina

pag.PAUSE = 1
pag.press('win') # Pressiona a tecla windows.
sleep(0.5)

pag.write('Edge') # Escreve o nome do programa que deseja abrir.

pag.press('enter') # Pressiona enter.
sleep(1)

ppc.copy('https://outlook.live.com/owa/') # Copia o link que deseja acessar.

pag.hotkey('ctrl','v')# Cola o link copiado.

pag.press('enter')#//

pag.click(x=1502, y=165) # Clica na posição determinada por você.
sleep(2)

pag.press('esc') # Pressiona Esc.
sleep(0.5)

pag.write('SEU EMAIL@hotmail.com') # Escreve o seu E-mail .
sleep(0.5)

pag.click(x=1121, y=663)#//
sleep(0.5)

pag.write('SUA SENHA') # Escreve sua senha ATENÇÃO: não publique o código com sua senha.

pag.click(x=1111, y=721)#//
sleep(2)

pag.click(x=276, y=215)#//
sleep(1)

pag.write('EMAIL DO DESTINATARIO@hotmail.com') # Escreve o E-mail do destinatário.
sleep(0.5)

pag.press('tab') # Pressiona Tab.

pag.press('tab') #//

ppc.copy('Olá este é um teste de envio de e-mail.') # Copia o título da mensagem.

pag.hotkey('ctrl','v') # Cola o título.
sleep(0.5)

pag.press('tab') #//
sleep(0.5)

ppc.copy('''Olá,
Estou enviando uma mensagem com automação em Python.
isso é muito legal!!!

A sorte é a desculpa dos fracassados!
''') # --- Copia a mensagem de texto que você quer enviar.

pag.hotkey('ctrl','v') # Cola a mensagem de texto.
sleep(1)

pag.hotkey('ctrl','enter') # Envia a mensagem.

print('Fim da automação!!!!')