import pyautogui
import time

print("Mova o mouse para a posição desejada e aguarde 2 segundos...")
time.sleep(2)
posicao_atual = pyautogui.position()
print(f"A posição do mouse é: {posicao_atual}")