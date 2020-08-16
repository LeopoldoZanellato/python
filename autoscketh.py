import pyautogui
from os import listdir
from os.path import isfile, join
import time

"""API criado para a automatização de tarefas com a utilização do AutoScetch. Mostra
a possibilidade em transfomar arquivos em PDF/DWG/DXF de uma forma automatizada.
Foi utilizadao o AutoScetch 10 versão trial"""

mypath = "C:/Users/lpl_z/desktop/desenhos"
onlyfiles = [f for f in listdir(mypath) if isfile(join(mypath, f))]
i = 0

dwg = []
skf = []
arquivos_faltantes = []
for n in onlyfiles:
    print(n)
    tipoarquivo = n[-3:].lower()
    if tipoarquivo == "dwg":
        dwg.append(n[:-4])
    elif tipoarquivo == "skf":
        skf.append(n[:-4])

for arquivo in skf:
    if arquivo not in dwg:
        arquivos_faltantes.append(arquivo + ".SKF")

print(f"Arquivos dwg: {dwg}")
print(f"Arquivos sqf: {skf}")
print(f"Arquivos faltantes: {arquivos_faltantes}")


for n in arquivos_faltantes:
    time.sleep(4)
    pyautogui.moveTo(40, 58, duration=0.5)  #clicar na pasta abrir
    pyautogui.click()
    time.sleep(3)
    pyautogui.write(arquivos_faltantes[i])  #escrever o arquivo da posição 0
    time.sleep(3)
    pyautogui.press('enter')  # abrir o arquivo


# salvando o arquivo como DXF

    pyautogui.moveTo(34, 30, duration=0.5)  #mover para file
    pyautogui.click()
    time.sleep(3)

    pyautogui.moveTo(91, 179)  #save as
    pyautogui.click()
    time.sleep(3)

    pyautogui.moveTo(1160, y=701, duration=0.5)
    pyautogui.click()
    time.sleep(3)

    pyautogui.moveTo(1047, 792, duration=0.5)  #tipo dwg ou dxf
    pyautogui.click()
    time.sleep(3)

    pyautogui.moveTo(1213, 671, duration=0.5)  #posição save
    pyautogui.click()

    time.sleep(3)
    pyautogui.moveTo(1909, y=31, duration=0.5)  #posição close
    pyautogui.click()
    i = i + 1



