import pyautogui
import time
import openpyxl


workbook = openpyxl.load_workbook('E:\\gpa.xlsx')
time.sleep(1)
sheet = workbook["Sheet1"]


i = 30

while i <= 68:

    pyautogui.click(14, 231)
    pyautogui.hotkey('end')
    time.sleep(0.5)

    (x, y) = pyautogui.locateCenterOnScreen('E:\\과정.png')
    pyautogui.click(x+221, y)
    time.sleep(2)
    pyautogui.click(x+130, y)
    time.sleep(0.5)
    pyautogui.hotkey('enter')

    time.sleep(1)

    (x, y) = pyautogui.locateCenterOnScreen('E:\\전공명.png')
    pyautogui.click(x+130, y)
    time.sleep(1.5)
    pyautogui.hotkey('enter')

    time.sleep(0.5)

    (x, y) = pyautogui.locateCenterOnScreen('E:\\재수강여부.png')
    pyautogui.click(x+130, y)
    time.sleep(0.5)
    pyautogui.hotkey('down')
    time.sleep(0.5)
    pyautogui.hotkey('enter')

    time.sleep(0.5)

    # 학년도
    a = sheet[f'B{i}'].value
    (x, y) = pyautogui.locateCenterOnScreen('E:\\수강년도.png')
    pyautogui.click(x+130, y)
    time.sleep(0.5)
    b = 2018-a
    print(b)
    c = 1
    while c <= b:
        pyautogui.hotkey('down')
        c += 1
        time.sleep(0.2)
    pyautogui.hotkey('enter')
    학년도 = a
    # 학기
    (x, y) = pyautogui.locateCenterOnScreen('E:\\학기.png')
    pyautogui.click(x+130, y)
    time.sleep(0.5)
    if 'Fall' in sheet[f'C{i}'].value:
        pyautogui.hotkey('down')
    else:
        pass
    pyautogui.hotkey('enter')

    # 교양
    (x, y) = pyautogui.locateCenterOnScreen('E:\\과목유형.png')
    pyautogui.click(x+130, y)
    time.sleep(1.5)
    if sheet[f'D{i}'].value == 1:
        pyautogui.hotkey('down')
    else:
        pass
    pyautogui.hotkey('enter')
    time.sleep(0.5)

    # 재수강
    # sheet[f'E{i}'].value

    # 과목명
    a = sheet[f'F{i}'].value
    (x, y) = pyautogui.locateCenterOnScreen('E:\\과목명.png')
    pyautogui.click(x+130, y)
    time.sleep(0.5)
    pyautogui.typewrite(a)

    # 학점
    a = sheet[f'G{i}'].value
    (x, y) = pyautogui.locateCenterOnScreen('E:\\취득학점.png')
    pyautogui.click(x+130, y)
    time.sleep(0.5)
    b = 1
    while b < a:
        pyautogui.hotkey('down')
        b += 1
    pyautogui.hotkey('enter')

    # 등급
    (x, y) = pyautogui.locateCenterOnScreen('E:\\성적.png')
    pyautogui.click(x+130, y)
    time.sleep(1.5)
    x = sheet[f'H{i}'].value
    if x == 'aa':
        n = 0
    elif x == 'a':
        n = 1
    elif x == 'b':
        n = 4
    elif x == 'c':
        n = 7
    else:
        print('등급 error')

    b = 0
    while b < n:
        pyautogui.hotkey('down')
        b += 1
        time.sleep(0.5)
    pyautogui.hotkey('enter')
    time.sleep(0.5)
    # 추가누르기
    pyautogui.click(pyautogui.locateCenterOnScreen(
        "E:\\추가.png"))
    time.sleep(3)

    pyautogui.click(pyautogui.locateCenterOnScreen(
        "E:\\저장.png"))
    time.sleep(1)
    pyautogui.hotkey('enter')
    time.sleep(1)
    pyautogui.hotkey('enter')
    time.sleep(1)

    i += 1
