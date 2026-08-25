import pyautogui 

pyautogui.hotkey('win')
pyautogui.write('outlook', interval=0.25) 
pyautogui.keyDown('enter')
pyautogui.PAUSE = 2.5
pyautogui.hotkey('win', 'pageup')
pyautogui.PAUSE = 1.5
pyautogui.click(button='left')
pyautogui.hotkey('ctrl', 'e')
pyautogui.PAUSE = 0.5
pyautogui.write('declaracao', interval=0.25) 
pyautogui.PAUSE = 0.5
pyautogui.keyDown('enter')
pyautogui.keyDown('enter')