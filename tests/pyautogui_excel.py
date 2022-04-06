#!/usr/local/bin python3

import pyautogui as pai
from pynput.mouse import Button, Controller

import time

def excel_process():
  filename='mock?data.xlsx'
  mouse=Controller()

  time.sleep(1)
  
  # execute finder
  pai.press('enter')
  pai.hotkey('fn','a') # launch dock
  pai.write('finder',interval=0.1)
  pai.press('enter')

  time.sleep(6)
  
  # open excel file
  pai.hotkey('command','f') # launch dock
  time.sleep(1)
  pai.typewrite(filename,interval=0.15)
  pai.moveTo(248, 218, 2, pai.easeOutQuad)
  time.sleep(0.1)
  mouse.click(Button.left,2)

  time.sleep(15)
  
  # select table data
  pai.moveTo(421, 214, 2, pai.easeOutQuad)
  mouse.click(Button.left)  
  with pai.hold('command'):
    pai.moveTo(576, 216, 2, pai.easeOutQuad) 
    mouse.click(Button.left)

  time.sleep(2)

  # create graph
  pai.moveTo(107, 67, 2, pai.easeOutQuad)
  mouse.click(Button.left)
  pai.moveTo(693, 111, 2, pai.easeOutQuad)
  mouse.click(Button.left)
  pai.moveTo(823, 205, 2, pai.easeOutQuad)
  mouse.click(Button.left)

  time.sleep(2) 

def main():
  excel_process()

if __name__ == '__main__':
  main()
