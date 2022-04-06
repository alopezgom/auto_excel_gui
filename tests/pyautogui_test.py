#!/usr/local/bin python3

import pyautogui as pai

import time

def launch_youtube():
  time.sleep(1)
  
  # execute firefox
  pai.press('enter')
  pai.hotkey('fn','a') # launch dock
  pai.write('firefox',interval=0.1) # go to firefox icon
  pai.press('enter')

  time.sleep(15)
  
  # enter to youtube and select the search text box
  pai.typewrite('youtube.com',interval=0.1)
  pai.press('enter')
  time.sleep(5)
  pai.moveTo(454, 168, 2, pai.easeOutQuad)
  pai.click(button='left')

  time.sleep(10)

  # search for song and play it
  pai.typewrite('leprous the sky is red',interval=0.15)
  pai.press('enter')
  time.sleep(5)
  pai.moveTo(699, 273, 2, pai.easeOutQuad)
  pai.click(button='left')

  time.sleep(15)

  # quit firefox
  pai.hotkey('command','q')
  pai.press('enter')
  

def main():
  launch_youtube()

if __name__ == '__main__':
  main()
