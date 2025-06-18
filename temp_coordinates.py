import pyautogui
import time

print("Move your mouse to the desired tab in Power BI...")
time.sleep(5)  # Gives you 5 seconds to switch to Power BI

# Shows the mouse position every second for 10 seconds
for i in range(10):
    x, y = pyautogui.position()
    print(f"Position {i+1}: X={x}, Y={y}")
    time.sleep(1)
