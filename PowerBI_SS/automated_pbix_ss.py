import os
import sys
import time
import win32com.client
import win32gui
import win32con
import subprocess
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.image import MIMEImage
from email.mime.text import MIMEText
from PIL import ImageGrab
import logging
from datetime import datetime
import pyautogui
import maiiler




class PowerBIScreenshotEmailer:
    def __init__(self, dashboard_path, email_config, log_path=None):
        self.dashboard_path = dashboard_path
        self.email_config = email_config
        
        logging.basicConfig(
            filename=log_path or os.path.join(os.getcwd(), 'powerbi_screenshot.log'),
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s: %(message)s'
        )

    def refresh_powerbi_data(self):
        """
        Simulate keyboard press Alt + H + R to refresh PowerBI data
        """
        try:
            print("Starting data refresh...")
            # Make sure PowerBI is in the foreground
            hwnd = win32gui.FindWindow(None, "Power BI Desktop")
            if hwnd:
                win32gui.SetForegroundWindow(hwnd)
                time.sleep(2)
                print("PowerBI window found and brought to foreground")
            else:
                print("PowerBI window not found")
            
            # Simulate pressing Alt+H (opens Home tab), then R (refresh)
            pyautogui.keyDown('alt')  # Press 'Alt' key
            pyautogui.keyUp('alt')
            time.sleep(0.5)  # Small delay
            pyautogui.press('h')  # Press 'H' to open the Home ribbon
            time.sleep(0.5)
            pyautogui.press('r')
            print("Refresh command sent")

            print("Waiting for refresh to complete (10 seconds)...")
            time.sleep(10)
            pyautogui.press('tab')  # Close any dialog box if it's open 

            logging.info("PowerBI data refreshed successfully")
            print("PowerBI data refreshed successfully")
            return True
        except Exception as e:
            logging.error(f"Error refreshing PowerBI data: {e}")
            print(f"Error refreshing PowerBI data: {e}")
            return False

    def switch_to_tab_by_click(self, x, y):
        """
        Click on a specific location to switch tabs.
        Coordinates must match the position of the desired tab.
        """
        try:
            hwnd = win32gui.FindWindow(None, "Power BI Desktop")
            if hwnd:
                win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)
                win32gui.SetForegroundWindow(hwnd)
                time.sleep(2)

            pyautogui.moveTo(x, y)
            pyautogui.click()
            logging.info(f"Switched tab by clicking at ({x}, {y})")
            print(f"Switched tab by clicking at ({x}, {y})")
            return True
        except Exception as e:
            logging.error(f"Error switching tab: {e}")
            print(f"Error switching tab: {e}")
            return False
        

    def save_the_bi(self):
        """
        Simulate keyboard press to save PowerBI data
        """
        try:
            print("Starting file save...")
            # Make sure PowerBI is in the foreground
            hwnd = win32gui.FindWindow(None, "Power BI Desktop")
            if hwnd:
                win32gui.SetForegroundWindow(hwnd)
                time.sleep(2)
                print("PowerBI window found and brought to foreground")
            else:
                print("PowerBI window not found")
            
            pyautogui.keyDown('alt')  # Press 'Alt' key
            pyautogui.keyUp('alt')
            time.sleep(0.5)

            # Simulate pressing Ctrl + S to save the Power BI dashboard
            pyautogui.hotkey('ctrl', 's')  # Simulates pressing 'Ctrl + S'
            time.sleep(1)
            print("Save command sent")

            # Simulate pressing Alt+1
            pyautogui.keyDown('alt')  # Press 'Alt' key
            pyautogui.keyUp('alt')
            time.sleep(0.5)  # Small delay
            pyautogui.press('1')
            time.sleep(0.5)

            pyautogui.hotkey('ctrl', 's')  # Simulates pressing 'Ctrl + S'
            print("Alt+1 command sent")

            logging.info("PowerBI data saved successfully")
            print("PowerBI data saved successfully")
            return True
        except Exception as e:
            logging.error(f"Error saving PowerBI data: {e}")
            print(f"Error saving PowerBI data: {e}")
            return False


    def open_and_refresh_powerbi(self):
        """
        Open PowerBI and attempt refresh using subprocess
        """
        try:
            print("Closing any existing PowerBI instances...")
            os.system('taskkill /F /IM "PBIDesktop.exe" 2>nul')
            time.sleep(2)
            
            print(f"Opening PowerBI file: {self.dashboard_path}")
            # Open PowerBI file
            subprocess.Popen([
                'start', 
                'powershell', 
                '-Command', 
                f'Start-Process "{self.dashboard_path}"'
            ], shell=True)
            
            # Wait for PowerBI to open
            print("Waiting for PowerBI to open (30 seconds)...")
            time.sleep(30)  # Increased wait time
            
            logging.info("PowerBI opened successfully")
            print("PowerBI opened successfully")
            return True
        
        except Exception as e:
            logging.error(f"PowerBI open error: {e}")
            print(f"PowerBI open error: {e}")
            return False

    def capture_dashboard_screenshot(self, page):
        try:
            print("Preparing to capture screenshot...")

            if page == 1:
                # Bring PowerBI to foreground
                hwnd = win32gui.FindWindow(None, "Power BI Desktop")
                if hwnd:
                    win32gui.SetForegroundWindow(hwnd)
                    time.sleep(5)
                    print("PowerBI window found and brought to foreground")
                else:
                    print("PowerBI window not found")

            if page == 1:
                screenshot_folder = r"D:\Utils_scripts\PowerBI_SS\screenshots"
                analyzer_ss_folder = r"D:\Utils_scripts\PowerBI_SS\analysis_ss"

            if page == 2:
                screenshot_folder = r"D:\Utils_scripts\PowerBI_SS\trend_screenshots"
                analyzer_ss_folder = r"D:\Utils_scripts\PowerBI_SS\analysis_ss"

            # Ensure screenshot directory exists
            os.makedirs(screenshot_folder, exist_ok=True)
            
            """print("Capturing screenshot...")
            region = (65, 200, 1421, 970)

            # Capture screen region
            screenshot = ImageGrab.grab(bbox=region)"""

            hwnd = win32gui.FindWindow(None, "Power BI Desktop")
            if hwnd:
                win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)
                win32gui.SetForegroundWindow(hwnd)
                time.sleep(5)

            # High-res screenshot of region
            if page == 2:
                screenshot = pyautogui.screenshot(region=(60, 200, 1360, 773))
            
            if page == 1:
                screenshot = pyautogui.screenshot(region=(65, 200, 1360, 773))

            # High-res screenshot of region
            a_screenshot = pyautogui.screenshot(region=(65, 200, 850, 773))

            # Save screenshot
            screenshot_path = os.path.join(
                screenshot_folder,
                f'powerbi_dashboard_{datetime.now().strftime("%Y%m%d_%H%M%S")}.png'
            )

            # Save screenshot
            if page == 1:
                analysis_screenshot_path = os.path.join(
                    analyzer_ss_folder,
                    f'powerbi_dashboard_{datetime.now().strftime("%Y%m%d_%H%M%S")}.png'
                )

            screenshot.save(screenshot_path)
            
            if page == 1:
                a_screenshot.save(analysis_screenshot_path)
            
            logging.info(f"Screenshot saved: {screenshot_path}")
            print(f"Screenshot saved: {screenshot_path}")
            
            if page == 1:
                logging.info(f"Screenshot for analysis saved: {analysis_screenshot_path}")
                print(f"Screenshot for analysis saved: {analysis_screenshot_path}")

            return screenshot_path
        
        except Exception as e:
            logging.error(f"Screenshot capture error: {e}")
            print(f"Screenshot capture error: {e}")
            return None
    
    def send_screenshot_email(self, screenshot_path):
        """
        Send email with dashboard screenshot
        
        :param screenshot_path: Path to screenshot image
        """
        try:
            print("Preparing to send email...")
            # Create multipart message
            msg = MIMEMultipart()
            msg['From'] = self.email_config['sender_email']
            msg['To'] = self.email_config['recipient_email']
            msg['Subject'] = f"PowerBI Dashboard Screenshot - {datetime.now().strftime('%Y-%m-%d %H:%M')}"
            
            # Email body
            body = "Attached is the latest PowerBI dashboard screenshot."
            msg.attach(MIMEText(body, 'plain'))
            
            # Attach screenshot
            with open(screenshot_path, 'rb') as file:
                img = MIMEImage(file.read(), name=os.path.basename(screenshot_path))
                msg.attach(img)
            
            print("Connecting to SMTP server...")
            # SMTP setup
            server = smtplib.SMTP(
                self.email_config['smtp_server'], 
                self.email_config['smtp_port']
            )
            server.starttls()
            server.login(
                self.email_config['sender_email'], 
                self.email_config['email_password']
            )
            
            print("Sending email...")
            server.send_message(msg)
            server.quit()
            
            logging.info("Screenshot email sent successfully")
            print("Screenshot email sent successfully")
            return True
        
        except Exception as e:
            logging.error(f"Email sending error: {e}")
            print(f"Email sending error: {e}")
            return False
    
    def run_screenshot_workflow(self):
        """Complete screenshot and email workflow"""
        try:
            print("Starting PowerBI Screenshot Workflow...")
            
            # Open and refresh PowerBI
            if not self.open_and_refresh_powerbi():
                print("Failed to open PowerBI. Exiting workflow.")
                return None
            
            # Refresh data using keyboard shortcuts
            if not self.refresh_powerbi_data():
                print("Failed to refresh PowerBI data. Continuing workflow.")

            time.sleep(2)
            if not self.save_the_bi():
                print("Failed to save PowerBI file. Continuing workflow.")
            
            # Capture screenshot
            screenshot_paths = []

            screenshot_path = self.capture_dashboard_screenshot(1)
            screenshot_paths.append(screenshot_path)
            if not screenshot_path:
                print("Failed to capture screenshot. Exiting workflow.")
                return None
            
            page2 = self.switch_to_tab_by_click(474, 1033)
            time.sleep(2)

            # Capture screenshot
            if page2:
                print("Toggled to the second (trend) page")
                screenshot_path = self.capture_dashboard_screenshot(2)
                screenshot_paths.append(screenshot_path)
                if not screenshot_path:
                    print("Failed to capture screenshot. Exiting workflow.")
                    return None
            
            page1 = self.switch_to_tab_by_click(276, 1022)
            time.sleep(2)

            if page1:
                print("Reverted back to the original sheet")
            
            result = maiiler.sender(screenshot_paths)
            print("Email sending result:", "Success" if result else "Failed")

            print(f"Workflow completed successfully. Screenshot saved at: {screenshot_path}")
            return screenshot_path
            
            
        
        except Exception as e:
            logging.error(f"Workflow error: {e}")
            print(f"Workflow error: {e}")
            return None
        
        finally:
            print("Closing PowerBI...")
            # Close PowerBI silently
            os.system('taskkill /F /IM "PBIDesktop.exe" 2>nul')
            print("PowerBI closed")


def main():
    print("Starting PowerBI Screenshot Script...")
    
    # Email and dashboard configuration
    email_config = {
        'sender_email': 'your_email@example.com',
        'recipient_email': 'recipient@example.com',
        'smtp_server': 'smtp.gmail.com',
        'smtp_port': 587,
        'email_password': 'your_app_password'
    }
    
    screenshot_emailer = PowerBIScreenshotEmailer(
        dashboard_path=r'C:\Users\Vserve-User\OneDrive - Building Controls and Solutions, LLC\Documents\Report stats.pbix',
        email_config=email_config,
        log_path=r'D:\utils_scripts\logger.log'
    )
    
    # Run the workflow directly (no background process)
    saved_path = screenshot_emailer.run_screenshot_workflow()
    
    print("Script execution completed!")
    if saved_path:
        print(f"Final screenshot location: {saved_path}")
    else:
        print("No screenshot was captured.")


if __name__ == "__main__":
    main()