import win32com.client
import time

PPT = win32com.client.Dispatch('PowerPoint.Application')
PPT.Visible = True
PPT.Presentations.Open()

time.sleep(1)
PPT.quit()
