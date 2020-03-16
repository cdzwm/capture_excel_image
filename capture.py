import win32com.client
from PIL import Image, ImageGrab

xlapp = win32com.client.Dispatch("Excel.Application")
xlapp.Visible = 1
xlapp.Workbooks.Add()
xlapp.Workbooks.Open(Filename=r"test.xlsx")
xlapp.Selection.CopyPicture(Appearance=1, Format=2)
xlapp.Quit()
im=ImageGrab.grabclipboard()
im.save('1.jpg')
