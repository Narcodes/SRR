from PIL import ImageGrab
import win32gui


def get_screenshot(tab_name):
    
    toplist, winlist = [], []
    def enum_cb(hwnd, results):
        winlist.append((hwnd, win32gui.GetWindowText(hwnd)))
    win32gui.EnumWindows(enum_cb, toplist)

    screen = [(hwnd, title) for hwnd, title in winlist if tab_name.lower() in title.lower()]
    # just grab the hwnd for first window matching firefox
    screen = screen[0]
    hwnd = screen[0]

    win32gui.SetForegroundWindow(hwnd)
    # if bbox_required:
    #     bbox = win32gui.GetWindowRect(hwnd)
    # else:
    #     bbox = False

    return ImageGrab.grab()

if __name__ == "__main__":
    img = get_screenshot('srr')
    img.save('Input_Parameters.png', 'PNG')
    # img.show()

    from openpyxl import Workbook, drawing

    wb = Workbook()
    ws = wb.active
    ws.title = 'Input'

    img = drawing.image.Image(r'C:\Users\h.fernandez.muriano\OneDrive - Accenture\Python Projects\2018_11 SRR-IC Tool\Input_Parameters.png')
    img.anchor = 'D1'
    
    img.width = 800
    img.height = 600
    
    ws.add_image(img)

    wb.save(r'C:\Users\h.fernandez.muriano\OneDrive - Accenture\Python Projects\2018_11 SRR-IC Tool\Test.xlsx')