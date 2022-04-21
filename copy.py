import win32clipboard

# set clipboard data
def set():
  win32clipboard.OpenClipboard()
  win32clipboard.EmptyClipboard()
  win32clipboard.SetClipboardText('testing 123')
  win32clipboard.CloseClipboard()

# get clipboard data
def get():
  win32clipboard.OpenClipboard()
  data = win32clipboard.GetClipboardData()
  win32clipboard.CloseClipboard()
  print(data)