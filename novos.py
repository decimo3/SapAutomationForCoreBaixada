from os import listdir, makedirs
from os.path import isfile, join
import time
from win10toast import ToastNotifier
import shutil

def fileWatcher(watchDirectory: str, pollTime: int):
  shutil.rmtree(watchDirectory)
  makedirs(watchDirectory)
  toaster = ToastNotifier()
  while True:
    if 'watching' not in locals(): #Check if this is the first time the function has run
      previousFileList = fileInDirectory(watchDirectory)
      watching = 1
      print('First Time')
      print(previousFileList)
    time.sleep(pollTime)
    newFileList = fileInDirectory(watchDirectory)
    fileDiff = listComparison(previousFileList, newFileList)
    previousFileList = newFileList
    if len(fileDiff) == 0: continue
    toaster.show_toast(str(fileDiff))
def fileInDirectory(my_dir: str):
    onlyfiles = [f for f in listdir(my_dir) if isfile(join(my_dir, f))]
    return(onlyfiles)
def listComparison(OriginalList: list, NewList: list):
    differencesList = [x for x in NewList if x not in OriginalList]
    return(differencesList)

fileWatcher("C:\\Users\\ruan.camello\\Documents\\Temporario", 5)