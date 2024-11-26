import os

dirpath = r'png'
for name in os.listdir(dirpath):
    print(name)
    if name[:4].isdigit():
        print(name)
        newName = name[:4] + '.png'
        oldpath = os.path.join(dirpath, name)
        newPath = os.path.join(dirpath, newName)
        os.rename(oldpath, newPath)