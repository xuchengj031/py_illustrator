'''
去掉白边
按比例缩小并存为png文件
dependency: 系统安装有imagemagick，并加入环境变量Path
input: ROOT/raw/images/*.jpg,*.png
output: ROOT/src/images/*.png
'''
import os

ROOT = os.getcwd()
DIR_SRC = os.path.join(ROOT, "raw", "images")
DIR_DST = os.path.join(ROOT, "src", "images")
print(DIR_DST)
if not os.path.exists(DIR_DST):
    os.mkdir(DIR_DST)


def trim(img, w, h, base):
    dst = os.path.join(DIR_DST, base.split('.')[0].upper() + ".png")
    cmd = r'magick convert "{0}" -fuzz 5% -trim -resize {1}x{2} "{3}"'.format(
        img, w, h, dst)
    os.system(cmd)

for dirpath, dirnames, filenames in os.walk(DIR_SRC):
    for filepath in filenames:
        ext = os.path.splitext(filepath)[1]
        base = os.path.basename(filepath)
        img = os.path.join(dirpath, filepath)
        if ext == ".jpg" or ext == ".png":
            trim(img, 480, 560, base)
