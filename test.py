import platform
import win32api, win32con
import Test_Scripts.EasyShell_Lib as EasyshellLib
import Library.CommonLib as CommonLib
import os
import time
import traceback
import pysnooper
from PIL import ImageGrab, ImageDraw, ImageFont
import subprocess

import numpy as np
from PIL import Image

def test():
    image_file = 'son.png'
    height = 100

    img = Image.open(image_file)
    img_width, img_height = img.size
    width = 2 * height * img_width // img_height  # 假定字符的高度是宽度的2倍
    img = img.resize((width, height), Image.ANTIALIAS)
    pixels = np.array(img.convert('L'))
    print(pixels.shape)
    print(pixels)
    chars = "MNHQ$OC?7>!:-;. "
    N = len(chars)
    step = 256 // N
    print(N)
    result = ''
    for i in range(height):
        for j in range(width):
            result += chars[pixels[i][j] // step]
        result += '\n'
    with open('text.txt', mode='w') as f:
        f.write(result)
test()