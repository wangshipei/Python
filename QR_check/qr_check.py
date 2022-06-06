import re

import cv2
# from pylibdmtx.pylibdmtx import decode
import numpy as np
from pylibdmtx import pylibdmtx

img = cv2.imread('2.jpg')
det = cv2.QRCodeDetector()
st_code = det.detectAndDecode(img)
print(st_code[0])

image = cv2.imread('1.jpg', cv2.IMREAD_UNCHANGED)
gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
ret, thresh = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY | cv2.THRESH_OTSU)
msg = pylibdmtx.decode(thresh)
pttn = re.compile(r'Decoded\(data=b(.*)\\x1d91EE')
for i in msg:
    # print(i)
    j = str(i)
    k = ''.join(re.findall(pttn, j))
    k1 = k.replace('\"', '')
    k2 = k1.replace("\'", '')
    print(k2)

