import cv2
import numpy as np
from pyimagesearch.shapedetector import ShapeDetector
from pyimagesearch.colorlabeler import ColorLabeler
import imutils
import array
import xlwt


workbook = xlwt.Workbook()
sheet = workbook.add_sheet("Notes QCM")
style = xlwt.easyxf('font: bold 1')
sheet.write(0, 0, 'Les Etudiants', style)
sheet.write(0, 1, 'Notes', style)

def alph(i):
    switcher = {
        0: 'a',
        1: 'b',
        2: 'c',
        3: 'd',
    }
    return switcher.get(i)
def tonum(i):
    switcher = {
        'a': 1,
        'b': 2,
        'c': 3,
        'd': 0,
    }
    return switcher.get(i)

#arrrep list dial correct answers
arrrep = ['a','d','b','a','a']

arr= array.array('i', [])

#Here u add Exam image that you want to correct
image = cv2.imread('part1c.jpg')


gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
blur = cv2.medianBlur(gray, 5)
sharpen_kernel = np.array([[-1,-1,-1], [-1,10,-1], [-1,-1,-1]])
sharpen = cv2.filter2D(blur, -1, sharpen_kernel)

thresh = cv2.threshold(sharpen,160,250, cv2.THRESH_BINARY_INV)[1]
kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (3,3))
close = cv2.morphologyEx(thresh, cv2.MORPH_CLOSE, kernel, iterations=2)

cnts = cv2.findContours(close, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
cnts = cnts[0] if len(cnts) == 2 else cnts[1]

min_area = 1800
max_area = 1900
image_number = 0
answers=0
i=0
image1 = image
for c in cnts:
    area = cv2.contourArea(c)
    if area > min_area and area < max_area:
        x,y,w,h = cv2.boundingRect(c)
        ROI = image[y:y+h, x:x+h]
        answers += 1
for c in reversed(cnts):
    area = cv2.contourArea(c)
    if area > min_area and area < max_area:
        x, y, w, h = cv2.boundingRect(c)
        ROI = image[y:y + h, x:x + h]
        cv2.rectangle(image1, (x, y), (x + w, y + h), (255, 255, 255), 2)
        image_number += 1
        resized = imutils.resize(ROI, width=300)
        ratio = image.shape[0] / float(resized.shape[0])

        # Blurring
        blurred = cv2.GaussianBlur(resized, (5, 5), 0)
        gray = cv2.cvtColor(blurred, cv2.COLOR_BGR2GRAY)
        lab = cv2.cvtColor(blurred, cv2.COLOR_BGR2LAB)
        thresh = cv2.threshold(gray, 25, 255, cv2.THRESH_BINARY)[1]
        #cv2.imshow("Thresh", thresh)
        #cv2.waitKey(0)

        # n9elbo 3la l countours o ndiro thresh
        cnts = cv2.findContours(thresh.copy(), cv2.RETR_EXTERNAL,cv2.CHAIN_APPROX_SIMPLE)
        cnts = imutils.grab_contours(cnts)
        cl = ColorLabeler()
        # loop over the contours
        for c in cnts:
            # compute the center of the contour
            M = cv2.moments(c)
            # detect the shape of the contour and label the color
            color = cl.label(lab, c)
            # color on the image
            text = "{}".format(color)

            if (text == 'white'):
                arr.append(0)

            else:
                if(tonum(arrrep[i])==image_number%4):
                    allo = cv2.rectangle(image, (x, y), (x + w, y + h), (36, 255, 12), 3)
                else :
                    allo = cv2.rectangle(image, (x, y), (x + w, y + h), (12, 36, 255), 3)
                i += 1
                arr.append(1)

cv2.imwrite('Correction/imagesol.jpg'.format(image_number), image)

j=0
i=0

#arrcorrection list implementation
arrc = []

for i in range(i,answers,4):
    itt=0
    for i in range (j,answers):
        if (arr[i]==1):
            char = alph(itt)
            arrc.append(char)
            itt+=1
            j+=1
            if(itt==4):
                break
        else:
            itt+=1
            j += 1
            if (itt == 4):
                break


# easy stuff here
note=0
for i in range(0,len(arrc)):
    print("la reponse de la question selectionné par l etudiant ",i+1,"est : ",arrc[i])
print("\r")

for i in range(0,len(arrc)):
    if(arrc[i] == arrrep[i]):
       note +=4


print("la note de l'etudiant est : ",note)
sheet.write(1, 0, 'Etudiant 1')
sheet.write(1, 1, note)
workbook.save("Liste Etudiants/sample.xls")
