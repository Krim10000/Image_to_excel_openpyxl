

#listado de las imagenes en el presente directorio
#0
print("Starting")


import os
from PIL import Image
import imageio
import numpy as np



#Search tif images
#1
print(" Searching tif images in the current directory")
Search_tif = [x for x in os.listdir() if x.endswith(".tif") ]  

#2
print("  "+  str(len(Search_tif))+" .tif images were found")


#Convierte las imagenes a png

for Itif in Search_tif:
    im = imageio.imread(Itif)
    img_uint8 = im.astype(np.uint8)
    imageio.imwrite(Itif[0:-4]+ ".png",img_uint8)
#3
print ( "   "+str(len(Search_tif)) + " images .tif were converted to .png" )
#4
print("    Searching images png & jpg in the current directory")
Search_images = [x for x in os.listdir() if x.endswith(".png") or x.endswith(".jpg")]# or x.endswith(".tif") ]  # search png, jpg  , to add a new format add" or x.endswith(".EXTENTION") " to the []

#5
print("     "+  str(len(Search_images))+" .png & jpg images were found")


#crea un archivo excel y adjunta las imagenes

from openpyxl import Workbook

from openpyxl.drawing.image import Image

wb = Workbook()
ws = wb.active
ws.title = "Images"
dest_filename = 'name and images.xlsx'

for R in range(1,1000):
    ws.row_dimensions[R].height  = 200

wcolA = 25
#6
print("      Writing excel")
i=0
for I in Search_images:
    i=i+1
    img = Image(I)
    img.width = 200
    img.height = 200
    Cell1 = ("A"+str(i))
    Cell2 = ("B"+str(i))
    ws[Cell1] = I
    
    ws.add_image(img, Cell2)





ws.column_dimensions["A"].width = wcolA
ws.column_dimensions["B"].width = wcolA
ws.column_dimensions["C"].width = wcolA
ws.column_dimensions["D"].width = wcolA
wb.save(filename = dest_filename)
#7
print("       File ready")
