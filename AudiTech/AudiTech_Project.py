import requests
import time
import os, sys, xlwt, os.path
import matplotlib.pyplot as plt
import pandas as pd
from xlwt import Workbook
from matplotlib.patches import Polygon
from PIL import Image
from io import BytesIO
import xlrd

# Add your Computer Vision subscription key and endpoint to your environment variables.
if 'COMPUTER_VISION_SUBSCRIPTION_KEY' in os.environ:
    subscription_key = os.environ['COMPUTER_VISION_SUBSCRIPTION_KEY']
else:
    print("\nSet the COMPUTER_VISION_SUBSCRIPTION_KEY environment variable.\n**Restart your shell or IDE for changes to take effect.**")
    sys.exit()

if 'COMPUTER_VISION_ENDPOINT' in os.environ:
    endpoint = os.environ['COMPUTER_VISION_ENDPOINT']

text_recognition_url = endpoint + "vision/v2.1/read/core/asyncBatchAnalyze"

q=0
dir_name = "C:\\Users\\shubh\\Pictures\\AudiTech_HackNJIT\\Receipt_Images\\"
# Set image_url to the URL of an image that you want to analyze.
file_count = os.listdir(dir_name)
wb = Workbook()
sheet1 = wb.add_sheet('Expenses Sheet')
sheet1.write(0,0,"Sr. No")
sheet1.write(0,1,"Name")
sheet1.write(0,2,"Date")
sheet1.write(0,3,"Total Amount")
for i in range(1, len(file_count)+1):
    image_path = "C:\\Users\\shubh\\Pictures\\AudiTech_HackNJIT\\Receipt_Images\\"+ str(i) +".jpg"

    time.sleep(2)
    headers = {'Ocp-Apim-Subscription-Key': subscription_key, 'Content-Type': 'application/octet-stream'}
    image_data = open(image_path, "rb").read()
    params = {'visualFeatures': 'Categories,Description,Color'}
    response = requests.post(text_recognition_url, headers=headers, params=params, data=image_data)
    response.raise_for_status()

    # Holds the URI used to retrieve the recognized text.
    operation_url = response.headers["Operation-Location"]

    # The recognized text isn't immediately available, so poll to wait for completion.
    analysis = {}
    poll = True
    while (poll):
        response_final = requests.get(
        response.headers["Operation-Location"], headers=headers)
        analysis = response_final.json()
        time.sleep(1)
        if ("recognitionResults" in analysis):
            poll = False
        if ("status" in analysis and analysis['status'] == 'Failed'):
            poll = False

    polygons = []
    if ("recognitionResults" in analysis):
        # Extract the recognized text, with bounding boxes.
        polygons = [(line["boundingBox"], line["text"])
                    for line in analysis["recognitionResults"][0]["lines"]]

    # Display the image and overlay it with the extracted text.
    plt.figure(figsize=(15, 15))
    image = Image.open(BytesIO(image_data))
    ax = plt.imshow(image)
    result = " "
    arr = []
    for polygon in polygons:
        vertices = [(polygon[0][i], polygon[0][i+1])
                    for i in range(0, len(polygon[0]), 2)]
        text = polygon[1]
        result = result + text + " "
        patch = Polygon(vertices, closed=True, fill=False, linewidth=2, color='y')
        ax.axes.add_patch(patch)
        plt.text(vertices[0][0], vertices[0][1], text, fontsize=20, va="top")
    array = result.split()

    z = array[0]
    for i in range(1,5):
        try:
            l = int(array[i])
            break
        except ValueError:
            z = z + " " + array[i]
        except TypeError:
            z = z + " " + array[i]
    arr.append(z)
    counter = 0
    for i in range(len(array)):
        if (array[i].upper() == "TOTAL" or array[i].upper() == "SUM" or array[i].upper() == "AMOUNT" or array[i].upper() == "TOTAL AMOUNT" or array[i].upper() == "SUM TOTAL" or array[i].upper() == "GRAND TOTAL"):
            if (array[i-1].upper() != "SUB"):
                if (type(array[i+1]) == str):
                    if '$' in array[i+1]:
                        if '.' in array[i+2]:
                            arr.append(array[i])
                            value = array[i+1] + array[i+2]
                            arr.append(value)
                        else:
                            arr.append(array[i])
                            arr.append(array[i+1])
                    elif '.' in array[i+1]:
                        arr.append(array[i])
                        value = array[i+1] + array[i+2]
                        arr.append(value)
                    elif ':' in array[i+1]:
                        if '$' in array[i+2]:
                            if '.' in array[i+3]:
                                arr.append(array[i])
                                value = array[i+2] + array[i+3]
                                arr.append(value)
                            else:
                                arr.append(array[i])
                                arr.append(array[i+2])
                        elif '.' in array[i+2]:
                            arr.append(array[i])
                            value = array[i+2] + array[i+3]
                            arr.append(value)
        if (array[i].upper() == "TOTAL" or array[i].upper() == "SUM" or array[i].upper() == "GRAND"):
            if (array[i+1].upper() == "AMOUNT" or array[i+1].upper() == "TOTAL"):
                if (array[i-1].upper() != "SUB"):
                    if (type(array[i+2]) == str):
                        if '$' in array[i+2]:
                            if '.' in array[i+3]:
                                arr.append(array[i] + " " + array[i+1])
                                value = array[i+2] + array[i+3]
                                arr.append(value)
                            else:
                                arr.append(array[i] + " " + array[i+1])
                                arr.append(array[i+2])
                        elif '.' in array[i+2]:
                            arr.append(array[i] + " " + array[i+1])
                            value = array[i+2] + array[i+3]
                            arr.append(value)
                        elif ':' in array[i+2]:
                            if '$' in array[i+3]:
                                if '.' in array[i+4]:
                                    arr.append(array[i] + " " + array[i+2])
                                    value = array[i+3] + array[i+4]
                                    arr.append(value)
                                else:
                                    arr.append(array[i] + " " + array[i+2])
                                    arr.append(array[i+3])
                            elif '.' in array[i+3]:
                                arr.append(array[i] + " " + array[i+2])
                                value = array[i+3] + array[i+4]
                                arr.append(value)
        try:
            if(counter == 0):
                lt = array[i].index('/', 0)
                if (array[i].index('/', lt+1) == (lt + 3)):
                    arr.append(array[i])
                    counter = 1
        except ValueError:
             lt = 1
        try:
            if(counter == 0):
                lt = array[i].index('-', 0)
                if (array[i].index('-', lt+1) == (lt + 3)):
                    arr.append(array[i])
                    counter = 1
        except ValueError:
            lt = 1
        try:
            if(counter == 0):
                lt = array[i].index('.', 0)
                if (array[i].index('.', lt+1) == (lt + 3)):
                    arr.append(array[i])
                    counter = 1
        except ValueError:
            lt = 1    
    print(arr)
    w=0
    while (w<len(arr)):
        j = q
        sheet1.write(j+1, 0, q+1)
        sheet1.write(j+1, 1, arr[w])
        if ("/" in arr[w+1]):
                sheet1.write(j+1, 2, arr[w+1])
        if ("/" in arr[w+2]):
                sheet1.write(j+1, 2, arr[w+2])
        if ("/" in arr[w+3]):
                sheet1.write(j+1, 2, arr[w+3])
        if (("$" in arr[w+1]) or ("." in arr[w+1])):
                sheet1.write(j+1, 3, arr[w+1])
        if (("$" in arr[w+2]) or ("." in arr[w+2])):
                sheet1.write(j+1, 3, arr[w+2])
        if (("$" in arr[w+3]) or ("." in arr[w+3])):
                sheet1.write(j+1, 3, arr[w+3])
        w=w+4
        q=q+1
        wb.save('C:\\Users\\shubh\\Pictures\\AudiTech_HackNJIT\\Generated_Balance_Sheet.xls')
        
sheet1 = pd.read_excel(r'C:\\Users\\shubh\\Pictures\\AudiTech_HackNJIT\\Balance_Sheet.xls') 
sheet2 = pd.read_excel(r'C:\\Users\\shubh\\Pictures\\AudiTech_HackNJIT\\Generated_Balance_Sheet.xls') 
# Iterating the Columns Names of both Sheets 
for i,j in zip(sheet1,sheet2): 
     
    # Creating empty lists to append the columns values     
    a,b =[],[] 
  
    # Iterating the columns values 
    for m, n in zip(sheet1[i],sheet2[j]): 
  
        # Appending values in lists 
        a.append(m) 
        b.append(n) 
  
    # Sorting the lists 
    a.sort() 
    b.sort() 
    counter = True
    # Iterating the list's values and comparing them 
    for m, n in zip(range(len(a)), range(len(b))): 
        if a[m] != b[n]: 
            print(str(a[m]) + " is not present in Balance Sheet from Row Number: " + str(m))
            counter = False
if (counter == True):
    print("Your Balance Sheet is Correct")