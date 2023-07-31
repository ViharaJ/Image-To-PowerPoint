from pptx import Presentation
from pptx.util import Cm, Inches
from pptx.enum.shapes import MSO_SHAPE_TYPE
import pandas as pd 
import os 
import sys

def canPlaceSP(imgName, idx):
    imgName = imgName[:-4]
    comp = imgName.split('_')
    
    if(len(comp) < 4):
        return False
    
    if idx == 0 and comp[-1] == "seite":
        return True
    elif idx == 1 and comp[-1] == "Partikel":
        return True
    
    return False

def canPlaceHD(imgName, idx):
    imgName = imgName[:-4]
    comp = imgName.split('_')
    
    if(len(comp) < 4):
        return False
    
    if idx == 0 and comp[-1] == "dunkel":
        return True
    elif idx == 1 and comp[-1] == "hell":
        return True
    
    return False


path = "C:/Users/v.jayaweera/Documents/Tim/Slides/TestSmall.pptx"
imagePath = "C:/Users/v.jayaweera/Documents/Tim/Slides/20230607_Proben"
hellDunkelPath = "C:/Users/v.jayaweera/Documents/Tim/Slides/20230607_Proben im Pulverbett"
excelPath = "C:/Users/v.jayaweera/Documents/Tim/Microscope/20230607_Versuchsplan.xlsx"
acceptedFileTypes = ["jpg", "png", "bmp", "tif", "JPG"]

dirPictures = os.listdir(imagePath)
hdPictures = os.listdir(hellDunkelPath)
kvDict = dict()

for filename in dirPictures: 
    #Check if file is of accepted type
    if( '.' in filename and filename.split('.')[-1] in acceptedFileTypes):
        nameComp = filename[:-4].split('_')
        if(len(nameComp) > 1):
            Versuch = nameComp[0]
            Run = nameComp[1]
        
            val = kvDict.get(Versuch + '_' + Run)
            
            if val:
                val.append(filename)
                kvDict.update({Versuch + '_' + Run: val})
            else:
                kvDict[Versuch + '_' + Run] = [filename]
                
for filename in hdPictures: 
    #Check if file is of accepted type
    if( '.' in filename and filename.split('.')[-1] in acceptedFileTypes):
        nameComp = filename[:-4].split('_')
        if(len(nameComp) > 1):
            Versuch = nameComp[0]
            Run = nameComp[1]
        
            val = kvDict.get(Versuch + '_' + Run)
            
            if val:
                val.append(filename)
                kvDict.update({Versuch + '_' + Run: val})
            else:
                kvDict[Versuch + '_' + Run] = [filename]
            

df = pd.read_excel(excelPath)
prs = Presentation(path)

for i in range(1,len(prs.slides)-1,2):
    slide1 = prs.slides[i]
    slide2 = prs.slides[i+1]
    title = slide1.shapes.title.text
    
    subDf = df.loc[df['Simulate'] == title]
    
    
    #Get unique runs, should return 3 items
    uniquePrefixes = (subDf.Versuch.astype(str) + '_' + subDf.Run.astype(str)).unique()
    
    #Slide 1
    left = 1
    Height = 6.38 
    Width = 10
    spacing = 1
    
    
    for j in range(len(uniquePrefixes)):
        #Images available to insert into slide
        imgSet = kvDict.get(uniquePrefixes[j])
        
              
        #if image set exists, insert
        if imgSet: 
            top = 3.6           
            for j in range(2):
                imgIdx = 0
                while(imgIdx < len(imgSet)):
                    currImage = imgSet[imgIdx]
                    if(canPlaceSP(currImage, j)):
                        slide1.shapes.add_shape(MSO_SHAPE_TYPE.PLACEHOLDER, left, top, Width, Height)
                        slide1.placeholders[-1].insert_picture(imagePath + '/'+currImage)
                        # slide1.shapes.add_picture(imagePath + '/'+currImage, Cm(left), Cm(top), height=Cm(Height))
                        print("Position ({}, {}) for image {}".format(left, top, currImage))
                        break
                    else:
                        #insert place holder?
                        slide1.shapes.add_shape(MSO_SHAPE_TYPE.PLACEHOLDER, left, top, Width, Height)
                        print("This image " + currImage + " doesn't go here")
                        imgIdx = imgIdx + 1 
                        
                    
                top = top + Height + 0.25
               

        left = left + Width + spacing
            
    #hell und dunkel
    left = 1
    for j in range(len(uniquePrefixes)):
        #Images available to insert into slide
        imgSet = kvDict.get(uniquePrefixes[j])
        
              
        #if image set exists, insert
        if imgSet: 
            top = 3.6           
            for j in range(2):
                imgIdx = 0
                while(imgIdx < len(imgSet)):
                    currImage = imgSet[imgIdx]
                    if(canPlaceHD(currImage, j)):
                        slide2.shapes.add_picture(hellDunkelPath + '/'+currImage, Cm(left), Cm(top), height=Cm(Height))
                        print("Position ({}, {}) for image {}".format(left, top, currImage))
                        break
                    else:
                        #insert place holder?
                        print("This image " + currImage + " doesn't go here")
                        imgIdx = imgIdx + 1 
                        
                    
                top = top + Height + 0.25
               

        left = left + Width + spacing
        
    

prs.save(path)