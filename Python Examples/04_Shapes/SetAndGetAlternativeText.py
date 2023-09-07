from spire.presentation import *

inputFile = "./Data/ShapeTemplate.pptx"
outputFile = "SetAlternativeText.pptx"
outputFile_txt = "GetAlternativeText.txt"

#Create an instance of presentation document
ppt = Presentation()
#Load file
ppt.LoadFromFile(inputFile)
#Get the first slide
slide = ppt.Slides[0]
#Set the alternative text (title and description)
slide.Shapes[0].AlternativeTitle = "Rectangle"
slide.Shapes[0].AlternativeText = "This is a Rectangle"
#Get the alternative text (title and description)
alternativeText = ""
title = slide.Shapes[0].AlternativeTitle
alternativeText += "Title: " + title + "\r\n"
description = slide.Shapes[0].AlternativeText
alternativeText += "Description: " + description
#Save the document
ppt.SaveToFile(outputFile, FileFormat.Pptx2013)
#Save the result file
f2=open(outputFile_txt,'w', encoding='UTF-8')
f2.write(alternativeText)
f2.close()
ppt.Dispose()