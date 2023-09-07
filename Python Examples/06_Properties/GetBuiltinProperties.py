from spire.presentation import *

inputFile = "./Data/GetProperties.pptx"
outputFile = "GetBuiltinProperties.txt"

#Create PPT document
presentation = Presentation()
#Load the PPT document from disk
presentation.LoadFromFile(inputFile)
#Get the builtin properties 
application = presentation.DocumentProperty.Application
author = presentation.DocumentProperty.Author
company = presentation.DocumentProperty.Company
keywords = presentation.DocumentProperty.Keywords
comments = presentation.DocumentProperty.Comments
category = presentation.DocumentProperty.Category
title = presentation.DocumentProperty.Title
subject = presentation.DocumentProperty.Subject
#Create StringBuilder to save 
content = []
content.append ("DocumentProperty.Application: " + application +"\r")
content.append("DocumentProperty.Author: " + author+"\r")
content.append("DocumentProperty.Company " + company+"\r")
content.append("DocumentProperty.Keywords: " + keywords+"\r")
content.append("DocumentProperty.Comments: " + comments+"\r")
content.append("DocumentProperty.Category: " + category+"\r")
content.append("DocumentProperty.Title: " + title+"\r")
content.append("DocumentProperty.Subject: " + subject+"\r")
#Save them to a txt file
#Save the result file
f2=open(outputFile,'w', encoding='UTF-8')
for item in content:
        f2.write(item)
f2.close()
presentation.Dispose()