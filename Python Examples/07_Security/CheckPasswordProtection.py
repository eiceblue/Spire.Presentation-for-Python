from spire.presentation import *

inputFile = "./Data/Template_Ppt_4.pptx"
outputFile = "CheckPasswordProtection.txt"

#Create Presentation
presentation = Presentation()
#Check whether a PPT document is password protected
isProtected = presentation.IsPasswordProtected(inputFile)
strB =[]
outString = "The file is " + ("password " if isProtected else "not password ") + "protected!"
strB.append (outString)
#Save the file
f2=open(outputFile,'w', encoding='UTF-8')
for item in strB:
        f2.write(item)
f2.close()
presentation.Dispose()