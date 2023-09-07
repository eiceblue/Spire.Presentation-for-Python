from spire.presentation import *

inputFile = "./Data/AddSection.pptx"
outputFile = "GetSectionIndex.txt"

#Create a PPT document
ppt = Presentation()
ppt.LoadFromFile(inputFile)
section = ppt.SectionList[0]
#Get the index of the section
index = ppt.SectionList.IndexOf(section)
f2=open(outputFile,'w', encoding='UTF-8')
f2.write("index: "+str(index))
f2.close()
ppt.Dispose()