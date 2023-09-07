from spire.presentation import *

inputFile = "./Data/BlankSample.pptx"
outputFile = "AddSection.pptx"

#Create a PPT document
ppt = Presentation()
ppt.LoadFromFile(inputFile)
#Get the second slide
slide = ppt.Slides[1]
#Append section with section name at the end
ppt.SectionList.Append("E-iceblue01")
#Add section with slide
ppt.SectionList.Add("section1", slide)
#Save to file.
ppt.SaveToFile(outputFile, FileFormat.Pptx2013)
ppt.Dispose()
