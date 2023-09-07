from spire.presentation import *

inputFile1 = "./Data/OperatePlaceholders.pptx"
inputFile2 = "./Data/Video.mp4"
inputFile3 = "./Data/E-iceblueLogo.png"
outputFile = "OperatePlaceholders_out.pptx"

#Create a PPT document
presentation = Presentation()
#Load the document from disk
presentation.LoadFromFile(inputFile1)
#Operate placeholders
for j, unusedItem in enumerate(presentation.Slides):
    slide = presentation.Slides[j]
    for i, unusedItem in enumerate(slide.Shapes):
        shape = slide.Shapes[i]
        if shape.Placeholder.Type == PlaceholderType.Media:
            shape.InsertVideo(inputFile2)
        elif shape.Placeholder.Type == PlaceholderType.Picture:
            shape.InsertPicture(inputFile3)
        elif shape.Placeholder.Type == PlaceholderType.Chart:
            shape.InsertChart(ChartType.ColumnClustered)
        elif shape.Placeholder.Type == PlaceholderType.Table:
            shape.InsertTable(3, 2)
        elif shape.Placeholder.Type == PlaceholderType.Diagram:
            shape.InsertSmartArt(SmartArtLayoutType.BasicBlockList)
#Save the document
presentation.SaveToFile(outputFile, FileFormat.Pptx2013)
presentation.Dispose()