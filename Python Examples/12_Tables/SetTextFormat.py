from spire.presentation.common import *
from spire.presentation import *


inputFile = "./Data/Table.pptx"
outputFile = "SetTextFormat.pptx"


#Create a PPT document
presentation = Presentation()

#Load PPT file from disk
presentation.LoadFromFile(inputFile)
#Get the first slide
slide = presentation.Slides[0]
strs = []
for shape in slide.Shapes:
    #Verify if it is table
    if isinstance(shape, ITable):
        table = shape

        cell1 = table.TableRows[0][0]
        #Set table cell's text alignment type 
        cell1.TextAnchorType = TextAnchorType.Top
        #Set italic style
        cell1.TextFrame.TextRange.Format.IsItalic = TriState.TTrue

        cell2 = table.TableRows[1][0]
        #Set table cell's foreground color
        cell2.TextFrame.TextRange.Fill.FillType = FillFormatType.Solid
        cell2.TextFrame.TextRange.Fill.SolidColor.Color = Color.get_Green()
        #Set table cell's background color
        cell2.FillFormat.FillType = FillFormatType.Solid
        cell2.FillFormat.SolidColor.Color = Color.get_LightGray()


        cell3 = table.TableRows[2][2]
        #Set table cell's font and font size
        cell3.TextFrame.TextRange.FontHeight = 12
        cell3.TextFrame.TextRange.LatinFont = TextFont("Arial Black")
        cell3.TextFrame.TextRange.HighlightColor.Color = Color.get_YellowGreen()


        cell4 = table.TableRows[2][1]
        #Set table cell's margin and borders
        cell4.MarginLeft = 20
        cell4.MarginTop = 30
        cell4.BorderTop.FillType = FillFormatType.Solid
        cell4.BorderTop.SolidFillColor.Color = Color.get_Red()
        cell4.BorderBottom.FillType = FillFormatType.Solid
        cell4.BorderBottom.SolidFillColor.Color = Color.get_Red()
        cell4.BorderLeft.FillType =FillFormatType.Solid
        cell4.BorderLeft.SolidFillColor.Color = Color.get_Red()
        cell4.BorderRight.FillType = FillFormatType.Solid
        cell4.BorderRight.SolidFillColor.Color = Color.get_Red()

presentation.SaveToFile(outputFile, FileFormat.Pptx2013)
presentation.Dispose()



    
