from spire.presentation.common import *
from spire.presentation import *


inputFile = "./Data/ChartSample4.pptx"
outputFile = "FillPictureInChartMarker.pptx"

#Create PPT document and load file
ppt = Presentation()
ppt.LoadFromFile(inputFile)

#Get chart on the first slide
Chart = ppt.Slides[0].Shapes[0] if isinstance(ppt.Slides[0].Shapes[0], IChart) else None

#Load image file in ppt
stream = Stream("Data/Logo.png")
IImage = ppt.Images.AppendStream (stream)
stream.Close()
        
#Create a ChartDataPoint object and specify the index
dataPoint = ChartDataPoint(Chart.Series[0])
dataPoint.Index = 0

#Fill picture in marker
dataPoint.MarkerFill.Fill.FillType = FillFormatType.Picture
dataPoint.MarkerFill.Fill.PictureFill.Picture.EmbedImage = IImage

#Set marker size
dataPoint.MarkerSize = 20

#Add the data point in series
Chart.Series[0].DataPoints.Add(dataPoint)

#Save the document
ppt.SaveToFile(outputFile, FileFormat.Pptx2010)
ppt.Dispose()


