from spire.presentation.common import *
from spire.presentation import *


outputFile = "SetMasterBackground.pptx"

#Create a PPT document
presentation = Presentation()

#Set the slide background of master
presentation.Masters[0].SlideBackground.Type = BackgroundType.Custom
presentation.Masters[0].SlideBackground.Fill.FillType =FillFormatType.Solid
presentation.Masters[0].SlideBackground.Fill.SolidColor.Color = Color.get_LightSalmon()

#Save the document
presentation.SaveToFile(outputFile, FileFormat.Pptx2013)
presentation.Dispose()

