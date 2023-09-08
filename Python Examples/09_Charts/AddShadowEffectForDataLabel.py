from spire.presentation.common import *
from spire.presentation import *


inputFile = "./Data/Template_Ppt_3.pptx"
outputFile = "AddShadowEffectForDataLabel.pptx"

#Create a PowerPoint document.
presentation = Presentation()

#Load the file from disk.
presentation.LoadFromFile(inputFile)

#Get the chart.
chart = presentation.Slides[0].Shapes[0] if isinstance(presentation.Slides[0].Shapes[0], IChart) else None

#Add a data label to the first chart series.
dataLabels = chart.Series[0].DataLabels
Label = dataLabels.Add()
Label.LabelValueVisible = True

#Add outer shadow effect to the data label.
Label.Effect.OuterShadowEffect = OuterShadowEffect()

#Set shadow color.
Label.Effect.OuterShadowEffect.ColorFormat.Color = Color.get_Yellow()

#Set blur.
Label.Effect.OuterShadowEffect.BlurRadius = 5

#Set distance.
Label.Effect.OuterShadowEffect.Distance = 10

#Set angle.
Label.Effect.OuterShadowEffect.Direction = 90

#Save to file.
presentation.SaveToFile(outputFile, FileFormat.Pptx2013)
presentation.Dispose()