from spire.presentation.common import *
from spire.presentation import *


inputImageFile="./Data/Logo.png"
outputFile ="InsertHtmlWithImage.pptx"
#Create an instance of presentation document
ppt = Presentation()
shapes = ppt.Slides[0].Shapes

shapes.AddFromHtml("<html><div><p>First paragraph</p><p><img src='"+inputImageFile+"'/></p><p>Second paragraph </p></html>")

#Save the document
ppt.SaveToFile(outputFile, FileFormat.Pptx2013)
ppt.Dispose()
