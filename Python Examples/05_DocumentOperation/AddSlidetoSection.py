from spire.presentation import *

inputFile = "./Data/Section.pptx"
outputFile = "AddSlidetoSection.pptx"

#Create a PPT document
presentation = Presentation()
presentation.LoadFromFile(inputFile)
#Add a new shape to the PPT document
presentation.Slides[0].Shapes.AppendShape(ShapeType.Rectangle, RectangleF.FromLTRB (200, 50, 500, 150))
#Create a new section and copy the first slide to it
NewSection = presentation.SectionList.Append("New Section")
NewSection.Insert(0, presentation.Slides[0])
presentation.SaveToFile(outputFile, FileFormat.Pptx2013)
presentation.Dispose()