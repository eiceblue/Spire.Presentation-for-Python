from spire.presentation import *

inputFile = "./Data/IsTextboxSample.pptx"
outputFile = "IsTextBox.txt"

#Create an instance of presentation document
ppt = Presentation()
#Load file
ppt.LoadFromFile(inputFile)
sb =[]
for slide in ppt.Slides:
    for shape in slide.Shapes:
        if isinstance(shape, IAutoShape):
            #Judge if the shape is textbox
            isTextbox = shape.IsTextBox
            sb.append ("shape is text box \r" if isTextbox else "shape is not text box \r")
#Save the result file
f2=open(outputFile,'w', encoding='UTF-8')
for item in sb:
        f2.write(item)
f2.close()
ppt.Dispose()