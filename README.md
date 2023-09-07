# Standalone PowerPoint Compatible Python API for Efficient Presentation Handling

[![Foo](https://i.imgur.com/dvNJk7D.png)](https://www.e-iceblue.com/Introduce/presentation-for-python.html)

[Product Page](https://www.e-iceblue.com/Introduce/presentation-for-python.html) | Documentation | Examples | [Forum](https://www.e-iceblue.com/forum/spire-presentation-f14.html) | [Temporary License](https://www.e-iceblue.com/TemLicense.html) | [Customized Demo](https://www.e-iceblue.com/Misc/customized-demo.html)

[Spire.Presentation for Python](https://www.e-iceblue.com/Introduce/presentation-for-python.html) is a comprehensive PowerPoint compatible API designed for developers to efficiently create, modify, read, and convert PowerPoint files within Python programs. It offers a broad spectrum of functions to manipulate PowerPoint documents without any external dependencies.

[Spire.Presentation for Python](https://www.e-iceblue.com/Introduce/presentation-for-python.html) supports a wide range of PowerPoint features, such as adding and formatting text, tables, charts, images, shapes, and other objects, inserting and modifying animations, transitions, and slide layouts, generating and managing master slides, and many more.

This professional Python API also enables developers to easily convert PowerPoint files to various formats with high quality, including PDF, SVG, image, HTML, XPS, and more.

### Support for Various PowerPoint Versions
- PPT - PowerPoint Presentation 97-2003
- PPS - PowerPoint SlideShow 97-2003
- PPTX - PowerPoint Presentation 2007/2010/2013/2016/2019
- PPSX - PowerPoint SlideShow 2007, 2010

### High-Quality and Efficient PowerPoint File Conversion
Spire.Presentation for Python allows conversion from PowerPoint files to images, PDF, HTML, XPS, and SVG and interconversion between PowerPoint Presentation formats.

### Support for Rich Presentation Manipulation Features
- Work with PowerPoint Charts
- Print PowerPoint Presentations
- Work with SmartArtImages and Shapes
- Audio and Video
- Protect Presentation Slides
- Text and Image Watermark
- Merge Split PowerPoint Document
- Comments and Notes
- Manage PowerPoint Tables
- Set Animations on Shapes
- Manage Hyperlink
- Extract Text and Image
- Replace Text

## Examples

### Create a PowerPoint document in Python
```Python
from spire.presentation.common import *
import math
from spire.presentation import *


outputFile ="HelloWorld.pptx"

#Create a PPT document
presentation = Presentation()

#Add a new shape to the PPT document
left = math.trunc(presentation.SlideSize.Size.Width / float(2))-250
rec = RectangleF.FromLTRB(left, 80, left+500, 150+80)
shape = presentation.Slides[0].Shapes.AppendShape(ShapeType.Rectangle, rec)
shape.ShapeStyle.LineColor.Color = Color.get_White()
shape.Fill.FillType = FillFormatType.none

#Add text to the shape
shape.AppendTextFrame("Hello World!")

#Set the font and fill style of the text
textRange = shape.TextFrame.TextRange
textRange.Fill.FillType = FillFormatType.Solid
textRange.Fill.SolidColor.Color = Color.get_CadetBlue()
textRange.FontHeight = 66
textRange.LatinFont = TextFont("Lucida Sans Unicode")

#Save the document
presentation.SaveToFile(outputFile, FileFormat.Pptx2010)
presentation.Dispose()
```

### Convert PowerPoint files to PDF
```Python
from spire.presentation.common import *
from spire.presentation import *


inputFile ="./Data/ToPDF.pptx"
outputFile = "ToPDF.pdf"

#Create a PPT document
presentation = Presentation()

#Load PPT file from disk
presentation.LoadFromFile(inputFile)

#Save the PPT to PDF file format
presentation.SaveToFile(outputFile, FileFormat.PDF)
presentation.Dispose()
```

### Convert PowerPoint files to Images
```Python
from spire.presentation.common import *
from spire.presentation import *

inputFile ="./Data/ToImage.pptx"

#Create PPT document
presentation = Presentation()

#Load PPT file from disk
presentation.LoadFromFile(inputFile)

#Save PPT document to images
for i, slide in enumerate(presentation.Slides):
    fileName ="ToImage_img_"+str(i)+".png"
    image = slide.SaveAsImage()
    image.Save(fileName)
    image.Dispose()

presentation.Dispose()
```

### Set passwords for PowerPoint presentations
```Python
from spire.presentation import *

inputFile = "./Data/Encrypt.pptx"
outputFile = "Encrypt_out.pptx"

#Create a PPT document
presentation = Presentation()
#Load the document from disk
presentation.LoadFromFile(inputFile)
#Get the password that the user entered
password = "e-iceblue"
#Encrypy the document with the password
presentation.Encrypt(password)
#Save the document
presentation.SaveToFile(outputFile, FileFormat.Pptx2010)
presentation.Dispose()
```

[Product Page](https://www.e-iceblue.com/Introduce/presentation-for-python.html) | Documentation | Examples | [Forum](https://www.e-iceblue.com/forum/spire-presentation-f14.html) | [Temporary License](https://www.e-iceblue.com/TemLicense.html) | [Customized Demo](https://www.e-iceblue.com/Misc/customized-demo.html)
