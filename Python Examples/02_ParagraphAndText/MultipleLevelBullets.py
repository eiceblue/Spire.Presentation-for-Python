from spire.presentation.common import *
from spire.presentation import *


inputFile ="./Data/Bullets2.pptx"
outputFile ="MultipleLevelBullets.pptx"

#Create a PPT document
presentation = Presentation()

#Load PPT file from disk
presentation.LoadFromFile(inputFile)
#Get the first slide
slide = presentation.Slides[0]

#Access the first placeholder in the slide and typecasting it as AutoShape
tf1 = (slide.Shapes[1]).TextFrame

#Access the first Paragraph and set bullet style
para = tf1.Paragraphs[0]
para.BulletType = TextBulletType.Symbol
para.BulletChar = 8226
para.Depth = 0

#Access the second Paragraph and set bullet style
para = tf1.Paragraphs[1]
para.BulletType = TextBulletType.Symbol
para.BulletChar = 45
para.Depth = 1

#Access the third Paragraph and set bullet style
para = tf1.Paragraphs[2]
para.BulletType = TextBulletType.Symbol
para.BulletChar =8226
para.Depth = 2

#Access the fourth Paragraph and set bullet style
para = tf1.Paragraphs[3]
para.BulletType = TextBulletType.Symbol
para.BulletChar = 45
para.Depth = 3

presentation.SaveToFile(outputFile, FileFormat.Pptx2013)
presentation.Dispose()