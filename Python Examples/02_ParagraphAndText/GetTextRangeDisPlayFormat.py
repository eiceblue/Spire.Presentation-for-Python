from spire.presentation.common import *
from spire.presentation import *

# Create a Presentation object
presentation = Presentation()
presentation.LoadFromFile("ShapeToImage.pptx")
#Get the first shape
shape = presentation.Slides[0].Shapes[0]
#Retrieve the text area in the first paragraph of the shape
textrange = shape.TextFrame.Paragraphs[0].TextRanges[0]
#Obtain the DisplayFormat property of the text area
displayformat = textrange.DisplayFormat

sb = []
sb.append(f"text ：{textrange.Text}")
sb.append(f"is bold ：{displayformat.IsBold}")
sb.append(f"is italic ：{displayformat.IsItalic}")
sb.append(f"latin_font FontName = ：{displayformat.LatinFont.FontName}")

with open("out.txt", 'a', encoding='utf-8') as file:
    for line in sb:
        file.write(line + '\n')

presentation.Dispose()