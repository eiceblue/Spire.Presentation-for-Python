# Python代码核心功能提取

# Spire.Presentation Python Hello World
## Create a simple presentation with "Hello World" text
```python
import math
from spire.presentation import *

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
```

---

# Spire.Presentation Python Paragraph
## Add and format paragraph in PowerPoint presentation
```python
#Create an instance of presentation document
ppt = Presentation()

#Append a new shape
shape = ppt.Slides[0].Shapes.AppendShape(ShapeType.Rectangle, RectangleF.FromLTRB (50, 70, 670, 220))
shape.Fill.FillType = FillFormatType.none
shape.ShapeStyle.LineColor.Color = Color.get_White()

#Set the alignment of paragraph
shape.TextFrame.Paragraphs[0].Alignment = TextAlignmentType.Left
#Set the indent of paragraph
shape.TextFrame.Paragraphs[0].Indent = 50
#Set the linespacing of paragraph
shape.TextFrame.Paragraphs[0].LineSpacing = 150
#Set the text of paragraph
shape.TextFrame.Text = "This powerful component suite contains the most up-to-date versions of all python components offered by E-iceblue."

#Set the Font
shape.TextFrame.Paragraphs[0].TextRanges[0].LatinFont = TextFont("Arial Rounded MT Bold")
shape.TextFrame.Paragraphs[0].TextRanges[0].Fill.FillType = FillFormatType.Solid
shape.TextFrame.Paragraphs[0].TextRanges[0].Fill.SolidColor.Color = Color.get_Black()
```

---

# Spire.Presentation Text Alignment
## Set different text alignment types for paragraphs in a PowerPoint shape
```python
#Get the related shape and set the text alignment
shape = presentation.Slides[0].Shapes[1]
shape.TextFrame.Paragraphs[0].Alignment = TextAlignmentType.Left
shape.TextFrame.Paragraphs[1].Alignment = TextAlignmentType.Center
shape.TextFrame.Paragraphs[2].Alignment = TextAlignmentType.Right
shape.TextFrame.Paragraphs[3].Alignment = TextAlignmentType.Justify
shape.TextFrame.Paragraphs[4].Alignment = TextAlignmentType.none
```

---

# spire.presentation python HTML
## Append HTML content to PowerPoint shapes
```python
#Add a shape 
shape = ppt.Slides[0].Shapes.AppendShape(ShapeType.Rectangle, RectangleF.FromLTRB (150, 100, 350, 300))

#Clear default paragraphs 
shape.TextFrame.Paragraphs.Clear()

code = "<html><body><p>This is a paragraph</p></body></html>"

#Append HTML, and generate a paragraph with default style in PPT document.
shape.TextFrame.Paragraphs.AddFromHtml(code)
codeColor = "<html><body><p style=\" color:black \">This is a paragraph</p></body></html>"
#Append HTML with black setting
shape.TextFrame.Paragraphs.AddFromHtml(codeColor)

#Add another shape
shape1 = ppt.Slides[0].Shapes.AppendShape(ShapeType.Rectangle, RectangleF.FromLTRB (350, 100, 550, 300))

#Clear default paragraph 
shape1.TextFrame.Paragraphs.Clear()

#Change the fill format of shape
shape1.Fill.FillType = FillFormatType.Solid
shape1.Fill.SolidColor.Color = Color.get_White()

#Append HTML
shape1.TextFrame.Paragraphs.AddFromHtml(code)
par = shape1.TextFrame.Paragraphs[0]
#Change the fill color for paragraph
for tr in par.TextRanges:
    tr.Fill.FillType = FillFormatType.Solid
    tr.Fill.SolidColor.Color = Color.get_Black()
```

---

# Spire.Presentation AutoFit Text or Shape
## Demonstrates how to set autofit options for text in shapes
```python
#Set the AutofitType property to Shape
textShape2 = ppt.Slides[0].Shapes.AppendShape(ShapeType.Rectangle, RectangleF.FromLTRB (150, 100, 300, 180))
textShape2.TextFrame.Text = "Resize shape to fit text."
textShape2.TextFrame.AutofitType = TextAutofitType.Shape

#Set the AutofitType property to Normal
textShape1 = ppt.Slides[0].Shapes.AppendShape(ShapeType.Rectangle, RectangleF.FromLTRB (400, 100, 550, 180))
textShape1.TextFrame.Text = "Shrink text to fit shape. Shrink text to fit shape. Shrink text to fit shape. Shrink text to fit shape."
textShape1.TextFrame.AutofitType = TextAutofitType.Normal
```

---

# Spire.Presentation Python Borders and Shading
## Set borders, gradient fill, and shadow effects for presentation shapes
```python
# Assuming presentation is already loaded
shape = presentation.Slides[0].Shapes[0]

#Set line color and width of the border
shape.Line.FillType = FillFormatType.Solid
shape.Line.Width = 3
shape.Line.SolidFillColor.Color = Color.get_LightYellow()

#Set the gradient fill color of shape
shape.Fill.FillType = FillFormatType.Gradient
shape.Fill.Gradient.GradientShape = GradientShapeType.Linear
shape.Fill.Gradient.GradientStops.AppendByKnownColors(1, KnownColors.LightBlue)
shape.Fill.Gradient.GradientStops.AppendByKnownColors(0, KnownColors.LightSkyBlue)

#Set the shadow for the shape
shadow = OuterShadowEffect()
shadow.BlurRadius = 20
shadow.Direction = 30
shadow.Distance = 8
shadow.ColorFormat.Color = Color.get_LightSeaGreen()
shape.EffectDag.OuterShadowEffect = shadow
```

---

# spire.presentation python bullets
## add bullets to paragraphs in presentation
```python
shape = presentation.Slides[0].Shapes[1]

for para in shape.TextFrame.Paragraphs:
    #Add the bullets
    para.BulletType = TextBulletType.Numbered
    para.BulletStyle = NumberedBulletStyle.BulletRomanLCPeriod
```

---

# spire.presentation python text styling
## change text style in presentation slides
```python
shape = presentation.Slides[0].Shapes[0]
paras = shape.TextFrame.Paragraphs

#Set the style for the text content in the first paragraph
for tr in paras[0].TextRanges:
    tr.Fill.FillType = FillFormatType.Solid
    tr.Fill.SolidColor.Color = Color.get_ForestGreen()
    tr.LatinFont = TextFont("Lucida Sans Unicode")
    tr.FontHeight = 14

#Set the style for the text content in the third paragraph
for tr in paras[2].TextRanges:
    tr.Fill.FillType = FillFormatType.Solid
    tr.Fill.SolidColor.Color = Color.get_CornflowerBlue()
    tr.LatinFont = TextFont("Calibri")
    tr.FontHeight = 16
    tr.TextUnderlineType = TextUnderlineType.Dashed
```

---

# Copy Paragraph Between PowerPoint Presentations
## This code demonstrates how to copy text from a shape in one PowerPoint presentation to a shape in another PowerPoint presentation.
```python
#Get the text from the first shape on the first slide
sourceshp = ppt1.Slides[0].Shapes[0]
text = (sourceshp).TextFrame.Text

#Get the first shape on the first slide from the target file
destshp = ppt2.Slides[0].Shapes[0]

#Add the text to the target file
(destshp).TextFrame.Text += "\n\n" + text
```

---

# Spire.Presentation Custom Bullet Numbering
## Customize bullet numbering for paragraphs in a presentation
```python
#Get the first slide
slide = presentation.Slides[0]

#Access the first placeholder in the slide and typecasting it as AutoShape
tf1 = (slide.Shapes[1]).TextFrame

#Access the first Paragraph and set bullet style
para = tf1.Paragraphs[0]
para.Depth = 0
para.BulletType = TextBulletType.Numbered
para.BulletStyle = NumberedBulletStyle.BulletArabicPeriod
para.BulletNumber = 2

#Access the second Paragraph and set bullet style
para = tf1.Paragraphs[1]
para.Depth = 0
para.BulletType = TextBulletType.Numbered
para.BulletStyle = NumberedBulletStyle.BulletArabicPeriod
para.BulletNumber = 4

#Access the third Paragraph and set bullet style
para = tf1.Paragraphs[2]
para.Depth = 0
para.BulletType = TextBulletType.Numbered
para.BulletStyle = NumberedBulletStyle.BulletArabicPeriod
para.BulletNumber = 6

#Access the fourth Paragraph and set bullet style
para = tf1.Paragraphs[3]
para.Depth = 0
para.BulletType = TextBulletType.Numbered
para.BulletStyle = NumberedBulletStyle.BulletArabicPeriod
para.BulletNumber = 7
```

---

# Spire.Presentation Edit Prompt Text
## Edit placeholder text in PowerPoint slides
```python
# Iterate through the slide
for shape in presentation.Slides[0].Shapes:
    if shape.Placeholder is not None and isinstance(shape, IAutoShape):
        text = ""
        # Set the text of the title
        if shape.Placeholder.Type == PlaceholderType.CenteredTitle:
            text = "custom title create by Spire"
        # Set text of the subtitle.
        elif shape.Placeholder.Type == PlaceholderType.Subtitle:
            text = "custom subtitle create by Spire"

        ( shape if isinstance(shape, IAutoShape) else None).TextFrame.Text = text
```

---

# Spire.Presentation text extraction
## Extract text from PowerPoint presentation slides
```python
# Iterate through slides and extract text
sb = []
for slide in presentation.Slides:
    for shape in slide.Shapes:
        if isinstance(shape, IAutoShape):
            for tp in (shape if isinstance(shape, IAutoShape) else None).TextFrame.Paragraphs:
                sb.append(tp.Text)
```

---

# Get Text Frame Effective Data
## Extract text frame format properties from a PowerPoint slide shape
```python
#Get the first slide from a presentation
slide = presentation.Slides[0]
#Get a shape from the slide
shape = slide.Shapes[0] if isinstance(slide.Shapes[0], IAutoShape) else None

#Extract text frame format properties
textFrameFormat = shape.TextFrame
sb = []
sb.append ("Anchoring type: " + str(textFrameFormat.AnchoringType))
sb.append("Autofit type: " + str(textFrameFormat.AutofitType))
sb.append("Text vertical type: " + str(textFrameFormat.VerticalTextType))
sb.append("Margins")
sb.append("   Left: " + str(textFrameFormat.MarginLeft))
sb.append("   Top: " + str(textFrameFormat.MarginTop))
sb.append("   Right: " + str(textFrameFormat.MarginRight))
sb.append("   Bottom: " + str(textFrameFormat.MarginBottom))
```

---

# Spire.Presentation Text Style Extraction
## Extract paragraph and text range style information from PowerPoint slides

```python
# Get the first slide
slide = presentation.Slides[0]
# Get a shape 
shape = presentation.Slides[0].Shapes[0] if isinstance(presentation.Slides[0].Shapes[0], IAutoShape) else None

for p, unusedItem in enumerate(shape.TextFrame.Paragraphs):
    paragraph = shape.TextFrame.Paragraphs[p]
    # Get the paragraph style
    paragraphIndent = paragraph.Indent
    paragraphAlignment = paragraph.Alignment
    paragraphFontAlignment = paragraph.FontAlignment
    paragraphHangingPunctuation = paragraph.HangingPunctuation
    paragraphLineSpacing = paragraph.LineSpacing
    paragraphSpaceBefore = paragraph.SpaceBefore
    paragraphSpaceAfter = paragraph.SpaceAfter
    
    for r, unusedItem in enumerate(paragraph.TextRanges):
        textRange = paragraph.TextRanges[r]
        # Get the text range style
        textRangeFontHeight = textRange.FontHeight
        textRangeLanguage = textRange.Language
        textRangeFont = textRange.LatinFont.FontName
```

---

# spire.presentation python text highlighting
## highlight specified text in presentation
```python
#Get the specified shape
shape = ppt.Slides[0].Shapes[1]

options = TextHighLightingOptions()
options.WholeWordsOnly = True
options.CaseSensitive = True

shape.TextFrame.HighLightText("Spire", Color.get_Yellow(), options)
```

---

# spire.presentation python paragraph indentation
## set paragraph indentation in presentation slides
```python
shape = presentation.Slides[0].Shapes[0]
paras = shape.TextFrame.Paragraphs

#Set the paragraph style for first paragraph
paras[0].Indent = 20
paras[0].LeftMargin = 10
paras[0].SpaceAfter = 10

#Set the paragraph style of the third paragraph 
paras[2].Indent = -100
paras[2].LeftMargin = 40
paras[2].SpaceBefore = 0
paras[2].SpaceAfter = 0
```

---

# Spire.Presentation HTML with Image
## Insert HTML content with image into presentation slide
```python
# Create an instance of presentation document
ppt = Presentation()
shapes = ppt.Slides[0].Shapes

# Add HTML content with image to the slide
shapes.AddFromHtml("<html><div><p>First paragraph</p><p><img src='image_path'/></p><p>Second paragraph </p></html>")
```

---

# Spire.Presentation Line Spacing
## Set line spacing for paragraphs in PowerPoint presentations
```python
#Create a PPT document
presentation = Presentation()

#Get the first slide
slide = presentation.Slides[0]
#Add a shape 
shape = presentation.Slides[0].Shapes.AppendShape(ShapeType.Rectangle, RectangleF.FromLTRB (50, 100, presentation.SlideSize.Size.Width - 50, 400))
shape.ShapeStyle.LineColor.Color = Color.get_White()
shape.Fill.FillType = FillFormatType.none
shape.TextFrame.Paragraphs.Clear()

#Add text
shape.AppendTextFrame("Spire.Presentation for Python is a professional presentation processing API that is highly compatible with PowerPoint. It is a completely independent class library that developers can use to create, edit, convert, and save PowerPoint presentations efficiently without installing Microsoft PowerPoint.")
#Set font and color of text
textRange = shape.TextFrame.TextRange
textRange.Fill.FillType = FillFormatType.Solid
textRange.Fill.SolidColor.Color = Color.get_BlueViolet()
textRange.FontHeight = 20
textRange.LatinFont = TextFont("Lucida Sans Unicode")

#Set properties of paragraph
shape.TextFrame.Paragraphs[0].SpaceBefore = 100
shape.TextFrame.Paragraphs[0].SpaceAfter = 100
shape.TextFrame.Paragraphs[0].LineSpacing = 150
```

---

# Spire.Presentation Python Font Styling
## Apply mixed font styles to specific words in presentation text
```python
#Get the second shape of the first slide
shape = ppt.Slides[0].Shapes[1] if isinstance(ppt.Slides[0].Shapes[1], IAutoShape) else None
#Get the text from the shape 
originalText = shape.TextFrame.Text

#Split the string by specified words and return substrings to a string array
splitArray = originalText.split("bold")

#Remove the paragraph from TextRange
tp = shape.TextFrame.Paragraphs[0]
tp.TextRanges.Clear()

#Append normal text that is in front of 'bold' to the paragraph
tr = TextRange(splitArray[0])
tp.TextRanges.Append(tr)
#Set font style of the text 'bold' as bold
tr = TextRange("bold")
tr.IsBold = TriState.TTrue
tp.TextRanges.Append(tr)

splitArray = splitArray[1].split("red")
#Append normal text that is in front of 'red' to the paragraph
tr = TextRange(splitArray[0])
tp.TextRanges.Append(tr)
#Set the color of the text 'red' as red
tr = TextRange("red")
tr.Fill.FillType = FillFormatType.Solid
tr.Format.Fill.SolidColor.Color = Color.get_Red()
tp.TextRanges.Append(tr)
splitArray = splitArray[1].split("underlined")
#Append normal text that is in front of 'underlined' to the paragraph
tr = TextRange(splitArray[0])
tp.TextRanges.Append(tr)
#Underline the text 'undelined'
tr = TextRange("underlined")
tr.TextUnderlineType = TextUnderlineType.Single
tp.TextRanges.Append(tr)

splitArray = splitArray[1].split("bigger font size")
#Append normal text that is in front of 'bigger font size' to the paragraph
tr = TextRange(splitArray[0])
tp.TextRanges.Append(tr)
#Set a large font for the text 'bigger font size'
tr = TextRange("bigger font size")
tr.FontHeight = 35
tp.TextRanges.Append(tr)

#Append other normal text
tr = TextRange(splitArray[1])
tp.TextRanges.Append(tr)
```

---

# Spire.Presentation Multiple Level Bullets
## Create multiple level bullet styles in a presentation
```python
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
```

---

# spire.presentation python paragraphs
## create multiple paragraphs with different text formatting in a PowerPoint shape
```python
# Access the first slide
slide = presentation.Slides[0]

# Add an AutoShape of rectangle type
left = math.trunc(presentation.SlideSize.Size.Width / float(2)) - 250
rec = RectangleF.FromLTRB (left, 150, 500+left, 300)
shape = presentation.Slides[0].Shapes.AppendShape(ShapeType.Rectangle, rec)

# Access TextFrame of the AutoShape
tf = shape.TextFrame

# Create Paragraphs and TextRanges with different text formats
para0 = tf.Paragraphs[0]
textRange1 = TextRange()
textRange2 = TextRange()
para0.TextRanges.Append(textRange1)
para0.TextRanges.Append(textRange2)

para1 = TextParagraph()
tf.Paragraphs.Append(para1)
textRange11 = TextRange()
textRange12 = TextRange()
textRange13 = TextRange()
para1.TextRanges.Append(textRange11)
para1.TextRanges.Append(textRange12)
para1.TextRanges.Append(textRange13)

para2 = TextParagraph()
tf.Paragraphs.Append(para2)
textRange21 = TextRange()
textRange22 = TextRange()
textRange23 = TextRange()
para2.TextRanges.Append(textRange21)
para2.TextRanges.Append(textRange22)
para2.TextRanges.Append(textRange23)

for i in range(0, 3):
    for j in range(0, 3):
        tf.Paragraphs[i].TextRanges[j].Text = "TextRange " + str(j)
        if j == 0:
            tf.Paragraphs[i].TextRanges[j].Fill.FillType = FillFormatType.Solid
            tf.Paragraphs[i].TextRanges[j].Fill.SolidColor.Color = Color.get_LightBlue()
            tf.Paragraphs[i].TextRanges[j].Format.IsBold = TriState.TTrue
            tf.Paragraphs[i].TextRanges[j].FontHeight = 15
        elif j == 1:
            tf.Paragraphs[i].TextRanges[j].Fill.FillType = FillFormatType.Solid
            tf.Paragraphs[i].TextRanges[j].Fill.SolidColor.Color = Color.get_Blue()
            tf.Paragraphs[i].TextRanges[j].Format.IsItalic = TriState.TTrue
            tf.Paragraphs[i].TextRanges[j].FontHeight = 18
```

---

# spire.presentation custom bullet style
## set custom picture bullet style for paragraphs
```python
#Get the second shape on the first slide
shape = ppt.Slides[0].Shapes[1] if isinstance(ppt.Slides[0].Shapes[1], IAutoShape) else None

#Traverse through the paragraphs in the shape
for paragraph in shape.TextFrame.Paragraphs:
    #Set the bullet style of paragraph as picture
    paragraph.BulletType = TextBulletType.Picture
    #Load a picture
    fileStream = Stream("./Data/icon.png")
    paragraph.BulletPicture.EmbedImage = ppt.Images.AppendStream(fileStream)
    fileStream.Close()
```

---

# Spire.Presentation Remove TextBox
## Remove text boxes from PowerPoint slides
```python
#Get the first slide
slide = ppt.Slides[0]
#Traverse all the shapes in slide
i = 0
while i < slide.Shapes.Count:
    #Remove all shapes
    shape = slide.Shapes[i] if isinstance(slide.Shapes[i], IAutoShape) else None
    slide.Shapes.Remove(shape)
```

---

# spire.presentation python text replacement
## replace text in presentation slides
```python
def ReplaceTags(pSlide, TagValues):
    for curShape in pSlide.Shapes:
        if isinstance(curShape, IAutoShape):
            for tp in ( curShape if isinstance(curShape, IAutoShape) else None).TextFrame.Paragraphs:
                for curKey in TagValues.keys():
                    tp.Text = tp.Text.replace(curKey, TagValues[curKey])
```

---

# spire.presentation python text replacement
## replace text in presentation while retaining style
```python
ppt.Slides[0].ReplaceFirstText("use", "test", True)
ppt.Slides[1].ReplaceAllText("Spire", "new spire", True)
```

---

# spire.presentation python replace text
## replace text with regex in presentation
```python
#Regex for all words
regex = Regex("\\d+.\\d+|\\w+")

#New string value
newvalue = "This is the test!"

#Loop and replace
for slide in presentation.Slides:
    for shape in slide.Shapes:
        shape.ReplaceTextWithRegex(regex, newvalue)
```

---

# spire.presentation python text rotation
## rotate text in presentation shape
```python
#Get a shape 
shape = presentation.Slides[0].Shapes[0] if isinstance(presentation.Slides[0].Shapes[0], IAutoShape) else None

#Set text rotation to 270 degrees
shape.TextFrame.VerticalTextType = VerticalTextType.Vertical270
```

---

# Spire.Presentation 3D Text Effect
## Set 3D effect for text in presentation slides
```python
#Create a new presentation object
ppt = Presentation()

#Get the first slide
slide = ppt.Slides[0]

#Append a new shape to slide and set the line color and fill type
shape = slide.Shapes.AppendShape(ShapeType.Rectangle, RectangleF.FromLTRB (30, 40, 680, 240))
shape.ShapeStyle.LineColor.Color = Color.get_White()
shape.Fill.FillType = FillFormatType.none

#Add text to the shape
shape.AppendTextFrame("This demo shows how to add 3D effect text to Presentation slide")

#Set the color of text in shape
textRange = shape.TextFrame.TextRange
textRange.Fill.FillType = FillFormatType.Solid
textRange.Fill.SolidColor.Color = Color.get_LightBlue()

#Set the Font of text in shape
textRange.FontHeight = 40
textRange.LatinFont = TextFont("Gulim")

#Set 3D effect for text
shape.TextFrame.TextThreeD.ShapeThreeD.PresetMaterial = PresetMaterialType.Matte
shape.TextFrame.TextThreeD.LightRig.PresetType = PresetLightRigType.Sunrise
shape.TextFrame.TextThreeD.ShapeThreeD.TopBevel.PresetType = BevelPresetType.Circle
shape.TextFrame.TextThreeD.ShapeThreeD.ContourColor.Color = Color.get_Green()
shape.TextFrame.TextThreeD.ShapeThreeD.ContourWidth = 3
```

---

# spire.presentation python text frame
## set anchor of text frame
```python
#Get the first slide
slide = presentation.Slides[0]
#Get a shape 
shape = presentation.Slides[0].Shapes[0] if isinstance(presentation.Slides[0].Shapes[0], IAutoShape) else None
shape.TextFrame.AnchoringType = TextAnchorType.Bottom
```

---

# Spire.Presentation Python Font Formatting
## Set paragraph font properties including font family, bold, italic, and color
```python
#Get the first slide
slide = presentation.Slides[0]

#Access the first and second placeholder in the slide and typecasting it as AutoShape
tf1 = (slide.Shapes[0]).TextFrame
tf2 = (slide.Shapes[1]).TextFrame

# Access the first Paragraph
para1 = tf1.Paragraphs[0]
para2 = tf2.Paragraphs[0]

#Justify the paragraph
para2.Alignment = TextAlignmentType.Justify

#Access the first text range
textRange1 = para1.FirstTextRange
textRange2 = para2.FirstTextRange

#Define new fonts
fd1 = TextFont("Elephant")
fd2 = TextFont("Castellar")

# Assign new fonts to text range
textRange1.LatinFont = fd1
textRange2.LatinFont = fd2

# Set font to Bold
textRange1.Format.IsBold = TriState.TTrue
textRange2.Format.IsBold = TriState.TFalse

# Set font to Italic
textRange1.Format.IsItalic = TriState.TFalse
textRange2.Format.IsItalic = TriState.TTrue

# Set font color
textRange1.Fill.FillType = FillFormatType.Solid
textRange1.Fill.SolidColor.Color = Color.get_Purple()
textRange2.Fill.FillType = FillFormatType.Solid
textRange2.Fill.SolidColor.Color = Color.get_Peru()
```

---

# Spire.Presentation Right-to-Left Columns
## Set text frame columns to right-to-left direction
```python
#Get the second shape
shape = ppt.Slides[0].Shapes[1] if isinstance(ppt.Slides[0].Shapes[1], IAutoShape) else None
#Set columns style to right-to-left
shape.TextFrame.RightToLeftColumns = True
```

---

# Spire.Presentation Text Shadow Effect
## Apply shadow effect to text in presentation slides
```python
#Create an instance of presentation document
ppt = Presentation()

#Get reference of the slide
slide = ppt.Slides[0]

#Add a new rectangle shape to the first slide
shape = slide.Shapes.AppendShape(ShapeType.Rectangle, RectangleF.FromLTRB (120, 100, 570, 300))
shape.Fill.FillType = FillFormatType.none

#Add the text to the shape and set the font for the text
shape.AppendTextFrame("Text shading on slides")
shape.TextFrame.Paragraphs[0].TextRanges[0].LatinFont = TextFont("Arial Black")
shape.TextFrame.Paragraphs[0].TextRanges[0].FontHeight = 21
shape.TextFrame.Paragraphs[0].TextRanges[0].Fill.FillType = FillFormatType.Solid
shape.TextFrame.Paragraphs[0].TextRanges[0].Fill.SolidColor.Color = Color.get_Black()

#Add outer shadow and set all necessary parameters
Shadow = OuterShadowEffect()

Shadow.BlurRadius = 0
Shadow.Direction = 50
Shadow.Distance = 10
Shadow.ColorFormat.Color = Color.get_LightBlue()

shape.TextFrame.TextRange.EffectDag.OuterShadowEffect = Shadow
```

---

# spire.presentation python text direction
## Set text direction in presentation slides
```python
#Append a shape with text to the first slide
textboxShape = ppt.Slides[0].Shapes.AppendShape(ShapeType.Rectangle, RectangleF.FromLTRB (250, 70, 350, 470))
textboxShape.ShapeStyle.LineColor.Color = Color.get_Transparent()
textboxShape.Fill.FillType = FillFormatType.Solid
textboxShape.Fill.SolidColor.Color = Color.get_LightBlue()
textboxShape.TextFrame.Text = "You Are Welcome Here"
#Set the text direction to vertical
textboxShape.TextFrame.VerticalTextType = VerticalTextType.Vertical

#Append another shape with text to the slide
textboxShape = ppt.Slides[0].Shapes.AppendShape(ShapeType.Rectangle, RectangleF.FromLTRB (350, 70, 450, 470))
textboxShape.ShapeStyle.LineColor.Color = Color.get_Transparent()
textboxShape.Fill.FillType = FillFormatType.Solid
textboxShape.Fill.SolidColor.Color = Color.get_LightGray()
#Append some asian characters
textboxShape.TextFrame.Text = "欢迎光临"
#Set the VerticalTextType as EastAsianVertical to avoid rotating text 90 degrees
textboxShape.TextFrame.VerticalTextType = VerticalTextType.EastAsianVertical
```

---

# spire.presentation python text formatting
## Set text font properties in PowerPoint presentation
```python
#Add text to the shape
shape.AppendTextFrame("Welcome to use Spire.Presentation")

textRange = shape.TextFrame.TextRange
#Set the font
textRange.LatinFont = TextFont("Times New Roman")
#Set bold property of the font
textRange.IsBold = TriState.TTrue

#Set italic property of the font
textRange.IsItalic = TriState.TTrue

#Set underline property of the font
textRange.TextUnderlineType = TextUnderlineType.Single

#Set the height of the font
textRange.FontHeight = 50

#Set the color of the font
textRange.Fill.FillType = FillFormatType.Solid
textRange.Fill.SolidColor.Color = Color.get_CadetBlue()
```

---

# spire.presentation python text margins
## set text margins for a shape in presentation
```python
#Append a new shape
shape = ppt.Slides[0].Shapes.AppendShape(ShapeType.Rectangle, RectangleF.FromLTRB (50, 100, 500, 250))

#Set margins for text inside shapes
shape.Fill.FillType = FillFormatType.none
shape.ShapeStyle.LineColor.Color = Color.get_LightBlue()
shape.TextFrame.Paragraphs[0].Alignment = TextAlignmentType.Justify
shape.TextFrame.Text = "Spire.Presentation for Python is a professional presentation processing API that is highly compatible with PowerPoint. It is a completely independent class library that developers can use to create, edit, convert, and save PowerPoint presentations efficiently without installing Microsoft PowerPoint."
shape.TextFrame.Paragraphs[0].TextRanges[0].LatinFont = TextFont("Arial Rounded MT Bold")
shape.TextFrame.Paragraphs[0].TextRanges[0].Fill.FillType = FillFormatType.Solid
shape.TextFrame.Paragraphs[0].TextRanges[0].Fill.SolidColor.Color = Color.get_Black()

#Set the margins for the text frame
shape.TextFrame.MarginTop = 10
shape.TextFrame.MarginBottom = 35
shape.TextFrame.MarginLeft = 15
shape.TextFrame.MarginRight = 30
```

---

# Spire.Presentation Text Transparency
## Set text transparency with different alpha values in a presentation
```python
# Create an instance of presentation document
ppt = Presentation()

# Add a shape
textboxShape = ppt.Slides[0].Shapes.AppendShape(ShapeType.Rectangle, RectangleF.FromLTRB (100, 100, 400, 220))
textboxShape.ShapeStyle.LineColor.Color = Color.get_Transparent()
textboxShape.Fill.FillType = FillFormatType.none

# Remove default blank paragraphs
textboxShape.TextFrame.Paragraphs.Clear()

# Add three paragraphs, apply color with different alpha values to text
alpha = 55
for i in range(0, 3):
    textboxShape.TextFrame.Paragraphs.Append(TextParagraph())
    textboxShape.TextFrame.Paragraphs[i].TextRanges.Append(TextRange("Text Transparency"))
    textboxShape.TextFrame.Paragraphs[i].TextRanges[0].Fill.FillType = FillFormatType.Solid
    textboxShape.TextFrame.Paragraphs[i].TextRanges[0].Fill.SolidColor.Color = Color.FromArgb(alpha, Color.get_Purple().R,Color.get_Purple().G,Color.get_Purple().B)
    alpha += 100
```

---

# spire.presentation python superscript and subscript
## create superscript and subscript text in presentation slides
```python
# Add a shape for superscript
shape = presentation.Slides[0].Shapes.AppendShape(ShapeType.Rectangle, RectangleF.FromLTRB (150, 100, 350, 150))
shape.ShapeStyle.LineColor.Color = Color.get_White()
shape.Fill.FillType = FillFormatType.none
shape.TextFrame.Paragraphs.Clear()

shape.AppendTextFrame("Test")
tr = TextRange("superscript")
shape.TextFrame.Paragraphs[0].TextRanges.Append(tr)

# Set superscript text
shape.TextFrame.Paragraphs[0].TextRanges[1].Format.ScriptDistance = 30

textRange = shape.TextFrame.Paragraphs[0].TextRanges[0]
textRange.Fill.FillType = FillFormatType.Solid
textRange.Fill.SolidColor.Color = Color.get_Black()
textRange.FontHeight = 20
textRange.LatinFont = TextFont("Lucida Sans Unicode")

textRange = shape.TextFrame.Paragraphs[0].TextRanges[1]
textRange.Fill.FillType = FillFormatType.Solid
textRange.Fill.SolidColor.Color = Color.get_CadetBlue()
textRange.LatinFont = TextFont("Lucida Sans Unicode")

# Add a shape for subscript
shape = presentation.Slides[0].Shapes.AppendShape(ShapeType.Rectangle, RectangleF.FromLTRB (150, 150, 350, 200))
shape.ShapeStyle.LineColor.Color = Color.get_White()
shape.Fill.FillType = FillFormatType.none
shape.TextFrame.Paragraphs.Clear()

shape.AppendTextFrame("Test")
tr = TextRange("subscript")
shape.TextFrame.Paragraphs[0].TextRanges.Append(tr)

# Set subscript text
shape.TextFrame.Paragraphs[0].TextRanges[1].Format.ScriptDistance = -25

textRange = shape.TextFrame.Paragraphs[0].TextRanges[0]
textRange.Fill.FillType = FillFormatType.Solid
textRange.Fill.SolidColor.Color = Color.get_Black()
textRange.FontHeight = 20
textRange.LatinFont = TextFont("Lucida Sans Unicode")

textRange = shape.TextFrame.Paragraphs[0].TextRanges[1]
textRange.Fill.FillType = FillFormatType.Solid
textRange.Fill.SolidColor.Color = Color.get_CadetBlue()
textRange.LatinFont = TextFont("Lucida Sans Unicode")
```

---

# spire.presentation python slide master
## add image to slide master
```python
#Get the master collection
master = presentation.Masters[0]

#Append image to slide master
image = "./Data/Logo.png"
rff = RectangleF.FromLTRB (40, 40, 130, 130)
pic = master.Shapes.AppendEmbedImageByPath(ShapeType.Rectangle, image, rff)
pic.Line.FillFormat.FillType = FillFormatType.none
```

---

# Spire.Presentation slide management
## Append slides with master layouts in PowerPoint presentations
```python
#Get the master
master = presentation.Masters[0]
#Get master layout slides
masterLayouts = master.Layouts
layoutSlide = masterLayouts[1]
#Append a rectangle to the layout slide
shape = layoutSlide.Shapes.AppendShape(ShapeType.Rectangle, RectangleF.FromLTRB (10, 50, 110, 130))
#Add a text into the shape and set the style
shape.Fill.FillType = FillFormatType.none
shape.AppendTextFrame("Layout slide 1")
shape.TextFrame.Paragraphs[0].TextRanges[0].LatinFont = TextFont("Arial Black")
shape.TextFrame.Paragraphs[0].TextRanges[0].Fill.FillType = FillFormatType.Solid
shape.TextFrame.Paragraphs[0].TextRanges[0].Fill.SolidColor.Color = Color.get_CadetBlue()
#Append new slide with master layout
presentation.Slides.Append(presentation.Slides[0], master.Layouts[1])
#Another way to append new slide with master layout
presentation.Slides.Insert(2, presentation.Slides[1], master.Layouts[1])
```

---

# Spire.Presentation Slide Master Customization
## Apply and customize slide master in PowerPoint presentation

```python
#Get the first slide master from the presentation
masterSlide = ppt.Masters[0]
#Customize the background of the slide master
rect = RectangleF.FromLTRB (0, 0, ppt.SlideSize.Size.Width, ppt.SlideSize.Size.Height)
masterSlide.SlideBackground.Fill.FillType = FillFormatType.Picture
# Append an image to the slide master background
image = masterSlide.Shapes.AppendEmbedImageByPath (ShapeType.Rectangle, "path_to_image", rect)
masterSlide.SlideBackground.Fill.PictureFill.Picture.EmbedImage = image.PictureFill.Picture.EmbedImage
#Change the color scheme
masterSlide.Theme.ColorScheme.Accent1.Color = Color.get_Red()
masterSlide.Theme.ColorScheme.Accent2.Color = Color.get_RosyBrown()
masterSlide.Theme.ColorScheme.Accent3.Color = Color.get_Ivory()
masterSlide.Theme.ColorScheme.Accent4.Color = Color.get_Lavender()
masterSlide.Theme.ColorScheme.Accent5.Color = Color.get_Black()
```

---

# Spire.Presentation Python Slide Position
## Change the position of slides in a presentation
```python
#Create a PPT document
presentation = Presentation()
#Move the first slide to the second slide position
slide = presentation.Slides[0]
slide.SlideNumber = 2
```

---

# Clone PowerPoint Presentation
## Clone slides from one PowerPoint presentation and append them to the end of another presentation
```python
#Load source document from disk
sourcePPT = Presentation()
sourcePPT.LoadFromFile(inputFile_1)
#Load destination document from disk
destPPT = Presentation()
destPPT.LoadFromFile(inputFile_2)
#Loop through all slides of source document
for slide in sourcePPT.Slides:
    #Append the slide at the end of destination document
    destPPT.Slides.AppendBySlide(slide)
#Save the document
destPPT.SaveToFile(outputFile, FileFormat.Pptx2013)
destPPT.Dispose()
```

---

# Spire.Presentation Python Master Cloning
## Clone PowerPoint masters from one presentation to another
```python
inputFile_1 = "./Data/CloneMaster1.pptx"
inputFile_2 = "./Data/CloneMaster2.pptx"
outputFile = "ClonePPTMasterToAnother.pptx"
# Load PPT1 from disk
presentation1 = Presentation()
presentation1.LoadFromFile(inputFile_1)
# Load PPT2 from disk
presentation2 = Presentation()
presentation2.LoadFromFile(inputFile_2)
# Add masters from PPT1 to PPT2
for masterSlide in presentation1.Masters:
    presentation2.Masters.AppendSlide(masterSlide)
# Save the document
presentation2.SaveToFile(outputFile, FileFormat.Pptx2013)
presentation2.Dispose()
```

---

# Spire.Presentation slide cloning
## Clone a slide and append it at the end of the presentation
```python
#Get the first slide
slide = presentation.Slides[0]
#Append the slide at the end of the document
presentation.Slides.AppendBySlide(slide)
```

---

# spire.presentation python slide cloning
## Clone a slide from one presentation to another
```python
#Create presentations
presentation = Presentation()
ppt1 = Presentation()
#Choose the first slide to be cloned from the source presentation
slide1 = ppt1.Slides[0]
#Insert the slide to the specified index in the destination presentation
index = 1
presentation.Slides.Insert(index, slide1)
```

---

# spire.presentation python slide cloning
## clone a slide within the same presentation
```python
#Create an instance of presentation document
ppt = Presentation()
#Get a list of slides and choose the first slide to be cloned
slide = ppt.Slides[0]
#Insert the desired slide to the specified index in the same presentation
index = 1
ppt.Slides.Insert(index, slide)
ppt.Dispose()
```

---

# spire.presentation python slide creation
## create and format presentation slides
```python
#Create PPT document
presentation = Presentation()
#Add new slide
presentation.Slides.Append()
#Set the background image
for i in range(0, 2):
    rect = RectangleF.FromLTRB (0, 0, presentation.SlideSize.Size.Width, presentation.SlideSize.Size.Height)
    presentation.Slides[i].Shapes.AppendEmbedImageByPath (ShapeType.Rectangle, ImageFile, rect)
    presentation.Slides[i].Shapes[0].Line.FillFormat.SolidFillColor.Color = Color.get_FloralWhite()
#Add title
left =math.trunc(presentation.SlideSize.Size.Width / float(2)) - 200
rec_title = RectangleF.FromLTRB (left, 70, 400+left, 120)
shape_title = presentation.Slides[0].Shapes.AppendShape(ShapeType.Rectangle, rec_title)
shape_title.ShapeStyle.LineColor.Color = Color.get_White()
shape_title.Fill.FillType = FillFormatType.none
para_title = TextParagraph()
para_title.Text = "E-iceblue"
para_title.Alignment = TextAlignmentType.Center
para_title.TextRanges[0].LatinFont = TextFont("Myriad Pro Light")
para_title.TextRanges[0].FontHeight = 36
para_title.TextRanges[0].Fill.FillType = FillFormatType.Solid
para_title.TextRanges[0].Fill.SolidColor.Color = Color.get_Black()
shape_title.TextFrame.Paragraphs.Append(para_title)
#Append new shape
shape = presentation.Slides[0].Shapes.AppendShape(ShapeType.Rectangle, RectangleF.FromLTRB (50, 150, 650, 430))
shape.ShapeStyle.LineColor.Color = Color.get_White()
shape.Fill.FillType = FillFormatType.none
shape.Line.FillType = FillFormatType.none
#Add text to shape
shape.AppendTextFrame("Welcome to use Spire.Presentation for .NET.")
#Add new paragraph
pare = TextParagraph()
pare.Text = ""
shape.TextFrame.Paragraphs.Append(pare)
#Add new paragraph
pare = TextParagraph()
pare.Text = "Spire.Presentation for .NET is a professional PowerPoint compatible component that enables developers to create, read, write, modify, convert and Print PowerPoint documents from any .NET(C#, VB.NET, ASP.NET) platform. As an independent PowerPoint .NET component, Spire.Presentation for .NET doesn't need Microsoft PowerPoint installed on the machine."
shape.TextFrame.Paragraphs.Append(pare)
#Set the Font
for para in shape.TextFrame.Paragraphs:
    para.TextRanges[0].LatinFont = TextFont("Myriad Pro")
    para.TextRanges[0].FontHeight = 24
    para.TextRanges[0].Fill.FillType = FillFormatType.Solid
    para.TextRanges[0].Fill.SolidColor.Color = Color.get_Black()
    para.Alignment = TextAlignmentType.Left
#Append new shape - SixPointedStar
shape = presentation.Slides[1].Shapes.AppendShape(ShapeType.SixPointedStar, RectangleF.FromLTRB (100, 100, 200, 200))
shape.Fill.FillType = FillFormatType.Solid
shape.Fill.SolidColor.Color = Color.get_Orange()
shape.ShapeStyle.LineColor.Color = Color.get_White()
#Append new shape
shape = presentation.Slides[1].Shapes.AppendShape(ShapeType.Rectangle, RectangleF.FromLTRB (50, 250, 650, 300))
shape.ShapeStyle.LineColor.Color = Color.get_White()
shape.Fill.FillType = FillFormatType.none
#Add text to shape
shape.AppendTextFrame("This is newly added Slide.")
#Set the Font
shape.TextFrame.Paragraphs[0].TextRanges[0].LatinFont = TextFont("Myriad Pro")
shape.TextFrame.Paragraphs[0].TextRanges[0].FontHeight = 24
shape.TextFrame.Paragraphs[0].TextRanges[0].Fill.FillType = FillFormatType.Solid
shape.TextFrame.Paragraphs[0].TextRanges[0].Fill.SolidColor.Color = Color.get_Black()
shape.TextFrame.Paragraphs[0].Alignment = TextAlignmentType.Left
shape.TextFrame.Paragraphs[0].Indent = 35
```

---

# Spire.Presentation Slide Master Management
## Create slide masters and apply them to slides with background images
```python
#Create an instance of presentation document
ppt = Presentation()
ppt.SlideSize.Type = SlideSizeType.Screen16x9
#Add slides
for i in range(0, 4):
    ppt.Slides.Append()
#Get the first default slide master
first_master = ppt.Masters[0]
#Append another slide master
ppt.Masters.AppendSlide(first_master)
second_master = ppt.Masters[1]
#Set different background image for the two slide masters
#The first slide master
rect = RectangleF.FromLTRB (0, 0, ppt.SlideSize.Size.Width, ppt.SlideSize.Size.Height)
first_master.SlideBackground.Fill.FillType = FillFormatType.Picture
image1 = first_master.Shapes.AppendEmbedImageByPath (ShapeType.Rectangle, "Data/bg.png", rect)
first_master.SlideBackground.Fill.PictureFill.Picture.EmbedImage = image1.PictureFill.Picture.EmbedImage
#The second slide master
second_master.SlideBackground.Fill.FillType = FillFormatType.Picture
image2 = second_master.Shapes.AppendEmbedImageByPath (ShapeType.Rectangle, "Data/Setbackground.png", rect)
second_master.SlideBackground.Fill.PictureFill.Picture.EmbedImage = image2.PictureFill.Picture.EmbedImage
#Apply the first master with layout to the first slide
ppt.Slides[0].Layout = first_master.Layouts[1]
#Apply the second master with layout to other slides
for i in range(1, ppt.Slides.Count):
    ppt.Slides[i].Layout = second_master.Layouts[8]
```

---

# spire.presentation theme detection
## detect themes used in presentation slides
```python
#Create an instance of presentation document
ppt = Presentation()
#Get the theme name of each slide in the document
for slide in ppt.Slides:
    themeName = slide.Theme.Name
```

---

# spire.presentation python slide access
## get slides by index or ID
```python
#Get slide by index 0
slide1 = presentation.Slides[0]
#Get slide by slide ID
slide2 = presentation.FindSlide(presentation.Slides[1].SlideID)
```

---

# Spire.Presentation Hide Slide
## Hide a specific slide in a PowerPoint presentation
```python
# Hide the second slide
ppt.Slides[1].Hidden = True
```

---

# Spire.Presentation Python Merge Slides
## Merge selected slides from multiple presentations into one
```python
# Create an instance of presentation document
ppt = Presentation()
# Remove the first slide
ppt.Slides.RemoveAt(0)
# Load two PPT files
ppt1 = Presentation()
ppt1.LoadFromFile(inputFile_1)
ppt2 = Presentation()
ppt2.LoadFromFile(inputFile_2)
# Append all slides in ppt1 to ppt
for i, unusedItem in enumerate(ppt1.Slides):
    ppt.Slides.AppendBySlide(ppt1.Slides[i])
# Append the second slide in ppt2 to ppt
ppt.Slides.AppendBySlide(ppt2.Slides[1])
# Save the document
ppt.SaveToFile(outputFile, FileFormat.Pptx2013)
ppt.Dispose()
```

---

# Spire.Presentation Python Remove Slide
## Remove slides from presentation by index or reference
```python
# Remove slide by index
presentation.Slides.RemoveAt(0)
# Remove slide by its reference
slide = presentation.Slides[1]
presentation.Slides.Remove(slide)
```

---

# spire.presentation python remove unused layouts
## Remove unused layout masters from PowerPoint presentation
```python
#Create an array list
layouts = []
for i, unusedItem in enumerate(ppt.Slides):
    #Get the layout used by slide
    layout = ppt.Slides[i].Layout
    layouts.append(layout.SlideID)
#Loop through masters and layouts
for i, unusedItem in enumerate(ppt.Masters):
    masterlayouts = ppt.Masters[i].Layouts
    for j in range(masterlayouts.Count - 1, -1, -1):
        if not masterlayouts[j].SlideID in layouts:
            #Remove unused layout
            masterlayouts.RemoveMasterLayout(j)
```

---

# Spire.Presentation Python slide numbering
## Set starting number for slides
```python
#Create PPT document
presentation = Presentation()
#Set 5 as the starting number
presentation.FirstSlideNumber = 5
```

---

# Spire.Presentation slide title management
## Get and set slide titles in a presentation
```python
#Get the first slide
slide = ppt.Slides[0]
#Get the title of the first slide
slideTitle = slide.Title
#Set the title of the second slide
ppt.Slides[1].Title = "Second Slide"
```

---

# Spire.Presentation for Python
## Add different layout slides to a presentation
```python
#Create a PPT document
presentation = Presentation()

#Remove the default slide
presentation.Slides.RemoveAt(0)

#Loop through slide layouts
for slideLayoutType in SlideLayoutType:
    #Append slide by specifing slide layout
    presentation.Slides.AppendByLayoutType(slideLayoutType)
```

---

# Spire.Presentation Python Slide Layout
## Change slide layout in a PowerPoint presentation
```python
#Create a PPT document
presentation = Presentation()

#Change the layout of slide
presentation.Slides[1].Layout = presentation.Masters[0].Layouts[4]
```

---

# spire.presentation python slide layout
## get slide layout names from presentation
```python
#Create a PPT document
presentation = Presentation()

#Load the document from disk
presentation.LoadFromFile("presentation_file_path")

builder = []

#Loop through the slides of PPT document
for i, unusedItem in enumerate(presentation.Slides):
    #Get the name of slide layout
    name = presentation.Slides[i].Layout.Name
    builder.append ("The name of slide "+str(i)+" layout is: "+name)
```

---

# Spire.Presentation slide layout
## Set slide layout and add content
```python
#Create an instance of presentation document
ppt = Presentation()

#Remove the first slide
ppt.Slides.RemoveAt(0)

#Append a slide and set the layout for slide
slide = ppt.Slides.AppendByLayoutType(SlideLayoutType.Title)

#Add content for Title and Text
shape = slide.Shapes[0] if isinstance(slide.Shapes[0], IAutoShape) else None
shape.TextFrame.Text = "Hello Wolrd! -> This is title"

shape = slide.Shapes[1] if isinstance(slide.Shapes[1], IAutoShape) else None
shape.TextFrame.Text = "E-iceblue Support Team -> This is content"
```

---

# spire.presentation python slide transitions
## Set different transition types and timings for slides in a PowerPoint presentation
```python
#Set the first slide transition as circle
presentation.Slides[0].SlideShowTransition.Type = TransitionType.Circle

# Set the transition time of 3 seconds
presentation.Slides[0].SlideShowTransition.AdvanceOnClick = True
presentation.Slides[0].SlideShowTransition.AdvanceAfterTime = 3000

#Set the second slide transition as comb and set the speed 
presentation.Slides[1].SlideShowTransition.Type = TransitionType.Comb
presentation.Slides[1].SlideShowTransition.Speed = TransitionSpeed.Slow

# Set the transition time of 5 seconds
presentation.Slides[1].SlideShowTransition.AdvanceOnClick = True
presentation.Slides[1].SlideShowTransition.AdvanceAfterTime = 5000

# Set the third slide transition as zoom
presentation.Slides[2].SlideShowTransition.Type = TransitionType.Zoom

# Set the transition time of 7 seconds
presentation.Slides[2].SlideShowTransition.AdvanceOnClick = True
presentation.Slides[2].SlideShowTransition.AdvanceAfterTime = 7000
```

---

# Spire.Presentation Transition Effects
## Set transition effects for presentation slides
```python
# Set effects
presentation.Slides[0].SlideShowTransition.Type = TransitionType.Cut
presentation.Slides[0].SlideShowTransition.Value.FromBlack = True
```

---

# spire.presentation python transitions
## Set slide transitions in a presentation
```python
#Set the first slide transition as push and sound mode
presentation.Slides[0].SlideShowTransition.Type = TransitionType.Push
presentation.Slides[0].SlideShowTransition.SoundMode = TransitionSoundMode.StartSound

#Set the second slide transition as fade and set the speed 
presentation.Slides[1].SlideShowTransition.Type = TransitionType.Fade
presentation.Slides[1].SlideShowTransition.Speed = TransitionSpeed.Slow
```

---

# Spire.Presentation Python Add Line
## Add a line to a PowerPoint slide using Spire.Presentation library
```python
#Create a PPT document
presentation = Presentation()
#Get the first slide
slide = presentation.Slides[0]
#Add a line in the slide
line = slide.Shapes.AppendShape(ShapeType.Line, RectangleF.FromLTRB (50, 100, 350, 100))
#Set color of the line
line.ShapeStyle.LineColor.Color = Color.get_Red()
```

---

# Spire.Presentation for Python
## Add lines with arrows to PowerPoint slides
```python
#Create an instance of presentation document
ppt = Presentation()

#Add a line to the slides and set its color to red
shape = ppt.Slides[0].Shapes.AppendShape(ShapeType.Line, RectangleF.FromLTRB (150, 100, 250, 200))
shape.ShapeStyle.LineColor.Color = Color.get_Red()
#Set the line end type as StealthArrow
shape.Line.LineEndType = LineEndType.StealthArrow

#Add a line to the slides and use default color
shape = ppt.Slides[0].Shapes.AppendShape(ShapeType.Line, RectangleF.FromLTRB (300, 150, 400, 250))
shape.Rotation = -45
#Set the line end type as TriangleArrowHead
shape.Line.LineEndType = LineEndType.TriangleArrowHead

#Add a line to the slides and set its color to Green
shape = ppt.Slides[0].Shapes.AppendShape(ShapeType.Line, RectangleF.FromLTRB (450, 100, 550, 200))
shape.ShapeStyle.LineColor.Color = Color.get_Green()
shape.Rotation = 90
#Set the line begin type as TriangleArrowHead
shape.Line.LineBeginType = LineEndType.StealthArrow
```

---

# spire.presentation python shapes
## add lines with two points to presentation
```python
ppt = Presentation()
#Get the first slide
slide = ppt.Slides[0]
#Add line with two points
line = slide.Shapes.AppendShapeByPoint(ShapeType.Line, PointF(50.0, 50.0), PointF(150.0, 150.0))
line.ShapeStyle.LineColor.Color = Color.get_Red()
line = slide.Shapes.AppendShapeByPoint(ShapeType.Line, PointF(150.0, 150.0), PointF(250.0, 50.0))
line.ShapeStyle.LineColor.Color = Color.get_Blue()
```

---

# Spire.Presentation Python Round Corner Rectangle
## Add a round corner rectangle shape to a PowerPoint slide
```python
#Create an instance of presentation document
ppt = Presentation()
#Append a round corner rectangle and set its radius
shape = ppt.Slides[0].Shapes.AppendRoundRectangle(300, 90, 100, 200, 80)
#Set the color and fill style of shape
shape.Fill.FillType = FillFormatType.Solid
shape.Fill.SolidColor.Color = Color.get_LightBlue()
shape.ShapeStyle.LineColor.Color = Color.get_SkyBlue()
#Rotate the shape to 90 degree
shape.Rotation = 90
```

---

# spire.presentation python shapes
## add various shapes to presentation slides
```python
#Create PPT document
presentation = Presentation()
#Append new shape - Triangle and set style
shape = presentation.Slides[0].Shapes.AppendShape(ShapeType.Triangle, RectangleF.FromLTRB (115, 130, 215, 230))
shape.Fill.FillType = FillFormatType.Solid
shape.Fill.SolidColor.Color = Color.get_LightGreen()
shape.ShapeStyle.LineColor.Color = Color.get_White()
#Append new shape - Ellipse
shape = presentation.Slides[0].Shapes.AppendShape(ShapeType.Ellipse, RectangleF.FromLTRB (290, 130, 440, 230))
shape.Fill.FillType = FillFormatType.Solid
shape.Fill.SolidColor.Color = Color.get_LightSkyBlue()
shape.ShapeStyle.LineColor.Color = Color.get_White()
#Append new shape - Heart
shape = presentation.Slides[0].Shapes.AppendShape(ShapeType.Heart, RectangleF.FromLTRB (470, 130, 600, 230))
shape.Fill.FillType = FillFormatType.Solid
shape.Fill.SolidColor.Color = Color.get_Red()
shape.ShapeStyle.LineColor.Color = Color.get_LightGray()
#Append new shape - FivePointedStar
shape = presentation.Slides[0].Shapes.AppendShape(ShapeType.FivePointedStar, RectangleF.FromLTRB (90, 270, 240, 420))
shape.Fill.FillType = FillFormatType.Gradient
shape.Fill.SolidColor.Color = Color.get_Black()
shape.ShapeStyle.LineColor.Color = Color.get_White()
#Append new shape - Rectangle
shape = presentation.Slides[0].Shapes.AppendShape(ShapeType.Rectangle, RectangleF.FromLTRB (320, 290, 420, 410))
shape.Fill.FillType = FillFormatType.Solid
shape.Fill.SolidColor.Color = Color.get_Pink()
shape.ShapeStyle.LineColor.Color = Color.get_LightGray()
#Append new shape - BentUpArrow
shape = presentation.Slides[0].Shapes.AppendShape(ShapeType.BentUpArrow, RectangleF.FromLTRB (470, 300, 720, 400))
#Set the color of shape
shape.Fill.FillType = FillFormatType.Gradient
shape.Fill.Gradient.GradientStops.AppendByKnownColors(1, KnownColors.Olive)
shape.Fill.Gradient.GradientStops.AppendByKnownColors(0, KnownColors.PowderBlue)
shape.ShapeStyle.LineColor.Color = Color.get_Red()
```

---

# Spire.Presentation Python Shape Arrangement
## Arrange shapes in presentation slides
```python
#Create an instance of presentation document
ppt = Presentation()
#Get the specified shape
shape = ppt.Slides[0].Shapes[0]
#Bring the shape forward through SetShapeArrange method
shape.SetShapeArrange(ShapeArrange.BringForward)
```

---

# spire.presentation background
## Set background image for PowerPoint slide
```python
#Set background Image
ImageFile = "./Data/backgroundImg.png"
rect = RectangleF.FromLTRB(0, 0, presentation.SlideSize.Size.Width, presentation.SlideSize.Size.Height)
presentation.Slides[0].Shapes.AppendEmbedImageByPath (ShapeType.Rectangle, ImageFile, rect)
#Add title
left  = math.trunc(presentation.SlideSize.Size.Width / float(2)) - 200
rec_title = RectangleF.FromLTRB (left, 70, 380+left, 120)
shape_title = presentation.Slides[0].Shapes.AppendShape(ShapeType.Rectangle, rec_title)
shape_title.Line.FillType = FillFormatType.none
shape_title.Fill.FillType =FillFormatType.none
para_title = TextParagraph()
para_title.Text = "Background Sample"
para_title.Alignment = TextAlignmentType.Center
para_title.TextRanges[0].LatinFont = TextFont("Lucida Sans Unicode")
para_title.TextRanges[0].FontHeight = 36
para_title.TextRanges[0].Fill.FillType = FillFormatType.Solid
para_title.TextRanges[0].Fill.SolidColor.Color = Color.get_DarkSlateBlue()
shape_title.TextFrame.Paragraphs.Append(para_title)
#Add new shape to PPT document
left = math.trunc(presentation.SlideSize.Size.Width / float(2)) - 300
rec = RectangleF.FromLTRB (left, 155, 600+left, 355)
shape = presentation.Slides[0].Shapes.AppendShape(ShapeType.Rectangle, rec)
shape.Line.FillType = FillFormatType.none
shape.Fill.FillType = FillFormatType.none
para = TextParagraph()
para.Text = "Spire.Presentation for .NET support PPT, PPS, PPTX and PPSX presentation formats. It provides functions such as managing text, image, shapes, tables, animations, audio and video on slides. It also support exporting presentation slides to EMF, JPG, TIFF, PDF format etc."
para.TextRanges[0].Fill.FillType = FillFormatType.Solid
para.TextRanges[0].Fill.SolidColor.Color = Color.get_CadetBlue()
para.TextRanges[0].FontHeight = 26
shape.TextFrame.Paragraphs.Append(para)
```

---

# Spire.Presentation Python Shape Operations
## Copy shapes between slides
```python
#Define the source slide and target slide
sourceSlide = ppt.Slides[0]
targetSlide = ppt.Slides[1]
#Copy the first shape from the source slide to the target slide
targetSlide.Shapes.AddShape(sourceSlide.Shapes[0])
```

---

# Spire.Presentation Python Gradient Shape Fill
## Fill a shape with gradient colors in PowerPoint presentation
```python
ppt = Presentation()
#Get the first shape and set the style to be Gradient
GradientShape = ppt.Slides[0].Shapes[0] if isinstance(ppt.Slides[0].Shapes[0], IAutoShape) else None
GradientShape.Fill.FillType = FillFormatType.Gradient
GradientShape.Fill.Gradient.GradientStops.AppendByColor (0, Color.get_LightSkyBlue())
GradientShape.Fill.Gradient.GradientStops.AppendByColor(1, Color.get_LightGray())
```

---

# Spire.Presentation Python Pattern Fill
## Fill a shape with a pattern and set line properties
```python
#Set the pattern fill format 
shape.Fill.FillType = FillFormatType.Pattern
shape.Fill.Pattern.PatternType = PatternFillType.Trellis
shape.Fill.Pattern.BackgroundColor.Color = Color.get_DarkGray()
shape.Fill.Pattern.ForegroundColor.Color = Color.get_Yellow()
#Set the fill format of line
shape.Line.FillType = FillFormatType.Solid
shape.Line.SolidFillColor.Color = Color.get_Transparent()
```

---

# Spire.Presentation Fill Shape with Picture
## Fill a presentation shape with a picture
```python
# Get a shape from the slide
shape = ppt.Slides[0].Shapes[0] if isinstance(ppt.Slides[0].Shapes[0], IAutoShape) else None
# Fill the shape with picture
shape.Fill.FillType = FillFormatType.Picture
shape.Fill.PictureFill.Picture.Url = picUrl
shape.Fill.PictureFill.FillType = PictureFillType.Stretch
```

---

# Fill Shape with Solid Color
## Demonstrates how to fill a shape with solid color in a PowerPoint presentation
```python
#Create a PPT document
presentation = Presentation()
#Get the first slide
slide = presentation.Slides[0]
#Add a rectangle
left = int(presentation.SlideSize.Size.Width / 2) - 50
rect = RectangleF.FromLTRB (left, 100, 100+left, 200)
shape = slide.Shapes.AppendShape(ShapeType.Rectangle, rect)
#Fill shape with solid color
shape.Fill.FillType = FillFormatType.Solid
shape.Fill.SolidColor.Color = Color.get_Yellow()
#Set the fill format of line
shape.Line.FillType = FillFormatType.Solid
shape.Line.SolidFillColor.Color = Color.get_Gray()
```

---

# Find Shape by Alternative Text
## Core functionality to find a shape in a presentation slide by its alternative text
```python
#Get the first slide
slide = presentation.Slides[0]
#Find shape in the slide
for shape in slide.Shapes:
    #Find the shape whose alternative text is altText
    if shape.AlternativeText=="Shape1":
        #Process the found shape
        shape.Name
```

---

# spire.presentation python title extraction
## Extract all titles from a PowerPoint presentation
```python
# Instantiate a list of IShape objects
shapelist = []
# Loop through all slides and all shapes on each slide
for slide in ppt.Slides:
    for shape in slide.Shapes:
        if shape.Placeholder is not None:
            # Get all titles
            if shape.Placeholder.Type == PlaceholderType.Title:
                shapelist.append(shape)
            elif shape.Placeholder.Type == PlaceholderType.CenteredTitle:
                shapelist.append(shape)
            elif shape.Placeholder.Type == PlaceholderType.Subtitle:
                shapelist.append(shape)
# Loop through the list and get the inner text of all shapes in the list
sb = []
sb.append("Below are all the obtained titles:")
for i, unusedItem in enumerate(shapelist):
    shape1 = shapelist[i] if isinstance(shapelist[i], IAutoShape) else None
    sb.append (shape1.TextFrame.Text)
```

---

# spire.presentation python get shape group alt text
## Extract alternative text from shape groups in a presentation
```python
#Loop through slides and shapes
for slide in presentation.Slides:
    for shape in slide.Shapes:
        if isinstance(shape, GroupShape):
            #Find the shape group
            groupShape = shape if isinstance(shape, GroupShape) else None
            for gShape in groupShape.Shapes:
                #Append the alternative text in builder
                builder.append (gShape.AlternativeText)
```

---

# Spire.Presentation Python Shapes
## Get shapes by placeholder and extract text
```python
# Get placeholder from a slide
placeholder = slide.Shapes[0].Placeholder
# Get shapes by placeholder
shapes = slide.GetPlaceholderShapes(placeholder)
# Extract text from shapes
text = ""
for shape in shapes:
    # If shape is IAutoShape
    if isinstance(shape, IAutoShape):
        autoShape = shape
        if autoShape.TextFrame is not None:
            text += autoShape.TextFrame.Text + "\r\n"
```

---

# spire.presentation python shapes
## group shapes in presentation
```python
#Create two shapes in the slide
rectangle = slide.Shapes.AppendShape(ShapeType.Rectangle, RectangleF.FromLTRB (250, 180, 450, 220))
rectangle.Fill.FillType = FillFormatType.Solid
rectangle.Fill.SolidColor.KnownColor = KnownColors.SkyBlue
rectangle.Line.Width = 0.1
ribbon = slide.Shapes.AppendShape(ShapeType.Ribbon2, RectangleF.FromLTRB (290, 155, 410, 235))
ribbon.Fill.FillType = FillFormatType.Solid
ribbon.Fill.SolidColor.KnownColor = KnownColors.LightPink
ribbon.Line.Width = 0.1
#Add the two shape objects to an array list
arr = []
arr.append(rectangle)
arr.append(ribbon)
#Group the shapes in the list
ppt.Slides[0].GroupShapes(arr)
```

---

# spire.presentation python hide shape
## hide shape by alternative text
```python
#Loop through slides
for slide in presentation.Slides:
    #Loop through shapes in the slide
    for shape in slide.Shapes:
        #Find the shape whose alternative text is Shape1
        if shape.AlternativeText=="Shape1":
            #Hide the shape
            shape.IsHidden = True
```

---

# spire.presentation python IsTextBox
## determine if a shape is a text box in PowerPoint
```python
# Create an instance of presentation document
ppt = Presentation()
# Iterate through slides and shapes
for slide in ppt.Slides:
    for shape in slide.Shapes:
        if isinstance(shape, IAutoShape):
            # Judge if the shape is textbox
            isTextbox = shape.IsTextBox
```

---

# Operate placeholders in presentation
## This code demonstrates how to operate on different types of placeholders in a PowerPoint presentation
```python
for j, unusedItem in enumerate(presentation.Slides):
    slide = presentation.Slides[j]
    for i, unusedItem in enumerate(slide.Shapes):
        shape = slide.Shapes[i]
        if shape.Placeholder.Type == PlaceholderType.Media:
            shape.InsertVideo("video_path")
        elif shape.Placeholder.Type == PlaceholderType.Picture:
            shape.InsertPicture("image_path")
        elif shape.Placeholder.Type == PlaceholderType.Chart:
            shape.InsertChart(ChartType.ColumnClustered)
        elif shape.Placeholder.Type == PlaceholderType.Table:
            shape.InsertTable(3, 2)
        elif shape.Placeholder.Type == PlaceholderType.Diagram:
            shape.InsertSmartArt(SmartArtLayoutType.BasicBlockList)
```

---

# spire.presentation shape locking
## prevent or allow changing shape properties
```python
#The changes of selection and rotation are allowed
shape.Locking.RotationProtection = False
shape.Locking.SelectionProtection = False
#The changes of size, position, shape type, aspect ratio, text editing and ajust handles are not allowed 
shape.Locking.ResizeProtection = True
shape.Locking.PositionProtection = True
shape.Locking.ShapeTypeProtection = True
shape.Locking.AspectRatioProtection = True
shape.Locking.TextEditingProtection = True
shape.Locking.AdjustHandlesProtection = True
```

---

# Spire.Presentation Python Shape Management
## Remove shapes from PowerPoint slides based on alternative text
```python
# Loop through slides
for i, unusedItem in enumerate(presentation.Slides):
    slide = presentation.Slides[i]
    # Loop through shapes
    j = 0
    while j < slide.Shapes.Count:
        shape = slide.Shapes[j]
        # Find the shapes whose alternative text contain "Shape"
        if shape.AlternativeText.find("Shape") != -1:
            slide.Shapes.Remove(shape)
            j -= 1
        j += 1
```

---

# Reorder Overlapping Shapes in PowerPoint
## This code demonstrates how to reorder overlapping shapes in a PowerPoint presentation using the ZOrder method.
```python
#Create an instance of presentation document
ppt = Presentation()
#Get the first shape of the first slide
shape = ppt.Slides[0].Shapes[0]
#Change the shape's zorder
ppt.Slides[0].Shapes.ZOrder(1, shape)
```

---

# Reset Position of Placeholders in PowerPoint
## This code demonstrates how to reset the position of date and slide number placeholders in a PowerPoint slide.
```python
#Get the first slide from the presentation.
slide = presentation.Slides[0]
for shapeToMove in slide.Shapes:
    #Reset the position of the slide number to the left.
    if shapeToMove.Name.find ("Slide Number Placeholder") != -1:
        shapeToMove.Left = 0
    elif shapeToMove.Name.find ("Date Placeholder") != -1:
        #Reset the position of the date time to the center.
        shapeToMove.Left = math.trunc(presentation.SlideSize.Size.Width / float(2))
        #Reset the date time display style.
        ( shapeToMove if isinstance(shapeToMove, IAutoShape) else None).TextFrame.TextRange.Paragraph.Text = DateTime.get_Now().ToString("dd.MM.yyyy")
        ( shapeToMove if isinstance(shapeToMove, IAutoShape) else None).TextFrame.IsCentered = True
```

---

# spire.presentation python shape manipulation
## adapt shapes to new slide size by calculating size and position ratios
```python
#Define the original slide size
currentHeight = ppt.SlideSize.Size.Height
currentWidth = ppt.SlideSize.Size.Width
#Change the slide size as A3
ppt.SlideSize.Type = SlideSizeType.A3
#Define the new slide size
newHeight = ppt.SlideSize.Size.Height
newWidth = ppt.SlideSize.Size.Width
#Define the ratio from the old and new slide size
ratioHeight = newHeight / currentHeight
ratioWidth = newWidth / currentWidth
#Reset the size and position of the shape on the slide
for slide in ppt.Slides:
    for shape in slide.Shapes:
        shape.Height = shape.Height * ratioHeight
        shape.Width = shape.Width * ratioWidth
        shape.Left = shape.Left * ratioHeight
        shape.Top = shape.Top * ratioWidth
```

---

# Spire.Presentation Python Shape Rotation
## Rotate shapes in a PowerPoint presentation
```python
#Get the shapes 
shape = ppt.Slides[0].Shapes[0] if isinstance(ppt.Slides[0].Shapes[0], IAutoShape) else None
#Set the rotation
shape.Rotation = 60
(ppt.Slides[0].Shapes[1] if isinstance(ppt.Slides[0].Shapes[1], IAutoShape) else None).Rotation = 120
(ppt.Slides[0].Shapes[2] if isinstance(ppt.Slides[0].Shapes[2], IAutoShape) else None).Rotation = 180
(ppt.Slides[0].Shapes[3] if isinstance(ppt.Slides[0].Shapes[3], IAutoShape) else None).Rotation = 240
```

---

# Spire.Presentation 3D Effects
## Set 3D effects for shapes in PowerPoint presentations
```python
# Create shape with 3D effect
shape = ppt.Slides[0].Shapes.AppendShape(ShapeType.RoundCornerRectangle, RectangleF.FromLTRB (150, 150, 300, 300))
shape.Fill.FillType = FillFormatType.Solid
shape.Fill.SolidColor.KnownColor = KnownColors.SkyBlue

# Apply 3D effect properties
effect = shape.ThreeD.ShapeThreeD
effect.PresetMaterial = PresetMaterialType.Powder
effect.TopBevel.PresetType = BevelPresetType.ArtDeco
effect.TopBevel.Height = 4
effect.TopBevel.Width = 12
effect.BevelColorMode = BevelColorType.Contour
effect.ContourColor.KnownColor = KnownColors.LightBlue
effect.ContourWidth = 3.5

# Create another shape with different 3D effect
shape2 = ppt.Slides[0].Shapes.AppendShape(ShapeType.Pentagon, RectangleF.FromLTRB (400, 150, 550, 300))
shape2.Fill.FillType = FillFormatType.Solid
shape2.Fill.SolidColor.KnownColor = KnownColors.LightGreen

# Apply different 3D effect properties
effect2 = shape2.ThreeD.ShapeThreeD
effect2.PresetMaterial = PresetMaterialType.SoftEdge
effect2.TopBevel.PresetType = BevelPresetType.SoftRound
effect2.TopBevel.Height = 12
effect2.TopBevel.Width = 12
effect2.BevelColorMode = BevelColorType.Contour
effect2.ContourColor.KnownColor = KnownColors.LawnGreen
effect2.ContourWidth = 5
```

---

# spire.presentation python alternative text
## set and get alternative text for shapes in presentation
```python
#Get the first slide
slide = ppt.Slides[0]
#Set the alternative text (title and description)
slide.Shapes[0].AlternativeTitle = "Rectangle"
slide.Shapes[0].AlternativeText = "This is a Rectangle"
#Get the alternative text (title and description)
alternativeText = ""
title = slide.Shapes[0].AlternativeTitle
alternativeText += "Title: " + title + "\r\n"
description = slide.Shapes[0].AlternativeText
alternativeText += "Description: " + description
```

---

# spire.presentation python ellipse formatting
## set ellipse shape format in presentation
```python
#Create a PPT document
presentation = Presentation()
#Get the first slide
slide = presentation.Slides[0]
#Add an ellipse shape
left = 100  # Simplified position calculation
rect = RectangleF.FromLTRB(left, 100, 200+left, 200)
shape = slide.Shapes.AppendShape(ShapeType.Ellipse, rect)
#Set the fill format of shape
shape.Fill.FillType = FillFormatType.Solid
shape.Fill.SolidColor.Color = Color.get_CadetBlue()
#Set the line format of shape
shape.Line.FillType = FillFormatType.Solid
shape.Line.SolidFillColor.Color = Color.get_DimGray()
```

---

# Spire.Presentation Python Line Formatting
## Demonstrates how to set format for lines of shapes in a presentation
```python
#Add a rectangle shape to the slide
shape = ppt.Slides[0].Shapes.AppendShape(ShapeType.Rectangle, RectangleF.FromLTRB (100, 150, 300, 250))
#Set the fill color of the rectangle shape
shape.Fill.FillType = FillFormatType.Solid
shape.Fill.SolidColor.Color = Color.get_White()
#Apply some formatting on the line of the rectangle
shape.Line.Style = TextLineStyle.ThickThin
shape.Line.Width = 5
shape.Line.DashStyle = LineDashStyleType.Dash
#Set the color of the line of the rectangle
shape.ShapeStyle.LineColor.Color = Color.get_SkyBlue()
#Add a ellipse shape to the slide
shape = ppt.Slides[0].Shapes.AppendShape(ShapeType.Ellipse, RectangleF.FromLTRB (400, 150, 600, 250))
#Set the fill color of the ellipse shape
shape.Fill.FillType = FillFormatType.Solid
shape.Fill.SolidColor.Color = Color.get_White()
#Apply some formatting on the line of the ellipse
shape.Line.Style = TextLineStyle.ThickBetweenThin
shape.Line.Width = 5
shape.Line.DashStyle = LineDashStyleType.DashDot
#Set the color of the line of the ellipse
shape.ShapeStyle.LineColor.Color = Color.get_OrangeRed()
```

---

# Spire.Presentation Line Join Styles
## Set different line join styles for shapes in a presentation
```python
#Add three shapes
shape1 = slide.Shapes.AppendShape(ShapeType.Rectangle, RectangleF.FromLTRB (50, 150, 200, 200))
shape2 = slide.Shapes.AppendShape(ShapeType.Rectangle, RectangleF.FromLTRB (250, 150, 400, 200))
shape3 = slide.Shapes.AppendShape(ShapeType.Rectangle, RectangleF.FromLTRB (450, 150, 600, 200))
#Fill shapes
shape1.Fill.FillType = FillFormatType.Solid
shape1.Fill.SolidColor.Color = Color.get_CadetBlue()
shape2.Fill.FillType = FillFormatType.Solid
shape2.Fill.SolidColor.Color = Color.get_CadetBlue()
shape3.Fill.FillType = FillFormatType.Solid
shape3.Fill.SolidColor.Color = Color.get_CadetBlue()
#Fill lines of shapes
shape1.Line.FillType = FillFormatType.Solid
shape1.Line.SolidFillColor.Color = Color.get_DarkGray()
shape2.Line.FillType = FillFormatType.Solid
shape2.Line.SolidFillColor.Color = Color.get_DarkGray()
shape3.Line.FillType = FillFormatType.Solid
shape3.Line.SolidFillColor.Color = Color.get_DarkGray()
#Set the line width
shape1.Line.Width = 10
shape2.Line.Width = 10
shape3.Line.Width = 10
#Set the join styles of lines
shape1.Line.JoinStyle = LineJoinType.Bevel
shape2.Line.JoinStyle = LineJoinType.Miter
shape3.Line.JoinStyle = LineJoinType.Round
#Add text in shapes
shape1.TextFrame.Text = "Bevel Join Style"
shape2.TextFrame.Text = "Miter Join Style"
shape3.TextFrame.Text = "Round Join Style"
```

---

# Spire.Presentation Shape Outline and Effects
## Setting outline colors and effects for presentation shapes
```python
#Get the first slide
slide = ppt.Slides[0]

#Draw a Rectangle shape
shape = slide.Shapes.AppendShape(ShapeType.Rectangle, RectangleF.FromLTRB (150, 180, 250, 230))
shape.Fill.FillType = FillFormatType.Solid
shape.Fill.SolidColor.Color = Color.get_SkyBlue()
#Set outline color
shape.ShapeStyle.LineColor.Color = Color.get_Red()
#Set shadow effect
shadow = PresetShadow()
shadow.ColorFormat.Color = Color.get_LightSkyBlue()
shadow.Preset = PresetShadowValue.FrontRightPerspective
shadow.Distance = 10.0
shadow.Direction = 225.0
shape.EffectDag.PresetShadowEffect = shadow

#Draw a Ellipse shape
shape = slide.Shapes.AppendShape(ShapeType.Ellipse, RectangleF.FromLTRB (400, 150, 500, 250))
shape.Fill.FillType = FillFormatType.Solid
shape.Fill.SolidColor.Color = Color.get_SkyBlue()
#Set outline color
shape.ShapeStyle.LineColor.Color = Color.get_Yellow()
#Set shadow effect
glow = GlowEffect()
glow.ColorFormat.Color = Color.get_LightPink()
glow.Radius = 20.0
shape.EffectDag.GlowEffect = glow
```

---

# Spire.Presentation Python Rounded Rectangle
## Set radius and properties of rounded rectangles in PowerPoint slides
```python
#Create a PPT document
presentation = Presentation()
#Insert a rounded rectangle and set its radius
presentation.Slides[0].Shapes.InsertRoundRectangle(0, 160, 180, 100, 200, 10)
#Append a rounded rectangle and set its radius
shape = presentation.Slides[0].Shapes.AppendRoundRectangle(380, 180, 100, 200, 100)
#Set the color and fill style of shape
shape.Fill.FillType = FillFormatType.Solid
shape.Fill.SolidColor.Color = Color.get_SeaGreen()
shape.ShapeStyle.LineColor.Color = Color.get_White()
#Rotate the shape to 90 degree
shape.Rotation = 90
```

---

# spire.presentation python shape formatting
## set rectangle format in presentation
```python
#Create a PPT document
presentation = Presentation()
#Add a shape
left = math.trunc(presentation.SlideSize.Size.Width / float(2)) - 100
rect = RectangleF.FromLTRB(left, 100, 200+left, 200)
shape = presentation.Slides[0].Shapes.AppendShape(ShapeType.Rectangle, rect)
#Set the fill format of shape
shape.Fill.FillType = FillFormatType.Solid
shape.Fill.SolidColor.Color = Color.get_CadetBlue()
#Set the fill format of line
shape.Line.FillType = FillFormatType.Solid
shape.Line.SolidFillColor.Color = Color.get_DimGray()
```

---

# Spire.Presentation Python Shadow Effect
## Set shadow effect for a shape in presentation
```python
#Add a shape to slide
rect1 = RectangleF.FromLTRB (200, 150, 500, 270)
shape = slide.Shapes.AppendShape(ShapeType.Rectangle, rect1)
shape.Fill.FillType = FillFormatType.Solid
shape.Fill.SolidColor.Color = Color.get_LightBlue()
shape.Line.FillType = FillFormatType.none
shape.TextFrame.Text = "This demo shows how to apply shadow effect to shape."
shape.TextFrame.TextRange.Fill.FillType = FillFormatType.Solid
shape.TextFrame.TextRange.Fill.SolidColor.Color = Color.get_Black()

#Create an inner shadow effect through InnerShadowEffect object
innerShadow = InnerShadowEffect()
innerShadow.BlurRadius = 20
innerShadow.Direction = 0
innerShadow.Distance = 0
innerShadow.ColorFormat.Color = Color.get_Black()

#Apply the shadow effect to shape
shape.EffectDag.InnerShadowEffect = innerShadow
```

---

# Spire.Presentation Shape to Image Conversion
## Convert shapes in a PowerPoint slide to image files
```python
inputFile = "./Data/ShapeToImage.pptx"
outputFolder = "output"

# Create a PPT document
presentation = Presentation()
presentation.LoadFromFile(inputFile)
for i, unusedItem in enumerate(presentation.Slides[0].Shapes):
    fileName =outputFolder + "//" + "ShapeToImage-"+str(i)+".png"
    # Save shapes as images
    image = presentation.Slides[0].Shapes.SaveAsImage(i)
    image.Save(fileName)
    image.Dispose()
presentation.Dispose()
```

---

# Spire.Presentation Python Shape Operations
## Ungroup shapes in a PowerPoint presentation
```python
groupShape = ppt.Slides[0].Shapes[0] if isinstance(ppt.Slides[0].Shapes[0], GroupShape) else None
#Ungroup the shapes
ppt.Slides[0].Ungroup(groupShape)
```

---

# Spire.Presentation Python Animation
## Add exit animation to a shape
```python
#Create an instance of presentation document
ppt = Presentation()
#Get the first slide
slide = ppt.Slides[0]
#Add a shape to the slide
starShape = slide.Shapes.AppendShape(ShapeType.FivePointedStar, RectangleF.FromLTRB (250, 100, 450, 300))
starShape.Fill.FillType = FillFormatType.Solid
starShape.Fill.SolidColor.KnownColor = KnownColors.LightBlue
#Add random bars effect to the shape
effect = slide.Timeline.MainSequence.AddEffect(starShape, AnimationEffectType.RandomBars)
#Change effect type from entrance to exit
effect.PresetClassType = TimeNodePresetClassType.Exit
```

---

# Spire.Presentation Animation Effects
## Create and configure animations for PowerPoint slides and shapes
```python
# Set slide transition animation
presentation.Slides[0].SlideShowTransition.Type = TransitionType.Circle

# Add animation effect to a triangle shape
shape = presentation.Slides[0].Shapes.AppendShape(ShapeType.Triangle, RectangleF.FromLTRB (100, 280, 180, 360))
shape.Fill.FillType = FillFormatType.Solid
shape.Fill.SolidColor.Color = Color.get_CadetBlue()
shape.ShapeStyle.LineColor.Color = Color.get_White()
shape.Slide.Timeline.MainSequence.AddEffect(shape, AnimationEffectType.Path4PointStar)

# Add animation effect to a rectangle shape
shape = presentation.Slides[0].Shapes.AppendShape(ShapeType.Rectangle, RectangleF.FromLTRB (210, 280, 360, 360))
shape.Fill.FillType = FillFormatType.Solid
shape.Fill.SolidColor.Color = Color.get_CadetBlue()
shape.ShapeStyle.LineColor.Color = Color.get_White()
shape.AppendTextFrame("Animated Shape")
shape.Slide.Timeline.MainSequence.AddEffect(shape, AnimationEffectType.FadedSwivel)

# Add animation effect to a cloud shape
shape = presentation.Slides[0].Shapes.AppendShape(ShapeType.Cloud, RectangleF.FromLTRB (390, 280, 470, 360))
shape.Fill.FillType = FillFormatType.Solid
shape.Fill.SolidColor.Color = Color.get_White()
shape.ShapeStyle.LineColor.Color = Color.get_CadetBlue()
shape.Slide.Timeline.MainSequence.AddEffect(shape, AnimationEffectType.FadedZoom)
```

---

# spire.presentation python animation
## apply animation effect to shape
```python
#Create an instance of presentation document
ppt = Presentation()
#Get the first slide
slide = ppt.Slides[0]
#Insert a rectangle in the slide and fill the shape
shape = slide.Shapes.AppendShape(ShapeType.Rectangle, RectangleF.FromLTRB (100, 150, 300, 230))
shape.Fill.FillType = FillFormatType.Solid
shape.Fill.SolidColor.Color = Color.get_LightBlue()
shape.ShapeStyle.LineColor.Color = Color.get_White()
shape.AppendTextFrame("Animated Shape")
#Apply FadedSwivel animation effect to the shape
shape.Slide.Timeline.MainSequence.AddEffect(shape, AnimationEffectType.FadedSwivel)
```

---

# Spire.Presentation Python Animation
## Apply animation to text in PowerPoint presentation
```python
#Create an instance of presentation document
ppt = Presentation()
#Get the first slide
slide = ppt.Slides[0]
#Add a shape to the slide
shape = slide.Shapes.AppendShape(ShapeType.Rectangle, RectangleF.FromLTRB (250, 150, 450, 250))
shape.Fill.FillType = FillFormatType.Solid
shape.Fill.SolidColor.Color = Color.get_LightBlue()
shape.ShapeStyle.LineColor.Color = Color.get_White()
shape.AppendTextFrame("This demo shows how to apply animation on text in PPT document.")
#Apply animation to the text in shape
animation = shape.Slide.Timeline.MainSequence.AddEffect(shape, AnimationEffectType.Float)
animation.SetStartEndParagraphs(0, 0)
```

---

# spire.presentation custom path animation
## create custom path animation for shapes in PowerPoint
```python
#Add shape
shape = ppt.Slides[0].Shapes.AppendShape(ShapeType.Rectangle, RectangleF.FromLTRB(0, 0, 200, 200))
#Add animation
effect = ppt.Slides[0].Timeline.MainSequence.AddEffect(shape, AnimationEffectType.PathUser)
common = effect.CommonBehaviorCollection
motion = common[0]
motion.Origin = AnimationMotionOrigin.Layout
motion.PathEditMode = AnimationMotionPathEditMode.Relative
#Add motion path
moinPath = MotionPath()
p1=PointF(0.0,0.0)
p2=PointF(0.1,0.1)
p3=PointF(-0.1,0.2)
moinPath.Add(MotionCommandPathType.MoveTo, [p1], MotionPathPointsType.CurveAuto, True)
moinPath.Add(MotionCommandPathType.LineTo, [p2], MotionPathPointsType.CurveAuto, True)
moinPath.Add(MotionCommandPathType.LineTo, [p3], MotionPathPointsType.CurveAuto, True)
moinPath.Add(MotionCommandPathType.End, [], MotionPathPointsType.CurveStraight, True)
motion.Path = moinPath
```

---

# Spire.Presentation Animation Timing Control
## Set duration and delay time for animations in a presentation
```python
# Get the first slide
slide = presentation.Slides[0]
animations = slide.Timeline.MainSequence
# Get duration time of animation
durationTime = animations[0].Timing.Duration
# Set new duration time of animation
animations[0].Timing.Duration = 0.8
# Get delay time of animation
delayTime = animations[0].Timing.TriggerDelayTime
# Set new delay time of animation
animations[0].Timing.TriggerDelayTime = 0.6
```

---

# spire.presentation animation effect info
## get animation effect information from presentation slides
```python
# Iterate through each slide
for slide in presentation.Slides:
    for effect in slide.Timeline.MainSequence:
        # Get the animation effect type
        animationEffectType = effect.AnimationEffectType
        # Get the slide number where the animation is located
        slideNumber = slide.SlideNumber
        # Get the shape name
        shapeName = effect.ShapeTarget.Name
```

---

# Get Animation Motion Paths from PowerPoint
## Extract motion path data from animations in a PowerPoint presentation
```python
presentation = Presentation()
presentation.LoadFromFile(inputFile)
slide = presentation.Slides[0]
#Get the first shape
shape = slide.Shapes[0]
#Create a list to save the tracks
sb = []
i = 1
#Traverse all animations
for effect in shape.Slide.Timeline.MainSequence:
    if effect.ShapeTarget.Id==shape.Id:
        #Get MotionPath
        path = (effect.CommonBehaviorCollection[0]).Path        
        #Get all points in the path
        for motionCmdPath in path:
            points = motionCmdPath.Points
            comType = motionCmdPath.CommandType
            if points is not None:
                for point in points:
                    sb.append(str(i) + "  MotionType: " + str(comType )+ " -> X: " + str(point.X) + ", Y: " + str(point.Y))
                i += 1
fp = open(outputFile,"w")
for s in sb:
    fp.write(s + "\n")
presentation.Dispose()
```

---

# spire.presentation python animation
## set animation for animate text
```python
#Create an instance of presentation document
ppt = Presentation()
#Set the AnimateType as Letter
ppt.Slides[0].Timeline.MainSequence[0].IterateType = AnimateType.Letter
#Set the IterateTimeValue for the animate text
ppt.Slides[0].Timeline.MainSequence[0].IterateTimeValue = 10
```

---

# spire.presentation python animation
## set animation repeat type
```python
#Get the first slide
slide = presentation.Slides[0]
animations = slide.Timeline.MainSequence
animations[0].Timing.AnimationRepeatType = AnimationRepeatType.UtilEndOfSlide
```

---

# Spire.Presentation Section Management
## Add sections to PowerPoint presentation
```python
# Create a PPT document
ppt = Presentation()
# Get the second slide
slide = ppt.Slides[1]
# Append section with section name at the end
ppt.SectionList.Append("E-iceblue01")
# Add section with slide
ppt.SectionList.Add("section1", slide)
```

---

# spire.presentation python section management
## Add slide to section in PowerPoint presentation
```python
#Add a new shape to the PPT document
presentation.Slides[0].Shapes.AppendShape(ShapeType.Rectangle, RectangleF.FromLTRB (200, 50, 500, 150))
#Create a new section and copy the first slide to it
NewSection = presentation.SectionList.Append("New Section")
NewSection.Insert(0, presentation.Slides[0])
```

---

# spire.presentation python delete sections
## delete all sections from a powerpoint presentation
```python
#Create a PPT document
ppt = Presentation()
#Remove all sections
ppt.SectionList.RemoveAll()
```

---

# Get section index in PowerPoint presentation
## Retrieve the index of a specific section in a PowerPoint document
```python
# Get the first section
section = ppt.SectionList[0]
# Get the index of the section
index = ppt.SectionList.IndexOf(section)
```

---

# Spire.Presentation Python Load From Stream
## Load PowerPoint document from stream
```python
#Create an instance of presentation document
ppt = Presentation()
#Load PowerPoint file from stream
from_stream = Stream("path_to_file.pptx")
ppt.LoadFromStream(from_stream, FileFormat.Pptx2013)
from_stream.Dispose()
ppt.Dispose()
```

---

# Spire.Presentation Python Loop Presentation
## Configure PowerPoint presentation to loop continuously with animations and narrations
```python
#Set the Boolean value of ShowLoop as true
ppt.ShowLoop = True
#Set the PowerPoint document to show animation and narration
ppt.ShowAnimation = True
ppt.ShowNarration = True
#Use slide transition timings to advance slide
ppt.UseTimings = True
```

---

# spire.presentation python page setup
## Set up slide size, orientation, and type in a presentation
```python
#Create PPT document
presentation = Presentation()
#Set the size of slides
size = SizeF(600.0,600.0)
presentation.SlideSize.Size = size
presentation.SlideSize.Orientation = SlideOrienation.Portrait
presentation.SlideSize.Type = SlideSizeType.Custom
#Set background image
rect = RectangleF.FromLTRB (0, 0, presentation.SlideSize.Size.Width, presentation.SlideSize.Size.Height)
presentation.Slides[0].Shapes.AppendEmbedImageByPath(ShapeType.Rectangle, "path_to_image", rect)
presentation.Slides[0].Shapes[0].Line.FillFormat.SolidFillColor.Color = Color.get_FloralWhite()
#Append new shape
left = math.trunc(presentation.SlideSize.Size.Width / float(2)) - 200
rec = RectangleF.FromLTRB (left, 150, 400+left, 350)
shape = presentation.Slides[0].Shapes.AppendShape(ShapeType.Rectangle, rec)
shape.ShapeStyle.LineColor.Color = Color.get_White()
shape.Fill.FillType = FillFormatType.none
#Add text to shape
shape.AppendTextFrame("The sample demonstrates how to set slide size.")
shape.TextFrame.Paragraphs[0].TextRanges[0].LatinFont = TextFont("Myriad Pro")
shape.TextFrame.Paragraphs[0].TextRanges[0].FontHeight = 24
shape.TextFrame.Paragraphs[0].TextRanges[0].Fill.FillType = FillFormatType.Solid
shape.TextFrame.Paragraphs[0].TextRanges[0].Fill.SolidColor.Color = Color.FromArgb(255,36, 64, 97)
```

---

# Spire.Presentation Save to Stream
## Demonstrates how to save a PowerPoint presentation to a stream
```python
#Create PowerPoint presentation
presentation = Presentation()
#Save to Stream
stream = Stream()
presentation.SaveToFile(stream, FileFormat.Pptx2013)
stream.Close()
presentation.Dispose()
```

---

# spire.presentation python kiosk mode
## set presentation show type as kiosk
```python
#Create an instance of presentation document
ppt = Presentation()
#Specify the presentation show type as kiosk
ppt.ShowType = SlideShowType.Kiosk
```

---

# Spire.Presentation Python Split PPT
## Split a PowerPoint presentation into individual slides
```python
#Create an instance of presentation document
ppt = Presentation()
#Load file
ppt.LoadFromFile(inputFile)
for i, slide in enumerate(ppt.Slides):
    #Initialize another instance of Presentation, and remove the blank slide
    newppt = Presentation()
    newppt.Slides.RemoveAt(0)
    #Append the specified slide from old presentation to the new one
    newppt.Slides.AppendBySlide(slide)
    #Save the document
    result = outputFolder + "//" + "SplitPPT-" + str(i) + ".pptx"
    newppt.SaveToFile(result, FileFormat.Pptx2010)
    newppt.Dispose()
ppt.Dispose()
```

---

# spire.presentation python builtin properties
## get builtin properties from presentation
```python
#Create PPT document
presentation = Presentation()
#Load the PPT document from disk
presentation.LoadFromFile(inputFile)
#Get the builtin properties 
application = presentation.DocumentProperty.Application
author = presentation.DocumentProperty.Author
company = presentation.DocumentProperty.Company
keywords = presentation.DocumentProperty.Keywords
comments = presentation.DocumentProperty.Comments
category = presentation.DocumentProperty.Category
title = presentation.DocumentProperty.Title
subject = presentation.DocumentProperty.Subject
```

---

# Spire.Presentation Mark Document as Final
## This code demonstrates how to mark a PowerPoint presentation as final using Spire.Presentation for Python
```python
#Create a PPT document
presentation = Presentation()
#Load the document from disk
presentation.LoadFromFile(inputFile)
#Mark the document as final
presentation.DocumentProperty.MarkAsFinal = True
#Save the document
presentation.SaveToFile(outputFile, FileFormat.Pptx2010)
presentation.Dispose()
```

---

# spire.presentation properties
## set document properties for PowerPoint presentation
```python
#Set the DocumentProperty of PPT document
presentation.DocumentProperty.Application = "Spire.Presentation"
presentation.DocumentProperty.Author = "E-iceblue"
presentation.DocumentProperty.Company = "E-iceblue Co., Ltd."
presentation.DocumentProperty.Keywords = "Demo File"
presentation.DocumentProperty.Comments = "This file is used to test Spire.Presentation."
presentation.DocumentProperty.Category = "Demo"
presentation.DocumentProperty.Title = "This is a demo file."
presentation.DocumentProperty.Subject = "Test"
```

---

# Spire.Presentation document properties
## Set properties for presentation template
```python
def SetPropertiesForTemplate(filePath, fileFormat):
    # Create a document
    presentation = Presentation()
    # Set the DocumentProperty 
    presentation.DocumentProperty.Application = "Spire.Presentation"
    presentation.DocumentProperty.Author = "E-iceblue"
    presentation.DocumentProperty.Company = "E-iceblue Co., Ltd."
    presentation.DocumentProperty.Keywords = "Demo File"
    presentation.DocumentProperty.Comments = "This file is used to test Spire.Presentation."
    presentation.DocumentProperty.Category = "Demo"
    presentation.DocumentProperty.Title = "This is a demo file."
    presentation.DocumentProperty.Subject = "Test"
    # Save to template file
    presentation.SaveToFile(filePath, fileFormat)
    presentation.Dispose()
```

---

# Spire.Presentation Password Protection Check
## Check if a PowerPoint file is password protected
```python
# Create Presentation
presentation = Presentation()
# Check whether a PPT document is password protected
isProtected = presentation.IsPasswordProtected(inputFile)
presentation.Dispose()
```

---

# Spire.Presentation Encryption
## Encrypt a PowerPoint presentation with password
```python
#Create a PPT document
presentation = Presentation()
#Load the document from disk
presentation.LoadFromFile(inputFile)
#Get the password that the user entered
password = "e-iceblue"
#Encrypt the document with the password
presentation.Encrypt(password)
```

---

# Spire.Presentation Password Modification
## Core functionality to modify the password of an encrypted PowerPoint presentation
```python
#Create a PowerPoint document.
presentation = Presentation()
#Remove the encryption.
presentation.RemoveEncryption()
#Protect the document by setting a new password.
presentation.Protect("654321")
```

---

# Spire.Presentation Python Open Encrypted PPT
## How to open an encrypted PowerPoint presentation and save it as a new file
```python
inputFile = "./Data/OpenEncryptedPPT.pptx"
outputFile = "OpenEncryptedPPT_out.pptx"

#Create a PPT document
presentation = Presentation()
#Load the PPT with password
presentation.LoadFromFile(inputFile, FileFormat.Pptx2010, "123456")
#Save as a new PPT with original password
presentation.SaveToFile(outputFile, FileFormat.Pptx2010)
presentation.Dispose()
```

---

# Spire.Presentation Digital Signature Removal
## Remove all digital signatures from a PowerPoint presentation
```python
#Create a PowerPoint document.
ppt = Presentation()
#Remove all digital signatures
if ppt.IsDigitallySigned == True:
    ppt.RemoveAllDigitalSignatures()
```

---

# spire.presentation remove encryption
## Remove encryption from PowerPoint presentation
```python
#Create a PowerPoint document.
presentation = Presentation()
#Load the encrypted file from disk.
presentation.LoadFromFile(inputFile, "123456")
#Remove encryption.
presentation.RemoveEncryption()
#Save to file.
presentation.SaveToFile(outputFile, FileFormat.Pptx2013)
presentation.Dispose()
```

---

# Spire.Presentation Document Protection
## Set PowerPoint document to read-only with password protection
```python
# Get the password that the user entered
password = "e-iceblue"
# Protect the document with the password
presentation.Protect(password)
```

---

# Spire.Presentation Background Setting
## Set different types of backgrounds for PowerPoint slides
```python
# Set the background of the first slide to Gradient color
presentation.Slides[0].SlideBackground.Type = BackgroundType.Custom
presentation.Slides[0].SlideBackground.Fill.FillType = FillFormatType.Gradient
presentation.Slides[0].SlideBackground.Fill.Gradient.GradientShape = GradientShapeType.Linear
presentation.Slides[0].SlideBackground.Fill.Gradient.GradientStyle = GradientStyle.FromCorner1
presentation.Slides[0].SlideBackground.Fill.Gradient.GradientStops.AppendByKnownColors(1, KnownColors.SkyBlue)
presentation.Slides[0].SlideBackground.Fill.Gradient.GradientStops.AppendByKnownColors(0, KnownColors.White)
```

## Set solid color background
```python
# Set the background of the second slide to Solid color
presentation.Slides[1].SlideBackground.Type = BackgroundType.Custom
presentation.Slides[1].SlideBackground.Fill.FillType = FillFormatType.Solid
presentation.Slides[1].SlideBackground.Fill.SolidColor.Color = Color.get_SkyBlue()
```

## Set picture background
```python
# Set the background of the third slide to picture
presentation.Slides[2].SlideBackground.Type = BackgroundType.Custom
presentation.Slides[2].SlideBackground.Fill.FillType = FillFormatType.Picture
presentation.Slides[2].SlideBackground.Fill.PictureFill.FillType = PictureFillType.Stretch
presentation.Slides[2].SlideBackground.Fill.PictureFill.Picture.EmbedImage = imageData
```

---

# spire.presentation python gradient background
## set gradient background for presentation slide
```python
#Get the first slide
slide = presentation.Slides[0]
#Set the background to gradient
slide.SlideBackground.Type = BackgroundType.Custom
slide.SlideBackground.Fill.FillType = FillFormatType.Gradient
#Add gradient stops
slide.SlideBackground.Fill.Gradient.GradientStops.AppendByColor(0.1, Color.get_LightSeaGreen())
slide.SlideBackground.Fill.Gradient.GradientStops.AppendByColor(0.7, Color.get_LightCyan())
#Set gradient shape type
slide.SlideBackground.Fill.Gradient.GradientShape = GradientShapeType.Linear
#Set the angle
slide.SlideBackground.Fill.Gradient.LinearGradientFill.Angle = 45
```

---

# spire.presentation master background
## set master slide background with solid color
```python
#Create a PPT document
presentation = Presentation()

#Set the slide background of master
presentation.Masters[0].SlideBackground.Type = BackgroundType.Custom
presentation.Masters[0].SlideBackground.Fill.FillType = FillFormatType.Solid
presentation.Masters[0].SlideBackground.Fill.SolidColor.Color = Color.get_LightSalmon()
```

---

# Spire.Presentation Error Bars Formatting
## Add and format vertical and horizontal error bars in PowerPoint charts
```python
# Get the column chart on the first slide and set chart title.
columnChart = presentation.Slides[0].Shapes[0] if isinstance(presentation.Slides[0].Shapes[0], IChart) else None
columnChart.ChartTitle.TextProperties.Text = "Vertical Error Bars"

# Add Y (Vertical) Error Bars.
# Get Y error bars of the first chart series.
errorBarsYFormat1 = columnChart.Series[0].ErrorBarsYFormat

# Set end cap.
errorBarsYFormat1.ErrorBarNoEndCap = False

# Specify direction.
errorBarsYFormat1.ErrorBarSimType = ErrorBarSimpleType.Plus

# Specify error amount type.
errorBarsYFormat1.ErrorBarvType = ErrorValueType.StandardError

# Set value.
errorBarsYFormat1.ErrorBarVal = 0.3

# Set line format.
errorBarsYFormat1.Line.FillType = FillFormatType.Solid
errorBarsYFormat1.Line.SolidFillColor.Color = Color.get_MediumVioletRed()
errorBarsYFormat1.Line.Width = 1

# Get the bubble chart on the second slide and set chart title.
bubbleChart = presentation.Slides[1].Shapes[0] if isinstance(presentation.Slides[1].Shapes[0], IChart) else None
bubbleChart.ChartTitle.TextProperties.Text = "Vertical and Horizontal Error Bars"

# Add X (Horizontal) and Y (Vertical) Error Bars.
# Get X error bars of the first chart series.
errorBarsXFormat = bubbleChart.Series[0].ErrorBarsXFormat

# Set end cap.
errorBarsXFormat.ErrorBarNoEndCap = False

# Specify direction.
errorBarsXFormat.ErrorBarSimType = ErrorBarSimpleType.Both

# Specify error amount type.
errorBarsXFormat.ErrorBarvType = ErrorValueType.StandardError

# Set value.
errorBarsXFormat.ErrorBarVal = 0.3

# Get Y error bars of the first chart series.
errorBarsYFormat2 = bubbleChart.Series[0].ErrorBarsYFormat

# Set end cap.
errorBarsYFormat2.ErrorBarNoEndCap = False

# Specify direction.
errorBarsYFormat2.ErrorBarSimType = ErrorBarSimpleType.Both

# Specify error amount type.
errorBarsYFormat2.ErrorBarvType = ErrorValueType.StandardError

# Set value.
errorBarsYFormat2.ErrorBarVal = 0.3
```

---

# Spire.Presentation Python Error Bars
## Add custom error bars to a chart in PowerPoint presentation
```python
#Get the bubble chart on the first slide
bubbleChart = ppt.Slides[0].Shapes[0] if isinstance(ppt.Slides[0].Shapes[0], IChart) else None

#Get X error bars of the first chart series
errorBarsXFormat = bubbleChart.Series[0].ErrorBarsXFormat

#Specify error amount type as custom error bars
errorBarsXFormat.ErrorBarvType = ErrorValueType.CustomErrorBars

#Set the minus and plus value of the X error bars
errorBarsXFormat.MinusVal = 0.5
errorBarsXFormat.PlusVal = 0.5

#Get Y error bars of the first chart series
errorBarsYFormat = bubbleChart.Series[0].ErrorBarsYFormat

#Specify error amount type as custom error bars
errorBarsYFormat.ErrorBarvType = ErrorValueType.CustomErrorBars

#Set the minus and plus value of the Y error bars
errorBarsYFormat.MinusVal = 1
errorBarsYFormat.PlusVal = 1
```

---

# Spire.Presentation Python Chart
## Add secondary value axis to chart
```python
#Get the chart from the PowerPoint file.
chart = presentation.Slides[0].Shapes[0] if isinstance(presentation.Slides[0].Shapes[0], IChart) else None

#Add a secondary axis to display the value of Series 3.
chart.Series[2].UseSecondAxis = True

#Set the grid line of secondary axis as invisible.
chart.SecondaryValueAxis.MajorGridTextLines.FillType = FillFormatType.none
```

---

# Spire.Presentation Python Chart
## Add shadow effect for data label in chart
```python
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
```

---

# Spire.Presentation Python Chart Trendline
## Add trendline to chart series in PowerPoint presentation
```python
#Get the target chart, add trendline for the first data series of the chart and specify the trendline type.
chart = presentation.Slides[0].Shapes[0] if isinstance(presentation.Slides[0].Shapes[0], IChart) else None
it = chart.Series[0].AddTrendLine(TrendlinesType.Linear)

#Set the trendline properties to determine what should be displayed.
it.displayEquation = False
it.displayRSquaredValue = False
```

---

# Spire.Presentation Python Chart
## Create and configure a pie chart with auto vary color setting
```python
ppt = Presentation()
rect1 = RectangleF.FromLTRB (40, 100, 550+40, 320+100)

#Add a pie chart
chart = ppt.Slides[0].Shapes.AppendChartInit (ChartType.Pie, rect1, False)
chart.ChartTitle.TextProperties.Text = "Sales by Quarter"
chart.ChartTitle.TextProperties.IsCentered = True
chart.ChartTitle.Height = 30
chart.HasTitle = True

#Set whether auto vary color, default value is true
chart.Series[0].IsVaryColor = False
chart.Series[0].Distance = 15
```

---

# Change Legend Color in PowerPoint Chart
## Demonstrates how to change the color and style of a chart legend in a PowerPoint presentation
```python
# Get chart on the first slide
Chart = ppt.Slides[0].Shapes[0] if isinstance(ppt.Slides[0].Shapes[0], IChart) else None

# Change the fill color
Chart.ChartLegend.TextProperties.Paragraphs[0].DefaultCharacterProperties.Fill.FillType = FillFormatType.Solid
Chart.ChartLegend.TextProperties.Paragraphs[0].DefaultCharacterProperties.Fill.SolidColor.Color = Color.get_Blue()

# Use italic for the paragraph
Chart.ChartLegend.TextProperties.Paragraphs[0].DefaultCharacterProperties.IsItalic = TriState.TTrue
```

---

# spire.presentation python chart data table
## change font size for chart data table
```python
#Get chart on the first slide
Chart = ppt.Slides[0].Shapes[0] if isinstance(ppt.Slides[0].Shapes[0], IChart) else None
Chart.HasDataTable = True

#Add a new paragraph in data table
tp = TextParagraph()
Chart.ChartDataTable.Text.Paragraphs.Append(tp)

#Change the font size
Chart.ChartDataTable.Text.Paragraphs[0].DefaultCharacterProperties.FontHeight = 15
```

---

# spire.presentation python chart legend
## change font size for chart legend
```python
# Get chart on the first slide
Chart = ppt.Slides[0].Shapes[0] if isinstance(ppt.Slides[0].Shapes[0], IChart) else None

# Change legend font size
Chart.ChartLegend.TextProperties.Paragraphs[0].DefaultCharacterProperties.FontHeight = 17
```

---

# Spire.Presentation Chart Series Name Modification
## Change the name of a chart series in a presentation
```python
#Get chart on the first slide
Chart = ppt.Slides[0].Shapes[0] if isinstance(ppt.Slides[0].Shapes[0], IChart) else None

#Get the ranges of series label 
cr = Chart.Series.SeriesLabel

#Change the value
cr[0].Text = "Changed series name"
```

---

# Spire.Presentation Python Trendline Equation Modification
## Modify font size and position of a trendline equation in a chart
```python
#Get chart on the first slide
chart = presentation.Slides[0].Shapes[0] if isinstance(presentation.Slides[0].Shapes[0], IChart) else None

#Get the first trendline 
trendline = chart.Series[0].TrendLines[0]

#Change font size for trendline Equation text
for para in trendline.TrendLineLabel.TextFrameProperties.Paragraphs:
    para.DefaultCharacterProperties.FontHeight = 20
    for range in para.TextRanges:
        range.FontHeight = 20

#Change position for trendline Equation
trendline.TrendLineLabel.OffsetX = -0.1
trendline.TrendLineLabel.OffsetY = -0.05
```

---

# spire.presentation python chart
## change text font in chart elements
```python
#Get the chart
chart = ppt.Slides[0].Shapes[0] if isinstance(ppt.Slides[0].Shapes[0], IChart) else None

#Change the font of title
chart.ChartTitle.TextProperties.Paragraphs[0].DefaultCharacterProperties.LatinFont = TextFont("Lucida Sans Unicode")
chart.ChartTitle.TextProperties.Paragraphs[0].DefaultCharacterProperties.Fill.SolidColor.KnownColor = KnownColors.Blue
chart.ChartTitle.TextProperties.Paragraphs[0].DefaultCharacterProperties.FontHeight = 30

#Change the font of legend
chart.ChartLegend.TextProperties.Paragraphs[0].DefaultCharacterProperties.Fill.SolidColor.KnownColor = KnownColors.DarkGreen
chart.ChartLegend.TextProperties.Paragraphs[0].DefaultCharacterProperties.LatinFont = TextFont("Lucida Sans Unicode")

#Change the font of series
chart.PrimaryCategoryAxis.TextProperties.Paragraphs[0].DefaultCharacterProperties.Fill.SolidColor.KnownColor = KnownColors.Red
chart.PrimaryCategoryAxis.TextProperties.Paragraphs[0].DefaultCharacterProperties.Fill.FillType = FillFormatType.Solid
chart.PrimaryCategoryAxis.TextProperties.Paragraphs[0].DefaultCharacterProperties.FontHeight = 10
chart.PrimaryCategoryAxis.TextProperties.Paragraphs[0].DefaultCharacterProperties.LatinFont = TextFont("Lucida Sans Unicode")
```

---

# spire.presentation python chart axis
## configure chart axis properties
```python
#Get the chart
chart = ppt.Slides[0].Shapes[0] if isinstance(ppt.Slides[0].Shapes[0], IChart) else None

#Add a secondary axis to display the value of Series 3
chart.Series[2].UseSecondAxis = True

#Set the grid line of secondary axis as invisible
chart.SecondaryValueAxis.MajorGridTextLines.FillType = FillFormatType.none

#Set bounds of axis value. Before we assign values, we must set IsAutoMax and IsAutoMin as false, otherwise MS PowerPoint will automatically set the values.
chart.PrimaryValueAxis.IsAutoMax = False
chart.PrimaryValueAxis.IsAutoMin = False
chart.SecondaryValueAxis.IsAutoMax = False
chart.SecondaryValueAxis.IsAutoMax = False
chart.PrimaryValueAxis.MinValue = 0
chart.PrimaryValueAxis.MaxValue = 5.0
chart.SecondaryValueAxis.MinValue = 0
chart.SecondaryValueAxis.MaxValue = 1.0

#Set axis line format
chart.PrimaryValueAxis.MinorGridLines.FillType = FillFormatType.Solid
chart.SecondaryValueAxis.MinorGridLines.FillType = FillFormatType.Solid
chart.PrimaryValueAxis.MinorGridLines.Width = 0.1
chart.SecondaryValueAxis.MinorGridLines.Width = 0.1
chart.PrimaryValueAxis.MinorGridLines.SolidFillColor.Color = Color.get_LightGray()
chart.SecondaryValueAxis.MinorGridLines.SolidFillColor.Color = Color.get_LightGray()
chart.PrimaryValueAxis.MinorGridLines.DashStyle = LineDashStyleType.Dash
chart.SecondaryValueAxis.MinorGridLines.DashStyle = LineDashStyleType.Dash
chart.PrimaryValueAxis.MajorGridTextLines.Width = 0.3
chart.PrimaryValueAxis.MajorGridTextLines.SolidFillColor.Color = Color.get_LightSkyBlue()
chart.SecondaryValueAxis.MajorGridTextLines.Width = 0.3
chart.SecondaryValueAxis.MajorGridTextLines.SolidFillColor.Color = Color.get_LightSkyBlue()
```

---

# spire.presentation python chart copy
## copy chart between PowerPoint presentations
```python
#Create a PPT document
presentation1 = Presentation()

#Load the file from disk.
presentation1.LoadFromFile(inputFile_1)

#Get the chart that is going to be copied.
chart = presentation1.Slides[0].Shapes[0] if isinstance(presentation1.Slides[0].Shapes[0], IChart) else None

#Load the second PowerPoint document.
presentation2 = Presentation()
presentation2.LoadFromFile(inputFile_2)

#Copy chart from the first document to the second document.
presentation2.Slides.Append()
presentation2.Slides[1].Shapes.CreateChart(chart, RectangleF.FromLTRB (100, 100, 600, 400), -1)
```

---

# spire.presentation copy chart within presentation
## copy a chart from one slide to another within the same PowerPoint presentation
```python
#Get the chart that is going to be copied.
chart = ppt.Slides[0].Shapes[0] if isinstance(ppt.Slides[0].Shapes[0], IChart) else None

#Copy the chart from the first slide to the specified location of the second slide within the same document.
slide1 = ppt.Slides.Append()
slide1.Shapes.CreateChart(chart, RectangleF.FromLTRB (100, 100, 600, 400), 0)
```

---

# Spire.Presentation 100% Stacked Bar Chart
## Create and configure a 100% stacked bar chart in PowerPoint
```python
#Create a PowerPoint document.
presentation = Presentation()

#Set slide size and get the first slide
presentation.SlideSize.Type = SlideSizeType.Screen16x9
slidesize = presentation.SlideSize.Size
slide = presentation.Slides[0]

#Append a chart
rect = RectangleF.FromLTRB (20, 20, slidesize.Width - 20, slidesize.Height - 20)
chart = slide.Shapes.AppendChart(ChartType.Bar100PercentStacked, rect)

#Set up chart data
columnlabels = ["Series 1", "Series 2", "Series 3"]
rowlabels = ["Category 1", "Category 2", "Category 3"]
values = [[ 20.83233, 10.34323, -10.354667 ], [ 10.23456, -12.23456, 23.34456 ], [ 12.34345, -23.34343, -13.23232 ]]

#Insert the column labels
c = 0
while c < len(columnlabels):
    chart.ChartData[0,c + 1].Text = columnlabels[c]
    c += 1

#Insert the row labels
r = 0
while r < len(rowlabels):
    chart.ChartData[r + 1,0].Text = rowlabels[r]
    r += 1

#Insert the values
value = 0.0
r = 0
while r < len(rowlabels):
    c = 0
    while c < len(columnlabels):
        value = round(values[r][c], 2)
        chart.ChartData[r + 1,c + 1].NumberValue = value
        c += 1
    r += 1

chart.Series.SeriesLabel = chart.ChartData[0,1,0,len(columnlabels)]
chart.Categories.CategoryLabels = chart.ChartData[1,0,len(rowlabels),0]

#Set the position of category axis
chart.PrimaryCategoryAxis.Position = AxisPositionType.Left
chart.SecondaryCategoryAxis.Position = AxisPositionType.Left
chart.PrimaryCategoryAxis.TickLabelPosition = TickLabelPositionType.TickLabelPositionLow

#Set the data, font and format for the series of each column
c = 0
while c < len(columnlabels):
    chart.Series[c].Values = chart.ChartData[1,c + 1,len(rowlabels),c + 1]
    chart.Series[c].Fill.FillType = FillFormatType.Solid
    chart.Series[c].InvertIfNegative = False
    r = 0
    while r < len(rowlabels):
        label = chart.Series[c].DataLabels.Add()
        label.LabelValueVisible = True
        chart.Series[c].DataLabels[r].HasDataSource = False
        chart.Series[c].DataLabels[r].NumberFormat = "0#\\%"
        chart.Series[c].DataLabels.TextProperties.Paragraphs[0].DefaultCharacterProperties.FontHeight = 12
        r += 1
    c += 1

#Set the color of the Series
chart.Series[0].Fill.SolidColor.Color = Color.get_YellowGreen()
chart.Series[1].Fill.SolidColor.Color = Color.get_Red()
chart.Series[2].Fill.SolidColor.Color = Color.get_Green()

#Set the font and size for chartlegend
font = TextFont("Tw Cen MT")
k = 0
while k < len(chart.ChartLegend.EntryTextProperties):
    chart.ChartLegend.EntryTextProperties[k].LatinFont = font
    chart.ChartLegend.EntryTextProperties[k].FontHeight = 20
    k += 1
```

---

# Spire.Presentation Box and Whisker Chart
## Create and configure a Box and Whisker chart in a PowerPoint presentation
```python
# Create a PPT document
ppt = Presentation()

# Insert a BoxAndWhisker chart to the first slide 
chart = ppt.Slides[0].Shapes.AppendChartInit(ChartType.BoxAndWhisker, RectangleF.FromLTRB(50, 50, 550, 450), False)

# Series labels
seriesLabel = ["Series 1", "Series 2", "Series 3"]
i = 0
while i < len(seriesLabel):
    chart.ChartData[0,i + 1].Text = "Series 1"
    i += 1

# Categories
categories = ["Category 1", "Category 1", "Category 1", "Category 1", "Category 1", "Category 1", "Category 1", "Category 2", "Category 2", "Category 2", "Category 2", "Category 2", "Category 2", "Category 3", "Category 3", "Category 3", "Category 3", "Category 3"]
i = 0
while i < len(categories):
    chart.ChartData[i + 1,0].Text = categories[i]
    i += 1

# Values
values = [[-7, -3, -24], [-10, 1, 11], [-28, -6, 34], [47, 2, -21], [35, 17, 22], [-22, 15, 19], [17, -11, 25], [-30, 18, 25], [49, 22, 56], [37, 22, 15], [-55, 25, 31], [14, 18, 22], [18, -22, 36], [-45, 25, -17], [-33, 18, 22], [18, 2, -23], [-33, -22, 10], [10, 19, 22]]
i = 0
while i < len(seriesLabel):
    j = 0
    while j < len(categories):
        chart.ChartData[j + 1,i + 1].NumberValue = values[j][i]
        j += 1
    i += 1

# Set series data
chart.Series.SeriesLabel = chart.ChartData[0,1,0,len(seriesLabel)]
chart.Categories.CategoryLabels = chart.ChartData[1,0,len(categories),0]
chart.Series[0].Values = chart.ChartData[1,1,len(categories),1]
chart.Series[1].Values = chart.ChartData[1,2,len(categories),2]
chart.Series[2].Values = chart.ChartData[1,3,len(categories),3]

# Configure series properties
chart.Series[0].ShowInnerPoints = False
chart.Series[0].ShowOutlierPoints = True
chart.Series[0].ShowMeanMarkers = True
chart.Series[0].ShowMeanLine = True
chart.Series[0].QuartileCalculationType = QuartileCalculation.ExclusiveMedian

chart.Series[1].ShowInnerPoints = False
chart.Series[1].ShowOutlierPoints = True
chart.Series[1].ShowMeanMarkers = True
chart.Series[1].ShowMeanLine = True
chart.Series[1].QuartileCalculationType = QuartileCalculation.InclusiveMedian

chart.Series[2].ShowInnerPoints = False
chart.Series[2].ShowOutlierPoints = True
chart.Series[2].ShowMeanMarkers = True
chart.Series[2].ShowMeanLine = True
chart.Series[2].QuartileCalculationType = QuartileCalculation.ExclusiveMedian

# Set chart title and legend
chart.HasLegend = True
chart.ChartTitle.TextProperties.Text = "BoxAndWhisker"
chart.ChartLegend.Position = ChartLegendPositionType.Top
```

---

# Spire.Presentation Bubble Chart Creation
## Create and configure a bubble chart in a PowerPoint presentation
```python
#Add bubble chart
rect1 = RectangleF.FromLTRB (90, 100, 640, 420)
chart = presentation.Slides[0].Shapes.AppendChartInit(ChartType.Bubble, rect1, False)

#Chart title
chart.ChartTitle.TextProperties.Text = "Bubble Chart"
chart.ChartTitle.TextProperties.IsCentered = True
chart.ChartTitle.Height = 30
chart.HasTitle = True

#Attach the data to chart
xdata = [7.7, 8.9, 1.0, 2.4]
ydata = [15.2, 5.3, 6.7, 8]
size = [1.1, 2.4, 3.7, 4.8]
chart.ChartData[0,0].Text = "X-Value"
chart.ChartData[0,1].Text = "Y-Value"
chart.ChartData[0,2].Text = "Size"
i = 0
while i < len(xdata):
    chart.ChartData[i + 1,0].NumberValue = xdata[i]
    chart.ChartData[i + 1,1].NumberValue = ydata[i]
    chart.ChartData[i + 1,2].NumberValue = size[i]
    i += 1

#Set series label
chart.Series.SeriesLabel = chart.ChartData["B1","B1"]
chart.Series[0].XValues = chart.ChartData["A2","A5"]
chart.Series[0].YValues = chart.ChartData["B2","B5"]
chart.Series[0].Bubbles.Add(chart.ChartData["C2"])
chart.Series[0].Bubbles.Add(chart.ChartData["C3"])
chart.Series[0].Bubbles.Add(chart.ChartData["C4"])
chart.Series[0].Bubbles.Add(chart.ChartData["C5"])
```

---

# Spire.Presentation Python Chart
## Create clustered column chart in PowerPoint presentation
```python
# Create a PPT file
presentation = Presentation()

# Add clustered column chart
rect1 = RectangleF.FromLTRB (90, 100, 640, 420)
chart = presentation.Slides[0].Shapes.AppendChartInit (ChartType.ColumnClustered, rect1, False)

# Chart title
chart.ChartTitle.TextProperties.Text = "Clustered Column Chart"
chart.ChartTitle.TextProperties.IsCentered = True
chart.ChartTitle.Height = 30
chart.HasTitle = True

# Data for series
Series1 = [7.7, 8.9, 1.0, 2.4]
Series2 = [15.2, 5.3, 6.7, 8]

# Set series text
chart.ChartData[0,1].Text = "Series1"
chart.ChartData[0,2].Text = "Series2"

# Set category text
chart.ChartData[1,0].Text = "Category 1"
chart.ChartData[2,0].Text = "Category 2"
chart.ChartData[3,0].Text = "Category 3"
chart.ChartData[4,0].Text = "Category 4"

# Fill data for chart
i = 0
while i < len(Series1):
    chart.ChartData[i + 1,1].NumberValue = Series1[i]
    chart.ChartData[i + 1,2].NumberValue = Series2[i]
    i += 1

# Set series label
chart.Series.SeriesLabel = chart.ChartData["B1","C1"]

# Set category label
chart.Categories.CategoryLabels = chart.ChartData["A2","A5"]

# Set values for series
chart.Series[0].Values = chart.ChartData["B2","B5"]
chart.Series[1].Values = chart.ChartData["C2","C5"]
```

---

# Spire.Presentation Python Chart
## Create a combination chart with column and line series
```python
#Create a presentation instance
presentation = Presentation()

#Insert a column clustered chart
rect = RectangleF.FromLTRB (100, 100, 650, 420)
chart = presentation.Slides[0].Shapes.AppendChart(ChartType.ColumnClustered, rect)

#Set chart title
chart.ChartTitle.TextProperties.Text = "Monthly Sales Report"
chart.ChartTitle.TextProperties.IsCentered = True
chart.ChartTitle.Height = 30
chart.HasTitle = True

#Set series labels
chart.Series.SeriesLabel = chart.ChartData["B1","C1"]

#Set categories labels    
chart.Categories.CategoryLabels = chart.ChartData["A2","A7"]

#Assign data to series values
chart.Series[0].Values = chart.ChartData["B2","B7"]
chart.Series[1].Values = chart.ChartData["C2","C7"]

#Change the chart type of serie 2 to line with markers
chart.Series[1].Type = ChartType.LineMarkers

#Plot data of series 2 on the secondary axis
chart.Series[1].UseSecondAxis = True

#Set the number format as percentage 
chart.SecondaryValueAxis.NumberFormat = "0%"

#Hide gridlinkes of secondary axis
chart.SecondaryValueAxis.MajorGridTextLines.FillType = FillFormatType.none

#Set overlap
chart.OverLap = -50

#Set gapwidth
chart.GapWidth = 200
```

---

# spire.presentation python chart
## Create 3D Cylinder Clustered Chart
```python
#Insert chart
left = math.trunc(presentation.SlideSize.Size.Width / float(2)) - 200
rect = RectangleF.FromLTRB (left, 85, 400+left, 485)
chart = presentation.Slides[0].Shapes.AppendChart(ChartType.Cylinder3DClustered, rect)

#Add chart Title
chart.ChartTitle.TextProperties.Text = "Report"
chart.ChartTitle.TextProperties.IsCentered = True
chart.ChartTitle.Height = 30
chart.HasTitle = True

#Configure chart data and appearance
chart.Series.SeriesLabel = chart.ChartData["B1","D1"]
chart.Categories.CategoryLabels = chart.ChartData["A2","A7"]
chart.Series[0].Values = chart.ChartData["B2","B7"]
chart.Series[0].Fill.FillType = FillFormatType.Solid
chart.Series[0].Fill.SolidColor.KnownColor = KnownColors.Brown
chart.Series[1].Values = chart.ChartData["C2","C7"]
chart.Series[1].Fill.FillType = FillFormatType.Solid
chart.Series[1].Fill.SolidColor.KnownColor = KnownColors.Green
chart.Series[2].Values = chart.ChartData["D2","D7"]
chart.Series[2].Fill.FillType = FillFormatType.Solid
chart.Series[2].Fill.SolidColor.KnownColor = KnownColors.Orange

#Set the 3D rotation
chart.RotationThreeD.XDegree = 10
chart.RotationThreeD.YDegree = 10
```

---

# Spire.Presentation Python Chart
## Create a doughnut chart in PowerPoint presentation
```python
#Create a ppt document
presentation = Presentation()
rect = RectangleF.FromLTRB (80, 100, 630, 420)

#Add a Doughnut chart
chart = presentation.Slides[0].Shapes.AppendChartInit(ChartType.Doughnut, rect, False)
chart.ChartTitle.TextProperties.Text = "Market share by country"
chart.ChartTitle.TextProperties.IsCentered = True
chart.ChartTitle.Height = 30
countries = ["Guba", "Mexico", "France", "German"]
sales = [1800, 3000, 5100, 6200]
chart.ChartData[0,0].Text = "Countries"
chart.ChartData[0,1].Text = "Sales"
i = 0
while i < len(countries):
    chart.ChartData[i + 1,0].Text = countries[i]
    chart.ChartData[i + 1,1].NumberValue = sales[i]
    i += 1
chart.Series.SeriesLabel = chart.ChartData["B1","B1"]
chart.Categories.CategoryLabels = chart.ChartData["A2","A5"]
chart.Series[0].Values = chart.ChartData["B2","B5"]
for i, item in enumerate(chart.Series[0].Values):
    cdp = ChartDataPoint(chart.Series[0])
    cdp.Index = i
    chart.Series[0].DataPoints.Add(cdp)

#Set the series color
chart.Series[0].DataPoints[0].Fill.FillType = FillFormatType.Solid
chart.Series[0].DataPoints[0].Fill.SolidColor.Color = Color.get_LightBlue()
chart.Series[0].DataPoints[1].Fill.FillType = FillFormatType.Solid
chart.Series[0].DataPoints[1].Fill.SolidColor.Color = Color.get_MediumPurple()
chart.Series[0].DataPoints[2].Fill.FillType = FillFormatType.Solid
chart.Series[0].DataPoints[2].Fill.SolidColor.Color = Color.get_DarkGray()
chart.Series[0].DataPoints[3].Fill.FillType = FillFormatType.Solid
chart.Series[0].DataPoints[3].Fill.SolidColor.Color = Color.get_DarkOrange()
chart.Series[0].DataLabels.LabelValueVisible = True
chart.Series[0].DataLabels.PercentValueVisible = True
chart.Series[0].DoughnutHoleSize = 60
```

---

# spire.presentation python funnel chart
## create a funnel chart in PowerPoint presentation
```python
#Create PPT document
ppt = Presentation()

#Create a Funnel chart to the first slide
chart = ppt.Slides[0].Shapes.AppendChartInit (ChartType.Funnel, RectangleF.FromLTRB (50, 50, 600, 450), False)

#Set series text
chart.ChartData[0,1].Text = "Series 1"

#Set category text
chart.ChartData[1,0].Text = "Category 1"
chart.ChartData[2,0].Text = "Category 2"
chart.ChartData[3,0].Text = "Category 3"

#Fill data for chart
chart.ChartData[1,1].NumberValue = 100
chart.ChartData[2,1].NumberValue = 75
chart.ChartData[3,1].NumberValue = 50

#Set series labels
chart.Series.SeriesLabel = chart.ChartData[0,1,0,1]

#Set categories labels 
chart.Categories.CategoryLabels = chart.ChartData[1,0,3,0]

#Assign data to series values
chart.Series[0].Values = chart.ChartData[1,1,3,1]

#Set the chart title
chart.ChartTitle.TextProperties.Text = "Funnel"
```

---

# spire.presentation python histogram chart
## create a histogram chart in powerpoint presentation
```python
#Create PPT document
ppt = Presentation()

#Add a Histogram chart
chart = ppt.Slides[0].Shapes.AppendChartInit (ChartType.Histogram, RectangleF.FromLTRB (50, 50, 550, 450), False)

#Set series text
chart.ChartData[0,0].Text = "Series 1"

#Set series label
chart.Series.SeriesLabel = chart.ChartData[0,0,0,0]

#Set values for series
chart.Series[0].Values = chart.ChartData[1,0,len(values),0]
chart.PrimaryCategoryAxis.NumberOfBins = 7
chart.PrimaryCategoryAxis.GapWidth = 20

#Chart title
chart.ChartTitle.TextProperties.Text = "Histogram"
chart.ChartLegend.Position = ChartLegendPositionType.Bottom
```

---

# Spire.Presentation Python Chart
## Create Line Markers Chart in PowerPoint
```python
#Create a PPT file
presentation = Presentation()

#Add line markers chart
rect1 = RectangleF.FromLTRB (90, 100, 640, 420)
chart = presentation.Slides[0].Shapes.AppendChartInit (ChartType.LineMarkers, rect1, False)

#Chart title
chart.ChartTitle.TextProperties.Text = "Line Makers Chart"
chart.ChartTitle.TextProperties.IsCentered = True
chart.ChartTitle.Height = 30
chart.HasTitle = True

#Set series text
chart.ChartData[0,1].Text = "Series1"
chart.ChartData[0,2].Text = "Series2"

#Set category text
chart.ChartData[1,0].Text = "Category 1"
chart.ChartData[2,0].Text = "Category 2"
chart.ChartData[3,0].Text = "Category 3"
chart.ChartData[4,0].Text = "Category 4"

#Fill data for chart
Series1 = [7.7, 8.9, 1.0, 2.4]
Series2 = [15.2, 5.3, 6.7, 8]
i = 0
while i < len(Series1):
    chart.ChartData[i + 1,1].NumberValue = Series1[i]
    chart.ChartData[i + 1,2].NumberValue = Series2[i]
    i += 1

#Set series label
chart.Series.SeriesLabel = chart.ChartData["B1","C1"]

#Set category label
chart.Categories.CategoryLabels = chart.ChartData["A2","A5"]

#Set values for series
chart.Series[0].Values = chart.ChartData["B2","B5"]
chart.Series[1].Values = chart.ChartData["C2","C5"]
```

---

# spire.presentation python map chart
## create a map chart in PowerPoint presentation
```python
#Create a PPT document
ppt = Presentation()

#Insert a Map chart to the first slide 
chart = ppt.Slides[0].Shapes.AppendChartInit (ChartType.Map, RectangleF.FromLTRB (50, 50, 500, 500), False)
chart.ChartData[0,1].Text = "series"

#Define some data.
countries = ["China", "Russia", "France", "Mexico", "United States", "India", "Australia"]
i = 0
while i < len(countries):
    chart.ChartData[i + 1,0].Text = countries[i]
    i += 1
values = [32, 20, 23, 17, 18, 6, 11]
i = 0
while i < len(values):
    chart.ChartData[i + 1,1].NumberValue = values[i]
    i += 1
chart.Series.SeriesLabel = chart.ChartData[0,1,0,1]
chart.Categories.CategoryLabels = chart.ChartData[1,0,7,0]
chart.Series[0].Values = chart.ChartData[1,1,7,1]
```

---

# spire.presentation python chart
## create Pareto chart in PowerPoint presentation
```python
#Create PPT document
ppt = Presentation()

#Create a Pareto chart in first slide
chart = ppt.Slides[0].Shapes.AppendChartInit (ChartType.Pareto, RectangleF.FromLTRB (50, 50, 550, 450), False)

#Set series text
chart.ChartData[0,1].Text = "Series 1"

#Set category text
categories = ["Category 1", "Category 2", "Category 4", "Category 3", "Category 4", "Category 2", "Category 1", "Category 1", "Category 3", "Category 2", "Category 4", "Category 2", "Category 3", "Category 1", "Category 3", "Category 2", "Category 4", "Category 1", "Category 1", "Category 3", "Category 2", "Category 4", "Category 1", "Category 1", "Category 3", "Category 2", "Category 4", "Category 1"]
i = 0
while i < len(categories):
    chart.ChartData[i + 1,0].Text = categories[i]
    i += 1

#Fill data for chart
values = [1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1]
i = 0
while i < len(values):
    chart.ChartData[i + 1,1].NumberValue = values[i]
    i += 1

#Configure chart
chart.Series.SeriesLabel = chart.ChartData[0,1,0,1]
chart.Categories.CategoryLabels = chart.ChartData[1,0,len(categories),0]
chart.Series[0].Values = chart.ChartData[1,1,len(values),1]
chart.PrimaryCategoryAxis.IsBinningByCategory = True
chart.Series[1].Line.FillFormat.FillType = FillFormatType.Solid
chart.Series[1].Line.FillFormat.SolidFillColor.Color = Color.get_Red()
chart.ChartTitle.TextProperties.Text = "Pareto"
chart.HasLegend = True
chart.ChartLegend.Position = ChartLegendPositionType.Bottom
```

---

# Spire.Presentation Python Pie Chart
## Create a pie chart with customized colors and data labels
```python
#Create a PPT document
presentation = Presentation()

#Insert a Pie chart to the first slide and set the chart title.
rect1 = RectangleF.FromLTRB (40, 100, 590, 420)
chart = presentation.Slides[0].Shapes.AppendChartInit (ChartType.Pie, rect1, False)
chart.ChartTitle.TextProperties.Text = "Sales by Quarter"
chart.ChartTitle.TextProperties.IsCentered = True
chart.ChartTitle.Height = 30
chart.HasTitle = True

#Set category labels, series label and series data.
chart.Series.SeriesLabel = chart.ChartData["B1","B1"]
chart.Categories.CategoryLabels = chart.ChartData["A2","A5"]
chart.Series[0].Values = chart.ChartData["B2","B5"]

#Add data points to series and fill each data point with different color.
for i, unusedItem in enumerate(chart.Series[0].Values):
    cdp = ChartDataPoint(chart.Series[0])
    cdp.Index = i
    chart.Series[0].DataPoints.Add(cdp)
chart.Series[0].DataPoints[0].Fill.FillType = FillFormatType.Solid
chart.Series[0].DataPoints[0].Fill.SolidColor.Color = Color.get_RosyBrown()
chart.Series[0].DataPoints[1].Fill.FillType = FillFormatType.Solid
chart.Series[0].DataPoints[1].Fill.SolidColor.Color = Color.get_LightBlue()
chart.Series[0].DataPoints[2].Fill.FillType = FillFormatType.Solid
chart.Series[0].DataPoints[2].Fill.SolidColor.Color = Color.get_LightPink()
chart.Series[0].DataPoints[3].Fill.FillType = FillFormatType.Solid
chart.Series[0].DataPoints[3].Fill.SolidColor.Color = Color.get_MediumPurple()

#Set the data labels to display label value and percentage value.
chart.Series[0].DataLabels.LabelValueVisible = True
chart.Series[0].DataLabels.PercentValueVisible = True
```

---

# spire.presentation python scatter chart
## create scatter chart in presentation
```python
#Create a presentation
pres = Presentation()

#Insert a chart and set chart title and chart type
rect1 = RectangleF.FromLTRB (90, 100, 640, 420)
chart = pres.Slides[0].Shapes.AppendChartInit (ChartType.ScatterMarkers, rect1, False)
chart.ChartTitle.TextProperties.Text = "ScatterMarker Chart"
chart.ChartTitle.TextProperties.IsCentered = True
chart.ChartTitle.Height = 30
chart.HasTitle = True

#Set chart data
xdata = [2.7, 8.9, 10.0, 12.4]
ydata = [3.2, 15.3, 6.7, 8]
chart.ChartData[0,0].Text = "X-Value"
chart.ChartData[0,1].Text = "Y-Value"
i = 0
while i < len(xdata):
    chart.ChartData[i + 1,0].NumberValue = xdata[i]
    chart.ChartData[i + 1,1].NumberValue = ydata[i]
    i += 1

#Set the series label
chart.Series.SeriesLabel = chart.ChartData["B1","B1"]

#Assign data to X axis, Y axis and Bubbles
chart.Series[0].XValues = chart.ChartData["A2","A5"]
chart.Series[0].YValues = chart.ChartData["B2","B5"]
```

---

# spire.presentation python chart
## create SunBurst chart
```python
#Create PPT document
ppt = Presentation()

#Create a SunBurst chart to the first slide
chart = ppt.Slides[0].Shapes.AppendChartInit (ChartType.SunBurst, RectangleF.FromLTRB (50, 50, 550, 450), False)

#Set series text
chart.ChartData[0,3].Text = "Series 1"

#Set category text
categories = [["Branch 1", "Stem 1", "Leaf 1"], ["Branch 1", "Stem 1", "Leaf 2"], ["Branch 1", "Stem 1", "Leaf 3"], ["Branch 1", "Stem 2", "Leaf 4"], ["Branch 1", "Stem 2", "Leaf 5"], ["Branch 1", "Leaf 6", None], ["Branch 1", "Leaf 7", None], ["Branch 2", "Stem 3", "Leaf 8"], ["Branch 2", "Leaf 9", None], ["Branch 2", "Stem 4", "Leaf 10"], ["Branch 2", "Stem 4", "Leaf 11"], ["Branch 2", "Stem 5", "Leaf 12"], ["Branch 3", "Stem 5", "Leaf 13"], ["Branch 3", "Stem 6", "Leaf 14"], ["Branch 3", "Leaf 15", None]]
for i in range(0, 15):
    for j in range(0, 3):
        chart.ChartData[i + 1,j].Text = categories[i][j]

#Fill data for chart
values = [17, 23, 48, 22, 76, 54, 77, 26, 44, 63, 10, 15, 48, 15, 51]
i = 0
while i < len(values):
    chart.ChartData[i + 1,3].NumberValue = values[i]
    i += 1

#Set series labels
chart.Series.SeriesLabel = chart.ChartData[0,3,0,3]

#Set categories labels 
chart.Categories.CategoryLabels = chart.ChartData[1,0,len(values),2]

#Assign data to series values
chart.Series[0].Values = chart.ChartData[1,3,len(values),3]
chart.Series[0].DataLabels.CategoryNameVisible = True
chart.ChartTitle.TextProperties.Text = "SunBurst"
chart.HasLegend = True
chart.ChartLegend.Position = ChartLegendPositionType.Top
```

---

# spire.presentation python chart
## create TreeMap chart in PowerPoint
```python
#Create PPT document
ppt = Presentation()

#Create a TreeMap chart to the first slide
chart = ppt.Slides[0].Shapes.AppendChartInit (ChartType.TreeMap, RectangleF.FromLTRB (50, 50, 550, 450), False)

#Set series text
chart.ChartData[0,3].Text = "Series 1"

#Set category text
categories = [["Branch 1", "Stem 1", "Leaf 1"], ["Branch 1", "Stem 1", "Leaf 2"], ["Branch 1", "Stem 1", "Leaf 3"], ["Branch 1", "Stem 2", "Leaf 4"], ["Branch 1", "Stem 2", "Leaf 5"], ["Branch 1", "Stem 2", "Leaf 6"], ["Branch 1", "Stem 2", "Leaf 7"], ["Branch 2", "Stem 3", "Leaf 8"], ["Branch 2", "Stem 3", "Leaf 9"], ["Branch 2", "Stem 4", "Leaf 10"], ["Branch 2", "Stem 4", "Leaf 11"], ["Branch 2", "Stem 5", "Leaf 12"], ["Branch 3", "Stem 5", "Leaf 13"], ["Branch 3", "Stem 6", "Leaf 14"], ["Branch 3", "Stem 6", "Leaf 15"]]
for i in range(0, 15):
    for j in range(0, 3):
        chart.ChartData[i + 1,j].Text = categories[i][j]

#Fill data for chart
values = [17, 23, 48, 22, 76, 54, 77, 26, 44, 63, 10, 15, 48, 15, 51]
i = 0
while i < len(values):
    chart.ChartData[i + 1,3].NumberValue = values[i]
    i += 1

#Set series labels
chart.Series.SeriesLabel = chart.ChartData[0,3,0,3]

#Set categories labels 
chart.Categories.CategoryLabels = chart.ChartData[1,0,len(values),2]

#Assign data to series values
chart.Series[0].Values = chart.ChartData[1,3,len(values),3]
chart.Series[0].DataLabels.CategoryNameVisible = True
chart.Series[0].TreeMapLabelOption = TreeMapLabelOption.Banner
chart.ChartTitle.TextProperties.Text = "TreeMap"
chart.HasLegend = True
chart.ChartLegend.Position = ChartLegendPositionType.Top
```

---

# spire.presentation python chart
## create WaterFall chart
```python
#Create PPT document
ppt = Presentation()

#Create a WaterFall chart to the first slide
chart = ppt.Slides[0].Shapes.AppendChartInit (ChartType.WaterFall, RectangleF.FromLTRB (50, 50, 550, 450), False)

#Set series labels
chart.Series.SeriesLabel = chart.ChartData[0,1,0,1]

#Set categories labels 
chart.Categories.CategoryLabels = chart.ChartData[1,0,7,0]

#Assign data to series values
chart.Series[0].Values = chart.ChartData[1,1,7,1]

#Set specific datapoints as totals
chartDataPoint = ChartDataPoint(chart.Series[0])
chartDataPoint.Index = 2
chartDataPoint.SetAsTotal = True
chart.Series[0].DataPoints.Add(chartDataPoint)

chartDataPoint2 = ChartDataPoint(chart.Series[0])
chartDataPoint2.Index = 5
chartDataPoint2.SetAsTotal = True
chart.Series[0].DataPoints.Add(chartDataPoint2)

#Configure chart appearance
chart.Series[0].ShowConnectorLines = True
chart.Series[0].DataLabels.LabelValueVisible = True
chart.ChartLegend.Position = ChartLegendPositionType.Right
chart.ChartTitle.TextProperties.Text = "WaterFall"
```

---

# spire.presentation python chart
## delete chart legend entries
```python
#Get the chart.
chart = presentation.Slides[0].Shapes[0] if isinstance(presentation.Slides[0].Shapes[0], IChart) else None

#Delete the first and the second legend entries from the chart.
chart.ChartLegend.DeleteEntry(0)
chart.ChartLegend.DeleteEntry(1)
```

---

# Spire.Presentation Doughnut Chart Hole Size
## Set the hole size of a doughnut chart in a PowerPoint presentation
```python
#Get the chart on the first slide
Chart = ppt.Slides[0].Shapes[0] if isinstance(ppt.Slides[0].Shapes[0], IChart) else None

#Set hole size
Chart.Series[0].DoughnutHoleSize = 55
```

---

# Edit Chart Data in PowerPoint
## Modify data point values in a chart
```python
#Get chart on the first slide
Chart = ppt.Slides[0].Shapes[0] if isinstance(ppt.Slides[0].Shapes[0], IChart) else None

#Change the value of the second datapoint of the first series
Chart.Series[0].Values[1].NumberValue = 6
```

---

# Spire.Presentation Python Chart
## Explode pie chart by setting series distance
```python
#Get the chart that needs to set the point explosion.
chart = presentation.Slides[0].Shapes[0] if isinstance(presentation.Slides[0].Shapes[0], IChart) else None

chart.Series[0].Distance = 15
```

---

# Spire.Presentation Chart Marker
## Fill picture in chart marker
```python
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
```

---

# Format Chart Data Labels
## Format chart data labels with custom text, font, position and color properties
```python
#Get the chart
chart = ppt.Slides[0].Shapes[0] if isinstance(ppt.Slides[0].Shapes[0], IChart) else None

#Get the chart series
sers = chart.Series

#Initialize four instances of series label and set parameters of each label
cd1 = sers[0].DataLabels.Add()
cd1.PercentageVisible = True
cd1.TextFrame.Text = "Custom Datalabel1"
cd1.TextFrame.TextRange.FontHeight = 12
cd1.TextFrame.TextRange.LatinFont = TextFont("Lucida Sans Unicode")
cd1.TextFrame.TextRange.Fill.FillType =FillFormatType.Solid
cd1.TextFrame.TextRange.Fill.SolidColor.Color = Color.get_Green()

cd2 = sers[0].DataLabels.Add()
cd2.Position = ChartDataLabelPosition.InsideEnd
cd2.PercentageVisible = True
cd2.TextFrame.Text = "Custom Datalabel2"
cd2.TextFrame.TextRange.FontHeight = 10
cd2.TextFrame.TextRange.LatinFont = TextFont("Arial")
cd2.TextFrame.TextRange.Fill.FillType = FillFormatType.Solid
cd2.TextFrame.TextRange.Fill.SolidColor.Color = Color.get_OrangeRed()

cd3 = sers[0].DataLabels.Add()
cd3.Position = ChartDataLabelPosition.Center
cd3.PercentageVisible = True
cd3.TextFrame.Text = "Custom Datalabel3"
cd3.TextFrame.TextRange.FontHeight = 14
cd3.TextFrame.TextRange.LatinFont = TextFont("Calibri")
cd3.TextFrame.TextRange.Fill.FillType = FillFormatType.Solid
cd3.TextFrame.TextRange.Fill.SolidColor.Color = Color.get_Blue()

cd4 = sers[0].DataLabels.Add()
cd4.Position = ChartDataLabelPosition.InsideBase
cd4.PercentageVisible = True
cd4.TextFrame.Text = "Custom Datalabel4"
cd4.TextFrame.TextRange.FontHeight = 12
cd4.TextFrame.TextRange.LatinFont = TextFont("Lucida Sans Unicode")
cd4.TextFrame.TextRange.Fill.FillType = FillFormatType.Solid
cd4.TextFrame.TextRange.Fill.SolidColor.Color = Color.get_OliveDrab()
```

---

# Spire.Presentation Chart Axis Values Extraction
## Extract values and units from chart axes in PowerPoint presentations
```python
#Get chart on the first slide
Chart = ppt.Slides[0].Shapes[0] if isinstance(ppt.Slides[0].Shapes[0], IChart) else None

#Get unit from primary category axis
majorUnit = Chart.PrimaryCategoryAxis.MajorUnit
majorUnitScale = Chart.PrimaryCategoryAxis.MajorUnitScale

#Get values from primary value axis
minValue = Chart.PrimaryValueAxis.MinValue
maxValue = Chart.PrimaryValueAxis.MaxValue
```

---

# Spire.Presentation Chart Axis Labels Grouping
## Group two-level axis labels in a chart
```python
#Get the chart from the first slide of the presentation.
chart = presentation.Slides[0].Shapes[0] if isinstance(presentation.Slides[0].Shapes[0], IChart) else None

#Get the category axis from the chart.
chartAxis = chart.PrimaryCategoryAxis

#Group the axis labels that have the same first-level label.
if chartAxis.HasMultiLvlLbl:
    chartAxis.IsMergeSameLabel = True
```

---

# Spire.Presentation Python Chart
## Hide axis and gridlines in a chart
```python
#Get chart on the first slide
Chart = ppt.Slides[0].Shapes[0] if isinstance(ppt.Slides[0].Shapes[0], IChart) else None

#Hide axis
Chart.PrimaryCategoryAxis.IsVisible = False
Chart.PrimaryValueAxis.IsVisible = False

#Remove gridline
Chart.PrimaryValueAxis.MajorGridTextLines.FillType = FillFormatType.none
```

---

# Spire.Presentation Chart Series Visibility
## Hide or show a series in a chart
```python
#Get the first slide.
slide = presentation.Slides[0]

#Get the first chart.
chart = slide.Shapes[0] if isinstance(slide.Shapes[0], IChart) else None

#Hide the first series of the chart.
chart.Series[0].IsHidden = True

#Show the first series of the chart.
#chart.Series[0].IsHidden = false
```

---

# spire.presentation chart invert if negative
## Set InvertIfNegative property for chart series
```python
#Get chart on the first slide
Chart = ppt.Slides[0].Shapes[0] if isinstance(ppt.Slides[0].Shapes[0], IChart) else None

#Set invert if negative
Chart.Series[0].InvertIfNegative = True
```

---

# Spire.Presentation Python Chart
## Modify chart category axis
```python
#Get chart on the first slide
Chart = ppt.Slides[0].Shapes[0] if isinstance(ppt.Slides[0].Shapes[0], IChart) else None

#Modify the major unit
Chart.PrimaryCategoryAxis.IsAutoMajor = False
Chart.PrimaryCategoryAxis.MajorUnit = 1
Chart.PrimaryCategoryAxis.MajorUnitScale = ChartBaseUnitType.Months
```

---

# Spire.Presentation Multiple Category Chart
## Create a PowerPoint presentation with a column clustered chart that has multiple category levels
```python
# Create a PPT file
presentation = Presentation()

# Add line markers chart
rect1 = RectangleF.FromLTRB (90, 100, 640, 420)
chart = presentation.Slides[0].Shapes.AppendChartInit (ChartType.ColumnClustered, rect1, False)

# Chart title
chart.ChartTitle.TextProperties.Text = "Muli-Category"
chart.ChartTitle.TextProperties.IsCentered = True
chart.ChartTitle.Height = 30
chart.HasTitle = True

# Data for series
Series1 = [7.7, 8.9, 7, 6, 7, 8]

# Set series text
chart.ChartData[0,2].Text = "Series1"

# Set category text
chart.ChartData[1,0].Text = "Grp 1"
chart.ChartData[3,0].Text = "Grp 2"
chart.ChartData[5,0].Text = "Grp 3"

chart.ChartData[1,1].Text = "A"
chart.ChartData[2,1].Text = "B"
chart.ChartData[3,1].Text = "C"
chart.ChartData[4,1].Text = "D"
chart.ChartData[5,1].Text = "E"
chart.ChartData[6,1].Text = "F"

# Fill data for chart
i = 0
while i < len(Series1):
    chart.ChartData[i + 1,2].NumberValue = Series1[i]
    i += 1

# Set series label
chart.Series.SeriesLabel = chart.ChartData["C1","C1"]
# Set category label
chart.Categories.CategoryLabels = chart.ChartData["A2","B7"]

# Set values for series
chart.Series[0].Values = chart.ChartData["C2","C7"]

# Set if the category axis has multiple levels
chart.PrimaryCategoryAxis.HasMultiLvlLbl = True
# Merge same label
chart.PrimaryCategoryAxis.IsMergeSameLabel = True
```

---

# spire.presentation python chart protection
## Protect chart data in PowerPoint presentation
```python
#Get the first shape from slide and convert it as IChart.
chart = presentation.Slides[0].Shapes[0] if isinstance(presentation.Slides[0].Shapes[0], IChart) else None

#Set the Boolean value of IChart.IsDataProtect as true.
chart.IsDataProtect = True
```

---

# spire.presentation python chart
## remove chart from PowerPoint slide
```python
#Get the first slide from the document.
slide = presentation.Slides[0]

#Remove chart from the slide.
for i, unusedItem in enumerate(slide.Shapes):
    shape = slide.Shapes[i] if isinstance(slide.Shapes[i], IShape) else None
    if isinstance(shape, IChart):
        slide.Shapes.Remove(shape)
```

---

# spire.presentation python chart
## remove tick marks from chart axis
```python
#Get the chart that need to be adjusted the number format and remove the tick marks.
chart = presentation.Slides[0].Shapes[0] if isinstance(presentation.Slides[0].Shapes[0], IChart) else None

#Set percentage number format for the axis value of chart.
chart.PrimaryValueAxis.NumberFormat = "0#\\%"

#Remove the tick marks for value axis and category axis.
chart.PrimaryValueAxis.MajorTickMark = TickMarkType.TickMarkNone
chart.PrimaryValueAxis.MinorTickMark = TickMarkType.TickMarkNone
chart.PrimaryCategoryAxis.MajorTickMark = TickMarkType.TickMarkNone
chart.PrimaryCategoryAxis.MinorTickMark = TickMarkType.TickMarkNone
```

---

# Save Chart as Image
## Convert a chart in a presentation slide to an image format
```python
# Save chart as image in .png format
image = ppt.Slides[0].Shapes.SaveAsImage(0)
```

---

# Scale Bubble Chart Size
## Change the bubble size scale in a bubble chart presentation
```python
#Get the chart from the first presentation slide.
chart = presentation.Slides[0].Shapes[0] if isinstance(presentation.Slides[0].Shapes[0], IChart) else None

#Scale the bubble size, the range value is from 0 to 300.
chart.BubbleScale = 50
```

---

# spire.presentation python axis position
## set chart axis position
```python
#Get chart on the first slide
Chart = ppt.Slides[0].Shapes[0] if isinstance(ppt.Slides[0].Shapes[0], IChart) else None

#Set axis position
Chart.PrimaryValueAxis.CrossBetweenType = CrossBetweenType.MidpointOfCategory
```

---

# spire.presentation python axis type
## set chart axis type as date axis with month scale
```python
#Get the chart
chart = ppt.Slides[0].Shapes[1] if isinstance(ppt.Slides[0].Shapes[1], IChart) else None

chart.PrimaryCategoryAxis.AxisType = AxisType.DateAxis
chart.PrimaryCategoryAxis.MajorUnitScale = ChartBaseUnitType.Months
```

---

# Spire.Presentation Python Chart Border Style
## Set border style for a chart in PowerPoint presentation
```python
#Get chart on the first slide
chart = presentation.Slides[0].Shapes[0] if isinstance(presentation.Slides[0].Shapes[0], IChart) else None

#Set border style
chart.Line.FillFormat.FillType = FillFormatType.Solid
chart.Line.FillFormat.SolidFillColor.Color = Color.get_Red()
chart.BorderRoundedCorners = True
```

---

# spire.presentation python chart
## Set chart data label range
```python
#Set data for the chart
cellRange = chart.ChartData["F1"]
cellRange.Text = "labelA"
cellRange = chart.ChartData["F2"]
cellRange.Text = "labelB"
cellRange = chart.ChartData["F3"]
cellRange.Text = "labelC"
cellRange = chart.ChartData["F4"]
cellRange.Text = "labelD"

#Set data label ranges
chart.Series[0].DataLabelRanges = chart.ChartData["F1","F4"]

#Add data label
dataLabel1 = chart.Series[0].DataLabels.Add()
dataLabel1.ID = 0
#Show the value
dataLabel1.LabelValueVisible = False
#Show the label string
dataLabel1.ShowDataLabelsRange = True
```

---

# Spire.Presentation Chart Number Format
## Set number formats for chart data, axis, and labels in a presentation
```python
#Get chart on the first slide
chart = presentation.Slides[0].Shapes[0] if isinstance(presentation.Slides[0].Shapes[0], IChart) else None

#Set the number format for Axis
chart.PrimaryValueAxis.NumberFormat = "#,##0.00"

#Set the DataLabels format for Axis
chart.Series[0].DataLabels.LabelValueVisible = True
chart.Series[0].DataLabels.PercentValueVisible = False
chart.Series[0].DataLabels.NumberFormat = "#,##0.00"
chart.Series[0].DataLabels.HasDataSource = False

#Set the number format for ChartData
for i in range(1, (chart.Series[0].Values.Count) + 1):
    chart.ChartData[i,1].NumberFormat = "#,##0.00"
```

---

# Spire.Presentation Python Trendline
## Set color and name for trendline in chart
```python
#Find the first trendline in the chart
trendline = chart.Series[0].TrendLines[0] if isinstance(chart.Series[0].TrendLines[0], ITrendlines) else None

#Set name for trendline
trendline.Name = "trendlineName"

#Set color for trendline
trendline.Line.FillType = FillFormatType.Solid
trendline.Line.SolidFillColor.Color = Color.get_Red()
```

---

# spire.presentation python chart
## set data label position
```python
#Get chart on the first slide
Chart = ppt.Slides[0].Shapes[0] if isinstance(ppt.Slides[0].Shapes[0], IChart) else None

#Add data label
label = Chart.Series[0].DataLabels.Add()
#Set the position of the label
label.X = 0.1
label.Y = 0.1
```

---

# Spire.Presentation Chart Data Point Coloring
## Set custom colors for data points in a chart

```python
#Get the chart
chart = ppt.Slides[0].Shapes[0] if isinstance(ppt.Slides[0].Shapes[0], IChart) else None

#Initialize an instances of dataPoint
cdp1 = ChartDataPoint(chart.Series[0])

#Specify the datapoint order
cdp1.Index = 0

#Set the color of the datapoint
cdp1.Fill.FillType = FillFormatType.Solid
cdp1.Fill.SolidColor.KnownColor = KnownColors.Orange

#Add the dataPoint to first series
chart.Series[0].DataPoints.Add(cdp1)

#Set the color for the other three data points
cdp2 = ChartDataPoint(chart.Series[0])
cdp2.Index = 1
cdp2.Fill.FillType = FillFormatType.Solid
cdp2.Fill.SolidColor.KnownColor = KnownColors.Gold
chart.Series[0].DataPoints.Add(cdp2)

cdp3 = ChartDataPoint(chart.Series[0])
cdp3.Index = 2
cdp3.Fill.FillType = FillFormatType.Solid
cdp3.Fill.SolidColor.KnownColor = KnownColors.MediumPurple
chart.Series[0].DataPoints.Add(cdp3)

cdp4 = ChartDataPoint(chart.Series[0])
cdp4.Index = 1
cdp4.Fill.FillType = FillFormatType.Solid
cdp4.Fill.SolidColor.KnownColor = KnownColors.ForestGreen
chart.Series[0].DataPoints.Add(cdp4)
```

---

# spire.presentation python set chart display unit
## Set display unit for chart value axis
```python
#Get chart on the first slide
Chart = ppt.Slides[0].Shapes[0] if isinstance(ppt.Slides[0].Shapes[0], IChart) else None

#Set the display unit
Chart.PrimaryValueAxis.DisplayUnit = ChartDisplayUnitType.Hundreds
```

---

# spire.presentation python chart
## set gap width for chart in presentation
```python
#Get chart on the first slide
Chart = ppt.Slides[0].Shapes[0] if isinstance(ppt.Slides[0].Shapes[0], IChart) else None

#Set gap width
Chart.GapWidth = 50
```

---

# spire.presentation python legend options
## set chart legend position and size
```python
#Get chart on the first slide
Chart = ppt.Slides[0].Shapes[0] if isinstance(ppt.Slides[0].Shapes[0], IChart) else None

#Set the legend positon
Chart.ChartLegend.Left = 20
Chart.ChartLegend.Top = 20

#Set the legend size
Chart.ChartLegend.Width = 250
Chart.ChartLegend.Height = 30
```

---

# Spire.Presentation Python Chart
## Set number format for chart axis
```python
#Get chart on the first slide
Chart = ppt.Slides[0].Shapes[0] if isinstance(ppt.Slides[0].Shapes[0], IChart) else None

#Set the number format
Chart.PrimaryCategoryAxis.NumberFormat = "yyyy"
```

---

# Spire Presentation Python Chart
## Set percentage for chart labels
```python
def _GetTotal(ranges):
    total = 0
    for i, unusedItem in enumerate(ranges):
        total += float(ranges[i].Text)
    return total

# Process chart to add percentage labels
dataPontPercent = 0

# Process each series in the chart
for i, unusedItem in enumerate(Chart.Series):
    series = Chart.Series[i]
    # Get the total number
    total = _GetTotal(series.Values)
    for j, unusedItem in enumerate(series.Values):
        # Get the percent
        dataPontPercent = float(series.Values[j].Text) / total * 100
        # Add data labels
        label = series.DataLabels.Add()
        label.LabelValueVisible = True
        # Set the percent text for the label
        label.TextFrame.Paragraphs[0].Text = "{0:.2F} %".format(dataPontPercent)
        label.TextFrame.Paragraphs[0].TextRanges[0].FontHeight = 12
```

---

# Spire.Presentation Chart Data Labels
## Set position and style of chart data labels in PowerPoint presentations
```python
#Add data label to chart and set its id.
label1 = chart.Series[0].DataLabels.Add()
label1.ID = 0

#Set custom position of data label. This position is relative to the default position.
label1.X = 0.1
label1.Y = -0.1

#Set label value visible
label1.LabelValueVisible = True

#Set legend key invisible
label1.LegendKeyVisible = False

#Set category name invisible
label1.CategoryNameVisible = False

#Set series name invisible
label1.SeriesNameVisible = False

#Set Percentage invisible
label1.PercentageVisible = False

#Set border style and fill style of data label
label1.Line.FillType = FillFormatType.Solid
label1.Line.SolidFillColor.Color = Color.get_Blue()
label1.Fill.FillType = FillFormatType.Solid
label1.Fill.SolidColor.Color = Color.get_Orange()
```

---

# spire.presentation python chart title rotation
## Set rotation angle for chart title
```python
#Get the chart
chart = ppt.Slides[0].Shapes[0] if isinstance(ppt.Slides[0].Shapes[0], IChart) else None

chart.ChartTitle.TextProperties.RotationAngle = -30
```

---

# Spire.Presentation for Python - Set Rotation for Data Labels
## This code demonstrates how to set rotation angle for data labels in a chart
```python
# Get chart from presentation
Chart = ppt.Slides[0].Shapes[0] if isinstance(ppt.Slides[0].Shapes[0], IChart) else None

# Set the rotation angle for the datalabels of first series
for i, unusedItem in enumerate(Chart.Series[0].Values):
    datalabel = Chart.Series[0].DataLabels.Add()
    datalabel.ID = i
    datalabel.RotationAngle = 45
```

---

# spire.presentation python chart
## set rotation angle for value axis text
```python
#Get chart on the first slide
Chart = ppt.Slides[0].Shapes[0] if isinstance(ppt.Slides[0].Shapes[0], IChart) else None

#Set the rotation angle for the text on the value axis
Chart.PrimaryValueAxis.TextRotationAngle = 45
```

---

# spire.presentation python chart
## set series overlap for chart
```python
#Create PPT document
ppt = Presentation()
#Get chart on the first slide
Chart = ppt.Slides[0].Shapes[0] if isinstance(ppt.Slides[0].Shapes[0], IChart) else None
#Set overlap
Chart.OverLap = 50
```

---

# Spire.Presentation Python Chart Marker Styling
## Set size and style for chart markers
```python
# Get chart from presentation
chart = ppt.Slides[0].Shapes[0] if isinstance(ppt.Slides[0].Shapes[0], IChart) else None

for i, unusedItem in enumerate(chart.Series[0].Values):
    # Create a ChartDataPoint object and specify the index
    dataPoint = ChartDataPoint(chart.Series[0])
    dataPoint.Index = i

    # Set the fill color of the data marker
    dataPoint.MarkerFill.Fill.FillType = FillFormatType.Solid
    dataPoint.MarkerFill.Fill.SolidColor.Color = Color.get_Yellow()

    # Set the line color of the data marker
    dataPoint.MarkerFill.Line.FillType = FillFormatType.Solid
    dataPoint.MarkerFill.Line.SolidFillColor.Color = Color.get_YellowGreen()

    # Set the size of the data marker
    dataPoint.MarkerSize = 20

    # Set the style of the data marker
    dataPoint.MarkerStyle = ChartMarkerType.Diamond
    chart.Series[0].DataPoints.Add(dataPoint)
```

---

# spire.presentation python chart
## set size for chart plot area
```python
#Get chart on the first slide
Chart = ppt.Slides[0].Shapes[0] if isinstance(ppt.Slides[0].Shapes[0], IChart) else None

#Set width and height for chart plot area
Chart.PlotArea.Width = 250
Chart.PlotArea.Height = 300
```

---

# Spire.Presentation Python Chart Title Font
## Set text font properties for chart title
```python
#Get the chart.
chart = presentation.Slides[0].Shapes[0] if isinstance(presentation.Slides[0].Shapes[0], IChart) else None

#Set the font for the text on chart title area.
chart.ChartTitle.TextProperties.Paragraphs[0].DefaultCharacterProperties.LatinFont = TextFont("Arial Unicode MS")
chart.ChartTitle.TextProperties.Paragraphs[0].DefaultCharacterProperties.Fill.SolidColor.KnownColor = KnownColors.Blue
chart.ChartTitle.TextProperties.Paragraphs[0].DefaultCharacterProperties.FontHeight = 50
```

---

# spire.presentation python chart
## set text font for legend and axis
```python
#Get the chart from a presentation (assuming presentation is already loaded)
chart = presentation.Slides[0].Shapes[0] if isinstance(presentation.Slides[0].Shapes[0], IChart) else None

#Set the font for the text on Chart Legend area.
chart.ChartLegend.TextProperties.Paragraphs[0].DefaultCharacterProperties.Fill.SolidColor.KnownColor = KnownColors.Green
chart.ChartLegend.TextProperties.Paragraphs[0].DefaultCharacterProperties.LatinFont = TextFont("Arial Unicode MS")

#Set the font for the text on Chart Axis area.
chart.PrimaryCategoryAxis.TextProperties.Paragraphs[0].DefaultCharacterProperties.Fill.SolidColor.KnownColor = KnownColors.Red
chart.PrimaryCategoryAxis.TextProperties.Paragraphs[0].DefaultCharacterProperties.Fill.FillType = FillFormatType.Solid
chart.PrimaryCategoryAxis.TextProperties.Paragraphs[0].DefaultCharacterProperties.FontHeight = 10
chart.PrimaryCategoryAxis.TextProperties.Paragraphs[0].DefaultCharacterProperties.LatinFont = TextFont("Arial Unicode MS")
```

---

# spire.presentation python chart
## set tick mark labels on category axis
```python
#Get the chart from the PowerPoint slide.
chart = presentation.Slides[0].Shapes[0] if isinstance(presentation.Slides[0].Shapes[0], IChart) else None

#Rotate tick labels.
chart.PrimaryCategoryAxis.TextRotationAngle = 45

#Specify interval between labels.
chart.PrimaryCategoryAxis.IsAutomaticTickLabelSpacing = False
chart.PrimaryCategoryAxis.TickLabelSpacing = 2

#Change position.
chart.PrimaryCategoryAxis.TickLabelPosition = TickLabelPositionType.TickLabelPositionHigh
```

---

# spire.presentation python tick marks
## set tick marks interval for chart axis
```python
#Create PPT document
ppt = Presentation()
#Get chart from presentation
chart = ppt.Slides[0].Shapes[0] if isinstance(ppt.Slides[0].Shapes[0], IChart) else None
chartAxis = chart.PrimaryCategoryAxis
chartAxis.TickMarkSpacing = 2
```

---

# spire.presentation python chart labels
## Show chart labels in PowerPoint presentation
```python
#Get chart on the first slide
Chart = ppt.Slides[0].Shapes[0] if isinstance(ppt.Slides[0].Shapes[0], IChart) else None

#Show data labels
Chart.Series[0].DataLabels.LabelValueVisible = True
Chart.Series[0].DataLabels.CategoryNameVisible = True
Chart.Series[0].DataLabels.SeriesNameVisible = True
```

---

# spire.presentation python chart
## Vary colors of data markers in the same series
```python
#Get the chart from the presentation.
chart = presentation.Slides[0].Shapes[0] if isinstance(presentation.Slides[0].Shapes[0], IChart) else None

#Create a ChartDataPoint object and specify the index.
dataPoint = ChartDataPoint(chart.Series[0])
dataPoint.Index = 0

#Set the fill color of the data marker.
dataPoint.MarkerFill.Fill.FillType = FillFormatType.Solid
dataPoint.MarkerFill.Fill.SolidColor.Color = Color.get_Red()

#Set the line color of the data marker.
dataPoint.MarkerFill.Line.FillType = FillFormatType.Solid
dataPoint.MarkerFill.Line.SolidFillColor.Color = Color.get_Red()

#Add the data point to the point collection of a series.
chart.Series[0].DataPoints.Add(dataPoint)

dataPoint = ChartDataPoint(chart.Series[0])
dataPoint.Index = 1
dataPoint.MarkerFill.Fill.FillType = FillFormatType.Solid
dataPoint.MarkerFill.Fill.SolidColor.Color = Color.get_Black()
dataPoint.MarkerFill.Line.FillType = FillFormatType.Solid
dataPoint.MarkerFill.Line.SolidFillColor.Color = Color.get_Black()
chart.Series[0].DataPoints.Add(dataPoint)

dataPoint = ChartDataPoint(chart.Series[0])
dataPoint.Index = 2
dataPoint.MarkerFill.Fill.FillType = FillFormatType.Solid
dataPoint.MarkerFill.Fill.SolidColor.Color = Color.get_Blue()
dataPoint.MarkerFill.Line.FillType = FillFormatType.Solid
dataPoint.MarkerFill.Line.SolidFillColor.Color = Color.get_Blue()
chart.Series[0].DataPoints.Add(dataPoint)
```

---

# spire.presentation conversion
## convert ODP to PDF
```python
inputFile ="./Data/toPdf.odp"
outputFile = "ConvertODPtoPDF.pdf"

#Create PPT document
presentation = Presentation()

#Load the PPT document from disk.
presentation.LoadFromFile(inputFile, FileFormat.ODP)

presentation.SaveToFile(outputFile,FileFormat.PDF)
presentation.Dispose()
```

---

# Spire.Presentation Python Conversion
## Convert PPS file to PPTX format
```python
#Create an instance of presentation document
ppt = Presentation()
#Load file
ppt.LoadFromFile(inputFile)

#Save the PPS document to PPTX file format
ppt.SaveToFile(outputFile, FileFormat.Pptx2013)
ppt.Dispose()
```

---

# Spire.Presentation PowerPoint to OFD Conversion
## Convert PowerPoint presentation to OFD format
```python
# Create an instance of presentation document
ppt = Presentation()
# Load file
ppt.LoadFromFile(inputFile)

ppt.SaveToFile(outputFile, FileFormat.OFD)
ppt.Dispose()
```

---

# spire.presentation python slide conversion
## convert individual slide to html format
```python
#Create PPT document
presentation = Presentation()

#Load the PPT document from disk
presentation.LoadFromFile(inputFile)

#Get the first slide
slide = presentation.Slides[0]

#Save the first slide to HTML 
slide.SaveToFile(outputFile, FileFormat.Html)
```

---

# Spire.Presentation file conversion
## Load and save DPS and DPT presentation files
```python
#Create PPT document
presentation = Presentation()

#Load the PPT document from disk.
presentation.LoadFromFile(inputFile, FileFormat.Dps)

presentation.SaveToFile(outputFile, FileFormat.Dps)
presentation.SaveToFile(outputFile2, FileFormat.Dpt)
presentation.Dispose()
```

---

# Spire.Presentation Python ODP to PDF Conversion
## Convert ODP (OpenDocument Presentation) files to PDF format
```python
#Create an instance of presentation document
ppt = Presentation()
#Load file
ppt.LoadFromFile(inputFile)
#Save the Odp document to PDF file format
ppt.SaveToFile(outputFile, FileFormat.PDF)
ppt.Dispose()
```

---

# Spire.Presentation Python Conversion
## Convert PowerPoint slide to SVG format
```python
#Create PPT document
presentation = Presentation()

#Load PPT file from disk
presentation.LoadFromFile(inputFile)

#Convert the second slide to SVG
svgStream = presentation.Slides[1].SaveToSVG()
svgStream.Save(outputFile)
presentation.Dispose()
```

---

# Spire.Presentation Python Conversion
## Convert PowerPoint slide to SVG format
```python
# Load presentation
presentation = Presentation()
presentation.LoadFromFile(inputFile)

# Get the first slide
slide = presentation.Slides[0]

# Save the slide to SVG bytes
svgStream = slide.SaveToSVG()

# Write the bytes to file
svgStream.Save(outputFile)
svgStream.Dispose()
```

---

# Spire.Presentation Python Slide to PDF Conversion
## Convert specific slide to PDF format
```python
#Create PPT document
presentation = Presentation()

#Load the PPT document from disk.
presentation.LoadFromFile(inputFile)

#Get the second slide
slide = presentation.Slides[1]

#Save the second slide to PDF
slide.SaveToFile(outputFile,FileFormat.PDF)
presentation.Dispose()
```

---

# spire.presentation python conversion
## convert PowerPoint presentation to HTML format
```python
#Create an instance of presentation document
ppt = Presentation()

#Load file
ppt.LoadFromFile(inputFile)

#Save the document to HTML format
ppt.SaveToFile(outputFile, FileFormat.Html)
ppt.Dispose()
```

---

# Spire.Presentation Convert PPT to Image
## Convert PowerPoint slides to image files
```python
#Create PPT document
presentation = Presentation()

#Load PPT file from disk
presentation.LoadFromFile(inputFile)

#Save PPT document to images
for i, slide in enumerate(presentation.Slides):
    image = slide.SaveAsImage()
    image.Save("ToImage_img_"+str(i)+".png")
    image.Dispose()

presentation.Dispose()
```

---

# Convert PPT to PDF
## Convert PowerPoint presentation to PDF format
```python
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

---

# Spire.Presentation PDF Conversion
## Convert PowerPoint to PDF with specific page size
```python
#Create a PPT document
presentation = Presentation()

#Set A4 page size
presentation.SlideSize.Type = SlideSizeType.A4

#Set landscape orientation
presentation.SlideSize.Orientation = SlideOrienation.Landscape
```

---

# PowerPoint PPT to PPTX Conversion
## Convert PowerPoint PPT files to PPTX format using Spire.Presentation
```python
inputFile = "./Data/ToPPTX.ppt"
outputFile = "ToPPTX.pptx"

#Create PPT document
presentation = Presentation()

#Load the PPT file from disk
presentation.LoadFromFile(inputFile)

#Save the PPT document to PPTX file format
presentation.SaveToFile(outputFile, FileFormat.Pptx2010)
presentation.Dispose()
```

---

# spire.presentation python conversion
## convert PowerPoint slide to specific size image
```python
#Create an instance of presentation document
ppt = Presentation()
#Load file
ppt.LoadFromFile(inputFile)

#Save the first slide to Image and set the image size to 600*400
img = ppt.Slides[0].SaveAsImageByWH(600, 400)
#Save image to file
img.Save(outputFile)
img.Dispose()
ppt.Dispose()
```

---

# Spire.Presentation Python Conversion
## Convert PowerPoint slides to SVG format
```python
# Create PPT document
presentation = Presentation()

# Load PPT file
presentation.LoadFromFile(inputFile)

# Retain notes when converting to SVG
presentation.IsNoteRetained = True

# Convert each slide to SVG
for index, slide in enumerate(presentation.Slides):
    fileName = outputFile + "ToSVG-" + str(index) + ".svg"
    svgStream = slide.SaveToSVG()
    svgStream.Save(fileName)

# Clean up
presentation.Dispose()
```

---

# PowerPoint to XPS Conversion
## Convert PowerPoint presentation to XPS format using Spire.Presentation
```python
# Create an instance of presentation document
ppt = Presentation()
# Load PowerPoint file
ppt.LoadFromFile(inputFile)
# Save to XPS file
ppt.SaveToFile(outputFile, FileFormat.XPS)
ppt.Dispose()
```

---

# Spire.Presentation Python Image Processing
## Change image size in PowerPoint presentation
```python
scale = 0.5
for slide in presentation.Slides:
    for shape in slide.Shapes:
        if isinstance(shape, IEmbedImage):
            shape.Width = shape.Width * scale
            shape.Height = shape.Height * scale
```

---

# spire.presentation python image extraction
## extract images from powerpoint presentation
```python
#Load a PPT document
ppt = Presentation()
ppt.LoadFromFile(inputFile)

for i, image in enumerate(ppt.Images):
    ImageName = outputFile+"Images_"+str(i)+".png"
    image.Image.Save(ImageName)
```

---

# Spire.Presentation Image Extraction
## Extract images from a specific slide in a PowerPoint presentation
```python
#Assuming ppt is a loaded Presentation object
#Traverse all shapes in the specific slide
for shape in ppt.Slides[1].Shapes:
    #Check if it's a SlidePicture object
    if isinstance(shape, SlidePicture):
        #Extract image from SlidePicture
        image = shape.PictureFill.Picture.EmbedImage.Image
    #Check if it's a PictureShape object
    elif isinstance(shape, PictureShape):
        #Extract image from PictureShape
        image = shape.EmbedImage.Image
```

---

# Spire Presentation Python EMF Image Insertion
## Insert EMF image into PowerPoint presentation slide
```python
# Define image size
img = Image.open(ImageFile)
width = img.width / 1.5
height = img.height / 1.5
rect = RectangleF.FromLTRB (100, 100, width+100, height+100)

# Append the EMF in slide
image = presentation.Slides[0].Shapes.AppendEmbedImageByPath (ShapeType.Rectangle, ImageFile, rect)
image.Line.FillType = FillFormatType.none
```

---

# spire.presentation python image insertion
## insert image into PowerPoint slide
```python
#Insert image to PPT
ImageFile2 = "./Data/Logo1.png"
left = math.trunc(presentation.SlideSize.Size.Width / float(2)) - 280
rect1 = RectangleF.FromLTRB (left, 140, 120+left, 260)
image = presentation.Slides[0].Shapes.AppendEmbedImageByPath (ShapeType.Rectangle, ImageFile2, rect1)
image.Line.FillType = FillFormatType.none
```

---

# spire.presentation python remove images
## remove all images from a slide
```python
# Create a presentation object
ppt = Presentation()

# Get the first slide
slide = ppt.Slides[0]

# Iterate through shapes in reverse order
for i in range(slide.Shapes.Count - 1, -1, -1):
    # Check if it is a SlidePicture object
    if isinstance(slide.Shapes[i], SlidePicture):
        # Remove the image
        slide.Shapes.RemoveAt(i)
```

---

# spire.presentation python image frame formatting
## Set format of an image frame in a presentation
```python
#Set the formatting of the image frame
pptImage.Line.FillFormat.FillType = FillFormatType.Solid
pptImage.Line.FillFormat.SolidFillColor.Color = Color.get_LightBlue()
pptImage.Line.Width = 5
pptImage.Rotation = -45
```

---

# Spire.Presentation Python Image Transparency
## Set transparency for an image in a PowerPoint presentation
```python
#Create an instance of presentation document
ppt = Presentation()

#Define image path and rectangle
imagePath = "./Data/Logo1.png"
rect1 = RectangleF.FromLTRB (200, 100, 400, 300)

#Add a shape
shape = ppt.Slides[0].Shapes.AppendShape(ShapeType.Rectangle, rect1)
shape.Line.FillType = FillFormatType.none
#Fill shape with image
shape.Fill.FillType = FillFormatType.Picture
shape.Fill.PictureFill.Picture.Url = imagePath
shape.Fill.PictureFill.FillType = PictureFillType.Stretch
#Set transparency on image
shape.Fill.PictureFill.Picture.Transparency = 50
```

---

# Spire.Presentation Python Image Update
## Update an image in a PowerPoint presentation

```python
#Create an instance of presentation document
ppt = Presentation()

#Get the first slide
slide = ppt.Slides[0]

#Append a new image to replace an existing image
stream = Stream("./Data/iceblueLogo.png")
image = ppt.Images.AppendStream(stream)
stream.Close()

#Replace the image which title is "image1" with the new image
for shape in slide.Shapes:
    if isinstance(shape, SlidePicture):
        if shape.AlternativeTitle == "image1":
            ( shape if isinstance(shape, SlidePicture) else None).PictureFill.Picture.EmbedImage = image
```

---

# Spire.Presentation Python Table
## Add image to table cell
```python
# Get the table from the slide
table = ppt.Slides[0].Shapes[0] if isinstance(ppt.Slides[0].Shapes[0], ITable) else None

# Load and append image to presentation
stream = Stream("./Data/Logo1.png")
pptImg = ppt.Images.AppendStream(stream)
stream.Close()

# Configure table cell to display the image
table[1,1].FillFormat.FillType = FillFormatType.Picture
table[1,1].FillFormat.PictureFill.Picture.EmbedImage = pptImg
table[1,1].FillFormat.PictureFill.FillType = PictureFillType.Stretch
```

---

# spire.presentation python table
## add row to table in PowerPoint
```python
#Get the table within the PowerPoint document.
table = presentation.Slides[0].Shapes[0] if isinstance(presentation.Slides[0].Shapes[0], ITable) else None

#Get the first row.
row = table.TableRows[1]

#Clone the row and add it to the end of table.
table.TableRows.Append(row)
rowCount = table.TableRows.Count

#Get the last row.
lastRow = table.TableRows[rowCount - 1]

#Set new data of the first cell of last row.
lastRow[0].TextFrame.Text = " The first added cell"

#Set new data of the second cell of last row.
lastRow[1].TextFrame.Text = " The second added cell"
```

---

# Spire.Presentation Python Table Operations
## Clone rows and columns in a PowerPoint table
```python
# Define columns with widths and rows with heights
widths = [110, 110, 110]
heights = [50, 30, 30, 30, 30]

# Add table shape to slide
table = presentation.Slides[0].Shapes.AppendTable(math.trunc(presentation.SlideSize.Size.Width / float(2)) - 275, 90, widths, heights)

# Add text to the row 1 cell 1
table[0,0].TextFrame.Text = "Row 1 Cell 1"

# Add text to the row 1 cell 2
table[1,0].TextFrame.Text = "Row 1 Cell 2"

# Clone row 1 at end of table
table.TableRows.Append(table.TableRows[0])

# Add text to the row 2 cell 1
table[0,1].TextFrame.Text = "Row 2 Cell 1"

# Add text to the row 2 cell 2
table[1,1].TextFrame.Text = "Row 2 Cell 2"

# Clone row 2 as the 4th row of table
table.TableRows.Insert(3, table.TableRows[1])

#Clone column 1 at end of table
table.ColumnsList.Add(table.ColumnsList[0])

#Clone the 2nd column at 4th column index
table.ColumnsList.Insert(3, table.ColumnsList[1])
```

---

# spire.presentation python table
## create and format table in presentation
```python
# Define table dimensions
widths = [100, 100, 150, 100, 100]
heights = [15, 15, 15, 15, 15, 15, 15, 15, 15, 15, 15, 15, 15]

# Add new table to PPT
left = math.trunc(presentation.SlideSize.Size.Width / float(2)) - 275
table = presentation.Slides[0].Shapes.AppendTable(left, 90, widths, heights)

# Add data to table
for i in range(0, 13):
    for j in range(0, 5):
        # Fill the table with data
        table[j,i].TextFrame.Text = dataStr[i][j]
        
        # Set the Font
        table[j,i].TextFrame.Paragraphs[0].TextRanges[0].LatinFont = TextFont("Arial Narrow")

# Set the alignment of the first row to Center
for i in range(0, 5):
    table[i,0].TextFrame.Paragraphs[0].Alignment = TextAlignmentType.Center

# Set the style of table
table.StylePreset = TableStylePreset.LightStyle3Accent1
```

---

# spire.presentation python table
## edit table data and style
```python
# Get the table in PowerPoint document
table = None
for shape in presentation.Slides[0].Shapes:
    if isinstance(shape, ITable):
        table = shape

# Change the style of table
table.StylePreset = TableStylePreset.LightStyle1Accent2

# Edit table data and style
for i, unusedItem in enumerate(table.ColumnsList):
    # Replace the data in cell
    table[i,2].TextFrame.Text = "new_data"
    
    # Set the highlight color
    table[i,2].TextFrame.TextRange.HighlightColor.Color = Color.get_BlueViolet()
```

---

# Spire.Presentation Table Cell Color Filling
## Fill all cells in a table with a solid color
```python
#Fill the table cells with color.
table = None
for shape in presentation.Slides[0].Shapes:
    if isinstance(shape, ITable):
        table = shape
        for row in table.TableRows:
            for cell in row:
                cell.FillFormat.FillType = FillFormatType.Solid
                cell.FillFormat.SolidColor.Color = Color.get_Pink()
```

---

# Spire.Presentation Python Table Formatting
## Fill a particular table row with color
```python
# Find the table in the first slide
table = None
for shape in presentation.Slides[0].Shapes:
    if isinstance(shape, ITable):
        table = shape

        # Get the second row (index 1)
        row = table.TableRows[1]
        # Fill each cell in the row with pink color
        for cell in row:
            cell.FillFormat.FillType = FillFormatType.Solid
            cell.FillFormat.SolidColor.Color = Color.get_Pink()
```

---

# Spire.Presentation Python Table Cell Border Colors
## Get border colors and display color of a table cell in PowerPoint

```python
# Get the table in the first slide
table = ppt.Slides[0].Shapes[0] if isinstance(ppt.Slides[0].Shapes[0], ITable) else None

# Get borders' color of the first cell
sb = []
sb.append("Color of left border:" + table[0,0].BorderLeftDisplayColor.ToString())
sb.append("Color of top border:" + table[0,0].BorderTopDisplayColor.ToString())
sb.append("Color of right border:" + table[0,0].BorderRightDisplayColor.ToString())
sb.append("Color of bottom border:" + table[0,0].BorderBottomDisplayColor.ToString())

# Get display color of the first cell
sb.append("Color of cell:" + table[0,0].DisplayColor.ToString())
```

---

# Identify Merged Cells in PowerPoint Tables
## This code identifies and provides information about merged cells in tables within a PowerPoint presentation.

```python
# Get the first slide
slide = presentation.Slides[0]
for shape in slide.Shapes:
    # Verify if it is table
    if isinstance(shape, ITable):
        table = shape
        for r, unusedItem in enumerate(table.TableRows):
            for c, unusedItem in enumerate(table.ColumnsList):
                # Get cell
                currentCell = table.TableRows[r][c]
                # Identify if it is merged cell
                if currentCell.RowSpan > 1 or currentCell.ColSpan > 1:
                    output = "Cell {0:s}:{1:s} is a part of merged cell with RowSpan={2:s} and ColSpan={3:s} starting from Cell {4:s}:{5:s}.".format(str(r),str( c), str(currentCell.RowSpan), str(currentCell.ColSpan), str(currentCell.FirstRowIndex), str(currentCell.FirstColumnIndex))
```

---

# spire.presentation table aspect ratio locking
## Lock aspect ratio of tables in PowerPoint presentation
```python
#Get the first slide
slide = presentation.Slides[0]
for shape in slide.Shapes:
    #Verify if it is table
    if isinstance(shape, ITable):
        table = shape
        #Lock aspect ratio
        table.ShapeLocking.AspectRatioProtection = True
```

---

# Spire.Presentation Table Cell Merging
## Merge cells in a PowerPoint table
```python
# Find table in slide shapes
for shape in presentation.Slides[0].Shapes:
    if isinstance(shape, ITable):
        table = shape
        
        # Merge the second row and third row of the first column
        table.MergeCells(table[0,1], table[0,2], False)
        
        table.MergeCells(table[3,4], table[4,4], True)
```

---

# spire.presentation python table manipulation
## remove rows and columns from a table in PowerPoint presentation
```python
#Get the table in PPT document
table = None
for shape in presentation.Slides[0].Shapes:
    if isinstance(shape, ITable):
        table = shape

        #Remove the second column
        table.ColumnsList.RemoveAt(1, False)

        #Remove the second row
        table.TableRows.RemoveAt(1, False)
```

---

# spire.presentation python table border
## remove table border style in PowerPoint presentation
```python
# Iterate through all slides
for slide in presentation.Slides:
    # Iterate through all shapes in each slide
    for shape in slide.Shapes:
        # Check if the shape is a table
        if isinstance(shape, ITable):
            # Iterate through all rows in the table
            for row in shape.TableRows:
                # Iterate through all cells in each row
                for cell in row:
                    # Remove border styles for each cell
                    cell.BorderTop.FillType = FillFormatType.none
                    cell.BorderBottom.FillType = FillFormatType.none
                    cell.BorderLeft.FillType = FillFormatType.none
                    cell.BorderRight.FillType = FillFormatType.none
```

---

# Spire.Presentation Python Table Removal
## Remove tables from PowerPoint slides
```python
#Get the tables within the PPT document.
shape_tems = []

for shape in presentation.Slides[0].Shapes:
    if isinstance(shape, ITable):
        #Add new table to table list.
        shape_tems.append(shape)

#Remove all the tables form the first slide.
for shape in shape_tems:
    presentation.Slides[0].Shapes.Remove(shape)
```

---

# Spire.Presentation Table Alignment
## Set horizontal and vertical alignment in table cells
```python
#Find table in the presentation
table = None
for shape in presentation.Slides[0].Shapes:
    if isinstance(shape, ITable):
        table = shape

        #Horizontal Alignment
        #Set the horizontal alignment for the cells in first column 
        table[0,1].TextFrame.Paragraphs[0].Alignment = TextAlignmentType.Left
        table[0,2].TextFrame.Paragraphs[0].Alignment = TextAlignmentType.Center
        table[0,3].TextFrame.Paragraphs[0].Alignment = TextAlignmentType.Right
        table[0,4].TextFrame.Paragraphs[0].Alignment = TextAlignmentType.Justify

        #Vertical Alignment
        #Set the vertical alignment for the cells in second column 
        table[1,1].TextAnchorType = TextAnchorType.Top
        table[1,2].TextAnchorType = TextAnchorType.Center
        table[1,3].TextAnchorType = TextAnchorType.Bottom
        table[1,4].TextAnchorType = TextAnchorType.none

        #Both orientations
        #Set the both horizontal and vertical alignment for the cells in the third column 
        table[2,1].TextFrame.Paragraphs[0].Alignment = TextAlignmentType.Left
        table[2,1].TextAnchorType = TextAnchorType.Top

        table[2,2].TextFrame.Paragraphs[0].Alignment = TextAlignmentType.Right
        table[2,2].TextAnchorType = TextAnchorType.Center

        table[2,3].TextFrame.Paragraphs[0].Alignment = TextAlignmentType.Justify
        table[2,3].TextAnchorType = TextAnchorType.Bottom

        table[2,4].TextFrame.Paragraphs[0].Alignment = TextAlignmentType.Center
        table[2,4].TextAnchorType = TextAnchorType.Top
```

---

# Set borders for existing table
## Set table border style and color in PowerPoint presentation
```python
#Get the table from the first slide of the sample document.
slide = presentation.Slides[0]
table = slide.Shapes[0] if isinstance(slide.Shapes[0], ITable) else None

#Set the border type as Inside and the border color as blue.
table.SetTableBorder(TableBorderType.Inside, 1, Color.get_Blue())
```

---

# Spire.Presentation Table Borders
## Set borders for newly created tables in a presentation
```python
# Create a PPT document
presentation = Presentation()

# Set the table width and height for each table cell.
tableWidth = [100, 100, 100, 100, 100]
tableHeight = [20, 20]

# Traverse all the border type of the table.
for item in TableBorderType:
    # Add a table to the presentation slide with the setting width and height
    itable = presentation.Slides.Append().Shapes.AppendTable(100, 100, tableWidth, tableHeight)

    # Add some text to the table cell.
    itable.TableRows[0][0].TextFrame.Text = "Row"
    itable.TableRows[1][0].TextFrame.Text = "Column"

    # Set the border type, border width and the border color for the table.
    itable.SetTableBorder(item, 1.5, Color.get_Red())
```

---

# Spire.Presentation Table Header
## Set the first row of a table as header in PowerPoint
```python
# Find table in the presentation
for shape in presentation.Slides[0].Shapes:
    if isinstance(shape, ITable):
        table = shape

# Set first row as header
table.FirstRow = True
```

---

# spire.presentation python table
## set row height and column width in a table
```python
#Get the table
table = None
for shape in ppt.Slides[0].Shapes:
    if isinstance(shape, ITable):
        table = shape

        #Set the height for the rows
        table.TableRows[0].Height = 100
        table.TableRows[1].Height = 80
        table.TableRows[2].Height = 60
        table.TableRows[3].Height = 40
        table.TableRows[4].Height = 20

        #Set the column width
        table.ColumnsList[0].Width = 60
        table.ColumnsList[1].Width = 80
        table.ColumnsList[2].Width = 120
        table.ColumnsList[3].Width = 140
        table.ColumnsList[4].Width = 160
```

---

# Setting Table Border Style in PowerPoint
## This code demonstrates how to set border styles for tables in a PowerPoint presentation
```python
# Find the table by looping through all the slides, and then set borders for it.
for slide in presentation.Slides:
    for shape in slide.Shapes:
        if isinstance(shape, ITable):
            for row in shape.TableRows:
                for cell in row:
                    cell.BorderTop.FillType = FillFormatType.Solid
                    cell.BorderBottom.FillType = FillFormatType.Solid
                    cell.BorderLeft.FillType = FillFormatType.Solid
                    cell.BorderRight.FillType = FillFormatType.Solid
```

---

# Spire.Presentation Python Table Style
## Set table style in PowerPoint presentation
```python
#Get the table
table = None
for shape in ppt.Slides[0].Shapes:
    if isinstance(shape, ITable):
        table = shape

        #Set the table style from TableStylePreset and apply it to selected table
        table.StylePreset = TableStylePreset.MediumStyle1Accent2
```

---

# Spire.Presentation Table Text Formatting
## Set text formatting for table cells in PowerPoint presentation
```python
# Get the first slide
slide = presentation.Slides[0]

# Find table in slide
for shape in slide.Shapes:
    if isinstance(shape, ITable):
        table = shape

        # Set text alignment and italic style
        cell1 = table.TableRows[0][0]
        cell1.TextAnchorType = TextAnchorType.Top
        cell1.TextFrame.TextRange.Format.IsItalic = TriState.TTrue

        # Set text color and cell background
        cell2 = table.TableRows[1][0]
        cell2.TextFrame.TextRange.Fill.FillType = FillFormatType.Solid
        cell2.TextFrame.TextRange.Fill.SolidColor.Color = Color.get_Green()
        cell2.FillFormat.FillType = FillFormatType.Solid
        cell2.FillFormat.SolidColor.Color = Color.get_LightGray()

        # Set font, size and highlight
        cell3 = table.TableRows[2][2]
        cell3.TextFrame.TextRange.FontHeight = 12
        cell3.TextFrame.TextRange.LatinFont = TextFont("Arial Black")
        cell3.TextFrame.TextRange.HighlightColor.Color = Color.get_YellowGreen()

        # Set cell margins and borders
        cell4 = table.TableRows[2][1]
        cell4.MarginLeft = 20
        cell4.MarginTop = 30
        cell4.BorderTop.FillType = FillFormatType.Solid
        cell4.BorderTop.SolidFillColor.Color = Color.get_Red()
        cell4.BorderBottom.FillType = FillFormatType.Solid
        cell4.BorderBottom.SolidFillColor.Color = Color.get_Red()
        cell4.BorderLeft.FillType = FillFormatType.Solid
        cell4.BorderLeft.SolidFillColor.Color = Color.get_Red()
        cell4.BorderRight.FillType = FillFormatType.Solid
        cell4.BorderRight.SolidFillColor.Color = Color.get_Red()
```

---

# Spire.Presentation Python Table
## Split a specific table cell in a PowerPoint presentation
```python
#Get the first slide.
slide = presentation.Slides[0]

#Get the table.
table = slide.Shapes[0]

#Split cell [1, 2] into 3 rows and 2 columns.
table[1,2].Split(3, 2)
```

---

# spire.presentation python table traversal
## Traverse through cells in a PowerPoint table
```python
#Get the table.
table = None
for shape in presentation.Slides[0].Shapes:
    if isinstance(shape, ITable):
        table = shape

        #Traverse through the cells of table.
        for row in table.TableRows:
            for cell in row:
                # Get text from each cell
                cell.TextFrame.Text
```

---

# Spire.Presentation Python Hyperlink
## Add hyperlink to image in PowerPoint presentation
```python
#Get the first slide.
slide = presentation.Slides[0]

#Add image to slide.
rect = RectangleF.FromLTRB (480, 350, 640, 510)
image = slide.Shapes.AppendEmbedImageByPath (ShapeType.Rectangle, "./Data/Logo1.png", rect)

#Add hyperlink to the image.
hyperlink = ClickHyperlink("https://www.e-iceblue.com")
image.Click = hyperlink
```

---

# Spire.Presentation Python Hyperlink
## Add hyperlink to SmartArt nodes
```python
#Get the smartArt shape
sr = ppt.Slides[0].Shapes[0] if isinstance(ppt.Slides[0].Shapes[0], ISmartArt) else None
#Add hyperlinks to the nodes
node = sr.Nodes[0]
node.Click = ClickHyperlink(ppt.Slides[1])
node = sr.Nodes[1]
node.Click = ClickHyperlink(ppt.Slides[2])
node = sr.Nodes[2]
node.Click = ClickHyperlink(ppt.Slides[3])
```

---

# spire.presentation python hyperlink
## add hyperlink to text in presentation
```python
#Find the text we want to add link to it.
shape = presentation.Slides[0].Shapes[0] if isinstance(presentation.Slides[0].Shapes[0], IAutoShape) else None
tp = shape.TextFrame.Paragraphs[0]
temp = tp.Text

#Split the original text.
textToLink = "Spire.Presentation"
strSplit = temp.split("Spire.Presentation")

#Clear all text.
tp.TextRanges.Clear()

#Add new text.
tr = TextRange(strSplit[0])
tp.TextRanges.Append(tr)

#Add the hyperlink.
tr = TextRange(textToLink)
tr.ClickAction.Address = "https://www.e-iceblue.com/Introduce/presentation-for-python.html"
tp.TextRanges.Append(tr)
```

---

# spire.presentation python hyperlink
## change hyperlink color in presentation
```python
#Get the first slide
slide = presentation.Slides[0]

#Get the theme of the slide
theme = slide.Theme

#Change the color of hyperlink to red
theme.ColorScheme.HyperlinkColor.Color = Color.get_Red()
```

---

# spire.presentation python get linked slide
## Get the linked slide information from a shape in a PowerPoint presentation
```python
#Get the second slide
slide = presentation.Slides[1]

#Get the first shape of the second slide
shape = slide.Shapes[0] if isinstance(slide.Shapes[0], IAutoShape) else None

#Get the linked slide index
if shape.Click.ActionType == HyperlinkActionType.GotoSlide:
    targetSlide = shape.Click.TargetSlide
    linkedSlideNumber = str(targetSlide.SlideNumber)
```

---

# Spire.Presentation Hyperlink with Outline Style
## Create a hyperlink with custom outline style in PowerPoint presentation
```python
#Create a PPT document
presentation = Presentation()

#Add new shape to PPT document
left = math.trunc(presentation.SlideSize.Size.Width / float(2)) - 255
rec = RectangleF.FromLTRB (left, 120, 400+left, 220)
shape = presentation.Slides[0].Shapes.AppendShape(ShapeType.Rectangle, rec)
shape.Fill.FillType = FillFormatType.none
shape.Line.FillType = FillFormatType.none

#Add a paragraph with hyperlink
para1 = TextParagraph()
tr1 = TextRange("Click to know more about Spire.Presentation")
tr1.ClickAction.Address = "https://www.e-iceblue.com/Introduce/presentation-for-python.html"
para1.TextRanges.Append(tr1)

#Set the format of textrange
tr1.Format.FontHeight = 20
tr1.IsItalic = TriState.TTrue

#Set the outline format of textrange
tr1.TextLineFormat.FillFormat.FillType = FillFormatType.Solid
tr1.TextLineFormat.FillFormat.SolidFillColor.Color = Color.get_LightSeaGreen()
tr1.TextLineFormat.JoinStyle = LineJoinType.Round
tr1.TextLineFormat.Width = 2

#Add the paragraph to shape
shape.TextFrame.Paragraphs.Append(para1)
shape.TextFrame.Paragraphs.Append(TextParagraph())
```

---

# spire.presentation python hyperlink
## create hyperlinks in PowerPoint presentation
```python
#Create a PPT document
presentation = Presentation()

#Add new shape to PPT document
left = math.trunc(presentation.SlideSize.Size.Width / float(2)) - 255
rec = RectangleF.FromLTRB (left, 120, 500+left, 400)
shape = presentation.Slides[0].Shapes.AppendShape(ShapeType.Rectangle, rec)
shape.Fill.FillType = FillFormatType.none
shape.Line.Width = 0

#Add some paragraphs with hyperlinks
para1 = TextParagraph()
tr = TextRange("E-iceblue")
tr.Fill.FillType = FillFormatType.Solid
tr.Fill.SolidColor.Color = Color.get_Blue()
para1.TextRanges.Append(tr)
para1.Alignment = TextAlignmentType.Center
shape.TextFrame.Paragraphs.Append(para1)
shape.TextFrame.Paragraphs.Append(TextParagraph())

#Add paragraph with web hyperlink
para2 = TextParagraph()
tr1 = TextRange("Click to know more about Spire.Presentation.")
tr1.ClickAction.Address = "http://www.e-iceblue.com/Introduce/presentation-for-net-introduce.html"
para2.TextRanges.Append(tr1)
shape.TextFrame.Paragraphs.Append(para2)
shape.TextFrame.Paragraphs.Append(TextParagraph())

para3 = TextParagraph()
tr2 = TextRange("Click to visit E-iceblue Home page.")
tr2.ClickAction.Address = "https://www.e-iceblue.com/"
para3.TextRanges.Append(tr2)
shape.TextFrame.Paragraphs.Append(para3)
shape.TextFrame.Paragraphs.Append(TextParagraph())

para4 = TextParagraph()
tr3 = TextRange("Click to go to the forum to raise questions.")
tr3.ClickAction.Address = "https://www.e-iceblue.com/forum/components-f5.html"
para4.TextRanges.Append(tr3)
shape.TextFrame.Paragraphs.Append(para4)
shape.TextFrame.Paragraphs.Append(TextParagraph())

#Add paragraph with email hyperlink
para5 = TextParagraph()
tr4 = TextRange("Click to contact our sales team via email.")
tr4.ClickAction.Address = "mailto:sales@e-iceblue.com"
para5.TextRanges.Append(tr4)
shape.TextFrame.Paragraphs.Append(para5)
shape.TextFrame.Paragraphs.Append(TextParagraph())

para6 = TextParagraph()
tr5 = TextRange("Click to contact our support team via email.")
tr5.ClickAction.Address = "mailto:support@e-iceblue.com"
para6.TextRanges.Append(tr5)
shape.TextFrame.Paragraphs.Append(para6)
```

---

# Spire.Presentation Python Hyperlink
## Create hyperlink to a specific slide in PowerPoint presentation
```python
# Create a PowerPoint document
presentation = Presentation()

# Append a slide to it
presentation.Slides.Append()

# Add a shape to the second slide
shape = presentation.Slides[1].Shapes.AppendShape(
    ShapeType.Rectangle, RectangleF.FromLTRB(10, 50, 210, 100))
shape.Fill.FillType = FillFormatType.none
shape.Line.FillType = FillFormatType.none
shape.TextFrame.Text = "Jump to the first slide"

# Create a hyperlink based on the shape and the text on it, linking to the first slide
hyperlink = ClickHyperlink(presentation.Slides[0])
shape.Click = hyperlink
shape.TextFrame.TextRange.ClickAction = hyperlink
```

---

# Spire.Presentation Hyperlink
## Create hyperlink to last viewed slide
```python
ppt = Presentation()
slide = ppt.Slides[0]
# Draw a shape
autoShape = slide.Shapes.AppendShape(
    ShapeType.Rectangle, RectangleF.FromLTRB(100, 100, 200, 200))
# Link to last viewed slide show
autoShape.Click = ClickHyperlink.get_LastVievedSlide()
```

---

# Spire.Presentation Hyperlink Modification
## Modify hyperlink address and text in PowerPoint presentation
```python
# Find the hyperlinks you want to edit.
shape = presentation.Slides[0].Shapes[0]

# Edit the link text and the target URL.
shape.TextFrame.TextRange.ClickAction.Address = "http://www.e-iceblue.com"
shape.TextFrame.TextRange.Text = "E-iceblue"
```

---

# Remove Hyperlink from PowerPoint
## This code demonstrates how to remove a hyperlink from a text in a PowerPoint slide
```python
# Create a PowerPoint document.
presentation = Presentation()

# Get the shape and its text with hyperlink.
shape = presentation.Slides[0].Shapes[0] if isinstance(
    presentation.Slides[0].Shapes[0], IAutoShape) else None

# Set the ClickAction property into null to remove the hyperlink.
shape.TextFrame.TextRange.ClickAction = None
```

---

# Spire.Presentation Audio Extraction
## Extract audio from PowerPoint presentation
```python
# Initialize audio data
AudioData = None

# Load a presentation
presentation = Presentation()
presentation.LoadFromFile(inputFile)

# Extract audio from shapes in the first slide
for shape in presentation.Slides[0].Shapes:
    if isinstance(shape, IAudio):
        audio = shape if isinstance(shape, IAudio) else None
        AudioData = audio.Data
        AudioData.SaveToFile(outputFile)

presentation.Dispose()
```

---

# Extract Video from PowerPoint
## This code demonstrates how to extract embedded videos from PowerPoint presentations
```python
# Create PPT document
presentation = Presentation()

# Define a counter for output files
i = 0

# Traverse all the slides of PPT file
for slide in presentation.Slides:
    # Traverse all the shapes of slides
    for shape in slide.Shapes:
        # If shape is IVideo
        if isinstance(shape, IVideo):
            # String for output file
            result = "ExtractVideo_" + str(i) + ".avi"
            # Save the video
            shape.EmbeddedVideoData.SaveToFile(result)
            i += 1
presentation.Dispose()
```

---

# Hide Audio During Presentation Show
## Hide audio objects on a slide during presentation
```python
# Get the first slide
slide = presentation.Slides[0]

# Hide Audio during show
for shape in slide.Shapes:
    if isinstance(shape, IAudio):
        shape.HideAtShowing = True
```

---

# Spire.Presentation Python Audio
## Insert audio into presentation
```python
# Add title
rec_title = RectangleF.FromLTRB(50, 240, 160+50, 50+240)
shape_title = presentation.Slides[0].Shapes.AppendShape(
    ShapeType.Rectangle, rec_title)
shape_title.ShapeStyle.LineColor.Color = Color.get_Transparent()

shape_title.Fill.FillType = FillFormatType.none
para_title = TextParagraph()
para_title.Text = "Audio:"
para_title.Alignment = TextAlignmentType.Center
para_title.TextRanges[0].LatinFont = TextFont("Myriad Pro Light")
para_title.TextRanges[0].FontHeight = 32
para_title.TextRanges[0].IsBold = TriState.TTrue
para_title.TextRanges[0].Fill.FillType = FillFormatType.Solid
para_title.TextRanges[0].Fill.SolidColor.Color = Color.FromArgb(
    255, 68, 68, 68)
shape_title.TextFrame.Paragraphs.Append(para_title)

# Insert audio into the document
audioRect = RectangleF.FromLTRB(220, 240, 80+220, 80+240)
presentation.Slides[0].Shapes.AppendAudioMedia(
    audio_path, audioRect)
```

---

# Spire.Presentation Python Video Insertion
## Insert video into PowerPoint presentation
```python
# Insert video into the document
left = math.trunc(presentation.SlideSize.Size.Width / float(2)) - 125
videoRect = RectangleF.FromLTRB(left, 240, 150+left, 150+240)
video = presentation.Slides[0].Shapes.AppendVideoMedia(
    "video_path.mp4", videoRect)
video.PictureFill.Picture.Url = "thumbnail_path.png"
```

---

# spire.presentation python sound effect extraction
## obtain sound effect properties from presentation
```python
# Create an instance of presentation document
ppt = Presentation()
# Load file
ppt.LoadFromFile(inputFile)

# Get the first slide
slide = ppt.Slides[0]

# Get the audio in a time node
audio = slide.Timeline.MainSequence[0].TimeNodeAudios[0]

# Get the properties of the audio, such as sound name, volume or detect if it's mute
text = []
text.append("SoundName: " + audio.SoundName)
text.append("Volume: " + str(audio.Volume))
text.append("IsMute: " + str(audio.IsMute))
```

---

# Spire.Presentation Video Replacement
## Replace videos in a PowerPoint presentation
```python
# Get videos collection
videos = presentation.Videos

# Traverse all the slides of PPT file
for sld in presentation.Slides:
    # Traverse all the shapes of slides
    for sp in sld.Shapes:
        # If shape is IVideo
        if isinstance(sp, IVideo):
            # Replace the video
            video = sp
            # Append video data from stream
            videoData = videos.AppendByStream(stream)
            video.EmbeddedVideoData = videoData
```

---

# spire.presentation python video play mode
## set video play mode to auto in PowerPoint slides
```python
# Find the video by looping through all the slides and set its play mode as auto.
for slide in presentation.Slides:
    for shape in slide.Shapes:
        if isinstance(shape, IVideo):
            (shape if isinstance(shape, IVideo)
             else None).PlayMode = VideoPlayMode.Auto
```

---

# Spire.Presentation Speaker Notes Management
## Add and retrieve speaker notes in PowerPoint slides
```python
# Get the first slide in the PowerPoint document
slide = presentation.Slides[0]

# Get the NotesSlide in the first slide, if there is no notes, we need to add it firstly
ns = slide.NotesSlide
if ns is None:
    ns = slide.AddNotesSlide()

# Add the text string as the notes
ns.NotesTextFrame.Text = "Speak notes added by Spire.Presentation"

# Get the speaker notes text
notes_text = ns.NotesTextFrame.Text
```

---

# Spire.Presentation Python Comment
## Add comment to PowerPoint slide
```python
# Comment author
author = presentation.CommentAuthors.AddAuthor("E-iceblue", "comment:")

# Add comment
point = PointF.Empty()
point.X = 18
point.Y = 25
presentation.Slides[0].AddComment(
    author, "Add comment", point, DateTime.get_Now())
```

---

# Spire.Presentation Python Example
## Add note to PowerPoint slide
```python
# Add note slide
notesSlide = slide.AddNotesSlide()

# Add paragraph in the notesSlide
paragraph = TextParagraph()
paragraph.Text = "Tips for making effective presentations:"
notesSlide.NotesTextFrame.Paragraphs.Append(paragraph)

paragraph = TextParagraph()
paragraph.Text = "Use the slide master feature to create a consistent and simple design template."
notesSlide.NotesTextFrame.Paragraphs.Append(paragraph)
# Set the bullet type for the paragraph in notesSlide
notesSlide.NotesTextFrame.Paragraphs[1].BulletType = TextBulletType.Numbered
notesSlide.NotesTextFrame.Paragraphs[1].BulletStyle = NumberedBulletStyle.BulletArabicPeriod

paragraph = TextParagraph()
paragraph.Text = "Simplify and limit the number of words on each screen."
notesSlide.NotesTextFrame.Paragraphs.Append(paragraph)
notesSlide.NotesTextFrame.Paragraphs[2].BulletType = TextBulletType.Numbered
notesSlide.NotesTextFrame.Paragraphs[2].BulletStyle = NumberedBulletStyle.BulletArabicPeriod

paragraph = TextParagraph()
paragraph.Text = "Use contrasting colors for text and background."
notesSlide.NotesTextFrame.Paragraphs.Append(paragraph)
notesSlide.NotesTextFrame.Paragraphs[3].BulletType = TextBulletType.Numbered
notesSlide.NotesTextFrame.Paragraphs[3].BulletStyle = NumberedBulletStyle.BulletArabicPeriod
```

---

# spire.presentation python delete comment
## delete comment from PowerPoint slide
```python
# Create a PPT document
presentation = Presentation()

# Delete the third comment
presentation.Slides[0].DeleteComment(presentation.Slides[0].Comments[2])
```

---

# Spire.Presentation Python Comment Extraction
## Extract comments from PowerPoint slides
```python
# Create a PowerPoint document.
presentation = Presentation()

# Load the file from disk.
presentation.LoadFromFile(inputFile)

# Get all comments from the first slide.
comments = presentation.Slides[0].Comments

# Process the comments
cs = []
i = 0
while i < len(comments):
    cs.append(comments[i].Text + "\r\n")
    i += 1
```

---

# Extract PowerPoint Slide Comments
## Retrieve comment text, author name, and posted time from a PowerPoint presentation
```python
# Loop through comments
for commentAuthor in presentation.CommentAuthors:
    for comment in commentAuthor.CommentsList:
        # Get comment information
        commentText = comment.Text
        authorName = comment.AuthorName
        time = comment.DateTime
```

---

# Spire.Presentation PowerPoint to SVG Conversion with Notes
## Convert PowerPoint slides to SVG format while retaining notes
```python
# Create a PowerPoint document.
presentation = Presentation()

# Load the PowerPoint file.
presentation.LoadFromFile("input.pptx")

# Retain the notes while converting PowerPoint file to svg file.
presentation.IsNoteRetained = True

# Convert presentation slides to svg file.
for slide in presentation.Slides:
    stream = slide.SaveToSVG()
    stream.Save("output.svg")
    stream.Close()
```

---

# Spire.Presentation Remove Notes from Slide
## Remove notes from a specific slide in a PowerPoint presentation
```python
# Get the first slide
slide = presentation.Slides[0]

# Get note slide
note = slide.NotesSlide
# Clear note text
note.NotesTextFrame.Text = ""
```

---

# Remove Speaker Notes from PowerPoint
## This code demonstrates how to remove speaker notes from a PowerPoint slide using Spire.Presentation
```python
# Get the first slide from the presentation
slide = presentation.Slides[0]

# Remove the first speaker note
slide.NotesSlide.NotesTextFrame.Paragraphs.RemoveAt(1)
```

---

# Spire.Presentation Header and Footer
## Set header and footer properties in PowerPoint presentation
```python
# Add footer
presentation.SetFooterText("Demo of Spire.Presentation")

# Set the footer visible
presentation.FooterVisible = True

# Set the page number visible
presentation.SlideNumberVisible = True

# Set the date visible
presentation.DateTimeVisible = True
```

---

# Spire.Presentation Python Header and Footer Management
## Manage header and footer in NotesMaster slide
```python
# Set the note Masters header and footer
noteMasterSlide = presentation.NotesMaster
if noteMasterSlide is not None:
    for shape in noteMasterSlide.Shapes:
        if shape.Placeholder is not None:
            if shape.Placeholder.Type is PlaceholderType.Header:
                (shape if isinstance(shape, IAutoShape)
                 else None).TextFrame.Text = "change the header by Spire"
            if shape.Placeholder.Type is PlaceholderType.Footer:
                (shape if isinstance(shape, IAutoShape)
                 else None).TextFrame.Text = "change the footer by Spire"
```

---

# Spire.Presentation SmartArt Node Access
## Access child nodes of SmartArt shapes in PowerPoint presentations
```python
for shape in presentation.Slides[0].Shapes:
    if isinstance(shape, ISmartArt):
        # Get the SmartArt and collect nodes
        sa = shape
        nodes = sa.Nodes

        position = 0
        # Access the parent node at position 0
        node = nodes[position]
        # Traverse through all child nodes inside SmartArt
        for i, node in enumerate(node.ChildNodes):
            # Access SmartArt child node at index i
            childnode = node
```

---

# Access SmartArt Nodes in Presentation
## Extract text, level, and position information from SmartArt nodes
```python
# Create and load presentation
presentation = Presentation()
presentation.LoadFromFile(inputFile)

# Access SmartArt nodes in the first slide
for shape in presentation.Slides[0].Shapes:
    if isinstance(shape, ISmartArt):
        # Get the SmartArt
        sa = shape

        # Access all nodes in the SmartArt
        nodes = sa.Nodes

        # Traverse through all nodes
        for i, unusedItem in enumerate(nodes):
            # Access SmartArt node at index i
            node = nodes[i]
            # Access node parameters
            node_text = node.TextFrame.Text  # Node text
            node_level = node.Level          # Node level
            node_position = node.Position    # Node position
```

---

# Access SmartArt Layout
## Extract SmartArt layout type from PowerPoint presentation
```python
# Iterate through shapes and check for SmartArt
for shape in presentation.Slides[0].Shapes:
    if isinstance(shape, ISmartArt):
        layout = str(shape.LayoutType)
```

---

# spire.presentation python smartart
## accessing specific child node in smartart
```python
# Create PPT document
presentation = Presentation()

for shape in presentation.Slides[0].Shapes:
    if isinstance(shape, ISmartArt):
        # Get the SmartArt
        sa = shape if isinstance(shape, ISmartArt) else None

        # Get SmartArt node collection
        nodes = sa.Nodes

        # Access SmartArt node at index 0
        node = nodes[0]

        # Access SmartArt child node at index 1
        childNode = node.ChildNodes[1]

        # Get the SmartArt child node parameters
        outString = "Node text = "+childNode.TextFrame.Text+", Node level = " + \
            str(childNode.Level)+", Node Position = "+str(childNode.Position)
```

---

# spire.presentation python smartart
## add nodes to SmartArt by position
```python
# Iterate through shapes in the slide
for shape in presentation.Slides[0].Shapes:
    if isinstance(shape, ISmartArt):
        # Get the SmartArt
        smartArt = shape if isinstance(shape, ISmartArt) else None

        position = 0
        # Add a new node at specific position
        node = smartArt.Nodes.AddNodeByPosition(position)
        # Add text and set the text style
        node.TextFrame.Text = "New Node"
        node.TextFrame.TextRange.Fill.FillType = FillFormatType.Solid
        node.TextFrame.TextRange.Fill.SolidColor.KnownColor = KnownColors.Red

        # Get a node
        node = smartArt.Nodes[1]
        position = 1
        # Add a new child node at specific position
        childNode = node.ChildNodes.AddNodeByPosition(position)
        # Add text and set the text style
        node.TextFrame.Text = "New child node"
        node.TextFrame.TextRange.Fill.FillType = FillFormatType.Solid
        node.TextFrame.TextRange.Fill.SolidColor.KnownColor = KnownColors.Blue
```

---

# Spire.Presentation SmartArt Node Management
## Add a new node to SmartArt and format its text
```python
# Get the SmartArt
sa = presentation.Slides[0].Shapes[0] if isinstance(
    presentation.Slides[0].Shapes[0], ISmartArt) else None

# Add a node
node = sa.Nodes.AddNode()
# Add text and set the text style
node.TextFrame.Text = "AddText"
node.TextFrame.TextRange.Fill.FillType = FillFormatType.Solid
node.TextFrame.TextRange.Fill.SolidColor.KnownColor = KnownColors.HotPink
```

---

# spire.presentation python SmartArt
## set nodes as assistant nodes in SmartArt
```python
# Find SmartArt shapes in the first slide
for shape in presentation.Slides[0].Shapes:
    if isinstance(shape, ISmartArt):
        smartArt = shape
        nodes = smartArt.Nodes

        # Traverse through all nodes inside SmartArt
        for i, unusedItem in enumerate(nodes):
            node = nodes[i]
            # Set non-assistant nodes as assistant nodes
            if not node.IsAssistant:
                node.IsAssistant = True
```

---

# spire.presentation python smartart
## change smartart node text
```python
for shape in presentation.Slides[0].Shapes:
    if isinstance(shape, ISmartArt):
        # Get the SmartArt and collect nodes
        smartArt = shape if isinstance(shape, ISmartArt) else None
        # Obtain the reference of a node by using its Index
        # select second root node
        node = smartArt.Nodes[1]
        # Set the text of the TextFrame
        node.TextFrame.Text = "Second root node"
```

---

# Change SmartArt Color Style
## Modify the color style of SmartArt objects in a PowerPoint presentation
```python
for shape in presentation.Slides[0].Shapes:
    if isinstance(shape, ISmartArt):
        # Get the SmartArt and collect nodes
        smartArt = shape if isinstance(shape, ISmartArt) else None
        # Check SmartArt color type
        if smartArt.ColorStyle == SmartArtColorType.ColoredFillAccent1:
            # Change SmartArt color type
            smartArt.ColorStyle = SmartArtColorType.ColorfulAccentColors
```

---

# Spire.Presentation SmartArt Style Change
## Change SmartArt shape style in PowerPoint presentation
```python
for shape in presentation.Slides[0].Shapes:
    if isinstance(shape, ISmartArt):
        # Get the SmartArt and collect nodes
        smartArt = shape if isinstance(shape, ISmartArt) else None
        # Check SmartArt style
        if smartArt.Style == SmartArtStyleType.SimpleFill:
            # Change SmartArt Style
            smartArt.Style = SmartArtStyleType.Cartoon
```

---

# spire.presentation SmartArt creation
## Create and customize SmartArt shapes in PowerPoint presentations
```python
# Append SmartArt shape to slide
sa = presentation.Slides[0].Shapes.AppendSmartArt(
    200, 60, 300, 300, SmartArtLayoutType.Gear)

# Set type and color of smartart
sa.Style = SmartArtStyleType.SubtleEffect
sa.ColorStyle = SmartArtColorType.GradientLoopAccent3

# Remove all shapes
to_remove = []

for a in sa.Nodes:
    to_remove.append(a)
for subnode in to_remove:
    sa.Nodes.RemoveNode(subnode)

# Add two custom shapes with text
node = sa.Nodes.AddNode()
sa.Nodes[0].TextFrame.Text = "aa"
node = sa.Nodes.AddNode()
node.TextFrame.Text = "bb"
node.TextFrame.TextRange.Fill.FillType = FillFormatType.Solid
node.TextFrame.TextRange.Fill.SolidColor.KnownColor = KnownColors.Black
```

---

# Spire.Presentation Python SmartArt
## Extract text from SmartArt shapes in PowerPoint presentations

```python
# Traverse through all the slides of the PPT file and find the SmartArt shapes.
st = []
for slide in presentation.Slides:
    for shape in slide.Shapes:
        if isinstance(shape, ISmartArt):
            # Extract text from SmartArt and append to the list.
            for node in shape.Nodes:
                st.append(node.TextFrame.Text)
```

---

# SmartArt Node Removal
## Remove a specific node from SmartArt in PowerPoint presentation
```python
# Get the SmartArt and collect nodes
sa = presentation.Slides[0].Shapes[0] if isinstance(
    presentation.Slides[0].Shapes[0], ISmartArt) else None
nodes = sa.Nodes

# Remove the node to specific position
nodes.RemoveNodeByPosition(2)
```

---

# spire.presentation python SmartArt
## set SmartArt link line outline
```python
# Get SmartArt from slide
smartArt = ppt.Slides[0].Shapes[0] if isinstance(
    ppt.Slides[0].Shapes[0], ISmartArt) else None
count = smartArt.Nodes.Count
node = None
# Loop through all SmartArt nodes
for i in range(0, count):
    node = smartArt.Nodes[i]
    # Set the line type
    node.LinkLine.FillType = FillFormatType.Solid
    # Set the line color
    node.LinkLine.SolidFillColor.Color = Color.get_Red()
    # Set the line width
    node.LinkLine.Width = 2
    # Set the line DashStyle
    node.LinkLine.DashStyle = LineDashStyleType.SystemDash
```

---

# Spire.Presentation SmartArt Node Outline
## Set outline properties for SmartArt nodes in a presentation
```python
ppt = Presentation()
smartArt = ppt.Slides[0].Shapes[0] if isinstance(
    ppt.Slides[0].Shapes[0], ISmartArt) else None
count = smartArt.Nodes.Count
node = None
# Loop through all nodes
for i in range(0, count):
    node = smartArt.Nodes[i]
    # Set the fill format type
    node.Line.FillType = FillFormatType.Solid
    # Set the line style
    node.Line.Style = TextLineStyle.ThinThin
    # Set the line color
    node.Line.SolidFillColor.Color = Color.get_Red()
    # Set the line width
    node.Line.Width = 2
```

---

# Spire.Presentation Image Watermark
## Add an image as a watermark to a PowerPoint slide
```python
# Create a PowerPoint document.
presentation = Presentation()

# Load the image stream and append to presentation
stream = Stream("Data/Logo.png")
image = presentation.Images.AppendStream(stream)
stream.Close()

# Set the properties of SlideBackground, and then fill the image as watermark.
presentation.Slides[0].SlideBackground.Type = BackgroundType.Custom
presentation.Slides[0].SlideBackground.Fill.FillType = FillFormatType.Picture
presentation.Slides[0].SlideBackground.Fill.PictureFill.FillType = PictureFillType.Stretch
presentation.Slides[0].SlideBackground.Fill.PictureFill.Picture.EmbedImage = image
```

---

# Spire.Presentation Python Watermark
## Add a text watermark to a PowerPoint slide
```python
# Define a rectangle range
left = (presentation.SlideSize.Size.Width - 336.4) / 2
top = (presentation.SlideSize.Size.Height - 110.8) / 2
rect = RectangleF(left, top, 336.4, 110.8)

# Add a rectangle shape with a defined range
shape = presentation.Slides[0].Shapes.AppendShape(
    ShapeType.Rectangle, rect)

# Set the style of the shape
shape.Fill.FillType = FillFormatType.none
shape.ShapeStyle.LineColor.Color = Color.get_White()
shape.Rotation = -45
shape.Locking.SelectionProtection = True
shape.Line.FillType = FillFormatType.none

# Add text to the shape
shape.TextFrame.Text = "E-iceblue"
textRange = shape.TextFrame.TextRange
# Set the style of the text range
textRange.Fill.FillType = FillFormatType.Solid
textRange.Fill.SolidColor.Color = Color.FromArgb(
    120, Color.get_HotPink().R, Color.get_HotPink().G, Color.get_HotPink().B)
textRange.FontHeight = 50
```

---

# Spire.Presentation Watermark Removal
## Remove text and image watermarks from PowerPoint slides
```python
# Remove text watermark by removing the shape which contains the text string "E-iceblue".
for i, unusedItem in enumerate(presentation.Slides):
    for j, unusedItem in enumerate(presentation.Slides[i].Shapes):
        if isinstance(presentation.Slides[i].Shapes[j], IAutoShape):
            shape = presentation.Slides[i].Shapes[j]
            if shape.TextFrame.Text.find("E-iceblue") != -1:
                presentation.Slides[i].Shapes.Remove(shape)

# Remove image watermark.
for i, unusedItem in enumerate(presentation.Slides):
    presentation.Slides[i].SlideBackground.Fill.FillType = FillFormatType.none
```

---

# Spire.Presentation OLE Embedding
## Embed Excel file as OLE object in PowerPoint presentation
```python
# Create a Presentation document
ppt = Presentation()

# Load an image file
stream = Stream("Data/EmbedExcelAsOLE.png")
oleImage = ppt.Images.AppendStream(stream)
stream.Close()

rec = RectangleF.FromLTRB(80, 60, oleImage.Width+80, oleImage.Height+60)
# Insert an OLE object to presentation based on the Excel data
oleStream = Stream("./Data/EmbedExcelAsOLE.xlsx")
oleObject = ppt.Slides[0].Shapes.AppendOleObject("excel", oleStream, rec)
oleObject.SubstituteImagePictureFillFormat.Picture.EmbedImage = oleImage
oleObject.ProgId = "Excel.Sheet.12"
oleStream.Close()
```

---

# Spire.Presentation for Python - OLE Object Embedding
## Embed a ZIP file into a PowerPoint presentation as an OLE object
```python
# Define rectangle for OLE object position and size
rec = RectangleF.FromLTRB(80, 60, 180, 160)

# Insert the zip object to presentation
ole = ppt.Slides[0].Shapes.AppendOleObject(filePath, stream, rec)
ole.ProgId = "Package"
oleImage = ppt.Images.AppendStream(imageStream)
ole.SubstituteImagePictureFillFormat.Picture.EmbedImage = oleImage
```

---

# Spire.Presentation Python OLE Object Extraction
## Extract OLE objects from PowerPoint presentations and save them as separate files based on their type
```python
# Loop through slides and shapes to find OLE objects
for slide in presentation.Slides:
    for shape in slide.Shapes:
        if isinstance(shape, IOleObject):
            oleObject = shape
            stream = oleObject.Data
            
            # Save OLE object data based on its type
            if oleObject.ProgId == "Excel.Sheet.8":
                stream.Save("ExtractOLEObject.xls")
            elif oleObject.ProgId == "Excel.Sheet.12":
                stream.Save("ExtractOLEObject.xlsx")
            elif oleObject.ProgId == "Word.Document.8":
                stream.Save("ExtractOLEObject.doc")
            elif oleObject.ProgId == "Word.Document.12":
                stream.Save("ExtractOLEObject.docx")
            elif oleObject.ProgId == "PowerPoint.Show.8":
                stream.Save("ExtractOLEObject.ppt")
            elif oleObject.ProgId == "PowerPoint.Show.12":
                stream.Save("ExtractOLEObject.pptx")
            
            stream.Dispose()
```

---

# Spire.Presentation OLE Data Modification
## Core functionality for modifying OLE object data in PowerPoint presentations
```python
# Loop through the slides and shapes
for slide in presentation.Slides:
    for shape in slide.Shapes:
        if isinstance(shape, IOleObject):
            # Find OLE object
            oleObject = shape if isinstance(shape, IOleObject) else None

            # Get its data
            stream = oleObject.Data
            stream2 = Stream()
            if oleObject.ProgId == "PowerPoint.Show.12":
                # Load the PPT stream
                ppt = Presentation()
                ppt.LoadFromStream(stream, FileFormat.Auto)
                # Append an image in slide
                ppt.Slides[0].Shapes.AppendEmbedImageByPath(
                    ShapeType.Rectangle, "Data/Logo.png", RectangleF.FromLTRB(50, 50, 150, 150))
                ppt.SaveToFile(stream2, FileFormat.Pptx2013)
                stream2.Position = 0
                # Modify the data
                oleObject.Data = stream2
```

---

# spire.presentation python vba
## remove VBA macros from PowerPoint presentation
```python
# Create a presentation object
presentation = Presentation()

# Load PPT file from disk
presentation.LoadFromFile(inputFile)
# Remove macros
# Note, at present it only can work on macros in PPT file, has not supported for PPTM file yet.
presentation.DeleteMacros()
presentation.SaveToFile(outputFile, FileFormat.PPT)
presentation.Dispose()
```

---

