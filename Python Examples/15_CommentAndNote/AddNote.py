from spire.presentation.common import *
from spire.presentation import *

inputFile = "./Data/AddNote.pptx"
outputFile = "AddNote.pptx"

# Create a PPT document and load file
ppt = Presentation()
ppt.LoadFromFile(inputFile)

slide = ppt.Slides[0]

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

# Save the file
ppt.SaveToFile(outputFile, FileFormat.Pptx2010)
ppt.Dispose()
