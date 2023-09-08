from spire.presentation.common import *
from spire.presentation import *

inputFile = "./Data/HeaderAndFooter.pptx"
outputFile = "HeaderAndFooter.pptx"

# Create a PPT document
presentation = Presentation()

presentation.LoadFromFile(inputFile)

# Add footer
presentation.SetFooterText("Demo of Spire.Presentation")

# Set the footer visible
presentation.FooterVisible = True

# Set the page number visible
presentation.SlideNumberVisible = True

# Set the date visible
presentation.DateTimeVisible = True

# Save the document
presentation.SaveToFile(outputFile, FileFormat.Pptx2010)
presentation.Dispose()
