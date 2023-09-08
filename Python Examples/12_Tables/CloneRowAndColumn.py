from spire.presentation.common import *
import math
from spire.presentation import *

outputFile = "CloneRowAndColumn.pptx"


presentation = Presentation()
# Access first slide
sld = presentation.Slides[0]

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

presentation.SaveToFile(outputFile, FileFormat.Pptx2010)
presentation.Dispose()

  

    
