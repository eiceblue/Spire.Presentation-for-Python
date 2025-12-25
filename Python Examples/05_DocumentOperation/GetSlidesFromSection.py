from spire.presentation.common import *
from spire.presentation import *

ppt = Presentation() 
ppt.LoadFromFile("AddSection.pptx") 
section=ppt.SectionList[0]  

# Get slide list
slides = section.GetSlides()

# Initialize an empty list for storing strings
sb = []

# Traverse the slide list
for i, slide in enumerate(slides):
    # Add slide ID information to the list
    sb.append(f"SlideID:{slide.SlideID}")  
 # Define output file path
outputFile = "output.txt"
   
# Merge the string elements in the list into a large string, separated by a line break
content = '\n'.join(sb)

# Open the file in append mode and write its contents
with open(outputFile, 'a', encoding='utf-8') as file:
    file.write(content)
ppt.Dispose()    
