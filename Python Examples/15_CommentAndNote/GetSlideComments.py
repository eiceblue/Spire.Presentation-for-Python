from spire.presentation.common import *
from spire.presentation import *


def AppendAllText(fname: str, text: List[str]):
    fp = open(fname, "w")
    for s in text:
        fp.write(s + "\n")
    fp.close()


inputFile = "./Data/Comments.pptx"
outputFile = "GetSlideComments.txt"

# Create a PPT document
presentation = Presentation()
cs = []

# Load document from disk
presentation.LoadFromFile(inputFile)

# Loop through comments
for commentAuthor in presentation.CommentAuthors:
    for comment in commentAuthor.CommentsList:
        # Get comment information
        commentText = comment.Text
        authorName = comment.AuthorName
        time = comment.DateTime
        cs.append("Comment text : " + commentText + "\n" + "Comment author : " +
                  authorName + "\n" + "Posted on time : " + time.ToString())

AppendAllText(outputFile, cs)
presentation.Dispose()
