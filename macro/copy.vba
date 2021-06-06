Sub CopyContent()

'User input
fileToOpen = InputBox("Enter file path here [C:\example.docx]: ")

startPage = InputBox("Enter start page here: ")

endPage = InputBox("Enter end page here: ")

chapterName = InputBox("Enter new chapter name here: ")

directory = InputBox("Enter the directory to save file here [c:\example]: ")

extention = InputBox("Enter file extention here [doc/docx]: ")

Documents.Open (fileToOpen)

'Select & Copy
Selection.GoTo What:=wdGoToPage, Which:=wdGoToAbsolute, Count:=startPage
Set rgePages = Selection.Range
Selection.GoTo What:=wdGoToPage, Which:=wdGoToAbsolute, Count:=endPage
rgePages.End = Selection.Bookmarks("\Page").Range.End
rgePages.Select
rgePages.Copy

'Add new document & paste contents
Set chapter = Documents.Add
ActiveDocument.Range.PasteSpecial

chapter.SaveAs directory & Application.PathSeparator & chapterName & "." & extention
chapter.Close SaveChanges:=False

'Reset previous selection
CutCopyMode = False

MsgBox "You new document is saved here " & directory & Application.PathSeparator & chapterName & "." & extention

ActiveDocument.Close SaveChanges:=False

Application.Quit

End Sub
