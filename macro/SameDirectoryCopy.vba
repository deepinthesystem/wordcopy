Sub CopyContent()

'User input

startPage = InputBox("Enter start page here: ")

endPage = InputBox("Enter end page here: ")

chapterName = InputBox("Enter new chapter name here: ")

extention = InputBox("Enter file extention here [doc/docx]: ")

directory = ActiveDocument.Path

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

End Sub