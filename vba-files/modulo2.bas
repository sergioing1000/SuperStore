Attribute VB_Name = "modulo2"

Sub ImportTextFile()
    Dim FilePath As String
    Dim Text As String
    Dim TextLines() As String
    Dim i As Integer
    Dim MyRange as Range
    
    FilePath = ThisWorkbook.Path & "\data\headers.dat"
    
    Open FilePath For Input As #1
    Text = Input(LOF(1), #1)
    Close #1

    TextLines = Split(Text, vbCrLf)

    Set MyRange = Range("A1").Resize(, CountLinesInFile(FilePath))

    i = 0
    for each cell in MyRange
            cell.Value = TextLines(i)
            i = i + 1
    next cell
    
End Sub

Function CountLinesInFile(filePath As String) As Long
    Dim fileNum As Integer
    Dim lineText As String
    Dim lineCount As Long
    
    fileNum = FreeFile
    Open filePath For Input As fileNum
    
    lineCount = 0
    
    Do While Not EOF(fileNum)
        Line Input #fileNum, lineText
        lineCount = lineCount + 1
    Loop
    
    Close fileNum
    
    CountLinesInFile = lineCount
End Function