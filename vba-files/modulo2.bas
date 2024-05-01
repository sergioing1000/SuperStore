Attribute VB_Name = "modulo2"

Sub ImportTextFile()
    Dim FilePath As String
    Dim Text As String
    Dim TextLines() As String
    Dim i As Integer
    Dim MyRange as Range
    
    ' Specify the file path
    ' FilePath = "C:\Path\To\Your\File\category.dat"
    FilePath = ThisWorkbook.Path & "\data\category.dat"

    
    ' Read the entire file into a string variable
    Open FilePath For Input As #1
    Text = Input(LOF(1), #1)
    Close #1

    ' Split the text into an array of lines
    TextLines = Split(Text, vbCrLf)

    Range("B3").Value = TextLines(0)
    Range("C3").Value = TextLines(1)
    Range("D3").Value = TextLines(2)

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
    
    ' Open the file
    fileNum = FreeFile
    Open filePath For Input As fileNum
    
    ' Initialize line count
    lineCount = 0
    
    ' Loop through each line in the file
    Do While Not EOF(fileNum)
        Line Input #fileNum, lineText
        lineCount = lineCount + 1
    Loop
    
    ' Close the file
    Close fileNum
    
    ' Return the line count
    CountLinesInFile = lineCount
End Function