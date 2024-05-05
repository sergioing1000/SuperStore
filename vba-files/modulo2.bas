Attribute VB_Name = "modulo2"

Sub ReadFile(inputString As String, ByRef outputArray() As String, ByRef outputVariable As Long)
    
    
    FilePath = ThisWorkbook.Path & inputString
    
    Dim i As Long
    Dim TextLines() As String
    Dim arrayContent As String


    Open FilePath For Input As #1
    Text = Input(LOF(1), #1)
    Close #1

    TextLines = Split(Text, vbCrLf)

    outputArray = TextLines

    For i = LBound(outputArray) To UBound(outputArray)
        arrayContent = arrayContent & outputArray(i) & vbCrLf
    Next i

    outputVariable = i
    
End Sub