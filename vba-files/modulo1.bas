Attribute VB_Name = "modulo1"

Sub Macro00_NewFile()

    'This macro creates a new empty Excel file
    'Change the first sheet name as "Sales SuperStore"
    
    Workbooks.Add
    ActiveSheet.Name = "Sales SuperStore"
    
End Sub

Sub Macro01_Headers()

    'This macro Inlcude headers in the "Sales SuperStore" sheet

    Dim i as Integer
    Dim path As String
    Dim ArrayFile() As String
    Dim numberOfLines As Long

    Dim headersRange as Range

    path = "\data\headers.dat"

    modulo2.ReadFile path, ArrayFile, numberOfLines

    i = CInt(numberOfLines)

    Set headersRange = Range("A1").Resize( , numberOfLines)

    i = 0
    for each cell in headersRange
        cell.Value = ArrayFile(i)
        i = i + 1
    next cell

End Sub

Sub Macro02_OrdersWithRANDQty()
    
    
    'This macro fills "Sales SuperStore" sheet with randomic PO's
    'Using the Variable Deep to define how many rows (Products) will be in the table.
    
    'It choose the name and last name of the customer from the list in a randomic way from
    'the Worksheets("Names") and the Worksheets("Last Names").
    
    'Each order will have the same customer name but could have different number of products.
    'Max number of products per order is defined by the variable ProPerOrder
        

    Randomize ' Initialize random-number generator.
    
    
    Application.ScreenUpdating = False
    
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim l As Long
    Dim m As Long
    
    Const Deep As Long = 180           '/// Define the qty of registers
    Const ProPerOrder As Integer = 20      '/// Define the qty of Products per Order
        
    Dim Name As Long
    Dim Last_Name As Long
    
    
    Dim PO As Long
    Dim store As Long
    
    Dim OrdersQty As Long
    
    Dim Cell As Range
    Dim Myrange As Range
    Dim MyRange2 As Range
    Dim MyRange3 As Range
    
    Dim test As String
    
    Dim PODate As Date
    
    
    
    Sheets("Sales SuperStore").Select
    
    
    Set Myrange = Range("D2:D" & Deep)
    Set MyRange2 = Range("P2:P" & Deep)
    Set MyRange3 = Range("U2:U" & Deep)
    
    
    
     
    '//////////////////////////////////////////RANGE1//////////////////////////////////////////
        
    PO = 31996
    l = 0
    m = 2
    
    PODate = 43037
    
    For Each Cell In Myrange
    
        
        If l = 0 Then
        
            k = Int(ProPerOrder * Rnd)  '// Defines the number of products per order
            
            store = Int(66 * Rnd) + 1
                        
            
            Name = Int(1013 * Rnd) + 1
            Last_Name = Int(2012 * Rnd) + 1
            
            j = 0
            l = 1
            
            
            
            If ((100 * Rnd) > 92) Then   '// Defines the date changing
            
                PODate = PODate + 1
            
            End If
            
            

        End If
        
        If j < k Then
        
        
                If PO < 10 Then
                    Cell.Value = "PO-00000" & PO
                                                            
                ElseIf PO < 100 Then
                    Cell.Value = "PO-0000" & PO
                                        
                ElseIf PO < 1000 Then
                    Cell.Value = "PO-000" & PO
                                    
                ElseIf PO < 10000 Then
                    Cell.Value = "PO-00" & PO
                                        
                ElseIf PO < 100000 Then
                    Cell.Value = "PO-0" & PO
                    
                ElseIf PO < 1000000 Then
                    Cell.Value = "PO-" & PO
                End If
                
                Cells(m, "O").Value = Worksheets("Cities-Stores").Cells(store + 1, "G") '//Postal Code
                Cells(m, "N").Value = Worksheets("Cities-Stores").Cells(store + 1, "B") '//City
                Cells(m, "M").Value = Worksheets("Cities-Stores").Cells(store + 1, "A") '//City order
                Cells(m, "L").Value = Worksheets("Cities-Stores").Cells(store + 1, "C") '//State
                Cells(m, "J").Value = Worksheets("Cities-Stores").Cells(store + 1, "H") '//Country
                Cells(m, "V").Value = Worksheets("Cities-Stores").Cells(store + 1, "E") '//Latitude
                Cells(m, "W").Value = Worksheets("Cities-Stores").Cells(store + 1, "F") '//Longitude
                Cells(m, "K").Value = Worksheets("Cities-Stores").Cells(store + 1, "I") '//Region
                Cells(m, "Y").Value = Worksheets("Cities-Stores").Cells(store + 1, "J") '//Region
                Cells(m, "X").Value = k + 1                                             '//Number of records
                
                
                Cells(m, "B").Value = Worksheets("Names").Cells(Name + 1, "B") & " " & Worksheets("Last Names").Cells(Last_Name + 1, "B") '//Name & Last_Name
                Cells(m, "C").Value = PODate                                            '//Date
                
                                                                
                PO = PO - 1
                
        ElseIf j = k Then
                              
                If PO < 10 Then
                    Cell.Value = "PO-00000" & PO
                                                            
                ElseIf PO < 100 Then
                    Cell.Value = "PO-0000" & PO
                                        
                ElseIf PO < 1000 Then
                    Cell.Value = "PO-000" & PO
                                    
                ElseIf PO < 10000 Then
                    Cell.Value = "PO-00" & PO
                                        
                ElseIf PO < 100000 Then
                    Cell.Value = "PO-0" & PO
                    
                ElseIf PO < 1000000 Then
                    Cell.Value = "PO-" & PO
                End If
                
                Cells(m, "O").Value = Worksheets("Cities-Stores").Cells(store + 1, "G") '//Postal Code
                Cells(m, "N").Value = Worksheets("Cities-Stores").Cells(store + 1, "B") '//City
                Cells(m, "M").Value = Worksheets("Cities-Stores").Cells(store + 1, "A") '//City order
                Cells(m, "L").Value = Worksheets("Cities-Stores").Cells(store + 1, "C") '//State
                Cells(m, "J").Value = Worksheets("Cities-Stores").Cells(store + 1, "H") '//Country
                Cells(m, "V").Value = Worksheets("Cities-Stores").Cells(store + 1, "E") '//Latitude
                Cells(m, "W").Value = Worksheets("Cities-Stores").Cells(store + 1, "F") '//Logitude
                Cells(m, "K").Value = Worksheets("Cities-Stores").Cells(store + 1, "I") '//Region
                Cells(m, "Y").Value = Worksheets("Cities-Stores").Cells(store + 1, "J") '//Region
                Cells(m, "X").Value = k + 1                                             '//Number of records
                
                
                Cells(m, "B").Value = Worksheets("Names").Cells(Name + 1, "B") & " " & Worksheets("Last Names").Cells(Last_Name + 1, "B") '//Name & Last_Name
                Cells(m, "C").Value = PODate
                                
                
                j = 0
                l = 0
                PO = PO + 1
        Else
        
            j = 0
            l = 0
            PO = PO + 1
            
               If PO < 10 Then
                    Cell.Value = "PO-00000" & PO
                                                            
                ElseIf PO < 100 Then
                    Cell.Value = "PO-0000" & PO
                                        
                ElseIf PO < 1000 Then
                    Cell.Value = "PO-000" & PO
                                    
                ElseIf PO < 10000 Then
                    Cell.Value = "PO-00" & PO
                                        
                ElseIf PO < 100000 Then
                    Cell.Value = "PO-0" & PO
                    
                ElseIf PO < 1000000 Then
                    Cell.Value = "PO-" & PO
                End If
                           
        End If
        
        
        If l = 1 Then
             PO = PO + 1
             j = j + 1
        End If
        
        m = m + 1
                
    Next Cell

    '//////////////////////////////////////////RANGE2//////////////////////////////////////////
    
    j = 1
    
               
    For Each Cell In MyRange2
    
        i = Int(1015 * Rnd) + 1
        
        
        If i > 0 And i <= 14 Then
                
            Cell.Value = "Accesories"
            Cells(j + 1, "A").Value = "Beauty"                                      '//Category
            Cells(j + 1, "E").Value = Worksheets("Accesories").Cells(i + 1, 2)      '//Product Name
            Cells(j + 1, "F").Value = Worksheets("Accesories").Cells(i + 1, 3)      '//Unit Price
            Cells(j + 1, "T").Value = Int((20 * Rnd) + 1)                           '//RAND Qty of product  Max     min
                                    
        ElseIf i > 14 And i <= 148 Then
            Cell.Value = "Appliances"
            Cells(j + 1, "A").Value = "Technology"                                  '//Category
            Cells(j + 1, "E").Value = Worksheets("Appliances").Cells(i - 13, 2)      '//Product Name
            Cells(j + 1, "F").Value = Worksheets("Appliances").Cells(i - 13, 3)     '//Unit Price
            Cells(j + 1, "T").Value = Int((5 * Rnd) + 1)                           '//RAND Qty of product  Max     min
                                                
        ElseIf i > 148 And i <= 287 Then
            Cell.Value = "Art"
            Cells(j + 1, "A").Value = "Office Supplies"                             '//Category
            Cells(j + 1, "E").Value = Worksheets("Art").Cells(i - 147, 2)           '//Product Name
            Cells(j + 1, "F").Value = Worksheets("Art").Cells(i - 147, 3)           '//Unit Price
            Cells(j + 1, "T").Value = Int((15 * Rnd) + 1)                           '//RAND Qty of product  Max     min
            
        ElseIf i > 287 And i <= 298 Then
            Cell.Value = "Binders"
            Cells(j + 1, "A").Value = "Office Supplies"                             '//Category
            Cells(j + 1, "E").Value = Worksheets("Binders").Cells(i - 286, 2)       '//Product Name
            Cells(j + 1, "F").Value = Worksheets("Binders").Cells(i - 286, 3)       '//Unit Price
            Cells(j + 1, "T").Value = Int((35 * Rnd) + 1)                           '//RAND Qty of product  Max     min
            
        ElseIf i > 298 And i <= 329 Then
            Cell.Value = "Bookcases"
            Cells(j + 1, "A").Value = "Furniture"                                   '//Category
            Cells(j + 1, "E").Value = Worksheets("Bookcases").Cells(i - 297, 2)     '//Product Name
            Cells(j + 1, "F").Value = Worksheets("Bookcases").Cells(i - 297, 3)      '//Unit Price
            Cells(j + 1, "T").Value = Int((4 * Rnd) + 1)                           '//RAND Qty of product  Max     min
            
        ElseIf i > 329 And i <= 487 Then
            Cell.Value = "Chairs"
            Cells(j + 1, "A").Value = "Furniture"                                   '//Category
            Cells(j + 1, "E").Value = Worksheets("Chairs").Cells(i - 328, 2)        '//Product Name
            Cells(j + 1, "F").Value = Worksheets("Chairs").Cells(i - 328, 3)        '//Unit Price
            Cells(j + 1, "T").Value = Int((10 * Rnd) + 1)                           '//RAND Qty of product  Max     min
            
        ElseIf i > 487 And i <= 607 Then
            Cell.Value = "Copiers"
            Cells(j + 1, "A").Value = "Technology"                                  '//Category
            Cells(j + 1, "E").Value = Worksheets("Copiers").Cells(i - 486, 2)       '//Product Name
            Cells(j + 1, "F").Value = Worksheets("Copiers").Cells(i - 486, 3)        '//Unit Price
            Cells(j + 1, "T").Value = Int((4 * Rnd) + 1)                           '//RAND Qty of product  Max     min
                   
        ElseIf i > 607 And i <= 637 Then
            Cell.Value = "Envelopes"
            Cells(j + 1, "A").Value = "Office Supplies"                             '//Category
            Cells(j + 1, "E").Value = Worksheets("Envelopes").Cells(i - 606, 2)     '//Product Name
            Cells(j + 1, "F").Value = Worksheets("Envelopes").Cells(i - 606, 3)     '//Unit Price
            Cells(j + 1, "T").Value = Int((250 * Rnd) + 1)                           '//RAND Qty of product  Max     min
        
        ElseIf i > 637 And i <= 719 Then
            Cell.Value = "Fasteners"
            Cells(j + 1, "A").Value = "Technology"                                  '//Category
            Cells(j + 1, "E").Value = Worksheets("Fasteners").Cells(i - 636, 2)     '//Product Name
            Cells(j + 1, "F").Value = Worksheets("Fasteners").Cells(i - 636, 3)     '//Unit Price
            Cells(j + 1, "T").Value = Int((800 * Rnd) + 1)                           '//RAND Qty of product  Max     min
        
        ElseIf i > 719 And i <= 877 Then
            Cell.Value = "Furnishings"
            Cells(j + 1, "A").Value = "Furniture"                                   '//Category
            Cells(j + 1, "E").Value = Worksheets("Furnishings").Cells(i - 718, 2)   '//Product Name
            Cells(j + 1, "F").Value = Worksheets("Furnishings").Cells(i - 718, 3)   '//Unit Price
            Cells(j + 1, "T").Value = Int((10 * Rnd) + 1)                           '//RAND Qty of product  Max     min
        
        ElseIf i > 877 And i <= 901 Then
            Cell.Value = "Labels"
            Cells(j + 1, "A").Value = "Office Supplies"                             '//Category
            Cells(j + 1, "E").Value = Worksheets("Labels").Cells(i - 876, 2)        '//Product Name
            Cells(j + 1, "F").Value = Worksheets("Labels").Cells(i - 876, 3)        '//Unit Price
            Cells(j + 1, "T").Value = Int((200 * Rnd) + 1)                           '//RAND Qty of product  Max     min
            
        ElseIf i > 901 And i <= 918 Then
            Cell.Value = "Gym Machines"
            Cells(j + 1, "A").Value = "Beauty"                                      '//Category
            Cells(j + 1, "E").Value = Worksheets("Gym Machines").Cells(i - 900, 2)  '//Product Name
            Cells(j + 1, "F").Value = Worksheets("Gym Machines").Cells(i - 900, 3)  '//Unit Price
            Cells(j + 1, "T").Value = Int((4 * Rnd) + 1)                           '//RAND Qty of product  Max     min
            
        ElseIf i > 918 And i <= 948 Then
            Cell.Value = "Papers"
            Cells(j + 1, "A").Value = "Office Supplies"                             '//Category
            Cells(j + 1, "E").Value = Worksheets("Papers").Cells(i - 917, 2)        '//Product Name
            Cells(j + 1, "F").Value = Worksheets("Papers").Cells(i - 917, 3)        '//Unit Price
            Cells(j + 1, "T").Value = Int((220 * Rnd) + 1)                           '//RAND Qty of product  Max     min
            
        ElseIf i > 948 And i <= 965 Then
            Cell.Value = "Storage"
            Cells(j + 1, "A").Value = "Furniture"                                   '//Category
            Cells(j + 1, "E").Value = Worksheets("Storage").Cells(i - 947, 2)       '//Product Name
            Cells(j + 1, "F").Value = Worksheets("Storage").Cells(i - 947, 3)      '//Unit Price
            Cells(j + 1, "T").Value = Int((35 * Rnd) + 1)                           '//RAND Qty of product  Max     min
            
        ElseIf i > 965 And i <= 968 Then
            Cell.Value = "Supplies"
            Cells(j + 1, "A").Value = "Office Supplies"                             '//Category
            Cells(j + 1, "E").Value = Worksheets("Supplies").Cells(i - 964, 2)      '//Product Name
            Cells(j + 1, "F").Value = Worksheets("Supplies").Cells(i - 964, 3)      '//Unit Price
            Cells(j + 1, "T").Value = Int((25 * Rnd) + 1)                           '//RAND Qty of product  Max     min
            
        ElseIf i > 968 And i <= 1016 Then
            Cell.Value = "Tables"
            Cells(j + 1, "A").Value = "Furniture"                                   '//Category
            Cells(j + 1, "E").Value = Worksheets("Tables").Cells(i - 967, 2)        '//Product Name
            Cells(j + 1, "F").Value = Worksheets("Tables").Cells(i - 967, 3)        '//Unit Price
            Cells(j + 1, "T").Value = Int((6 * Rnd) + 1)                           '//RAND Qty of product  Max     min
        
        End If
        
        j = j + 1

    Next Cell
    
    
    '//////////////////////////////////////////RANGE3////////////////////////////////////////// Total Price RANGE
        
        j = 2
        
        Dim ws As Worksheet
        Dim unit_price As Double
        Dim qty As Double
       
'        For Each Cell In MyRange3           '// Inserts the formula to calculate Total Price
'                Cell.Formula = "=F" & j & "*" & "T" & j
'                j = j + 1
'        Next Cell

        

        For Each Cell In MyRange3           '// Inserts the value calculated for Total Price
        
                unit_price = Cells(j, 6).Value
                qty = Cells(j, 20).Value
                Cell.Value = unit_price * qty
                j = j + 1
        Next Cell

        
    '//////////////////////////////////////////RANGE4////////////////////////////////////////// END OF FILE RANGE
        
    Set Myrange = Range("A" & Deep + 1 & ":Y" & Deep + 1)
        
    For Each Cell In Myrange
                
        Cell.Value = "END OF FILE"
        
    Next Cell
    
    
    Range("A2").Select
    ActiveWindow.FreezePanes = True
    
    ActiveWindow.Zoom = 80
    
    
    '///////////////////////////////////////////////////////////////////////////////////////////
    
    
    Columns("F:F").NumberFormat = "#,##0.00"
    Columns("U:U").NumberFormat = "#,##0.00"
    Columns("C:C").NumberFormat = "[$-409]d/mmm/yy;@"
    
    Columns("T:T").HorizontalAlignment = xlCenter
    
    
    Columns("A:B").EntireColumn.AutoFit
    
    Columns("D:F").EntireColumn.AutoFit
    
    Columns("J:J").EntireColumn.AutoFit
    Columns("L:M").EntireColumn.AutoFit
    
    
    Columns("N:P").EntireColumn.AutoFit
    
    Columns("U:X").EntireColumn.AutoFit
    
    
    Application.ScreenUpdating = True

End Sub

Sub Macro03_InsertCitiesSheet()

    Dim ws As Worksheet
    Dim i As Integer
    Dim Arreglo(0, 670) As String
        
    Dim Cell As Range
    Dim Myrange As Range

    Dim path As String
    Dim ArrayFile() As String
    Dim numberOfLines As Long

 
    Application.DisplayAlerts = False
    
    For Each ws In ActiveWorkbook.Worksheets
        
        If ws.Name = "Cities-Stores" Then
            ws.Delete
        End If
        
    Next ws
    
    Application.DisplayAlerts = True
            
    Sheets.Add After:=ActiveSheet
    ActiveSheet.Name = "Cities-Stores"
    
  
    path = "\data\cities.dat"

    modulo2.ReadFile path, ArrayFile, numberOfLines

    numberOfLinesAsString = CStr(numberOfLines/10)

    Set Myrange = Range("A1:J"& numberOfLinesAsString)
    i = 0
        
    For Each Cell In Myrange
        
        Cell.Value = ArrayFile(i)
        
        i = i + 1
            
    Next Cell
    
    Columns("A:J").EntireColumn.AutoFit
    
    
    Columns("A:A").HorizontalAlignment = xlCenter
    Columns("H:H").HorizontalAlignment = xlCenter
    Columns("G:G").HorizontalAlignment = xlRight
    Range("G1").HorizontalAlignment = xlCenter
    
    Range("A2").Select
    ActiveWindow.FreezePanes = True
    
    Set Myrange = Range("A1:J1")
    Myrange.Font.Bold = True
        
End Sub

Sub Macro04_InsertSheets()

    Dim path As String
    Dim ArrayFile() As String
    Dim numberOfLines As Long
    Dim i as Integer

    path = "\data\sheets.dat"

    modulo2.ReadFile path, ArrayFile, numberOfLines

    for i = 0 to numberOfLines-1

        Sheets.Add After:=ActiveSheet
        ActiveSheet.Name = ArrayFile (i)

    next i

End Sub

Sub Macro05_Fill_Categories()

    Sheets("Categories").Select

    Dim path As String
    Dim ArrayFile() As String
    Dim numberOfLines As Long
    Dim Myrange as Range
    Dim i as Integer

    path = "\data\categories.dat"

    modulo2.ReadFile path, ArrayFile, numberOfLines

    Set Myrange = Range("A1:B5")

    i=0

    for each cell in Myrange
        cell.value = ArrayFile (i)
        i=i+1
    next

    Range("A1:B1").Font.Bold = True
    Columns("A:B").EntireColumn.AutoFit
    Columns("A:A").HorizontalAlignment = xlCenter
    
    Range("A2").Select
    ActiveWindow.FreezePanes = True
        
End Sub
Sub Macro06_Fill_Sub_Category()

    Dim path As String
    Dim i As Integer
    Dim ArrayFile() As String
    Dim numberOfLines As Long
        
    Dim Myrange As Range
    
    Sheets("Sub-Categories").Select

    path = "\data\sub_categories.dat"
    modulo2.ReadFile path, ArrayFile, numberOfLines
    
    Set Myrange = Range("A1:B17")

    i=0

    for each cell in Myrange
        cell.value = ArrayFile (i)
        i=i+1
    next
    
    Range("A1:B1").Font.Bold = True

    Columns("A:B").EntireColumn.AutoFit
    Columns("A:A").HorizontalAlignment = xlCenter
    
    Range("A2").Select
    ActiveWindow.FreezePanes = True
        
End Sub

Sub Macro07_Fill_Accesories()
    
    Sheets("Accesories").Select

    Dim Myrange As Range
    Dim i As Long

    Dim path As String
    Dim ArrayFile() As String
    Dim numberOfLines As Long

    path = "\data\accesories.dat"
    modulo2.ReadFile path, ArrayFile, numberOfLines

    Randomize ' Initialize random-number generator.
  
    Set Myrange = Range("A1:B15")

    i=0

    for each cell in Myrange
        cell.value = ArrayFile (i)
        i=i+1
    next
    
    '/////////////////PRICES/////////////////

    With Range("C1")
        .Value = "Unit Price USD"
        .Font.Bold = True
        .WrapText = True
    End With

    Set Myrange = Range("C2:C15")
    
    For Each Cell In Myrange

        Cell.Value = Round((32 * Rnd) + 6.5, 2)       '////max     min
        
    Next Cell
    
   
    Columns("C:C").NumberFormat = "#,##0.00"
    
    '/////////////////PRICES/////////////////
    
    Range("A1:C1").Font.Bold = True

    Columns("A:B").EntireColumn.AutoFit
    Columns("A:A").HorizontalAlignment = xlCenter

    Columns("C:C").NumberFormat = "#,##0.00"

    Range("A2").Select
    ActiveWindow.FreezePanes = True

End Sub


Sub Macro08_Fill_Appliances()

    Sheets("Appliances").Select

    Dim Myrange As Range
    Dim i As Double
    ' Dim j As Double


    Dim path As String
    Dim ArrayFile() As String
    Dim numberOfLines As Long

    path = "\data\appliances.dat"
    modulo2.ReadFile path, ArrayFile, numberOfLines

    Randomize ' Initialize random-number generator.
    
    Set Myrange = Range("A1:B135")

    i=0

    for each cell in Myrange
        cell.value = ArrayFile (i)
        i=i+1
    next
    
    '/////////////////PRICES/////////////////
    
    Set Myrange = Range("C2:C135")
    
    For Each Cell In Myrange

        Cell.Value = Round((285 * Rnd) + 43, 2)         '////max     min
        
    Next Cell
    
    With Range("C1")
        .Value = "Unit Price USD"
        .Font.Bold = True
        .WrapText = True
    End With
    
    Columns("C:C").NumberFormat = "#,##0.00"
    
    '/////////////////PRICES/////////////////
    
    Range("A1:B1").Font.Bold = True

    Columns("A:B").EntireColumn.AutoFit
    Columns("A:A").HorizontalAlignment = xlCenter

    Range("A2").Select
    ActiveWindow.FreezePanes = True

End Sub
Sub Macro09_Fill_Binders()

    Sheets("Binders").Select

    Dim Myrange As Range
    Dim i As Double

    Dim path As String
    Dim ArrayFile() As String
    Dim numberOfLines As Long

    path = "\data\binders.dat"
    modulo2.ReadFile path, ArrayFile, numberOfLines
       
    Set Myrange = Range("A1:B12")

    i=0

    for each cell in Myrange
        cell.value = ArrayFile (i)
        i=i+1
    next
    
    '/////////////////PRICES/////////////////
    
    Set Myrange = Range("C2:C12")
    
    For Each Cell In Myrange

        Cell.Value = Round((23 * Rnd) + 0.1, 2)         '////max     min
        
    Next Cell
    
    With Range("C1")
        .Value = "Unit Price USD"
        .Font.Bold = True
        .WrapText = True
    End With
    
    Columns("C:C").NumberFormat = "#,##0.00"
    
    '/////////////////PRICES/////////////////

    Range("A1:C1").Font.Bold = True

    Columns("A:B").EntireColumn.AutoFit
    Columns("A:A").HorizontalAlignment = xlCenter

    Range("A2").Select
    ActiveWindow.FreezePanes = True

End Sub

Sub Macro10_Fill_Art()

    Sheets("Art").Select

    Dim Myrange As Range
    Dim Cell As Range
    Dim i As Double

    Set Myrange = Range("A1:B140")

    Dim path As String
    Dim ArrayFile() As String
    Dim numberOfLines As Long

    path = "\data\art.dat"
    modulo2.ReadFile path, ArrayFile, numberOfLines

    i=0

    for each cell in Myrange
        cell.value = ArrayFile (i)
        i=i+1
    next

    Randomize ' Initialize random-number generator.

    '/////////////////PRICES/////////////////
    
    Set Myrange = Range("C2:C140")
    
    For Each Cell In Myrange

        Cell.Value = Round((130 * Rnd) + 32, 2)         '////max     min
        
    Next Cell
    
    With Range("C1")
        .Value = "Unit Price USD"
        .Font.Bold = True
        .WrapText = True
    End With
    
    Columns("C:C").NumberFormat = "#,##0.00"
    
    '/////////////////PRICES/////////////////
    
    Range("A1:C1").Font.Bold = True

    Columns("A:B").EntireColumn.AutoFit
    Columns("A:A").HorizontalAlignment = xlCenter

    Range("A2").Select
    ActiveWindow.FreezePanes = True

End Sub
Sub Macro11_Fill_Bookcases()

    Sheets("Bookcases").Select

    Dim Myrange As Range
    Dim i As Double

    Dim path As String
    Dim ArrayFile() As String
    Dim numberOfLines As Long

    path = "\data\bookcases.dat"
    modulo2.ReadFile path, ArrayFile, numberOfLines

    Set Myrange = Range("A1:B32")

    i=0

    for each cell in Myrange
        cell.value = ArrayFile (i)
        i=i+1
    next

    '/////////////////PRICES/////////////////

    
    Set Myrange = Range("C2:C32")

    Randomize ' Initialize random-number generator.
    
    For Each Cell In Myrange

        Cell.Value = Round((730 * Rnd) + 150, 2)         '////max     min
        
    Next Cell
    
    With Range("C1")
        .Value = "Unit Price USD"
        .Font.Bold = True
        .WrapText = True
    End With
    
    Columns("C:C").NumberFormat = "#,##0.00"
    
    '/////////////////PRICES/////////////////

    Range("A1:C1").Font.Bold = True

    Columns("A:B").EntireColumn.AutoFit
    Columns("A:A").HorizontalAlignment = xlCenter

    Range("A2").Select
    ActiveWindow.FreezePanes = True

End Sub

Sub Macro12_Fill_Chairs()

    Sheets("Chairs").Select

    Dim Myrange As Range
    Dim Cell As Range
    Dim i As Double
    ' Dim j As Double


    Dim path As String
    Dim ArrayFile() As String
    Dim numberOfLines As Long

    path = "\data\chairs.dat"
    modulo2.ReadFile path, ArrayFile, numberOfLines

    Set Myrange = Range("A1:B159")

    i=0

    for each cell in Myrange
        cell.value = ArrayFile (i)
        i=i+1
    next

     '/////////////////PRICES/////////////////
    
    Set Myrange = Range("C2:C159")
    
    For Each Cell In Myrange

        Cell.Value = Round((435 * Rnd) + 35, 2)         '////max     min
        
    Next Cell
    
    With Range("C1")
        .Value = "Unit Price USD"
        .Font.Bold = True
        .WrapText = True
    End With
    
    Columns("C:C").NumberFormat = "#,##0.00"
    
    '/////////////////PRICES/////////////////
    
    Range("A1:C1").Font.Bold = True

    Columns("A:B").EntireColumn.AutoFit
    Columns("A:A").HorizontalAlignment = xlCenter

    Range("A2").Select
    ActiveWindow.FreezePanes = True


End Sub

Sub Macro13_Fill_Copiers()

    Sheets("Copiers").Select

    Dim Myrange As Range
    Dim i As Double

    Dim path As String
    Dim ArrayFile() As String
    Dim numberOfLines As Long

    path = "\data\copiers.dat"
    modulo2.ReadFile path, ArrayFile, numberOfLines

    Set Myrange = Range("A1:B121")

    i=0

    for each cell in Myrange
        cell.value = ArrayFile (i)
        i=i+1
    next
    
    '/////////////////PRICES/////////////////
    
    Set Myrange = Range("C2:C121")
    
    For Each Cell In Myrange

        Cell.Value = Round((2130 * Rnd) + 145, 2)         '////max     min
        
    Next Cell
    
    With Range("C1")
        .Value = "Unit Price USD"
        .Font.Bold = True
        .WrapText = True
    End With
    
    Columns("C:C").NumberFormat = "#,##0.00"
    
    '/////////////////PRICES/////////////////

    Range("A1:C1").Font.Bold = True

    Columns("A:B").EntireColumn.AutoFit
    Columns("A:A").HorizontalAlignment = xlCenter

    Range("A2").Select
    ActiveWindow.FreezePanes = True

End Sub
Sub Macro14_Fill_Envelopes()


    Sheets("Envelopes").Select

    Dim Myrange As Range
    Dim i As Double

    Dim path As String
    Dim ArrayFile() As String
    Dim numberOfLines As Long

    path = "\data\envelopes.dat"
    modulo2.ReadFile path, ArrayFile, numberOfLines

    Set Myrange = Range("A1:B31")

    i=0

    for each cell in Myrange
        cell.value = ArrayFile (i)
        i=i+1
    next
    
    '/////////////////PRICES/////////////////
    
    Set Myrange = Range("C2:C31")
    
    For Each Cell In Myrange

        Cell.Value = Round((31 * Rnd) + 0.2, 2)          '////max     min
        
    Next Cell
    
    With Range("C1")
        .Value = "Unit Price USD"
        .Font.Bold = True
        .WrapText = True
    End With
    
    Columns("C:C").NumberFormat = "#,##0.00"
    
    '/////////////////PRICES/////////////////
    
    Range("A1:C1").Font.Bold = True

    Columns("A:B").EntireColumn.AutoFit
    Columns("A:A").HorizontalAlignment = xlCenter

    Range("A2").Select
    ActiveWindow.FreezePanes = True

End Sub
Sub Macro15_Fill_Fasteners()

    Sheets("Fasteners").Select

    Dim Myrange As Range
    Dim i As Double

    Dim path As String
    Dim ArrayFile() As String
    Dim numberOfLines As Long

    path = "\data\fasteners.dat"
    modulo2.ReadFile path, ArrayFile, numberOfLines

    Set Myrange = Range("A1:B83")

    i=0

    for each cell in Myrange
        cell.value = ArrayFile (i)
        i=i+1
    next
    
    '/////////////////PRICES/////////////////
    
    Set Myrange = Range("C2:C83")
    
    For Each Cell In Myrange

        Cell.Value = Round((10 * Rnd) + 0.1, 2)          '////max     min
        
    Next Cell
    
    With Range("C1")
        .Value = "Unit Price USD"
        .Font.Bold = True
        .WrapText = True
    End With
    
    Columns("C:C").NumberFormat = "#,##0.00"
    
    '/////////////////PRICES/////////////////

    Range("A1:C1").Font.Bold = True

    Columns("A:B").EntireColumn.AutoFit
    Columns("A:A").HorizontalAlignment = xlCenter

    Range("A2").Select
    ActiveWindow.FreezePanes = True

End Sub

Sub Macro16_Fill_Furnishings()

    Sheets("Furnishings").Select

    Dim Myrange As Range
    Dim Cell As Range
    Dim i As Double

    Dim path As String
    Dim ArrayFile() As String
    Dim numberOfLines As Long

    path = "\data\furnishings.dat"
    modulo2.ReadFile path, ArrayFile, numberOfLines

    Set Myrange = Range("A1:B159")

    i=0

    for each cell in Myrange
        cell.value = ArrayFile (i)
        i=i+1
    next

    '/////////////////PRICES/////////////////
    
    Set Myrange = Range("C2:C159")
    
    For Each Cell In Myrange

        Cell.Value = Round((375 * Rnd) + 33, 2)          '////max     min
        
    Next Cell
    
    With Range("C1")
        .Value = "Unit Price USD"
        .Font.Bold = True
        .WrapText = True
    End With
    
    Columns("C:C").NumberFormat = "#,##0.00"
    
    '/////////////////PRICES/////////////////

    Range("A1:C1").Font.Bold = True

    Columns("A:B").EntireColumn.AutoFit
    Columns("A:A").HorizontalAlignment = xlCenter
    
    Range("A2").Select
    ActiveWindow.FreezePanes = True


End Sub
Sub Macro17_Fill_Labels()

    Sheets("Labels").Select

    Dim Myrange As Range
    Dim i As Double

    Dim path As String
    Dim ArrayFile() As String
    Dim numberOfLines As Long

    path = "\data\labels.dat"
    modulo2.ReadFile path, ArrayFile, numberOfLines

    Set Myrange = Range("A1:B25")

    i=0

    for each cell in Myrange
        cell.value = ArrayFile (i)
        i=i+1
    next
   
     '/////////////////PRICES/////////////////
    
    Set Myrange = Range("C2:C25")
    
    For Each Cell In Myrange

        Cell.Value = Round((28 * Rnd) + 2.2, 2)            '////max     min
        
    Next Cell
    
    With Range("C1")
        .Value = "Unit Price USD"
        .Font.Bold = True
        .WrapText = True
    End With
    
    Columns("C:C").NumberFormat = "#,##0.00"
    
    '/////////////////PRICES/////////////////
    
    Range("A1:C1").Font.Bold = True

    Columns("A:B").EntireColumn.AutoFit
    Columns("A:A").HorizontalAlignment = xlCenter

    Range("A2").Select
    ActiveWindow.FreezePanes = True


End Sub
Sub Macro18_Fill_Gym_Machines()

    Sheets("Gym Machines").Select

    Dim Myrange As Range
    Dim i As Double

    Dim path As String
    Dim ArrayFile() As String
    Dim numberOfLines As Long

    path = "\data\gymmachines.dat"
    modulo2.ReadFile path, ArrayFile, numberOfLines

    Set Myrange = Range("A1:B18")

    i=0

    for each cell in Myrange
        cell.value = ArrayFile (i)
        i=i+1
    next
    
    '/////////////////PRICES/////////////////
    
    Set Myrange = Range("C2:C18")
    
    For Each Cell In Myrange

        Cell.Value = Round((1940 * Rnd) + 330, 2)            '////max     min
        
    Next Cell
    
    With Range("C1")
        .Value = "Unit Price USD"
        .Font.Bold = True
        .WrapText = True
    End With
    
    Columns("C:C").NumberFormat = "#,##0.00"
    
    '/////////////////PRICES/////////////////

    Range("A1:C1").Font.Bold = True

    Columns("A:B").EntireColumn.AutoFit
    Columns("A:A").HorizontalAlignment = xlCenter

    Range("A2").Select
    ActiveWindow.FreezePanes = True


End Sub

Sub Macro19_Fill_Papers()


    Sheets("Papers").Select

    Dim Myrange As Range    
    Dim i As Double
    
    Dim path As String
    Dim ArrayFile() As String
    Dim numberOfLines As Long

    path = "\data\papers.dat"
    modulo2.ReadFile path, ArrayFile, numberOfLines

    Set Myrange = Range("A1:B31")

    i=0

    for each cell in Myrange
        cell.value = ArrayFile (i)
        i=i+1
    next
    
    '/////////////////PRICES/////////////////
    
    Set Myrange = Range("C2:C31")
    
    For Each Cell In Myrange

        Cell.Value = Round((28 * Rnd) + 3.15, 2)            '////max     min
        
    Next Cell
    
    With Range("C1")
        .Value = "Unit Price USD"
        .Font.Bold = True
        .WrapText = True
    End With
    
    Columns("C:C").NumberFormat = "#,##0.00"
    
    '/////////////////PRICES/////////////////

    Range("A1:C1").Font.Bold = True

    Columns("A:B").EntireColumn.AutoFit
    Columns("A:A").HorizontalAlignment = xlCenter

    Range("A2").Select
    ActiveWindow.FreezePanes = True

End Sub
Sub Macro20_Fill_Storage()


    Sheets("Storage").Select

    Dim Myrange As Range
    Dim i As Double

    Dim path As String
    Dim ArrayFile() As String
    Dim numberOfLines As Long

    path = "\data\storages.dat"
    modulo2.ReadFile path, ArrayFile, numberOfLines

    Set Myrange = Range("A1:B18")

    i=0

    for each cell in Myrange
        cell.value = ArrayFile (i)
        i=i+1
    next    
    
     '/////////////////PRICES/////////////////
    
    Set Myrange = Range("C2:C18")
    
    For Each Cell In Myrange

        Cell.Value = Round((46 * Rnd) + 18.4, 2)             '////max     min
        
    Next Cell
    
    With Range("C1")
        .Value = "Unit Price USD"
        .Font.Bold = True
        .WrapText = True
    End With
    
    Columns("C:C").NumberFormat = "#,##0.00"
    
    '/////////////////PRICES/////////////////

    Range("A1:C1").Font.Bold = True

    Columns("A:B").EntireColumn.AutoFit
    Columns("A:A").HorizontalAlignment = xlCenter

    Range("A2").Select
    ActiveWindow.FreezePanes = True


End Sub
Sub Macro21_Fill_Supplies()


    Dim Myrange As Range
    Dim Cell As Range
    
    Sheets("Supplies").Select
    
    Range("A1").Value = "Order"
    Range("A1").Font.Bold = True
    
    Range("B1").Value = "Supplies"
    Range("B1").Font.Bold = True
    
    Range("A2").Value = "1"
    Range("B2").Value = "Supplies 1"
    Range("A3").Value = "2"
    Range("B3").Value = "Supplies 2"
    Range("A4").Value = "3"
    Range("B4").Value = "Supplies 3"
    
    
     '/////////////////PRICES/////////////////
    
    Set Myrange = Range("C2:C4")
    
    For Each Cell In Myrange

        Cell.Value = Round((30 * Rnd) + 10, 2)             '////max     min
        
    Next Cell
    
    With Range("C1")
        .Value = "Unit Price USD"
        .Font.Bold = True
        .WrapText = True
    End With
    
    Columns("C:C").NumberFormat = "#,##0.00"
    
    '/////////////////PRICES/////////////////

    
    Columns("A:B").EntireColumn.AutoFit
    Columns("A:A").HorizontalAlignment = xlCenter
    
    Range("A2").Select
    ActiveWindow.FreezePanes = True
        
End Sub

Sub Macro22_Fill_Tables()


    Sheets("Tables").Select

    Dim Myrange As Range
    Dim i As Double
    
    Set Myrange = Range("A1:B48")

    Dim path As String
    Dim ArrayFile() As String
    Dim numberOfLines As Long

    path = "\data\tables.dat"
    modulo2.ReadFile path, ArrayFile, numberOfLines

    Set Myrange = Range("A1:B48")

    i=0

    for each cell in Myrange
        cell.value = ArrayFile (i)
        i=i+1
    next
    
    '/////////////////PRICES/////////////////
    
    Set Myrange = Range("C2:C48")
    
    For Each Cell In Myrange

        Cell.Value = Round((1450 * Rnd) + 95, 2)             '////max     min
        
    Next Cell
    
    With Range("C1")
        .Value = "Unit Price USD"
        .Font.Bold = True
        .WrapText = True
    End With
    
    Columns("C:C").NumberFormat = "#,##0.00"
    
    '/////////////////PRICES/////////////////

    Range("A1:C1").Font.Bold = True

    Columns("A:B").EntireColumn.AutoFit
    Columns("A:A").HorizontalAlignment = xlCenter

    Range("A2").Select
    ActiveWindow.FreezePanes = True

End Sub

Sub Macro23_Fill_Products()

    Sheets("Products").Select
    
    Dim Cat As Long
    Dim Sub_Cat As Long
    Dim SKU As Long
    
    Dim i As Long
    Dim j As Long
    Dim k As Long
    
    Dim A As String
    Dim B As String
    Dim D As String
    Dim E As String
        
    
    
    Dim Myrange As Range
    Dim Cell As Range
    
    Dim Arreglo(1, 15) As String
    
    
    Arreglo(0, 0) = "Accesories"
    Arreglo(1, 0) = "A2:B15"
    
    Arreglo(0, 1) = "Appliances"
    Arreglo(1, 1) = "A2:B135"
    
    Arreglo(0, 2) = "Art"
    Arreglo(1, 2) = "A2:B140"
    
    Arreglo(0, 3) = "Binders"
    Arreglo(1, 3) = "A2:B12"
    
    Arreglo(0, 4) = "Bookcases"
    Arreglo(1, 4) = "A2:B32"
    
    Arreglo(0, 5) = "Chairs"
    Arreglo(1, 5) = "A2:B159"
    
    Arreglo(0, 6) = "Copiers"
    Arreglo(1, 6) = "A2:B121"
    
    Arreglo(0, 7) = "Envelopes"
    Arreglo(1, 7) = "A2:B31"
    
    Arreglo(0, 8) = "Fasteners"
    Arreglo(1, 8) = "A2:B83"
    
    Arreglo(0, 9) = "Furnishings"
    Arreglo(1, 9) = "A2:B159"
    
    Arreglo(0, 10) = "Labels"
    Arreglo(1, 10) = "A2:B25"
    
    Arreglo(0, 11) = "Gym Machines"
    Arreglo(1, 11) = "A2:B18"
    
    Arreglo(0, 12) = "Papers"
    Arreglo(1, 12) = "A2:B31"
    
    Arreglo(0, 13) = "Storage"
    Arreglo(1, 13) = "A2:B18"
    
    Arreglo(0, 14) = "Supplies"
    Arreglo(1, 14) = "A2:B4"
    
    Arreglo(0, 15) = "Tables"
    Arreglo(1, 15) = "A2:B48"
    
    
    Range("A1").FormulaR1C1 = "Cat"
    Range("B1").FormulaR1C1 = "Sub-Cat"
    Range("C1").FormulaR1C1 = "SKU"
    Range("D1").FormulaR1C1 = "Product Order"
    Range("E1").FormulaR1C1 = "Product Name"
    Range("F1").FormulaR1C1 = "Product Code"
    Range("F1").FormulaR1C1 = "File"
        
    SKU = 1
    
    i = 1
    j = 0
    
    k = 1
    
        
    '///////////////////////////////////////////
    
    For j = 0 To 15
    
        Set Myrange = Worksheets(Arreglo(0, j)).Range(Arreglo(1, j))
        
        For Each Cell In Myrange
    
            If i = 1 Then
            
                Cells(SKU + 1, "D").Value = Cell.Value
                D = Cell.Value
                
                Cells(SKU + 1, "B").Value = "SC0000" & j + 1
                B = "SC0000" & j + 1
                
                Cells(SKU + 1, "C").Value = SKU
                                
                    If j = 4 Or j = 5 Or j = 9 Or j = 13 Or j = 15 Then
                        Cells(SKU + 1, "A").Value = "C0001"
                        A = "C0001"
                        
                    ElseIf j = 2 Or j = 3 Or j = 7 Or j = 10 Or j = 12 Or j = 14 Then
                        Cells(SKU + 1, "A").Value = "C0002"
                        A = "C0002"
                        
                    ElseIf j = 1 Or j = 6 Or j = 8 Then
                        Cells(SKU + 1, "A").Value = "C0003"
                        A = "C0003"
                    
                    ElseIf j = 0 Or j = 6 Or j = 11 Then
                        Cells(SKU + 1, "A").Value = "C0004"
                        A = "C0004"
                    
                    End If
                                
                SKU = SKU + 1
                
                i = i + 1
                
                
            ElseIf i = 2 Then
            
            
                Cells(SKU, "E").Value = Cell.Value
                E = Cell.Value
                            
                i = 1
                
                
            End If
                
            Cells(SKU, "F").Value = A & " - " & B & " - " & D & " - " & E
            
            
            If SKU < 10 Then
                    Cells(SKU, "G").Value = "'000" & SKU - 1
                                                            
            ElseIf SKU < 100 Then
                    Cells(SKU, "G").Value = "'00" & SKU - 1
                                        
            ElseIf SKU < 1000 Then
                    Cells(SKU, "G").Value = "'0" & SKU - 1
                                    
            ElseIf SKU < 10000 Then
                    Cells(SKU, "G").Value = SKU - 1
            End If
            
    
        Next Cell
        
        
    Next j

    '///////////////////////////////////////////
    
        
    Columns("A:F").EntireColumn.AutoFit
    
    Range("A1:F1").Font.Bold = True
    
    Range("A2").Select
    ActiveWindow.FreezePanes = True
    
    ActiveWindow.Zoom = 80
    
End Sub



Sub Macro25_Search_Products()

    UserForm1.Show
        
End Sub

Sub Macro24_Fill_Names()

    Sheets("Names").Select

    Dim Myrange As Range
    
    Dim i As Double

    Dim path As String
    Dim ArrayFile() As String
    Dim numberOfLines As Long

    path = "\data\names.dat"
    modulo2.ReadFile path, ArrayFile, numberOfLines

    Set Myrange = Range("A1:B1014")

    i=0

    for each cell in Myrange
        cell.value = ArrayFile (i)
        i=i+1
    next
    
    ' Dim Arreglo(0, 1012) As String


    ' Set Myrange = Range("A2:B1014")
    ' i = 0
    ' j = 0
    
    ' For Each Cell In Myrange

    '     If j Mod 2 = 0 Then

    '         Cell.Value = i + 1

    '         i = i + 1
    '         j = j + 1
    '     Else

    '         Cell.Value = Arreglo(0, i - 1)

    '         j = j + 1
    '     End If

    ' Next Cell
        
    Range("A1:C1").Font.Bold = True

    Columns("A:A").EntireColumn.AutoFit
    Columns("A:A").HorizontalAlignment = xlCenter

    Columns("B:B").EntireColumn.AutoFit

    Range("A2").Select
    ActiveWindow.FreezePanes = True


End Sub

Sub Macro25_Fill_LastNames()

    Sheets("Last Names").Select

    Dim Myrange As Range
    Dim i As Double
    

    Dim path As String
    Dim ArrayFile() As String
    Dim numberOfLines As Long

    path = "\data\lastnames.dat"
    modulo2.ReadFile path, ArrayFile, numberOfLines

    Set Myrange = Range("A1:B2013")

    i=0

    for each cell in Myrange
        cell.value = ArrayFile (i)
        i=i+1
    next

    Range("A1:C1").Font.Bold = True
    
    Columns("A:A").EntireColumn.AutoFit
    Columns("A:A").HorizontalAlignment = xlCenter
    
    Columns("B:B").EntireColumn.AutoFit
        
    Range("A2").Select
    ActiveWindow.FreezePanes = True
    
End Sub

Sub MacroN_All_TEMP()

    Call Macro00_NewFile
    Call Macro01_Headers
    Call Macro03_InsertCitiesSheet
    Call Macro04_InsertSheets

    Call Macro05_Fill_Categories
    Call Macro06_Fill_Sub_Category
    Call Macro07_Fill_Accesories
    Call Macro08_Fill_Appliances
    Call Macro09_Fill_Binders
    Call Macro10_Fill_Art
    Call Macro11_Fill_Bookcases
    Call Macro12_Fill_Chairs
    Call Macro13_Fill_Copiers
    Call Macro14_Fill_Envelopes
    Call Macro15_Fill_Fasteners
    Call Macro16_Fill_Furnishings
    Call Macro17_Fill_Labels
    Call Macro18_Fill_Gym_Machines
    Call Macro19_Fill_Papers
    Call Macro20_Fill_Storage
    Call Macro21_Fill_Supplies
    Call Macro22_Fill_Tables

    Call Macro23_Fill_Products

    Call Macro24_Fill_Names
    Call Macro25_Fill_LastNames

    Call Macro02_OrdersWithRANDQty

end Sub

Sub MacroN_All()

    Call Macro00_NewFile
    Call Macro01_Headers
    
    Call Macro03_InsertCitiesSheet
    Call Macro04_InsertSheets
    Call Macro05_Fill_Category
    Call Macro06_Fill_Sub_Category
           
    Call Macro07_Fill_Accesories
    Call Macro08_Fill_Appliances
    Call Macro09_Fill_Binders
    Call Macro10_Fill_Art
    Call Macro11_Fill_Bookcases
    Call Macro12_Fill_Chairs
    Call Macro13_Fill_Copiers
    Call Macro14_Fill_Envelopes
    Call Macro15_Fill_Fasteners
    Call Macro16_Fill_Furnishings
    Call Macro17_Fill_Labels
    Call Macro18_Fill_Gym_Machines
    Call Macro19_Fill_Papers
    Call Macro20_Fill_Storage
    Call Macro21_Fill_Supplies
    Call Macro22_Fill_Tables
    
    Call Macro23_Fill_Products
            
        Call Macro24_Fill_Names
        Call Macro25_Fill_LastNames
        
    Call Macro02_OrdersWithRANDQty
    
    Sheets("Sales SuperStore").Select
    
End Sub