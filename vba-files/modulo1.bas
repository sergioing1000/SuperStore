Attribute VB_Name = "modulo1"

Sub MacroN_All()
    Call Macro00_NewFile
    Call Macro01_Headers
End Sub

Sub Macro00_NewFile()

    'This macro creates a new empty Excel file
    'Change the first sheet name as "Sales SuperStore"
    
    Workbooks.Add
    ActiveSheet.Name = "Sales SuperStore"
    
End Sub

Sub Macro01_Headers()

    'This macro Inlcude headers in the "Sales SuperStore" sheet

    Dim i as Integer
    Dim My_range As Range
    Dim header() As String
    
    headers = Array("Category", "Customer ", "Order Date", "Order ID", "Product Name", "Unit Price", "Segment", "Ship Date", "Ship Mode", "Country", "Region", "State", "City order", "City", "Postal Code", "Sub-Category", "Maesure Names", "Discount", "Profit", "Quantity", "Total Price", "Latitude", "Longitude", "Number of records", "Sub-Region")

    Set My_range = Range("A1:Y1")

    i=0
    for each cell in My_range
        cell.value= headers (i)
        i=i +1
    next cell
    
    My_range.Font.Bold = True
        
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

    Arreglo(0, 1) = "Order"
    Arreglo(0, 2) = "City"
    Arreglo(0, 3) = "State"
    Arreglo(0, 4) = "ST"
    Arreglo(0, 5) = "Latitude"
    Arreglo(0, 6) = "Logitude"
    Arreglo(0, 7) = "ZIP CODE"
    Arreglo(0, 8) = "Country"
    Arreglo(0, 9) = "Region"
    Arreglo(0, 10) = "Sub - Region Number"
    
    Arreglo(0, 11) = "1"
    Arreglo(0, 12) = "Birmingham"
    Arreglo(0, 13) = "Alabama"
    Arreglo(0, 14) = "AL"
    Arreglo(0, 15) = "33.5186° N"
    Arreglo(0, 16) = "86.8104° W"
    Arreglo(0, 17) = "35203"
    Arreglo(0, 18) = "US"
    Arreglo(0, 19) = "South"
    Arreglo(0, 20) = "4"
    
    Arreglo(0, 21) = "2"
    Arreglo(0, 22) = "Montgomery"
    Arreglo(0, 23) = "Alabama"
    Arreglo(0, 24) = "AL"
    Arreglo(0, 25) = "32.3792° N"
    Arreglo(0, 26) = "86.3077° W"
    Arreglo(0, 27) = "36104"
    Arreglo(0, 28) = "US"
    Arreglo(0, 29) = "South"
    Arreglo(0, 30) = "4"
    
    Arreglo(0, 31) = "3"
    Arreglo(0, 32) = "Anchorage"
    Arreglo(0, 33) = "Alaska"
    Arreglo(0, 34) = "AK"
    Arreglo(0, 35) = "61.2181° N"
    Arreglo(0, 36) = "149.9003° W"
    Arreglo(0, 37) = "99513"
    Arreglo(0, 38) = "US"
    Arreglo(0, 39) = "West"
    Arreglo(0, 40) = "10"
    
    Arreglo(0, 41) = "4"
    Arreglo(0, 42) = "Tucson"
    Arreglo(0, 43) = "Arizona"
    Arreglo(0, 44) = "AZ"
    Arreglo(0, 45) = "32.2226° N"
    Arreglo(0, 46) = "110.9747° W"
    Arreglo(0, 47) = "85701"
    Arreglo(0, 48) = "US"
    Arreglo(0, 49) = "West"
    Arreglo(0, 50) = "9"
    
    Arreglo(0, 51) = "5"
    Arreglo(0, 52) = "Phoenix"
    Arreglo(0, 53) = "Arizona"
    Arreglo(0, 54) = "AZ"
    Arreglo(0, 55) = "33.4484° N"
    Arreglo(0, 56) = "112.0740° W"
    Arreglo(0, 57) = "85003"
    Arreglo(0, 58) = "US"
    Arreglo(0, 59) = "West"
    Arreglo(0, 60) = "9"
    
    Arreglo(0, 61) = "6"
    Arreglo(0, 62) = "Little Rock"
    Arreglo(0, 63) = "Arkansas"
    Arreglo(0, 64) = "AR"
    Arreglo(0, 65) = "34.7465° N"
    Arreglo(0, 66) = "92.2896° W"
    Arreglo(0, 67) = "72201"
    Arreglo(0, 68) = "US"
    Arreglo(0, 69) = "South"
    Arreglo(0, 70) = "6"
    
    Arreglo(0, 71) = "7"
    Arreglo(0, 72) = "San Francisco"
    Arreglo(0, 73) = "California"
    Arreglo(0, 74) = "CA"
    Arreglo(0, 75) = "37.7749° N"
    Arreglo(0, 76) = "122.4194° W"
    Arreglo(0, 77) = "94102"
    Arreglo(0, 78) = "US"
    Arreglo(0, 79) = "West"
    Arreglo(0, 80) = "9"
    
    Arreglo(0, 81) = "8"
    Arreglo(0, 82) = "Los Angeles"
    Arreglo(0, 83) = "California"
    Arreglo(0, 84) = "CA"
    Arreglo(0, 85) = "34.0522° N"
    Arreglo(0, 86) = "118.2437° W"
    Arreglo(0, 87) = "90013"
    Arreglo(0, 88) = "US"
    Arreglo(0, 89) = "West"
    Arreglo(0, 90) = "9"
    
    Arreglo(0, 91) = "9"
    Arreglo(0, 92) = "San Diego"
    Arreglo(0, 93) = "California"
    Arreglo(0, 94) = "CA"
    Arreglo(0, 95) = "32.7157° N"
    Arreglo(0, 96) = "117.1611° W"
    Arreglo(0, 97) = "92101"
    Arreglo(0, 98) = "US"
    Arreglo(0, 99) = "West"
    Arreglo(0, 100) = "9"
    
    Arreglo(0, 101) = "10"
    Arreglo(0, 102) = "Denver"
    Arreglo(0, 103) = "Colorado"
    Arreglo(0, 104) = "CO"
    Arreglo(0, 105) = "39.7392° N"
    Arreglo(0, 106) = "104.9903° W"
    Arreglo(0, 107) = "80261"
    Arreglo(0, 108) = "US"
    Arreglo(0, 109) = "Central"
    Arreglo(0, 110) = "8"
    
    Arreglo(0, 111) = "11"
    Arreglo(0, 112) = "Hartford"
    Arreglo(0, 113) = "Connecticut"
    Arreglo(0, 114) = "CT"
    Arreglo(0, 115) = "41.7658° N"
    Arreglo(0, 116) = "72.6734° W"
    Arreglo(0, 117) = "06161"
    Arreglo(0, 118) = "US"
    Arreglo(0, 119) = "East"
    Arreglo(0, 120) = "1"
    
    Arreglo(0, 121) = "12"
    Arreglo(0, 122) = "Wilmington"
    Arreglo(0, 123) = "Delaware"
    Arreglo(0, 124) = "DE"
    Arreglo(0, 125) = "34.2104° N"
    Arreglo(0, 126) = "77.8868° W"
    Arreglo(0, 127) = "28403"
    Arreglo(0, 128) = "US"
    Arreglo(0, 129) = "East"
    Arreglo(0, 130) = "3"
    
    Arreglo(0, 131) = "13"
    Arreglo(0, 132) = "Miami"
    Arreglo(0, 133) = "Florida"
    Arreglo(0, 134) = "FL"
    Arreglo(0, 135) = "25.7617° N"
    Arreglo(0, 136) = "80.1918° W"
    Arreglo(0, 137) = "33131"
    Arreglo(0, 138) = "US"
    Arreglo(0, 139) = "South"
    Arreglo(0, 140) = "4"
    
    Arreglo(0, 141) = "14"
    Arreglo(0, 142) = "Atlanta"
    Arreglo(0, 143) = "Georgia"
    Arreglo(0, 144) = "GA"
    Arreglo(0, 145) = "33.7490° N"
    Arreglo(0, 146) = "84.3880° W"
    Arreglo(0, 147) = "30335"
    Arreglo(0, 148) = "US"
    Arreglo(0, 149) = "South"
    Arreglo(0, 150) = "4"
    
    Arreglo(0, 151) = "15"
    Arreglo(0, 152) = "Columbus"
    Arreglo(0, 153) = "Georgia"
    Arreglo(0, 154) = "GA"
    Arreglo(0, 155) = "39.9612° N"
    Arreglo(0, 156) = "82.9988° W"
    Arreglo(0, 157) = "43215"
    Arreglo(0, 158) = "US"
    Arreglo(0, 159) = "South"
    Arreglo(0, 160) = "4"
    
    Arreglo(0, 161) = "16"
    Arreglo(0, 162) = "Boise"
    Arreglo(0, 163) = "Idaho"
    Arreglo(0, 164) = "ID"
    Arreglo(0, 165) = "43.6150° N"
    Arreglo(0, 166) = "116.2023° W"
    Arreglo(0, 167) = "83724"
    Arreglo(0, 168) = "US"
    Arreglo(0, 169) = "West"
    Arreglo(0, 170) = "10"
    
    Arreglo(0, 171) = "17"
    Arreglo(0, 172) = "Chicago"
    Arreglo(0, 173) = "Illinois"
    Arreglo(0, 174) = "IL"
    Arreglo(0, 175) = "41.8781° N"
    Arreglo(0, 176) = "87.6298° W"
    Arreglo(0, 177) = "60604"
    Arreglo(0, 178) = "US"
    Arreglo(0, 179) = "Central"
    Arreglo(0, 180) = "5"
    
    Arreglo(0, 181) = "18"
    Arreglo(0, 182) = "Indianapolis"
    Arreglo(0, 183) = "Indiana"
    Arreglo(0, 184) = "IN"
    Arreglo(0, 185) = "39.7684° N"
    Arreglo(0, 186) = "86.1581° W"
    Arreglo(0, 187) = "46207"
    Arreglo(0, 188) = "US"
    Arreglo(0, 189) = "Central"
    Arreglo(0, 190) = "5"
    
    Arreglo(0, 191) = "19"
    Arreglo(0, 192) = "Des Moines"
    Arreglo(0, 193) = "Iowa"
    Arreglo(0, 194) = "IA"
    Arreglo(0, 195) = "41.5868° N"
    Arreglo(0, 196) = "93.6250° W"
    Arreglo(0, 197) = "50392"
    Arreglo(0, 198) = "US"
    Arreglo(0, 199) = "Central"
    Arreglo(0, 200) = "7"
    
    Arreglo(0, 201) = "20"
    Arreglo(0, 202) = "Cedar Rapids"
    Arreglo(0, 203) = "Iowa"
    Arreglo(0, 204) = "IA"
    Arreglo(0, 205) = "41.9779° N"
    Arreglo(0, 206) = "91.6656° W"
    Arreglo(0, 207) = "52401"
    Arreglo(0, 208) = "US"
    Arreglo(0, 209) = "Central"
    Arreglo(0, 210) = "7"
    
    Arreglo(0, 211) = "21"
    Arreglo(0, 212) = "Topeka"
    Arreglo(0, 213) = "Kansas"
    Arreglo(0, 214) = "KS"
    Arreglo(0, 215) = "39.0473° N"
    Arreglo(0, 216) = "95.6752° W"
    Arreglo(0, 217) = "66612"
    Arreglo(0, 218) = "US"
    Arreglo(0, 219) = "Central"
    Arreglo(0, 220) = "7"
    
    Arreglo(0, 221) = "22"
    Arreglo(0, 222) = "Louisville"
    Arreglo(0, 223) = "Kentucky"
    Arreglo(0, 224) = "KY"
    Arreglo(0, 225) = "38.2527° N"
    Arreglo(0, 226) = "85.7585° W"
    Arreglo(0, 227) = "40202"
    Arreglo(0, 228) = "US"
    Arreglo(0, 229) = "South"
    Arreglo(0, 230) = "4"
    
    Arreglo(0, 231) = "23"
    Arreglo(0, 232) = "Lexington"
    Arreglo(0, 233) = "Kentucky"
    Arreglo(0, 234) = "KY"
    Arreglo(0, 235) = "38.0406° N"
    Arreglo(0, 236) = "84.5037° W"
    Arreglo(0, 237) = "40507"
    Arreglo(0, 238) = "US"
    Arreglo(0, 239) = "South"
    Arreglo(0, 240) = "4"
    
    Arreglo(0, 241) = "24"
    Arreglo(0, 242) = "New Orleans"
    Arreglo(0, 243) = "Louisiana"
    Arreglo(0, 244) = "LA"
    Arreglo(0, 245) = "29.9511° N"
    Arreglo(0, 246) = "90.0715° W"
    Arreglo(0, 247) = "70163"
    Arreglo(0, 248) = "US"
    Arreglo(0, 249) = "South"
    Arreglo(0, 250) = "6"
    
    Arreglo(0, 251) = "25"
    Arreglo(0, 252) = "Alexandria"
    Arreglo(0, 253) = "Louisiana"
    Arreglo(0, 254) = "LA"
    Arreglo(0, 255) = "31.3113° N"
    Arreglo(0, 256) = "92.4451° W"
    Arreglo(0, 257) = "71301"
    Arreglo(0, 258) = "US"
    Arreglo(0, 259) = "South"
    Arreglo(0, 260) = "6"
    
    Arreglo(0, 261) = "26"
    Arreglo(0, 262) = "Augusta"
    Arreglo(0, 263) = "Maine"
    Arreglo(0, 264) = "ME"
    Arreglo(0, 265) = "33.4735° N"
    Arreglo(0, 266) = "82.0105° W"
    Arreglo(0, 267) = "30904"
    Arreglo(0, 268) = "US"
    Arreglo(0, 269) = "East"
    Arreglo(0, 270) = "1"
    
    Arreglo(0, 271) = "27"
    Arreglo(0, 272) = "Baltimore"
    Arreglo(0, 273) = "Maryland"
    Arreglo(0, 274) = "MD"
    Arreglo(0, 275) = "39.2904° N"
    Arreglo(0, 276) = "76.6122° W"
    Arreglo(0, 277) = "21201"
    Arreglo(0, 278) = "US"
    Arreglo(0, 279) = "East"
    Arreglo(0, 280) = "3"
    
    Arreglo(0, 281) = "28"
    Arreglo(0, 282) = "Frederick"
    Arreglo(0, 283) = "Maryland"
    Arreglo(0, 284) = "MD"
    Arreglo(0, 285) = "39.4143° N"
    Arreglo(0, 286) = "77.4105° W"
    Arreglo(0, 287) = "21701"
    Arreglo(0, 288) = "US"
    Arreglo(0, 289) = "East"
    Arreglo(0, 290) = "3"
    
    Arreglo(0, 291) = "29"
    Arreglo(0, 292) = "Boston"
    Arreglo(0, 293) = "Massachusetts"
    Arreglo(0, 294) = "MA"
    Arreglo(0, 295) = "42.3601° N"
    Arreglo(0, 296) = "71.0589° W"
    Arreglo(0, 297) = "02203"
    Arreglo(0, 298) = "US"
    Arreglo(0, 299) = "East"
    Arreglo(0, 300) = "1"
    
    Arreglo(0, 301) = "30"
    Arreglo(0, 302) = "Worcester"
    Arreglo(0, 303) = "Massachusetts"
    Arreglo(0, 304) = "MA"
    Arreglo(0, 305) = "42.2626° N"
    Arreglo(0, 306) = "71.8023° W"
    Arreglo(0, 307) = "01608"
    Arreglo(0, 308) = "US"
    Arreglo(0, 309) = "East"
    Arreglo(0, 310) = "1"
    
    Arreglo(0, 311) = "31"
    Arreglo(0, 312) = "Detroit"
    Arreglo(0, 313) = "Michigan"
    Arreglo(0, 314) = "MI"
    Arreglo(0, 315) = "42.3314° N"
    Arreglo(0, 316) = "83.0458° W"
    Arreglo(0, 317) = "48226"
    Arreglo(0, 318) = "US"
    Arreglo(0, 319) = "Central"
    Arreglo(0, 320) = "5"
    
    Arreglo(0, 321) = "32"
    Arreglo(0, 322) = "Minneapolis"
    Arreglo(0, 323) = "Minnesota"
    Arreglo(0, 324) = "MN"
    Arreglo(0, 325) = "44.9778° N"
    Arreglo(0, 326) = "93.2650° W"
    Arreglo(0, 327) = "55401"
    Arreglo(0, 328) = "US"
    Arreglo(0, 329) = "Central"
    Arreglo(0, 330) = "5"
    
    Arreglo(0, 331) = "33"
    Arreglo(0, 332) = "Jackson"
    Arreglo(0, 333) = "Mississippi"
    Arreglo(0, 334) = "MS"
    Arreglo(0, 335) = "32.2988° N"
    Arreglo(0, 336) = "90.1848° W"
    Arreglo(0, 337) = "39201"
    Arreglo(0, 338) = "US"
    Arreglo(0, 339) = "South"
    Arreglo(0, 340) = "4"
    
    Arreglo(0, 341) = "34"
    Arreglo(0, 342) = "Kansas City"
    Arreglo(0, 343) = "Missouri"
    Arreglo(0, 344) = "MO"
    Arreglo(0, 345) = "39.0997° N"
    Arreglo(0, 346) = "94.5786° W"
    Arreglo(0, 347) = "64106"
    Arreglo(0, 348) = "US"
    Arreglo(0, 349) = "Central"
    Arreglo(0, 350) = "7"
    
    Arreglo(0, 351) = "35"
    Arreglo(0, 352) = "St. Louis"
    Arreglo(0, 353) = "Missouri"
    Arreglo(0, 354) = "MO"
    Arreglo(0, 355) = "38.6270° N"
    Arreglo(0, 356) = "90.1994° W"
    Arreglo(0, 357) = "63103"
    Arreglo(0, 358) = "US"
    Arreglo(0, 359) = "Central"
    Arreglo(0, 360) = "7"
    
    Arreglo(0, 361) = "36"
    Arreglo(0, 362) = "Billings"
    Arreglo(0, 363) = "Montana"
    Arreglo(0, 364) = "MT"
    Arreglo(0, 365) = "45.7833° N"
    Arreglo(0, 366) = "108.5007° W"
    Arreglo(0, 367) = "59101"
    Arreglo(0, 368) = "US"
    Arreglo(0, 369) = "Central"
    Arreglo(0, 370) = "8"
    
    Arreglo(0, 371) = "37"
    Arreglo(0, 372) = "Bellevue "
    Arreglo(0, 373) = "Nebraska"
    Arreglo(0, 374) = "NE"
    Arreglo(0, 375) = "41.1544° N"
    Arreglo(0, 376) = "95.9146° W"
    Arreglo(0, 377) = "68005"
    Arreglo(0, 378) = "US"
    Arreglo(0, 379) = "Central"
    Arreglo(0, 380) = "7"
    
    Arreglo(0, 381) = "38"
    Arreglo(0, 382) = "Carson City"
    Arreglo(0, 383) = "Nevada"
    Arreglo(0, 384) = "NV"
    Arreglo(0, 385) = "39.1638° N"
    Arreglo(0, 386) = "119.7674° W"
    Arreglo(0, 387) = "89701"
    Arreglo(0, 388) = "US"
    Arreglo(0, 389) = "West"
    Arreglo(0, 390) = "9"
    
    Arreglo(0, 391) = "39"
    Arreglo(0, 392) = "Las Vegas"
    Arreglo(0, 393) = "Nevada"
    Arreglo(0, 394) = "NV"
    Arreglo(0, 395) = "36.1699° N"
    Arreglo(0, 396) = "115.1398° W"
    Arreglo(0, 397) = "89101"
    Arreglo(0, 398) = "US"
    Arreglo(0, 399) = "West"
    Arreglo(0, 400) = "9"
    
    Arreglo(0, 401) = "40"
    Arreglo(0, 402) = "Reno"
    Arreglo(0, 403) = "Nevada"
    Arreglo(0, 404) = "NV"
    Arreglo(0, 405) = "39.5296° N"
    Arreglo(0, 406) = "119.8138° W"
    Arreglo(0, 407) = "89501"
    Arreglo(0, 408) = "US"
    Arreglo(0, 409) = "West"
    Arreglo(0, 410) = "9"
    
    Arreglo(0, 411) = "41"
    Arreglo(0, 412) = "Concord"
    Arreglo(0, 413) = "New Hampshire"
    Arreglo(0, 414) = "NH"
    Arreglo(0, 415) = "43.2081° N"
    Arreglo(0, 416) = "71.5376° W"
    Arreglo(0, 417) = "03301"
    Arreglo(0, 418) = "US"
    Arreglo(0, 419) = "East"
    Arreglo(0, 420) = "1"
    
    Arreglo(0, 421) = "42"
    Arreglo(0, 422) = "Atlantic City"
    Arreglo(0, 423) = "New Jersey"
    Arreglo(0, 424) = "NJ"
    Arreglo(0, 425) = "39.3643° N"
    Arreglo(0, 426) = "74.4229° W"
    Arreglo(0, 427) = "08401"
    Arreglo(0, 428) = "US"
    Arreglo(0, 429) = "East"
    Arreglo(0, 430) = "2"
    
    Arreglo(0, 431) = "43"
    Arreglo(0, 432) = "Albuquerque"
    Arreglo(0, 433) = "New Mexico"
    Arreglo(0, 434) = "NM"
    Arreglo(0, 435) = "35.0844° N"
    Arreglo(0, 436) = "106.6504° W"
    Arreglo(0, 437) = "87102"
    Arreglo(0, 438) = "US"
    Arreglo(0, 439) = "South"
    Arreglo(0, 440) = "6"
    
    Arreglo(0, 441) = "44"
    Arreglo(0, 442) = "New York City"
    Arreglo(0, 443) = "New York"
    Arreglo(0, 444) = "NY"
    Arreglo(0, 445) = "40.7128° N"
    Arreglo(0, 446) = "74.0060° W"
    Arreglo(0, 447) = "10007"
    Arreglo(0, 448) = "US"
    Arreglo(0, 449) = "East"
    Arreglo(0, 450) = "2"
    
    Arreglo(0, 451) = "45"
    Arreglo(0, 452) = "Buffalo"
    Arreglo(0, 453) = "New York"
    Arreglo(0, 454) = "NY"
    Arreglo(0, 455) = "42.8864° N"
    Arreglo(0, 456) = "78.8784° W"
    Arreglo(0, 457) = "14202"
    Arreglo(0, 458) = "US"
    Arreglo(0, 459) = "East"
    Arreglo(0, 460) = "2"
    
    Arreglo(0, 461) = "46"
    Arreglo(0, 462) = "Charlotte"
    Arreglo(0, 463) = "North Carolina"
    Arreglo(0, 464) = "NC"
    Arreglo(0, 465) = "35.2271° N"
    Arreglo(0, 466) = "80.8431° W"
    Arreglo(0, 467) = "28280"
    Arreglo(0, 468) = "US"
    Arreglo(0, 469) = "South"
    Arreglo(0, 470) = "4"
    
    Arreglo(0, 471) = "47"
    Arreglo(0, 472) = "Raleigh"
    Arreglo(0, 473) = "North Carolina"
    Arreglo(0, 474) = "NC"
    Arreglo(0, 475) = "35.7796° N"
    Arreglo(0, 476) = "78.6382° W"
    Arreglo(0, 477) = "27601"
    Arreglo(0, 478) = "US"
    Arreglo(0, 479) = "South"
    Arreglo(0, 480) = "4"
    
    Arreglo(0, 481) = "48"
    Arreglo(0, 482) = "Columbus"
    Arreglo(0, 483) = "Ohio"
    Arreglo(0, 484) = "OH"
    Arreglo(0, 485) = "39.9612° N"
    Arreglo(0, 486) = "82.9988° W"
    Arreglo(0, 487) = "43215"
    Arreglo(0, 488) = "US"
    Arreglo(0, 489) = "Central"
    Arreglo(0, 490) = "5"
    
    Arreglo(0, 491) = "49"
    Arreglo(0, 492) = "Oklahoma City"
    Arreglo(0, 493) = "Oklahoma"
    Arreglo(0, 494) = "OK"
    Arreglo(0, 495) = "35.4676° N"
    Arreglo(0, 496) = "97.5164° W"
    Arreglo(0, 497) = "73102"
    Arreglo(0, 498) = "US"
    Arreglo(0, 499) = "South"
    Arreglo(0, 500) = "6"
    
    Arreglo(0, 501) = "50"
    Arreglo(0, 502) = "Tulsa"
    Arreglo(0, 503) = "Oklahoma"
    Arreglo(0, 504) = "OK"
    Arreglo(0, 505) = "36.1540° N"
    Arreglo(0, 506) = "95.9928° W"
    Arreglo(0, 507) = "74103"
    Arreglo(0, 508) = "US"
    Arreglo(0, 509) = "South"
    Arreglo(0, 510) = "6"
    
    Arreglo(0, 511) = "51"
    Arreglo(0, 512) = "Portland"
    Arreglo(0, 513) = "Oregon"
    Arreglo(0, 514) = "OR"
    Arreglo(0, 515) = "45.5051° N"
    Arreglo(0, 516) = "122.6750° W"
    Arreglo(0, 517) = "97201"
    Arreglo(0, 518) = "US"
    Arreglo(0, 519) = "West"
    Arreglo(0, 520) = "10"
    
    Arreglo(0, 521) = "52"
    Arreglo(0, 522) = "Philadelphia"
    Arreglo(0, 523) = "Pennsylvania"
    Arreglo(0, 524) = "PA"
    Arreglo(0, 525) = "39.9526° N"
    Arreglo(0, 526) = "75.1652° W"
    Arreglo(0, 527) = "19107"
    Arreglo(0, 528) = "US"
    Arreglo(0, 529) = "East"
    Arreglo(0, 530) = "3"
    
    Arreglo(0, 531) = "53"
    Arreglo(0, 532) = "Providence"
    Arreglo(0, 533) = "Rhode Island"
    Arreglo(0, 534) = "RI"
    Arreglo(0, 535) = "41.8240° N"
    Arreglo(0, 536) = "71.4128° W"
    Arreglo(0, 537) = "02903"
    Arreglo(0, 538) = "US"
    Arreglo(0, 539) = "East"
    Arreglo(0, 540) = "1"
    
    Arreglo(0, 541) = "54"
    Arreglo(0, 542) = "Charleston"
    Arreglo(0, 543) = "South Carolina"
    Arreglo(0, 544) = "SC"
    Arreglo(0, 545) = "32.7765° N"
    Arreglo(0, 546) = "79.9311° W"
    Arreglo(0, 547) = "29401"
    Arreglo(0, 548) = "US"
    Arreglo(0, 549) = "South"
    Arreglo(0, 550) = "4"
    
    Arreglo(0, 551) = "55"
    Arreglo(0, 552) = "Sioux Falls"
    Arreglo(0, 553) = "South Dakota"
    Arreglo(0, 554) = "SD"
    Arreglo(0, 555) = "43.5473° N"
    Arreglo(0, 556) = "96.7283° W"
    Arreglo(0, 557) = "57104"
    Arreglo(0, 558) = "US"
    Arreglo(0, 559) = "Central"
    Arreglo(0, 560) = "8"
    
    Arreglo(0, 561) = "56"
    Arreglo(0, 562) = "Nashville"
    Arreglo(0, 563) = "Tennessee"
    Arreglo(0, 564) = "TN"
    Arreglo(0, 565) = "36.1627° N"
    Arreglo(0, 566) = "86.7816° W"
    Arreglo(0, 567) = "37219"
    Arreglo(0, 568) = "US"
    Arreglo(0, 569) = "South"
    Arreglo(0, 570) = "4"
    
    Arreglo(0, 571) = "57"
    Arreglo(0, 572) = "Houston"
    Arreglo(0, 573) = "Texas"
    Arreglo(0, 574) = "TX"
    Arreglo(0, 575) = "29.7604° N"
    Arreglo(0, 576) = "95.3698° W"
    Arreglo(0, 577) = "77002"
    Arreglo(0, 578) = "US"
    Arreglo(0, 579) = "South"
    Arreglo(0, 580) = "6"
    
    Arreglo(0, 581) = "58"
    Arreglo(0, 582) = "San Antonio"
    Arreglo(0, 583) = "Texas"
    Arreglo(0, 584) = "TX"
    Arreglo(0, 585) = "29.4241° N"
    Arreglo(0, 586) = "98.4936° W"
    Arreglo(0, 587) = "78205"
    Arreglo(0, 588) = "US"
    Arreglo(0, 589) = "South"
    Arreglo(0, 590) = "6"
    
    Arreglo(0, 591) = "59"
    Arreglo(0, 592) = "Dallas"
    Arreglo(0, 593) = "Texas"
    Arreglo(0, 594) = "TX"
    Arreglo(0, 595) = "32.7767° N"
    Arreglo(0, 596) = "96.7970° W"
    Arreglo(0, 597) = "75201"
    Arreglo(0, 598) = "US"
    Arreglo(0, 599) = "South"
    Arreglo(0, 600) = "6"
    
    Arreglo(0, 601) = "60"
    Arreglo(0, 602) = "Salt Lake City"
    Arreglo(0, 603) = "Utah"
    Arreglo(0, 604) = "UT"
    Arreglo(0, 605) = "40.7608° N"
    Arreglo(0, 606) = "111.8910° W"
    Arreglo(0, 607) = "84111"
    Arreglo(0, 608) = "US"
    Arreglo(0, 609) = "Central"
    Arreglo(0, 610) = "8"
    
    Arreglo(0, 611) = "61"
    Arreglo(0, 612) = "Burlington"
    Arreglo(0, 613) = "Vermont"
    Arreglo(0, 614) = "VT"
    Arreglo(0, 615) = "44.4759° N"
    Arreglo(0, 616) = "73.2121° W"
    Arreglo(0, 617) = "05401"
    Arreglo(0, 618) = "US"
    Arreglo(0, 619) = "East"
    Arreglo(0, 620) = "1"
    
    Arreglo(0, 621) = "62"
    Arreglo(0, 622) = "Virginia Beach"
    Arreglo(0, 623) = "Virginia"
    Arreglo(0, 624) = "VA"
    Arreglo(0, 625) = "36.8529° N"
    Arreglo(0, 626) = "75.9780° W"
    Arreglo(0, 627) = "23451"
    Arreglo(0, 628) = "US"
    Arreglo(0, 629) = "East"
    Arreglo(0, 630) = "3"
    
    Arreglo(0, 631) = "63"
    Arreglo(0, 632) = "Seattle"
    Arreglo(0, 633) = "Washington"
    Arreglo(0, 634) = "WA"
    Arreglo(0, 635) = "47.6062° N"
    Arreglo(0, 636) = "122.3321° W"
    Arreglo(0, 637) = "98164"
    Arreglo(0, 638) = "US"
    Arreglo(0, 639) = "West"
    Arreglo(0, 640) = "10"
    
    Arreglo(0, 641) = "64"
    Arreglo(0, 642) = "Charleston"
    Arreglo(0, 643) = "West Virginia"
    Arreglo(0, 644) = "WV"
    Arreglo(0, 645) = "38.3498° N"
    Arreglo(0, 646) = "81.6326° W"
    Arreglo(0, 647) = "25301"
    Arreglo(0, 648) = "US"
    Arreglo(0, 649) = "East"
    Arreglo(0, 650) = "3"
    
    Arreglo(0, 651) = "65"
    Arreglo(0, 652) = "Milwaukee"
    Arreglo(0, 653) = "Wisconsin"
    Arreglo(0, 654) = "WI"
    Arreglo(0, 655) = "43.0389° N"
    Arreglo(0, 656) = "87.9065° W"
    Arreglo(0, 657) = "53202"
    Arreglo(0, 658) = "US"
    Arreglo(0, 659) = "Central"
    Arreglo(0, 660) = "5"
    
    Arreglo(0, 661) = "66"
    Arreglo(0, 662) = "Laramie"
    Arreglo(0, 663) = "Wyoming"
    Arreglo(0, 664) = "WY"
    Arreglo(0, 665) = "41.3114° N"
    Arreglo(0, 666) = "105.5911° W"
    Arreglo(0, 667) = "82070"
    Arreglo(0, 668) = "US"
    Arreglo(0, 669) = "Central"
    Arreglo(0, 670) = "8"
 
    Application.DisplayAlerts = False
    
    For Each ws In ActiveWorkbook.Worksheets
        
        If ws.Name = "Cities-Stores" Then
            ws.Delete
        End If
        
    Next ws
    
    Application.DisplayAlerts = True
            
    Sheets.Add After:=ActiveSheet
    ActiveSheet.Name = "Cities-Stores"
    
    Set Myrange = Range("A1:J67")
    i = 1
        
    For Each Cell In Myrange
        
        Cell.Value = Arreglo(0, i)
        
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
    
'   Application.ErrorCheckingOptions.BackgroundChecking = False
        
End Sub

Sub Macro04_InsertSheets()

    Sheets.Add After:=ActiveSheet
    ActiveSheet.Name = "People"
    
    Sheets.Add After:=ActiveSheet
    ActiveSheet.Name = "Names"
    
    Sheets.Add After:=ActiveSheet
    ActiveSheet.Name = "Last Names"
    
    Sheets.Add After:=ActiveSheet
    ActiveSheet.Name = "Category"
    
    Sheets.Add After:=ActiveSheet
    ActiveSheet.Name = "Sub-Category"
    
    Sheets.Add After:=ActiveSheet
    ActiveSheet.Name = "Products"
    
    Sheets.Add After:=ActiveSheet
    ActiveSheet.Name = "Accesories"
    
    Sheets.Add After:=ActiveSheet
    ActiveSheet.Name = "Appliances"
    
    Sheets.Add After:=ActiveSheet
    ActiveSheet.Name = "Art"
    
    Sheets.Add After:=ActiveSheet
    ActiveSheet.Name = "Binders"
    
    Sheets.Add After:=ActiveSheet
    ActiveSheet.Name = "Bookcases"
        
    Sheets.Add After:=ActiveSheet
    ActiveSheet.Name = "Chairs"
    
    Sheets.Add After:=ActiveSheet
    ActiveSheet.Name = "Copiers"
    
    Sheets.Add After:=ActiveSheet
    ActiveSheet.Name = "Envelopes"
    
    Sheets.Add After:=ActiveSheet
    ActiveSheet.Name = "Fasteners"
    
    Sheets.Add After:=ActiveSheet
    ActiveSheet.Name = "Furnishings"
    
    Sheets.Add After:=ActiveSheet
    ActiveSheet.Name = "Labels"
    
    Sheets.Add After:=ActiveSheet
    ActiveSheet.Name = "Gym Machines"
    
    Sheets.Add After:=ActiveSheet
    ActiveSheet.Name = "Papers"
    
    Sheets.Add After:=ActiveSheet
    ActiveSheet.Name = "Storage"
    
    Sheets.Add After:=ActiveSheet
    ActiveSheet.Name = "Supplies"
    
    Sheets.Add After:=ActiveSheet
    ActiveSheet.Name = "Tables"
 
    Sheets.Add After:=ActiveSheet
    ActiveSheet.Name = "Ship Mode"
    
    Sheets.Add After:=ActiveSheet
    ActiveSheet.Name = "Segment"
    
    Sheets.Add After:=ActiveSheet
    ActiveSheet.Name = "Region"
    
    Sheets.Add After:=ActiveSheet
    ActiveSheet.Name = "State"
    
    
End Sub

Sub Macro05_Fill_Category()

    Sheets("Category").Select
    
    Range("A1").Value = "Order"
    Range("A1").Font.Bold = True
    
    Range("B1").Value = "Category"
    Range("B1").Font.Bold = True
    
    Range("A2").Value = "1"
    Range("B2").Value = "Furniture"
    Range("A3").Value = "2"
    Range("B3").Value = "Office Supplies"
    Range("A4").Value = "3"
    Range("B4").Value = "Technology"
    Range("A5").Value = "4"
    Range("B5").Value = "Beauty"
    
    Columns("A:B").EntireColumn.AutoFit
    Columns("A:A").HorizontalAlignment = xlCenter
    
    Range("A2").Select
    ActiveWindow.FreezePanes = True
        
End Sub
Sub Macro06_Fill_Sub_Category()

    
    Dim i As Integer
    Dim j As Integer
    Dim Arreglo(0, 15) As String
        
    Dim Cell As Range
    Dim Myrange As Range
    
    Set Myrange = Range("A2:B17")
    
    Sheets("Sub-Category").Select
    
    Arreglo(0, 0) = "Accesories"
    Arreglo(0, 1) = "Appliances"
    Arreglo(0, 2) = "Art"
    Arreglo(0, 3) = "Binders"
    Arreglo(0, 4) = "Bookcases"
    Arreglo(0, 5) = "Chairs"
    Arreglo(0, 6) = "Copiers"
    Arreglo(0, 7) = "Envelopes"
    Arreglo(0, 8) = "Fasteners"
    Arreglo(0, 9) = "Furnishings"
    Arreglo(0, 10) = "Labels"
    Arreglo(0, 11) = "Gym Machines"
    Arreglo(0, 12) = "Papers"
    Arreglo(0, 13) = "Storage"
    Arreglo(0, 14) = "Suppliers"
    Arreglo(0, 15) = "Tables"
    
    
    Set Myrange = Range("A2:B17")
    
    i = 0

    j = 0

    For Each Cell In Myrange

        If j Mod 2 = 0 Then

            Cell.Value = i + 1

            i = i + 1
            j = j + 1
        Else

            Cell.Value = Arreglo(0, i - 1)

            j = j + 1
        End If

    Next Cell

    
    Range("A1").Value = "Order"
    Range("A1").Font.Bold = True
    
    Range("B1").Value = "Sub-Category"
    Range("B1").Font.Bold = True
    
    Columns("A:B").EntireColumn.AutoFit
    Columns("A:A").HorizontalAlignment = xlCenter
    
    Range("A2").Select
    ActiveWindow.FreezePanes = True
        
End Sub


Sub Macro08_Fill_Accesories()
    
    Sheets("Accesories").Select

    Dim Myrange As Range
    Dim Cell As Range
    Dim i As Long
    Dim j As Long

    Dim Arreglo(0, 13) As String

    Arreglo(0, 0) = "Headbands"
    Arreglo(0, 1) = "Hair Clips"
    Arreglo(0, 2) = "Hair Pins"
    Arreglo(0, 3) = "Hair Extensions"
    Arreglo(0, 4) = "Hair Accessories"
    Arreglo(0, 5) = "Cosmetic Bags"
    Arreglo(0, 6) = "Nail Polish & Decoration Products"
    Arreglo(0, 7) = "Hair Care"
    Arreglo(0, 8) = "Makeup"
    Arreglo(0, 9) = "Skin Care"
    Arreglo(0, 10) = "Foot, Hand & Nail Care"
    Arreglo(0, 11) = "Tools & Accessories"
    Arreglo(0, 12) = "Shave & Hair Removal"
    Arreglo(0, 13) = "Personal Care"
  
  
    Set Myrange = Range("A2:B15")
    i = 0

    j = 0

    For Each Cell In Myrange

        If j Mod 2 = 0 Then

            Cell.Value = i + 1

            i = i + 1
            j = j + 1
        Else

            Cell.Value = Arreglo(0, i - 1)

            j = j + 1
        End If

    Next Cell
    
    '/////////////////PRICES/////////////////
    
    Set Myrange = Range("C2:C15")
    
    For Each Cell In Myrange

        Cell.Value = Round((32 * Rnd) + 6.5, 2)       '////max     min
        
    Next Cell
    
    With Range("C1")
        .Value = "Unit Price USD"
        .Font.Bold = True
        .WrapText = True
    End With
    
    Columns("C:C").NumberFormat = "#,##0.00"
    
    '/////////////////PRICES/////////////////
    
    

    Range("A1").Value = "Order"
    Range("A1").Font.Bold = True

    Range("B1").Value = "Accesories"
    Range("B1").Font.Bold = True
    
    Columns("A:B").EntireColumn.AutoFit
    Columns("A:A").HorizontalAlignment = xlCenter


    Columns("C:C").NumberFormat = "#,##0.00"

    Range("A2").Select
    ActiveWindow.FreezePanes = True

End Sub


Sub Macro09_Fill_Appliances()


    Sheets("Appliances").Select

    Dim Myrange As Range
    Dim Cell As Range
    Dim i As Double
    Dim j As Double

    Dim Arreglo(0, 133) As String


    Arreglo(0, 0) = "Air conditioner"
    Arreglo(0, 1) = "Air ioniser"
    Arreglo(0, 2) = "Air purifier"
    Arreglo(0, 3) = "Appliance plug"
    Arreglo(0, 4) = "Aroma lamp"
    Arreglo(0, 5) = "Attic fan"
    Arreglo(0, 6) = "Bachelor griller"
    Arreglo(0, 7) = "Bedside lamp"
    Arreglo(0, 8) = "Back boiler"
    Arreglo(0, 9) = "Beverage opener"
    Arreglo(0, 10) = "Blender"
    Arreglo(0, 11) = "Box fan"
    Arreglo(0, 12) = "Blade Assembly"
    Arreglo(0, 13) = "Calculator"
    Arreglo(0, 14) = "Camcorder"
    Arreglo(0, 15) = "Can opener"
    Arreglo(0, 16) = "Cassette player"
    Arreglo(0, 17) = "Ceiling fan"
    Arreglo(0, 18) = "Central vacuum cleaner"
    Arreglo(0, 19) = "Grandfather clock"
    Arreglo(0, 20) = "Wall clock"
    Arreglo(0, 21) = "Clothes dryer"
    Arreglo(0, 22) = "Clothes iron"
    Arreglo(0, 23) = "Coffee grinder"
    Arreglo(0, 24) = "Coffeemaker"
    Arreglo(0, 25) = "Coffee percolator"
    Arreglo(0, 26) = "Cold-pressed juicer"
    Arreglo(0, 27) = "Cooler"
    Arreglo(0, 28) = "Combo washer dryer"
    Arreglo(0, 29) = "Communal oven"
    Arreglo(0, 30) = "Convection microwave"
    Arreglo(0, 31) = "Convection oven 1000W"
    Arreglo(0, 32) = "Corn roaster"
    Arreglo(0, 33) = "Corn butterer"
    Arreglo(0, 34) = "Crepe maker"
    Arreglo(0, 35) = "Crepe machine"
    Arreglo(0, 36) = "Convection oven 1500W"
    Arreglo(0, 37) = "Deep fryer"
    Arreglo(0, 38) = "Dehumidifier"
    Arreglo(0, 39) = "Digital camera"
    Arreglo(0, 40) = "Dish drying cabinet"
    Arreglo(0, 41) = "Dishwasher"
    Arreglo(0, 42) = "Drawer dishwasher"
    Arreglo(0, 43) = "DVD player"
    Arreglo(0, 44) = "Edger"
    Arreglo(0, 45) = "Electric cooker"
    Arreglo(0, 46) = "Electric razor"
    Arreglo(0, 47) = "Electric toothbrush"
    Arreglo(0, 48) = "Electric water boiler"
    Arreglo(0, 49) = "Evaporative cooler"
    Arreglo(0, 50) = "Exhaust hood"
    Arreglo(0, 51) = "Fan heater"
    Arreglo(0, 52) = "Desk fan"
    Arreglo(0, 53) = "Fire detector"
    Arreglo(0, 54) = "Food processor"
    Arreglo(0, 55) = "Forced-air"
    Arreglo(0, 56) = "Freezer"
    Arreglo(0, 57) = "Futon dryer"
    Arreglo(0, 58) = "Garbage disposal unit"
    Arreglo(0, 59) = "Gas Stove"
    Arreglo(0, 60) = "Gramaphone"
    Arreglo(0, 61) = "Gravy strainer"
    Arreglo(0, 62) = "Hair dryer"
    Arreglo(0, 63) = "Hair iron"
    Arreglo(0, 64) = "Hearing aid"
    Arreglo(0, 65) = "Hob (hearth)"
    Arreglo(0, 66) = "Home Cake server"
    Arreglo(0, 67) = "Humidifier (Vaporizer)"
    Arreglo(0, 68) = "HVAC"
    Arreglo(0, 69) = "Icebox"
    Arreglo(0, 70) = "Juicer"
    Arreglo(0, 71) = "Microphone Karaoke Set"
    Arreglo(0, 72) = "Disco ball Karaoke Set"
    Arreglo(0, 73) = "Kimchi refrigerator"
    Arreglo(0, 74) = "Lawn mower"
    Arreglo(0, 75) = "Riding mower"
    Arreglo(0, 76) = "Leaf blower"
    Arreglo(0, 77) = "Lighter"
    Arreglo(0, 78) = "Light fixture"
    Arreglo(0, 79) = "Melting Chocolate Fountain"
    Arreglo(0, 80) = "Meat grinder"
    Arreglo(0, 81) = "Megaphone"
    Arreglo(0, 82) = "Micathermic heater"
    Arreglo(0, 83) = "Microwave oven"
    Arreglo(0, 84) = "Mixer"
    Arreglo(0, 85) = "Mogul lamp"
    Arreglo(0, 86) = "Mousetrap"
    Arreglo(0, 87) = "Nightlight"
    Arreglo(0, 88) = "Oil heater"
    Arreglo(0, 89) = "Oven Cleaner"
    Arreglo(0, 90) = "Panini press"
    Arreglo(0, 91) = "Pasta maker"
    Arreglo(0, 92) = "Patio heater"
    Arreglo(0, 93) = "Paper shredder"
    Arreglo(0, 94) = "Pencil sharpener"
    Arreglo(0, 95) = "Popcorn maker"
    Arreglo(0, 96) = "Pressure-cooker"
    Arreglo(0, 97) = "Radiator (heating)"
    Arreglo(0, 98) = "Radio receiver"
    Arreglo(0, 99) = "Interior refrigerator"
    Arreglo(0, 100) = "Thermal mass refrigerator"
    Arreglo(0, 101) = "Rotisserie"
    Arreglo(0, 102) = "Sewing machine"
    Arreglo(0, 103) = "Kitchen sink"
    Arreglo(0, 104) = "Separate sink spray"
    Arreglo(0, 105) = "Slow cooker"
    Arreglo(0, 106) = "Snowblower"
    Arreglo(0, 107) = "Space heater"
    Arreglo(0, 108) = "Steam mop"
    Arreglo(0, 109) = "Stereo"
    Arreglo(0, 110) = "Stove"
    Arreglo(0, 111) = "Sump pump"
    Arreglo(0, 112) = "Digital Phone"
    Arreglo(0, 113) = "Table lamp"
    Arreglo(0, 114) = "Television set Reciver"
    Arreglo(0, 115) = "Television Remote Control"
    Arreglo(0, 116) = "Television set Speaker"
    Arreglo(0, 117) = "Tie Rack"
    Arreglo(0, 118) = "Toaster oven"
    Arreglo(0, 119) = "Toaster"
    Arreglo(0, 120) = "Trash compactor"
    Arreglo(0, 121) = "Trouser Rack"
    Arreglo(0, 122) = "Manual vacuum cleaner"
    Arreglo(0, 123) = "Robotic vacuum cleaner"
    Arreglo(0, 124) = "Videocassette recorder"
    Arreglo(0, 125) = "Waffle iron"
    Arreglo(0, 126) = "Washing machine"
    Arreglo(0, 127) = "Water cooker"
    Arreglo(0, 128) = "Waterpik"
    Arreglo(0, 129) = "Water purifier"
    Arreglo(0, 130) = "Solar water heater"
    Arreglo(0, 131) = "Tankless water heater"
    Arreglo(0, 132) = "Weed Eater"
    Arreglo(0, 133) = "Window fan"

    
    
    Set Myrange = Range("A2:B135")
    i = 0

    j = 0

    For Each Cell In Myrange

        If j Mod 2 = 0 Then

            Cell.Value = i + 1

            i = i + 1
            j = j + 1
        Else

            Cell.Value = Arreglo(0, i - 1)

            j = j + 1
        End If

    Next Cell
    
    
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
    

    Range("A1").Value = "Order"
    Range("A1").Font.Bold = True

    Range("B1").Value = "Appliances"
    Range("B1").Font.Bold = True

    Columns("A:B").EntireColumn.AutoFit
    Columns("A:A").HorizontalAlignment = xlCenter

    Range("A2").Select
    ActiveWindow.FreezePanes = True

End Sub
Sub Macro10_Fill_Binders()


    Sheets("Binders").Select

    Dim Myrange As Range
    Dim Cell As Range
    Dim i As Double
    Dim j As Double

    Dim Arreglo(0, 10) As String



    Arreglo(0, 0) = "Hard cover binders"
    Arreglo(0, 1) = "Soft cover binders"
    Arreglo(0, 2) = "Decorative binders"
    Arreglo(0, 3) = "School Binders"
    Arreglo(0, 4) = "blue binder, "
    Arreglo(0, 5) = "Pink binder "
    Arreglo(0, 6) = "Colorful binders"
    Arreglo(0, 7) = "Binders with arch mechanisms"
    Arreglo(0, 8) = "3-ring binders "
    Arreglo(0, 9) = "Presentation binder"
    Arreglo(0, 10) = "Binder with zipper"
    
    Set Myrange = Range("A2:B12")
    i = 0

    j = 0

    For Each Cell In Myrange

        If j Mod 2 = 0 Then

            Cell.Value = i + 1

            i = i + 1
            j = j + 1
        Else

            Cell.Value = Arreglo(0, i - 1)

            j = j + 1
        End If

    Next Cell
    
    
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
    

    Range("A1").Value = "Order"
    Range("A1").Font.Bold = True

    Range("B1").Value = "Binders"
    Range("B1").Font.Bold = True

    Columns("A:B").EntireColumn.AutoFit
    Columns("A:A").HorizontalAlignment = xlCenter

    Range("A2").Select
    ActiveWindow.FreezePanes = True

End Sub

Sub Macro11_Fill_Art()

    Sheets("Art").Select

    Dim Myrange As Range
    Dim Cell As Range
    Dim i As Double
    Dim j As Double

    Dim Arreglo(0, 138) As String
    
    Arreglo(0, 0) = "Jewelry Making-Beading Supplies"
    Arreglo(0, 1) = "Jewelry Making-Beads & Bead Assortments"
    Arreglo(0, 2) = "Jewelry Making-Charms"
    Arreglo(0, 3) = "Jewelry Making-Jewelry Casting Supplies"
    Arreglo(0, 4) = "Jewelry Making-Jewelry Findings"
    Arreglo(0, 5) = "Jewelry Making-Jewelry Making Kits"
    Arreglo(0, 6) = "Jewelry Making-Jewelry Making Tools & Accessories"
    Arreglo(0, 7) = "Jewelry Making-Metal Stamping Tools"
    Arreglo(0, 8) = "Jewelry Making-Purse Making"
    Arreglo(0, 9) = "Jewelry Making-Storage"
    Arreglo(0, 10) = "Jewelry Making-Wax Molding Materials"
    Arreglo(0, 11) = "Crafting-Basket Making"
    Arreglo(0, 12) = "Crafting-Candle Making"
    Arreglo(0, 13) = "Crafting-Ceramics & Pottery"
    Arreglo(0, 14) = "Crafting-Craft Supplies"
    Arreglo(0, 15) = "Crafting-Doll Making"
    Arreglo(0, 16) = "Crafting-Fabric Ribbons"
    Arreglo(0, 17) = "Crafting-Floral Arranging"
    Arreglo(0, 18) = "Crafting-Mosaic Making"
    Arreglo(0, 19) = "Crafting-Paper & Paper Crafts"
    Arreglo(0, 20) = "Crafting-Picture Framing"
    Arreglo(0, 21) = "Crafting-Scratchboards & Foil Engraving"
    Arreglo(0, 22) = "Crafting-Sculpture Supplies"
    Arreglo(0, 23) = "Crafting-Soap Making"
    Arreglo(0, 24) = "Crafting-Weaving & Spinning"
    Arreglo(0, 25) = "Crafting-Woodcrafts"
    Arreglo(0, 26) = "Decorate Fabric"
    Arreglo(0, 27) = "Fabric Decorating-Dyes"
    Arreglo(0, 28) = "Fabric Decorating-Fabric & Textile Paints"
    Arreglo(0, 29) = "Fabric Decorating-Fabric Decorating Kits"
    Arreglo(0, 30) = "Fabric Decorating-Fixatives"
    Arreglo(0, 31) = "Gift Wrapping-Gift Bags"
    Arreglo(0, 32) = "Gift Wrapping-Gift Boxes"
    Arreglo(0, 33) = "Gift Wrapping-Gift Wrap Cellophane"
    Arreglo(0, 34) = "Gift Wrapping-Gift Wrap Cellophane Bags"
    Arreglo(0, 35) = "Gift Wrapping-Gift Wrap Paper"
    Arreglo(0, 36) = "Gift Wrapping-Gift Wrap Ribbons"
    Arreglo(0, 37) = "Gift Wrapping-Gift Wrap Tags"
    Arreglo(0, 38) = "Gift Wrapping-Wrapping Tissue"
    Arreglo(0, 39) = "Crochet Hooks"
    Arreglo(0, 40) = "Crochet Kits"
    Arreglo(0, 41) = "Crochet Thread"
    Arreglo(0, 42) = "Knitting & Crochet Notions"
    Arreglo(0, 43) = "Knitting Kits"
    Arreglo(0, 44) = "Knitting Needles"
    Arreglo(0, 45) = "Knitting Patterns"
    Arreglo(0, 46) = "Knitting & Crochet-Yarn"
    Arreglo(0, 47) = "Knitting & Crochet-Yarn Storage"
    Arreglo(0, 48) = "Counted Cross Stitch"
    Arreglo(0, 49) = "Count aida Cross Stitch"
    Arreglo(0, 50) = "Sullivans Embroidery Floss Skein"
    Arreglo(0, 51) = "PRYM Quilting Hoop Wood"
    Arreglo(0, 52) = "Art & Poster Tubes"
    Arreglo(0, 53) = "Art Tool & Sketch Boxes"
    Arreglo(0, 54) = "Beading Storage"
    Arreglo(0, 55) = "Craft & Sewing Supplies"
    Arreglo(0, 56) = "Embroidery Storage"
    Arreglo(0, 57) = "Scrapbooking Storage"
    Arreglo(0, 58) = "Sewing Storage"
    Arreglo(0, 59) = "Yarn Storage"
    Arreglo(0, 60) = "Drying & Print Racks"
    Arreglo(0, 61) = "Paint Brush Organizers"
    Arreglo(0, 62) = "Paint Brush Holders"
    Arreglo(0, 63) = "Pen, Pencil & Marker Cases"
    Arreglo(0, 64) = "Portfolios"
    Arreglo(0, 65) = "Art Storage Boxes "
    Arreglo(0, 66) = "Art Storage Organizers"
    Arreglo(0, 67) = "Storage Cabinets"
    Arreglo(0, 68) = "Art Paper-Artist Trading Cards"
    Arreglo(0, 69) = "Art Paper-Bristol Paper & Vellum"
    Arreglo(0, 70) = "Art Paper-Drawing Paper"
    Arreglo(0, 71) = "Art Paper-Easel Pads"
    Arreglo(0, 72) = "Art Paper-Pastel Paper"
    Arreglo(0, 73) = "Art Paper-Sketchbooks & Notebooks"
    Arreglo(0, 74) = "Art Paper-Tracing Paper"
    Arreglo(0, 75) = "Art Paper-Watercolor Paper"
    Arreglo(0, 76) = "Canvas-Canvas Boards & Panels"
    Arreglo(0, 77) = "Canvas-Canvas Tools & Accessories"
    Arreglo(0, 78) = "Boards -Hardboard"
    Arreglo(0, 79) = "Boards -Pastelboard"
    Arreglo(0, 80) = "Canvas-Pre-Stretched Canvas"
    Arreglo(0, 81) = "Boards -Wood Art Boards"
    Arreglo(0, 82) = "Drawing-Art Sets"
    Arreglo(0, 83) = "Drawing & Lettering Aids"
    Arreglo(0, 84) = "Drawing Media"
    Arreglo(0, 85) = "Drawing-Erasers"
    Arreglo(0, 86) = "Drawing-Sharpeners"
    Arreglo(0, 87) = "Drawing-Easels"
    Arreglo(0, 88) = "Painting-Airbrush Materials"
    Arreglo(0, 89) = "Painting-Kits"
    Arreglo(0, 90) = "Paint Finishes"
    Arreglo(0, 91) = "Paint Mediums & Additives"
    Arreglo(0, 92) = "Paint Pens, Markers & Daubers"
    Arreglo(0, 93) = "Paint Sponges"
    Arreglo(0, 94) = "Paint-By-Number Kits"
    Arreglo(0, 95) = "Paintbrushes"
    Arreglo(0, 96) = "Paints"
    Arreglo(0, 97) = "Palette Knives"
    Arreglo(0, 98) = "Palettes & Palette Cups"
    Arreglo(0, 99) = "Party Decorations-Balloons"
    Arreglo(0, 100) = "Party Decorations-Banners & Garlands"
    Arreglo(0, 101) = "Party Decorations-Card Boxes"
    Arreglo(0, 102) = "Party Decorations-Cardboard Cutouts"
    Arreglo(0, 103) = "Party Decorations-Centerpieces"
    Arreglo(0, 104) = "Party Decorations-Confetti"
    Arreglo(0, 105) = "Party Decorations-Guestbooks"
    Arreglo(0, 106) = "Party Decorations-Luminarias"
    Arreglo(0, 107) = "Party Decorations-Streamers"
    Arreglo(0, 108) = "Party Decorations-Tablecovers"
    Arreglo(0, 109) = "Party Decorations-Tissue Pom Poms"
    Arreglo(0, 110) = "Printmaking-Etching Supplies"
    Arreglo(0, 111) = "Printmaking-Printing Presses & Accessories"
    Arreglo(0, 112) = "Printmaking-Printmaking Inks"
    Arreglo(0, 113) = "Printmaking-Relief & Block Printing Materials"
    Arreglo(0, 114) = "Printmaking-Screen Printing"
    Arreglo(0, 115) = "Stamping-Adhesive Vinyl"
    Arreglo(0, 116) = "Stamping-Adhesives"
    Arreglo(0, 117) = "Stamping-Albums & Refills"
    Arreglo(0, 118) = "Stamping-Chipboard"
    Arreglo(0, 119) = "Stamping-Cutting Mats"
    Arreglo(0, 120) = "Stamping-Die-Cutting & Embossing"
    Arreglo(0, 121) = "Stamping-Kits"
    Arreglo(0, 122) = "Stamping-Paper & Card Stock"
    Arreglo(0, 123) = "Stamping-Paper Punches"
    Arreglo(0, 124) = "Stamping-Pens & Markers"
    Arreglo(0, 125) = "Scrapbooking Embellishments"
    Arreglo(0, 126) = "Scrapbooking Tools"
    Arreglo(0, 127) = "Scrapbooking -Stamps & Ink Pads"
    Arreglo(0, 128) = "Scrapbooking -Stickers & Sticker Machines"
    Arreglo(0, 129) = "Scrapbooking -Stencils & Templates"
    Arreglo(0, 130) = "Sewing-Quilting"
    Arreglo(0, 131) = "Sewing Machine Parts & Accessories"
    Arreglo(0, 132) = "Sewing Machines"
    Arreglo(0, 133) = "Sewing Notions & Supplies"
    Arreglo(0, 134) = "Sewing Patterns & Templates"
    Arreglo(0, 135) = "Sewing Project Kits"
    Arreglo(0, 136) = "Sewing-Storage & Furniture"
    Arreglo(0, 137) = "Sewing-Thread & Floss"
    Arreglo(0, 138) = "Sewing-Trim & Embellishments"
      
    Set Myrange = Range("A2:B140")
    i = 0

    j = 0

    For Each Cell In Myrange

        If j Mod 2 = 0 Then

            Cell.Value = i + 1

            i = i + 1
            j = j + 1
        Else

            Cell.Value = Arreglo(0, i - 1)

            j = j + 1
        End If

    Next Cell
    
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

    Range("A1").Value = "Order"
    Range("A1").Font.Bold = True

    Range("B1").Value = "Art"
    Range("B1").Font.Bold = True

    Columns("A:B").EntireColumn.AutoFit
    Columns("A:A").HorizontalAlignment = xlCenter

    Range("A2").Select
    ActiveWindow.FreezePanes = True

End Sub
Sub Macro12_Fill_Bookcases()

    Sheets("Bookcases").Select

    Dim Myrange As Range
    Dim Cell As Range
    Dim i As Double
    Dim j As Double

    Dim Arreglo(0, 30) As String
    
    Arreglo(0, 0) = "BESTÅ"
    Arreglo(0, 1) = "BILLY"
    Arreglo(0, 2) = "BILLY / BOTTNA"
    Arreglo(0, 3) = "BILLY / GNEDBY"
    Arreglo(0, 4) = "BILLY / OXBERG"
    Arreglo(0, 5) = "BRIMNES"
    Arreglo(0, 6) = "BRUSALI"
    Arreglo(0, 7) = "EDVALLA"
    Arreglo(0, 8) = "EKET"
    Arreglo(0, 9) = "ENERYDA"
    Arreglo(0, 10) = "GALANT"
    Arreglo(0, 11) = "GNEDBY"
    Arreglo(0, 12) = "GUBBARP"
    Arreglo(0, 13) = "HACKÅS"
    Arreglo(0, 14) = "HANVIKEN"
    Arreglo(0, 15) = "HAVSTA"
    Arreglo(0, 16) = "HEMNES"
    Arreglo(0, 17) = "KALLAX"
    Arreglo(0, 18) = "KLACKBERG"
    Arreglo(0, 19) = "LAXVIKEN"
    Arreglo(0, 20) = "LIATORP"
    Arreglo(0, 21) = "LOMMARP"
    Arreglo(0, 22) = "MÖLLARP"
    Arreglo(0, 23) = "MOSSARYD"
    Arreglo(0, 24) = "NANNARP"
    Arreglo(0, 25) = "NOTVIKEN"
    Arreglo(0, 26) = "ÖSTERNÄS"
    Arreglo(0, 27) = "OXBERG"
    Arreglo(0, 28) = "PS 2017"
    Arreglo(0, 29) = "RIKSVIKEN"
    Arreglo(0, 30) = "VASSVIKEN"

    Set Myrange = Range("A2:B32")
    i = 0

    j = 0

    For Each Cell In Myrange

        If j Mod 2 = 0 Then

            Cell.Value = i + 1

            i = i + 1
            j = j + 1
        Else

            Cell.Value = Arreglo(0, i - 1)

            j = j + 1
        End If

    Next Cell
    
    '/////////////////PRICES/////////////////
    
    Set Myrange = Range("C2:C32")
    
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

    Range("A1").Value = "Order"
    Range("A1").Font.Bold = True

    Range("B1").Value = "Bookcases"
    Range("B1").Font.Bold = True

    Columns("A:B").EntireColumn.AutoFit
    Columns("A:A").HorizontalAlignment = xlCenter

    Range("A2").Select
    ActiveWindow.FreezePanes = True


End Sub

Sub Macro13_Fill_Chairs()

    Sheets("Chairs").Select

    Dim Myrange As Range
    Dim Cell As Range
    Dim i As Double
    Dim j As Double

    Dim Arreglo(0, 157) As String

    Arreglo(0, 0) = "Aalto armchair"
    Arreglo(0, 1) = "Adirondack chair Low"
    Arreglo(0, 2) = "Aalto armchair 406"
    Arreglo(0, 3) = "Adirondack chair High"
    Arreglo(0, 4) = "Aeron chair"
    Arreglo(0, 5) = "Air chair"
    Arreglo(0, 6) = "Armchair"
    Arreglo(0, 7) = "Bachelor's chair"
    Arreglo(0, 8) = "Balans chair"
    Arreglo(0, 9) = "Ball Chair"
    Arreglo(0, 10) = "Bar stool"
    Arreglo(0, 11) = "Barcelona chair"
    Arreglo(0, 12) = "Bardic chair"
    Arreglo(0, 13) = "Barrel chair"
    Arreglo(0, 14) = "Bath chair"
    Arreglo(0, 15) = "Beach chair (Strandkorb)"
    Arreglo(0, 16) = "Bean bag chair"
    Arreglo(0, 17) = "Bench chair"
    Arreglo(0, 18) = "Bergère chair"
    Arreglo(0, 19) = "Bikini chair"
    Arreglo(0, 20) = "Bofinger chair"
    Arreglo(0, 21) = "Bosun's chair"
    Arreglo(0, 22) = "Breuer Chair"
    Arreglo(0, 23) = "Brewster Chair"
    Arreglo(0, 24) = "Bubble Chair"
    Arreglo(0, 25) = "Bungee chair"
    Arreglo(0, 26) = "Butterfly chair"
    Arreglo(0, 27) = "Campeche chair"
    Arreglo(0, 28) = "Cantilever chair"
    Arreglo(0, 29) = "Captain's chair"
    Arreglo(0, 30) = "Caquetoire Chair"
    Arreglo(0, 31) = "Car chair"
    Arreglo(0, 32) = "Carver chair"
    Arreglo(0, 33) = "Cathedra"
    Arreglo(0, 34) = "Chaise a bureau Chair"
    Arreglo(0, 35) = "Chaise longue"
    Arreglo(0, 36) = "Chesterfield chair"
    Arreglo(0, 37) = "Chiavari chair"
    Arreglo(0, 38) = "Club chair"
    Arreglo(0, 39) = "Cogswell chair"
    Arreglo(0, 40) = "Corner chair"
    Arreglo(0, 41) = "Curule chair"
    Arreglo(0, 42) = "Dante chair"
    Arreglo(0, 43) = "Deckchair"
    Arreglo(0, 44) = "Dentist chair"
    Arreglo(0, 45) = "Dining chair"
    Arreglo(0, 46) = "Director's chair"
    Arreglo(0, 47) = "Easy chair"
    Arreglo(0, 48) = "Eames Lounge Chair"
    Arreglo(0, 49) = "Egg chair"
    Arreglo(0, 50) = "Electric chair"
    Arreglo(0, 51) = "Elijah's chair"
    Arreglo(0, 52) = "Emeco 1006"
    Arreglo(0, 53) = "Farthingale chair"
    Arreglo(0, 54) = "Fauteuil Chair"
    Arreglo(0, 55) = "Fiddleback Chair"
    Arreglo(0, 56) = "Fighting Chair"
    Arreglo(0, 57) = "Folding Chair"
    Arreglo(0, 58) = "Folding seat"
    Arreglo(0, 59) = "Friendship bench chair"
    Arreglo(0, 60) = "Gaming chair"
    Arreglo(0, 61) = "Garden Egg chair"
    Arreglo(0, 62) = "Glastonbury chair"
    Arreglo(0, 63) = "Glider (or platform rocker)"
    Arreglo(0, 64) = "Hassock Chair"
    Arreglo(0, 65) = "High Chair"
    Arreglo(0, 66) = "Hanging Egg Chair"
    Arreglo(0, 67) = "Inflatable chair"
    Arreglo(0, 68) = "Ironing chair"
    Arreglo(0, 69) = "Jack and Jill chair"
    Arreglo(0, 70) = "Jump seat"
    Arreglo(0, 71) = "Kneeling chairs "
    Arreglo(0, 72) = "Knotted chair"
    Arreglo(0, 73) = "Ladderback chair"
    Arreglo(0, 74) = "Lambing chair"
    Arreglo(0, 75) = "Lawn chair"
    Arreglo(0, 76) = "Lifeguard chairs"
    Arreglo(0, 77) = "Lift chair"
    Arreglo(0, 78) = "Litter sedan chair"
    Arreglo(0, 79) = "Louis Ghost chair"
    Arreglo(0, 80) = "Massage chair"
    Arreglo(0, 81) = "Monobloc chair"
    Arreglo(0, 82) = "Morris chair"
    Arreglo(0, 83) = "Muskoka chair"
    Arreglo(0, 84) = "Navy chair"
    Arreglo(0, 85) = "No. 14 chair"
    Arreglo(0, 86) = "Nursing chair"
    Arreglo(0, 87) = "Office chair"
    Arreglo(0, 88) = "Orbiter seat "
    Arreglo(0, 89) = "ON Chair"
    Arreglo(0, 90) = "Ottoman Chair"
    Arreglo(0, 91) = "Ovalia Egg Chair"
    Arreglo(0, 92) = "Onit chair"
    Arreglo(0, 93) = "Panton Chair"
    Arreglo(0, 94) = "Papasan chair"
    Arreglo(0, 95) = "Parsons chair"
    Arreglo(0, 96) = "Patio chair"
    Arreglo(0, 97) = "Peacock chair"
    Arreglo(0, 98) = "Pew Chair"
    Arreglo(0, 99) = "Pew stacker chair"
    Arreglo(0, 100) = "Planter's chair"
    Arreglo(0, 101) = "Poäng armchair "
    Arreglo(0, 102) = "Poofbag chair"
    Arreglo(0, 103) = "Porter's chair"
    Arreglo(0, 104) = "Potty chair"
    Arreglo(0, 105) = "Pouffe seat"
    Arreglo(0, 106) = "Power chairs"
    Arreglo(0, 107) = "Pressback chair"
    Arreglo(0, 108) = "Pushchair or stroller"
    Arreglo(0, 109) = "Recliner chair "
    Arreglo(0, 110) = "Restraint chair"
    Arreglo(0, 111) = "Revolving chair"
    Arreglo(0, 112) = "Rex chair"
    Arreglo(0, 113) = "Ribbon Chair"
    Arreglo(0, 114) = "Rocking chair"
    Arreglo(0, 115) = "Rumble seat"
    Arreglo(0, 116) = "Saddle chair"
    Arreglo(0, 117) = "Savonarola chair"
    Arreglo(0, 118) = "Sedan chair"
    Arreglo(0, 119) = "Sgabello seat"
    Arreglo(0, 120) = "Shaker rocker rocking chair"
    Arreglo(0, 121) = "Shaker tilting chair"
    Arreglo(0, 122) = "Shower chair"
    Arreglo(0, 123) = "Side chair"
    Arreglo(0, 124) = "Sit-stand chair"
    Arreglo(0, 125) = "Sling chair"
    Arreglo(0, 126) = "Slumber chair"
    Arreglo(0, 127) = "Spinning chair"
    Arreglo(0, 128) = "Stacking chair"
    Arreglo(0, 129) = "Steno chair"
    Arreglo(0, 130) = "Step chair"
    Arreglo(0, 131) = "Stool a chair "
    Arreglo(0, 132) = "Sweetheart chair"
    Arreglo(0, 133) = "Swivel chairs"
    Arreglo(0, 134) = "Tarachair"
    Arreglo(0, 135) = "Tête-à-tête chair"
    Arreglo(0, 136) = "Throne ceremonial chair"
    Arreglo(0, 137) = "Toilet chair"
    Arreglo(0, 138) = "Tuffet"
    Arreglo(0, 139) = "Tulip chair"
    Arreglo(0, 140) = "Turned chair"
    Arreglo(0, 141) = "Two-slat post-and-rung shaving chair"
    Arreglo(0, 142) = "UP5 chair"
    Arreglo(0, 143) = "Visitor's chair"
    Arreglo(0, 144) = "Voyeuse chair"
    Arreglo(0, 145) = "Wainscot Chair"
    Arreglo(0, 146) = "Watchman's chair"
    Arreglo(0, 147) = "Wassily Chair"
    Arreglo(0, 148) = "Wheelchair"
    Arreglo(0, 149) = "Wicker chair"
    Arreglo(0, 150) = "Wiggle chair"
    Arreglo(0, 151) = "Windsor chair"
    Arreglo(0, 152) = "Wing chair"
    Arreglo(0, 153) = "Writing armchair"
    Arreglo(0, 154) = "X-chair"
    Arreglo(0, 155) = "Zaisu legless chair"
    Arreglo(0, 156) = "Zig-Zag Chair"
    Arreglo(0, 157) = "Zero-gravity chairs"

    Set Myrange = Range("A2:B159")
    i = 0

    j = 0

    For Each Cell In Myrange

        If j Mod 2 = 0 Then

            Cell.Value = i + 1

            i = i + 1
            j = j + 1
        Else

            Cell.Value = Arreglo(0, i - 1)

            j = j + 1
        End If

    Next Cell
    
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
    

    Range("A1").Value = "Order"
    Range("A1").Font.Bold = True

    Range("B1").Value = "Chairs"
    Range("B1").Font.Bold = True

    Columns("A:B").EntireColumn.AutoFit
    Columns("A:A").HorizontalAlignment = xlCenter

    Range("A2").Select
    ActiveWindow.FreezePanes = True


End Sub

Sub Macro14_Fill_Copiers()

    Sheets("Copiers").Select

    Dim Myrange As Range
    Dim Cell As Range
    Dim i As Double
    Dim j As Double

    Dim Arreglo(0, 119) As String

    Arreglo(0, 0) = "A. B. Dick Copier"
    Arreglo(0, 1) = "Advanced Matrix Technology Copier"
    Arreglo(0, 2) = "ALPS Copier"
    Arreglo(0, 3) = "AMT Datasouth Copier"
    Arreglo(0, 4) = "Avery Dennison Copier"
    Arreglo(0, 5) = "Apple Copier"
    Arreglo(0, 6) = "ASK Technology Copier"
    Arreglo(0, 7) = "Axonix Copier"
    Arreglo(0, 8) = "Bell-Mark Copier"
    Arreglo(0, 9) = "Benson, Inc. Copier"
    Arreglo(0, 10) = "Brother Copier"
    Arreglo(0, 11) = "Bull Copier"
    Arreglo(0, 12) = "Canon Copier"
    Arreglo(0, 13) = "Centronics Copier"
    Arreglo(0, 14) = "Checkpoint Meto Copier"
    Arreglo(0, 15) = "Citizen Copier"
    Arreglo(0, 16) = "Codimag Copier"
    Arreglo(0, 17) = "Cognitive Copier"
    Arreglo(0, 18) = "Compuprint Copier"
    Arreglo(0, 19) = "Computer Peripherals Inc Copier"
    Arreglo(0, 20) = "Comtec Copier"
    Arreglo(0, 21) = "Compress UV Printers Copier"
    Arreglo(0, 22) = "Copal Copier"
    Arreglo(0, 23) = "Control Data Corporation Copier"
    Arreglo(0, 24) = "DASCOM Copier"
    Arreglo(0, 25) = "Datamax-O'Neil Copier"
    Arreglo(0, 26) = "Dataproducts Copier"
    Arreglo(0, 27) = "Datasouth Copier"
    Arreglo(0, 28) = "Decision Data Copier"
    Arreglo(0, 29) = "Delphax Technologies inc Copier"
    Arreglo(0, 30) = "Diablo Copier"
    Arreglo(0, 31) = "Digital Equipment Corporation Copier"
    Arreglo(0, 32) = "Dell Copier"
    Arreglo(0, 33) = "Eastman Kodak Copier"
    Arreglo(0, 34) = "Eltron Copier"
    Arreglo(0, 35) = "Epson Copier"
    Arreglo(0, 36) = "Everex (Abaton div.) Copier"
    Arreglo(0, 37) = "Facit Copier"
    Arreglo(0, 38) = "Fargo Copier"
    Arreglo(0, 39) = "Fujifilm Copier"
    Arreglo(0, 40) = "Fujitsu Copier"
    Arreglo(0, 41) = "Fuji Xerox Copier"
    Arreglo(0, 42) = "GENICOM Copier"
    Arreglo(0, 43) = "GCC Printers Copier"
    Arreglo(0, 44) = "General Electric Copier"
    Arreglo(0, 45) = "Hitachi Copier"
    Arreglo(0, 46) = "Heidelberg Copier"
    Arreglo(0, 47) = "Hewlett-Packard Copier"
    Arreglo(0, 48) = "Imprint Digital Copier"
    Arreglo(0, 49) = "IBM Copier"
    Arreglo(0, 50) = "InfoPrint Copier"
    Arreglo(0, 51) = "Juki Copier"
    Arreglo(0, 52) = "Kentek Copier"
    Arreglo(0, 53) = "Kodak Copier"
    Arreglo(0, 54) = "Konica Copier"
    Arreglo(0, 55) = "Konica Minolta Copier"
    Arreglo(0, 56) = "Kyocera Mita Copier"
    Arreglo(0, 57) = "Lake Erie Systems Copier"
    Arreglo(0, 58) = "Lanier Copier"
    Arreglo(0, 59) = "Lenovo Copier"
    Arreglo(0, 60) = "Lexmark Copier"
    Arreglo(0, 61) = "LiPi Data sys. Copier"
    Arreglo(0, 62) = "Mannesmann Tally Copier"
    Arreglo(0, 63) = "MapleJet Copier"
    Arreglo(0, 64) = "Minolta Copier"
    Arreglo(0, 65) = "Minolta-QMS Copier"
    Arreglo(0, 66) = "Memorex Telex Copier"
    Arreglo(0, 67) = "Microcom Corporation Copier"
    Arreglo(0, 68) = "MTX Copier"
    Arreglo(0, 69) = "Nakajima Copier"
    Arreglo(0, 70) = "NEC Copier"
    Arreglo(0, 71) = "Nidec Copal Copier"
    Arreglo(0, 72) = "Nipson Copier"
    Arreglo(0, 73) = "Océ Copier"
    Arreglo(0, 74) = "Oki Data Copier"
    Arreglo(0, 75) = "Olivetti Copier"
    Arreglo(0, 76) = "Output Technology Copier"
    Arreglo(0, 77) = "Office Automation Systems Inc (OASYS) Copier"
    Arreglo(0, 78) = "Panasonic Copier"
    Arreglo(0, 79) = "Pentax Copier"
    Arreglo(0, 80) = "Printer System Corporation Copier"
    Arreglo(0, 81) = "Printek Copier"
    Arreglo(0, 82) = "Printer Systems International Copier"
    Arreglo(0, 83) = "Printronix Copier"
    Arreglo(0, 84) = "PSI Engineering Copier"
    Arreglo(0, 85) = "Prototype & Production Systems, Inc Copier"
    Arreglo(0, 86) = "QMS Copier"
    Arreglo(0, 87) = "Qume Copier"
    Arreglo(0, 88) = "Rank Xerox Copier"
    Arreglo(0, 89) = "Ricoh Copier"
    Arreglo(0, 90) = "Riso Kagaku Corporation Copier"
    Arreglo(0, 91) = "RJS Copier"
    Arreglo(0, 92) = "Samsung Copier"
    Arreglo(0, 93) = "Sato Copier"
    Arreglo(0, 94) = "Seiko Copier"
    Arreglo(0, 95) = "Seiko Epson Copier"
    Arreglo(0, 96) = "Sewoo Copier"
    Arreglo(0, 97) = "Sharp Copier"
    Arreglo(0, 98) = "Siemens Nixdorf Copier"
    Arreglo(0, 99) = "Source Technologies Copier"
    Arreglo(0, 100) = "Swecoin Copier"
    Arreglo(0, 101) = "Syscan Copier"
    Arreglo(0, 102) = "Star Copier"
    Arreglo(0, 103) = "Star Micronics Copier"
    Arreglo(0, 104) = "Tally Copier"
    Arreglo(0, 105) = "TallyGenicom Copier"
    Arreglo(0, 106) = "TEC Copier"
    Arreglo(0, 107) = "Tektronix Copier"
    Arreglo(0, 108) = "Teletype Copier"
    Arreglo(0, 109) = "Texas Instruments Copier"
    Arreglo(0, 110) = "Toshiba Copier"
    Arreglo(0, 111) = "Trilog Copier"
    Arreglo(0, 112) = "TVS Electronics Copier"
    Arreglo(0, 113) = "UBIX Corp. Copier"
    Arreglo(0, 114) = "Versatec Copier"
    Arreglo(0, 115) = "Xeikon Copier"
    Arreglo(0, 116) = "Xerox Copier"
    Arreglo(0, 117) = "Xerox International Partners Copier"
    Arreglo(0, 118) = "Wipro Technologies  Copier"
    Arreglo(0, 119) = "Zebra Copier"


    
    Set Myrange = Range("A2:B121")
    i = 0

    j = 0

    For Each Cell In Myrange

        If j Mod 2 = 0 Then

            Cell.Value = i + 1

            i = i + 1
            j = j + 1
        Else

            Cell.Value = Arreglo(0, i - 1)

            j = j + 1
        End If

    Next Cell
    
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

    Range("A1").Value = "Order"
    Range("A1").Font.Bold = True

    Range("B1").Value = "Copiers"
    Range("B1").Font.Bold = True

    Columns("A:B").EntireColumn.AutoFit
    Columns("A:A").HorizontalAlignment = xlCenter

    Range("A2").Select
    ActiveWindow.FreezePanes = True

   

End Sub
Sub Macro15_Fill_Envelopes()


    Sheets("Envelopes").Select

    Dim Myrange As Range
    Dim Cell As Range
    Dim i As Double
    Dim j As Double

    Dim Arreglo(0, 29) As String
    
    Arreglo(0, 0) = "4 Baronial Envelope - 92 x 130 mm"
    Arreglo(0, 1) = "5 1/2 Baronial Envelope - 111 x 146 mm"
    Arreglo(0, 2) = "6 Baronial Envelope - 212 x 165 mm"
    Arreglo(0, 3) = "Lee Envelope - 133 x 184 mm"
    Arreglo(0, 4) = "A2 Envelope - 111 x 146 mm"
    Arreglo(0, 5) = "A6 Envelope - 121 x 165 mm"
    Arreglo(0, 6) = "A7 Envelope - 133 x 184 mm"
    Arreglo(0, 7) = "A8 Envelope - 410 x 206 mm"
    Arreglo(0, 8) = "A10 Envelope - 152 x 241 mm"
    Arreglo(0, 9) = "Slimline Envelope - 98 x 225 mm"
    Arreglo(0, 10) = "5 Square Envelope - 127 x 127 mm"
    Arreglo(0, 11) = "5 1/2 Square Envelope - 140 x 140 mm"
    Arreglo(0, 12) = "6 Square Envelope - 152 x 152 mm"
    Arreglo(0, 13) = "6 1/2 Square Envelope - 165 x 165 mm"
    Arreglo(0, 14) = "7 Square Envelope - 178 x 178 mm"
    Arreglo(0, 15) = "7 1/2 Square Envelope - 190 x 190 mm"
    Arreglo(0, 16) = "8 Square Envelope - 203 x 203 mm"
    Arreglo(0, 17) = "8 1/2 Square Envelope - 216 x 216 mm"
    Arreglo(0, 18) = "6.75 Envelope - 165 x 92 mm"
    Arreglo(0, 19) = "Monarch Envelope - 190 x 98 mm"
    Arreglo(0, 20) = "No. 9 Commercial Envelope - 225 x 98 mm"
    Arreglo(0, 21) = "No. 10 Commercial Envelope - 241 x 105 mm"
    Arreglo(0, 22) = "No. 10 Square Envelope - 241 x 105 mm"
    Arreglo(0, 23) = "No. 10 Peel ? Seal Envelope - 241 x 105 mm"
    Arreglo(0, 24) = "No. 10 Commercial (Standard Poly Window) Envelope - 241 x 105 mm"
    Arreglo(0, 25) = "No. 10 Policy Envelope - 241 x 105 mm"
    Arreglo(0, 26) = "No. 10 Commercial (Canadian Window) Envelope - 241 x 105 mm"
    Arreglo(0, 27) = "DL Envelope - 220 x 110 mm"
    Arreglo(0, 28) = "9 x 12 Envelope - 229 x 305 mm"
    Arreglo(0, 29) = "10 x 13 Envelope - 254 x 330 mm"

    Set Myrange = Range("A2:B31")
    i = 0

    j = 0

    For Each Cell In Myrange

        If j Mod 2 = 0 Then

            Cell.Value = i + 1

            i = i + 1
            j = j + 1
        Else

            Cell.Value = Arreglo(0, i - 1)

            j = j + 1
        End If

    Next Cell
    
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

    Range("A1").Value = "Order"
    Range("A1").Font.Bold = True

    Range("B1").Value = "Envelopes"
    Range("B1").Font.Bold = True

    Columns("A:B").EntireColumn.AutoFit
    Columns("A:A").HorizontalAlignment = xlCenter

    Range("A2").Select
    ActiveWindow.FreezePanes = True


End Sub
Sub Macro16_Fill_Fasteners()

    Sheets("Fasteners").Select

    Dim Myrange As Range
    Dim Cell As Range
    Dim i As Double
    Dim j As Double

    Dim Arreglo(0, 81) As String
    
    Arreglo(0, 0) = "Carriage Bolts"
    Arreglo(0, 1) = "Hex Head Bolts"
    Arreglo(0, 2) = "Machine Screws"
    Arreglo(0, 3) = "Shoulder Bolts"
    Arreglo(0, 4) = "Socket Cap Screws"
    Arreglo(0, 5) = "Socket Set 'Grub' Screws"
    Arreglo(0, 6) = "Square Head Bolts"
    Arreglo(0, 7) = "Deck Screws"
    Arreglo(0, 8) = "Hex Lag Screws"
    Arreglo(0, 9) = "Self-Drilling Screws"
    Arreglo(0, 10) = "Sheet Metal Screws"
    Arreglo(0, 11) = "Wood Screws"
    Arreglo(0, 12) = "Cap Nuts"
    Arreglo(0, 13) = "Castle Nuts"
    Arreglo(0, 14) = "Coupling Nuts"
    Arreglo(0, 15) = "Flange Serrated Nuts"
    Arreglo(0, 16) = "Hex Finish Nuts"
    Arreglo(0, 17) = "Hex Jam Nuts"
    Arreglo(0, 18) = "Heavy Hex Nuts"
    Arreglo(0, 19) = "Hex Machine Nuts"
    Arreglo(0, 20) = "Hex Machine Nuts Small Pattern"
    Arreglo(0, 21) = "Keps-K Lock Nuts"
    Arreglo(0, 22) = "Knurled Thumb Nuts"
    Arreglo(0, 23) = "Nylon Hex Jam Nuts"
    Arreglo(0, 24) = "Nylon Insert Lock Nuts"
    Arreglo(0, 25) = "Prevailing Torque Lock Nuts (Stover)"
    Arreglo(0, 26) = "Slotted Hex Nuts"
    Arreglo(0, 27) = "Square Nuts"
    Arreglo(0, 28) = "Structural Heavy Hex Nuts"
    Arreglo(0, 29) = "T-Nuts"
    Arreglo(0, 30) = "Break Away or Shear Nuts"
    Arreglo(0, 31) = "Tri-Groove Nuts"
    Arreglo(0, 32) = "Wing Nuts"
    Arreglo(0, 33) = "Backup Rivet Washers"
    Arreglo(0, 34) = "Belleville Conical Washers"
    Arreglo(0, 35) = "Dock Washers"
    Arreglo(0, 36) = "Fender Washers"
    Arreglo(0, 37) = "Fender Washers - Extra Thick"
    Arreglo(0, 38) = "Finishing Cup Washers"
    Arreglo(0, 39) = "Flat Washers"
    Arreglo(0, 40) = "Flat Washers - Extra Thick"
    Arreglo(0, 41) = "Flat Washers - Military Standard"
    Arreglo(0, 42) = "Flat Washers - 900 Series"
    Arreglo(0, 43) = "Lock Washers - Split Ring"
    Arreglo(0, 44) = "Lock Washers - High Collar"
    Arreglo(0, 45) = "Lock Washers - External Tooth"
    Arreglo(0, 46) = "Lock Washers - Internal Tooth"
    Arreglo(0, 47) = "NAS Washers"
    Arreglo(0, 48) = "Neoprene EPDM Washers"
    Arreglo(0, 49) = "Structural Washers"
    Arreglo(0, 50) = "Square Washers"
    Arreglo(0, 51) = "POP Rivets (Open End)"
    Arreglo(0, 52) = "Closed End POP Rivets (Sealed)"
    Arreglo(0, 53) = "Large Flange POP Rivets"
    Arreglo(0, 54) = "Countersunk POP Rivets"
    Arreglo(0, 55) = "Colored Rivets"
    Arreglo(0, 56) = "Multi-Grip Rivets"
    Arreglo(0, 57) = "Structural Rivets"
    Arreglo(0, 58) = "Tri-Fold Rivets"
    Arreglo(0, 59) = "Acoustical Wedge Anchors"
    Arreglo(0, 60) = "Drop In Anchors"
    Arreglo(0, 61) = "Double Expansion Shield Anchors"
    Arreglo(0, 62) = "Hammer Drive Pin Anchors"
    Arreglo(0, 63) = "Kaptoggle Hollow Wall Anchors"
    Arreglo(0, 64) = "Lag Shield Expansion Anchors"
    Arreglo(0, 65) = "Machine Screw Anchors"
    Arreglo(0, 66) = "Masonry Screws"
    Arreglo(0, 67) = "Plastic Toggle Anchors"
    Arreglo(0, 68) = "Sammys Screws"
    Arreglo(0, 69) = "Sleeve Anchors"
    Arreglo(0, 70) = "Toggle Wing Hollow Wall Anchors"
    Arreglo(0, 71) = "Wedge Anchors"
    Arreglo(0, 72) = "Dowel Pins"
    Arreglo(0, 73) = "Helicoil Threaded Inserts"
    Arreglo(0, 74) = "E-Z Lok Threaded Inserts"
    Arreglo(0, 75) = "Keystock"
    Arreglo(0, 76) = "Threaded Rod"
    Arreglo(0, 77) = "Unthreaded Rod"
    Arreglo(0, 78) = "Bowed-E Retaining Rings"
    Arreglo(0, 79) = "E-Style Retaining Rings"
    Arreglo(0, 80) = "External Shaft Retaining Rings"
    Arreglo(0, 81) = "Internal Housing Retaining Rings"

    Set Myrange = Range("A2:B83")
    i = 0

    j = 0

    For Each Cell In Myrange

        If j Mod 2 = 0 Then

            Cell.Value = i + 1

            i = i + 1
            j = j + 1
        Else

            Cell.Value = Arreglo(0, i - 1)

            j = j + 1
        End If

    Next Cell
    
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

    Range("A1").Value = "Order"
    Range("A1").Font.Bold = True

    Range("B1").Value = "Fasteners"
    Range("B1").Font.Bold = True

    Columns("A:B").EntireColumn.AutoFit
    Columns("A:A").HorizontalAlignment = xlCenter

    Range("A2").Select
    ActiveWindow.FreezePanes = True




End Sub

Sub Macro17_Fill_Furnishings()

    Sheets("Furnishings").Select

    Dim Myrange As Range
    Dim Cell As Range
    Dim i As Double
    Dim j As Double

    Dim Arreglo(0, 157) As String
    
    Arreglo(0, 0) = "Pinch pleat (or tailored pleat) curtains"
    Arreglo(0, 1) = "Box pleat curtains"
    Arreglo(0, 2) = "Goblet pleat curtains"
    Arreglo(0, 3) = "Pencil pleat curtains"
    Arreglo(0, 4) = "Eyelet (grommet) curtains"
    Arreglo(0, 5) = "Rod-pocket curtains"
    Arreglo(0, 6) = "Tab-top curtains"
    Arreglo(0, 7) = "Chair Cushion"
    Arreglo(0, 8) = "Bench Cushion"
    Arreglo(0, 9) = "Chaise Cushion"
    Arreglo(0, 10) = "Rocking Chair Seat Cushion"
    Arreglo(0, 11) = "Window Seat Cushion"
    Arreglo(0, 12) = "Wicker Cushion"
    Arreglo(0, 13) = "Deep Seating Cushion"
    Arreglo(0, 14) = "Church Pew Cushion"
    Arreglo(0, 15) = "Boxed-Edge Cushion"
    Arreglo(0, 16) = "Knife-Edge Cushion"
    Arreglo(0, 17) = "Single-Welted Cushion"
    Arreglo(0, 18) = "Double-Welted Cushion"
    Arreglo(0, 19) = "Cushion Ties"
    Arreglo(0, 20) = "Throw Pillows Cushion"
    Arreglo(0, 21) = "BED SKIRT"
    Arreglo(0, 22) = "DUST RUFFLE"
    Arreglo(0, 23) = "MATTRESS"
    Arreglo(0, 24) = "MATTRESS PROTECTOR"
    Arreglo(0, 25) = "MATTRESS PAD"
    Arreglo(0, 26) = "TOPPER"
    Arreglo(0, 27) = "BOTTOM SHEET"
    Arreglo(0, 28) = "FITTED SHEET"
    Arreglo(0, 29) = "FLAT SHEET"
    Arreglo(0, 30) = "TOP SHEET "
    Arreglo(0, 31) = "DUVET COVER"
    Arreglo(0, 32) = "CONFORTER"
    Arreglo(0, 33) = "BLANKET"
    Arreglo(0, 34) = "PILLOW PROTECTOR"
    Arreglo(0, 35) = "PILLOW CASE"
    Arreglo(0, 36) = "COVERLET "
    Arreglo(0, 37) = "QUILT"
    Arreglo(0, 38) = "BED SPREAD"
    Arreglo(0, 39) = "THROW "
    Arreglo(0, 40) = "BED SCARF"
    Arreglo(0, 41) = "Amnesty-Sís-Pinton Tapestries"
    Arreglo(0, 42) = "Apocalypse Tapestry"
    Arreglo(0, 43) = "Armada tapestries"
    Arreglo(0, 44) = "Aubusson tapestry"
    Arreglo(0, 45) = "Bayeux Tapestry"
    Arreglo(0, 46) = "Bayeux Tapestry tituli"
    Arreglo(0, 47) = "Beauvais Manufactory"
    Arreglo(0, 48) = "St. Hedwig's Cathedral"
    Arreglo(0, 49) = "Brussels tapestry"
    Arreglo(0, 50) = "Les Chasses de Maximilien"
    Arreglo(0, 51) = "Christ in Glory in the Tetramorph"
    Arreglo(0, 52) = "Cloth of St Gereon"
    Arreglo(0, 53) = "The Death of Polydorus"
    Arreglo(0, 54) = "Devonshire Hunting Tapestries"
    Arreglo(0, 55) = "Franses Tapestry"
    Arreglo(0, 56) = "Game of Thrones Tapestry"
    Arreglo(0, 57) = "Gobelins Manufactory"
    Arreglo(0, 58) = "Great Tapestry of Scotland"
    Arreglo(0, 59) = "Grödinge tapestry"
    Arreglo(0, 60) = "Gunthertuch tapestry"
    Arreglo(0, 61) = "Hestia Tapestry"
    Arreglo(0, 62) = "The History of Constantine"
    Arreglo(0, 63) = "The Hunt of the Unicorn"
    Arreglo(0, 64) = "Hunting of Birds with a Hawk and a Bow"
    Arreglo(0, 65) = "Jagiellonian tapestries"
    Arreglo(0, 66) = "K'o-ssu Tapestry"
    Arreglo(0, 67) = "Kalaga Tapestry"
    Arreglo(0, 68) = "Kilim Tapestry"
    Arreglo(0, 69) = "The Lady and the Unicorn Tapestry"
    Arreglo(0, 70) = "Millefleur Tapestry"
    Arreglo(0, 71) = "Moravská gobelínová manufaktura Tapestry"
    Arreglo(0, 72) = "Mortlake Tapestry Works"
    Arreglo(0, 73) = "Navajo weaving Tapestry"
    Arreglo(0, 74) = "New World Tapestry"
    Arreglo(0, 75) = "Oseberg tapestry fragments"
    Arreglo(0, 76) = "Överhogdal tapestries"
    Arreglo(0, 77) = "The Pastoral Amusements Tapestry"
    Arreglo(0, 78) = "Pastrana Tapestries Tapestry"
    Arreglo(0, 79) = "Prestonpans Tapestry"
    Arreglo(0, 80) = "Quaker Tapestry"
    Arreglo(0, 81) = "Raphael Cartoons Tapestry"
    Arreglo(0, 82) = "Ros Tapestry Project"
    Arreglo(0, 83) = "Royal Tapestry Factory"
    Arreglo(0, 84) = "Ryijy Tapestry"
    Arreglo(0, 85) = "Sampul tapestry"
    Arreglo(0, 86) = "Scottish Diaspora Tapestry"
    Arreglo(0, 87) = "Scottish Royal tapestry collection"
    Arreglo(0, 88) = "Sheldon tapestries"
    Arreglo(0, 89) = "Siparium Tapestry"
    Arreglo(0, 90) = "Skog tapestry"
    Arreglo(0, 91) = "The Triumph of Fame"
    Arreglo(0, 92) = "Valois Tapestries"
    Arreglo(0, 93) = "Walsall Silver Thread Tapestries"
    Arreglo(0, 94) = "William Baumgarten & Co Tapestry"
    Arreglo(0, 95) = "The World Trade Center Tapestry"
    Arreglo(0, 96) = "Wool rugs"
    Arreglo(0, 97) = "Cotton rugs"
    Arreglo(0, 98) = "Jute and sisal rugs"
    Arreglo(0, 99) = "Silk and viscose rugs"
    Arreglo(0, 100) = "Nylon rugs"
    Arreglo(0, 101) = "Olefin rugs"
    Arreglo(0, 102) = "Polyester rugs"
    Arreglo(0, 103) = "Universal Chair cover"
    Arreglo(0, 104) = "Satin Chair cover"
    Arreglo(0, 105) = "Polyester Chair cover"
    Arreglo(0, 106) = "Spandex Chair cover"
    Arreglo(0, 107) = "Chiavari Chair cover"
    Arreglo(0, 108) = "The Ottoman Sofa"
    Arreglo(0, 109) = "The Armchair Sofa"
    Arreglo(0, 110) = "The Loveseat Sofa"
    Arreglo(0, 111) = "The Sectional Sofa"
    Arreglo(0, 112) = "Modular Sofa"
    Arreglo(0, 113) = "Sofa beds"
    Arreglo(0, 114) = "Futons Sofa"
    Arreglo(0, 115) = "Clik-claks Sofa "
    Arreglo(0, 116) = "Classic Round Arm Sofa"
    Arreglo(0, 117) = "Retro Square Arm Sofa"
    Arreglo(0, 118) = "Hard Wedge Arm Sofa"
    Arreglo(0, 119) = "Rounded Wedge Arm Sofa"
    Arreglo(0, 120) = "The Sloped Arm Sofa"
    Arreglo(0, 121) = "Belgian Roll Arm Sofa"
    Arreglo(0, 122) = "English Roll Arm Sofa"
    Arreglo(0, 123) = "No Arms Sofa"
    Arreglo(0, 124) = "Wooden Arms Sofa"
    Arreglo(0, 125) = "Straight Back Sofa"
    Arreglo(0, 126) = "Tuxedo Sofas"
    Arreglo(0, 127) = "High Back Sofa"
    Arreglo(0, 128) = "Round Back Sofa"
    Arreglo(0, 129) = "Camelback Sofa"
    Arreglo(0, 130) = "Wingback Sofa"
    Arreglo(0, 131) = "Barrelback Sofa"
    Arreglo(0, 132) = "Rollback Sofa"
    Arreglo(0, 133) = "Round bean bag"
    Arreglo(0, 134) = "Square bean bag"
    Arreglo(0, 135) = "Game chairs bean bag"
    Arreglo(0, 136) = "Novelty bean bag"
    Arreglo(0, 137) = "Elongated large bean bag"
    Arreglo(0, 138) = "Kids/Youth Bean Bags"
    Arreglo(0, 139) = "Large/Teen Bean Bags"
    Arreglo(0, 140) = "Extra Large Bean Bags"
    Arreglo(0, 141) = "Double Extra Large Bean Bags"
    Arreglo(0, 142) = "Polystyrene beads"
    Arreglo(0, 143) = "Shredded foam filler"
    Arreglo(0, 144) = "Nylon Carpets"
    Arreglo(0, 145) = "Olefin Carpets"
    Arreglo(0, 146) = "Polyester Carpets"
    Arreglo(0, 147) = "Acrylic Carpets"
    Arreglo(0, 148) = "Wool Carpets"
    Arreglo(0, 149) = "Triexta Carpets"
    Arreglo(0, 150) = "Microfiber Doormat"
    Arreglo(0, 151) = "Coir Mats Doormat"
    Arreglo(0, 152) = "Rubber Doormat"
    Arreglo(0, 153) = "Cast Iron Doormat"
    Arreglo(0, 154) = "Snow Doormat"
    Arreglo(0, 155) = "Ice-Melting  Doormat"
    Arreglo(0, 156) = "Weather Resistant Doormat"
    Arreglo(0, 157) = "Eco-Friendly Doormat"
    
    
    Set Myrange = Range("A2:B159")
    i = 0

    j = 0

    For Each Cell In Myrange

        If j Mod 2 = 0 Then

            Cell.Value = i + 1

            i = i + 1
            j = j + 1
        Else

            Cell.Value = Arreglo(0, i - 1)

            j = j + 1
        End If

    Next Cell
    
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

    Range("A1").Value = "Order"
    Range("A1").Font.Bold = True

    Range("B1").Value = "Furnishings"
    Range("B1").Font.Bold = True

    Columns("A:B").EntireColumn.AutoFit
    Columns("A:A").HorizontalAlignment = xlCenter
    
    Range("A2").Select
    ActiveWindow.FreezePanes = True


End Sub
Sub Macro18_Fill_Labels()

    Sheets("Labels").Select

    Dim Myrange As Range
    Dim Cell As Range
    Dim i As Double
    Dim j As Double

    Dim Arreglo(0, 23) As String

    Arreglo(0, 0) = "Products Labels"
    Arreglo(0, 1) = "Packaging Labels"
    Arreglo(0, 2) = "Assets Labels"
    Arreglo(0, 3) = "Textiles Labels"
    Arreglo(0, 4) = "Mailing Labels"
    Arreglo(0, 5) = "Notebook Labels"
    Arreglo(0, 6) = "Piggyback Labels"
    Arreglo(0, 7) = "Smart Labels"
    Arreglo(0, 8) = "Blockout Labels"
    Arreglo(0, 9) = "Radioactive Labels"
    Arreglo(0, 10) = "Laser / printer labels"
    Arreglo(0, 11) = "Security Labels"
    Arreglo(0, 12) = "Antimicrobial Labels"
    Arreglo(0, 13) = "Fold-out Labels"
    Arreglo(0, 14) = "Barcode Labels"
    Arreglo(0, 15) = "Paper Labels"
    Arreglo(0, 16) = "Nonwoven fabric Labels"
    Arreglo(0, 17) = "Latex Labels"
    Arreglo(0, 18) = "Plastics Labels"
    Arreglo(0, 19) = "Foil Labels"
    Arreglo(0, 20) = "Thermal Labels"
    Arreglo(0, 21) = "Thermal transfer Labels"
    Arreglo(0, 22) = "Thermal transfer ribbon Labels"
    Arreglo(0, 23) = "Substrate Silk Screen Labels"
    
    Set Myrange = Range("A2:B25")
    i = 0

    j = 0

    For Each Cell In Myrange

        If j Mod 2 = 0 Then

            Cell.Value = i + 1

            i = i + 1
            j = j + 1
        Else

            Cell.Value = Arreglo(0, i - 1)

            j = j + 1
        End If

    Next Cell
    
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
    

    Range("A1").Value = "Order"
    Range("A1").Font.Bold = True

    Range("B1").Value = "Labels"
    Range("B1").Font.Bold = True

    Columns("A:B").EntireColumn.AutoFit
    Columns("A:A").HorizontalAlignment = xlCenter

    Range("A2").Select
    ActiveWindow.FreezePanes = True


End Sub
Sub Macro19_Fill_Gym_Machines()

    Sheets("Gym Machines").Select

    Dim Myrange As Range
    Dim Cell As Range
    Dim i As Double
    Dim j As Double

    Dim Arreglo(0, 29) As String



    Arreglo(0, 0) = "Cardio equipment"
    Arreglo(0, 1) = "Spinning Bikes"
    Arreglo(0, 2) = "Cross-country ski machine"
    Arreglo(0, 3) = "Elliptical trainers"
    Arreglo(0, 4) = "Rowing machines"
    Arreglo(0, 5) = "Stair-steppers"
    Arreglo(0, 6) = "Stationary bicycle"
    Arreglo(0, 7) = "Exercise Bikes"
    Arreglo(0, 8) = "Assault Air Bike"
    Arreglo(0, 9) = "Treadmill"
    Arreglo(0, 10) = "Strength equipment"
    Arreglo(0, 11) = "Ankle weights"
    Arreglo(0, 12) = "Exercise mat"
    Arreglo(0, 13) = "Hand weights"
    Arreglo(0, 14) = "Resistance bands and tubing"
    Arreglo(0, 15) = "Bands"
    Arreglo(0, 16) = "Tubing"
    
    Set Myrange = Range("A2:B18")
    i = 0

    j = 0

    For Each Cell In Myrange

        If j Mod 2 = 0 Then

            Cell.Value = i + 1

            i = i + 1
            j = j + 1
        Else

            Cell.Value = Arreglo(0, i - 1)

            j = j + 1
        End If

    Next Cell
    
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

    Range("A1").Value = "Order"
    Range("A1").Font.Bold = True

    Range("B1").Value = "Gym Machines"
    Range("B1").Font.Bold = True

    Columns("A:B").EntireColumn.AutoFit
    Columns("A:A").HorizontalAlignment = xlCenter

    Range("A2").Select
    ActiveWindow.FreezePanes = True


End Sub

Sub Macro20_Fill_Papers()


    Sheets("Papers").Select

    Dim Myrange As Range
    Dim Cell As Range
    Dim i As Double
    Dim j As Double

    Dim Arreglo(0, 29) As String

    
    Arreglo(0, 0) = "Repro paper"
    Arreglo(0, 1) = "Coated paper"
    Arreglo(0, 2) = "Tissue paper "
    Arreglo(0, 3) = "Newsprint"
    Arreglo(0, 4) = "Cardboard"
    Arreglo(0, 5) = "Paperboard"
    Arreglo(0, 6) = "Fine art paper"
    Arreglo(0, 7) = "Bank paper"
    Arreglo(0, 8) = "Banana paper"
    Arreglo(0, 9) = "Bond paper"
    Arreglo(0, 10) = "Book paper"
    Arreglo(0, 11) = "Soft Coated paper"
    Arreglo(0, 12) = "Construction paper"
    Arreglo(0, 13) = "Sugar paper"
    Arreglo(0, 14) = "Cotton paper"
    Arreglo(0, 15) = "Fish paper "
    Arreglo(0, 16) = "Inkjet paper"
    Arreglo(0, 17) = "Kraft paper"
    Arreglo(0, 18) = "Laid paper"
    Arreglo(0, 19) = "Leather paper"
    Arreglo(0, 20) = "Mummy paper"
    Arreglo(0, 21) = "Oak tag paper"
    Arreglo(0, 22) = "Sandpaper"
    Arreglo(0, 23) = "Tyvek paper"
    Arreglo(0, 24) = "Wallpaper"
    Arreglo(0, 25) = "Washi paper"
    Arreglo(0, 26) = "Waterproof paper"
    Arreglo(0, 27) = "Wax paper"
    Arreglo(0, 28) = "Wove paper"
    Arreglo(0, 29) = "Xuan paper"

    Set Myrange = Range("A2:B31")
    i = 0

    j = 0

    For Each Cell In Myrange

        If j Mod 2 = 0 Then

            Cell.Value = i + 1

            i = i + 1
            j = j + 1
        Else

            Cell.Value = Arreglo(0, i - 1)

            j = j + 1
        End If

    Next Cell
    
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


    Range("A1").Value = "Order"
    Range("A1").Font.Bold = True

    Range("B1").Value = "Papers"
    Range("B1").Font.Bold = True

    Columns("A:B").EntireColumn.AutoFit
    Columns("A:A").HorizontalAlignment = xlCenter

    Range("A2").Select
    ActiveWindow.FreezePanes = True

End Sub
Sub Macro21_Fill_Storage()


    Sheets("Storage").Select

    Dim Myrange As Range
    Dim Cell As Range
    Dim i As Double
    Dim j As Double

    Dim Arreglo(0, 16) As String
    
    Arreglo(0, 0) = "Decorative Bins"
    Arreglo(0, 1) = "Baskets "
    Arreglo(0, 2) = "Drawer Organizers"
    Arreglo(0, 3) = "Plastic Bins "
    Arreglo(0, 4) = "Plastic Baskets"
    Arreglo(0, 5) = "Drawers"
    Arreglo(0, 6) = "Like-it System "
    Arreglo(0, 7) = "3 Sprouts"
    Arreglo(0, 8) = "Garage Storage Totes"
    Arreglo(0, 9) = "Storage Bags"
    Arreglo(0, 10) = "Storage Totes"
    Arreglo(0, 11) = "Cases of Storage"
    Arreglo(0, 12) = "The Home Edit Exclusive Collection"
    Arreglo(0, 13) = "Lego Storage"
    Arreglo(0, 14) = "Storage Benches & Seats"
    Arreglo(0, 15) = "SmartStore"
    Arreglo(0, 16) = "Trunks "
    
    Set Myrange = Range("A2:B18")
    i = 0

    j = 0

    For Each Cell In Myrange

        If j Mod 2 = 0 Then

            Cell.Value = i + 1

            i = i + 1
            j = j + 1
        Else

            Cell.Value = Arreglo(0, i - 1)

            j = j + 1
        End If

    Next Cell
    
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

    Range("A1").Value = "Order"
    Range("A1").Font.Bold = True

    Range("B1").Value = "Storage"
    Range("B1").Font.Bold = True

    Columns("A:B").EntireColumn.AutoFit
    Columns("A:A").HorizontalAlignment = xlCenter

    Range("A2").Select
    ActiveWindow.FreezePanes = True


End Sub
Sub Macro22_Fill_Supplies()


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

Sub Macro23_Fill_Tables()


    Sheets("Tables").Select

    Dim Myrange As Range
    Dim Cell As Range
    Dim i As Double
    Dim j As Double

    Dim Arreglo(0, 46) As String

    Arreglo(0, 0) = "Coffee Table"
    Arreglo(0, 1) = "Accent Table"
    Arreglo(0, 2) = "Console Table"
    Arreglo(0, 3) = "Side Table"
    Arreglo(0, 4) = "C-table"
    Arreglo(0, 5) = "Drink Table"
    Arreglo(0, 6) = "End Table"
    Arreglo(0, 7) = "Bunching Table"
    Arreglo(0, 8) = "Stacking Table"
    Arreglo(0, 9) = "Nesting table"
    Arreglo(0, 10) = "Drum Table"
    Arreglo(0, 11) = "Foyer Tables"
    Arreglo(0, 12) = "Ottoman Tables"
    Arreglo(0, 13) = "Dining Table"
    Arreglo(0, 14) = "Kitchen Table"
    Arreglo(0, 15) = "Bedside Table"
    Arreglo(0, 16) = "Nightstand Table"
    Arreglo(0, 17) = "Pub Table"
    Arreglo(0, 18) = "Patio Table"
    Arreglo(0, 19) = "Work Table"
    Arreglo(0, 20) = "Conference Table"
    Arreglo(0, 21) = "Computer Table"
    Arreglo(0, 22) = "Game Tables - Pool Table"
    Arreglo(0, 23) = "Game Tables - Ping Pong Table"
    Arreglo(0, 24) = "Game Tables - Foosball Table"
    Arreglo(0, 25) = "Game Tables - Card Table"
    Arreglo(0, 26) = "Square or Rectangle Table "
    Arreglo(0, 27) = "Polygon Table "
    Arreglo(0, 28) = "Industrial Table"
    Arreglo(0, 29) = "Farmhouse Table"
    Arreglo(0, 30) = "Shabby Chic Table"
    Arreglo(0, 31) = "Mid-Century Modern Table"
    Arreglo(0, 32) = "Scandinavian Table"
    Arreglo(0, 33) = "Antique Table"
    Arreglo(0, 34) = "Custom Table"
    Arreglo(0, 35) = "Big Box Self-Assembled Table"
    Arreglo(0, 36) = "Furniture Store Table"
    Arreglo(0, 37) = "DIY or Restoration Table"
    Arreglo(0, 38) = "Pre-loved Table"
    Arreglo(0, 39) = "Wood Veneer Table"
    Arreglo(0, 40) = "Laminate Table"
    Arreglo(0, 41) = "Marble Table"
    Arreglo(0, 42) = "Solid Wood Table"
    Arreglo(0, 43) = "Metal Table"
    Arreglo(0, 44) = "Glass Table"
    Arreglo(0, 45) = "Aesthetics Table"
    Arreglo(0, 46) = "Functional Table"

    Set Myrange = Range("A2:B48")
    i = 0

    j = 0

    For Each Cell In Myrange

        If j Mod 2 = 0 Then

            Cell.Value = i + 1

            i = i + 1
            j = j + 1
        Else

            Cell.Value = Arreglo(0, i - 1)

            j = j + 1
        End If

    Next Cell
    
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


    Range("A1").Value = "Order"
    Range("A1").Font.Bold = True

    Range("B1").Value = "Tables"
    Range("B1").Font.Bold = True

    Columns("A:B").EntireColumn.AutoFit
    Columns("A:A").HorizontalAlignment = xlCenter

    Range("A2").Select
    ActiveWindow.FreezePanes = True

End Sub

Sub Macro24_Fill_Products()

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

Sub MacroN_menos_2_Fill_Names()

    Sheets("Names").Select

    Dim Myrange As Range
    Dim Cell As Range
    Dim i As Double
    Dim j As Double

    Dim Arreglo(0, 1012) As String

    Arreglo(0, 0) = "Aaron"
    Arreglo(0, 1) = "Abbott"
    Arreglo(0, 2) = "Abel"
    Arreglo(0, 3) = "Abner"
    Arreglo(0, 4) = "Abraham"
    Arreglo(0, 5) = "Adam"
    Arreglo(0, 6) = "Addison"
    Arreglo(0, 7) = "Adler"
    Arreglo(0, 8) = "Adley"
    Arreglo(0, 9) = "Adrian"
    Arreglo(0, 10) = "Adrien"
    Arreglo(0, 11) = "Aedan"
    Arreglo(0, 12) = "Aiden"
    Arreglo(0, 13) = "Aiken"
    Arreglo(0, 14) = "Alan"
    Arreglo(0, 15) = "Allan"
    Arreglo(0, 16) = "Alastair"
    Arreglo(0, 17) = "Albern"
    Arreglo(0, 18) = "Albert"
    Arreglo(0, 19) = "Albion"
    Arreglo(0, 20) = "Alden"
    Arreglo(0, 21) = "Aldis"
    Arreglo(0, 22) = "Aldrich"
    Arreglo(0, 23) = "Alexander"
    Arreglo(0, 24) = "Alfie"
    Arreglo(0, 25) = "Alfred"
    Arreglo(0, 26) = "Algernon"
    Arreglo(0, 27) = "Alston"
    Arreglo(0, 28) = "Alton"
    Arreglo(0, 29) = "Alvin"
    Arreglo(0, 30) = "Ambrose"
    Arreglo(0, 31) = "Amery"
    Arreglo(0, 32) = "Amos"
    Arreglo(0, 33) = "Andrew"
    Arreglo(0, 34) = "Angus"
    Arreglo(0, 35) = "Ansel"
    Arreglo(0, 36) = "Anthony"
    Arreglo(0, 37) = "Archer"
    Arreglo(0, 38) = "Archibald"
    Arreglo(0, 39) = "Arlen"
    Arreglo(0, 40) = "Arnold"
    Arreglo(0, 41) = "Arthur"
    Arreglo(0, 42) = "Art"
    Arreglo(0, 43) = "Arvel"
    Arreglo(0, 44) = "Atwater"
    Arreglo(0, 45) = "Atwood"
    Arreglo(0, 46) = "Aubrey"
    Arreglo(0, 47) = "Austin"
    Arreglo(0, 48) = "Avery"
    Arreglo(0, 49) = "Axel"
    Arreglo(0, 50) = "Baird"
    Arreglo(0, 51) = "Baldwin"
    Arreglo(0, 52) = "Barclay"
    Arreglo(0, 53) = "Barnaby"
    Arreglo(0, 54) = "Baron"
    Arreglo(0, 55) = "Barrett"
    Arreglo(0, 56) = "Barry"
    Arreglo(0, 57) = "Bartholomew"
    Arreglo(0, 58) = "Basil"
    Arreglo(0, 59) = "Benedict"
    Arreglo(0, 60) = "Benjamin"
    Arreglo(0, 61) = "Benton"
    Arreglo(0, 62) = "Bernard"
    Arreglo(0, 63) = "Bert"
    Arreglo(0, 64) = "Bevis"
    Arreglo(0, 65) = "Blaine"
    Arreglo(0, 66) = "Blair"
    Arreglo(0, 67) = "Blake"
    Arreglo(0, 68) = "Bond"
    Arreglo(0, 69) = "Boris"
    Arreglo(0, 70) = "Bowen"
    Arreglo(0, 71) = "Braden"
    Arreglo(0, 72) = "Bradley"
    Arreglo(0, 73) = "Brandan"
    Arreglo(0, 74) = "Brendan"
    Arreglo(0, 75) = "Brendon"
    Arreglo(0, 76) = "Brent"
    Arreglo(0, 77) = "Bret"
    Arreglo(0, 78) = "Brett"
    Arreglo(0, 79) = "Brian"
    Arreglo(0, 80) = "Brice"
    Arreglo(0, 81) = "Brigham"
    Arreglo(0, 82) = "Brock"
    Arreglo(0, 83) = "Broderick"
    Arreglo(0, 84) = "Brooke"
    Arreglo(0, 85) = "Bruce"
    Arreglo(0, 86) = "Bruno"
    Arreglo(0, 87) = "Bryant"
    Arreglo(0, 88) = "Buck"
    Arreglo(0, 89) = "Bud"
    Arreglo(0, 90) = "Burgess"
    Arreglo(0, 91) = "Burton"
    Arreglo(0, 92) = "Byron"
    Arreglo(0, 93) = "Cadman"
    Arreglo(0, 94) = "Calvert"
    Arreglo(0, 95) = "Caldwell"
    Arreglo(0, 96) = "Caleb"
    Arreglo(0, 97) = "Calvin"
    Arreglo(0, 98) = "Carrick"
    Arreglo(0, 99) = "Carl"
    Arreglo(0, 100) = "Carlton"
    Arreglo(0, 101) = "Carney"
    Arreglo(0, 102) = "Carroll"
    Arreglo(0, 103) = "Carter"
    Arreglo(0, 104) = "Carver"
    Arreglo(0, 105) = "Cary"
    Arreglo(0, 106) = "Casey"
    Arreglo(0, 107) = "Casper"
    Arreglo(0, 108) = "Cecil"
    Arreglo(0, 109) = "Cedric"
    Arreglo(0, 110) = "Chad"
    Arreglo(0, 111) = "Chadwick"
    Arreglo(0, 112) = "Chalmers"
    Arreglo(0, 113) = "Chandler"
    Arreglo(0, 114) = "Channing"
    Arreglo(0, 115) = "Chapman"
    Arreglo(0, 116) = "Charles"
    Arreglo(0, 117) = "Chatwin"
    Arreglo(0, 118) = "Chester"
    Arreglo(0, 119) = "Christian"
    Arreglo(0, 120) = "Christopher"
    Arreglo(0, 121) = "Clarence"
    Arreglo(0, 122) = "Claude"
    Arreglo(0, 123) = "Clayton"
    Arreglo(0, 124) = "Clay"
    Arreglo(0, 125) = "Clifford"
    Arreglo(0, 126) = "Cliff"
    Arreglo(0, 127) = "Clive"
    Arreglo(0, 128) = "Clyde"
    Arreglo(0, 129) = "Coleman"
    Arreglo(0, 130) = "Colin"
    Arreglo(0, 131) = "Collier"
    Arreglo(0, 132) = "Conan"
    Arreglo(0, 133) = "Connell"
    Arreglo(0, 134) = "Connor"
    Arreglo(0, 135) = "Conrad"
    Arreglo(0, 136) = "Conroy"
    Arreglo(0, 137) = "Conway"
    Arreglo(0, 138) = "Corwin"
    Arreglo(0, 139) = "Crispin"
    Arreglo(0, 140) = "Crosby"
    Arreglo(0, 141) = "Culbert"
    Arreglo(0, 142) = "Culver"
    Arreglo(0, 143) = "Curt"
    Arreglo(0, 144) = "Curtis"
    Arreglo(0, 145) = "Cuthbert"
    Arreglo(0, 146) = "Craig"
    Arreglo(0, 147) = "Cyril"
    Arreglo(0, 148) = "Dale"
    Arreglo(0, 149) = "Daley"
    Arreglo(0, 150) = "Dalton"
    Arreglo(0, 151) = "Damon"
    Arreglo(0, 152) = "Daniel"
    Arreglo(0, 153) = "Darcy"
    Arreglo(0, 154) = "Darian"
    Arreglo(0, 155) = "Darell"
    Arreglo(0, 156) = "Darrel"
    Arreglo(0, 157) = "David"
    Arreglo(0, 158) = "Davin"
    Arreglo(0, 159) = "Dean"
    Arreglo(0, 160) = "Declan"
    Arreglo(0, 161) = "Delmar"
    Arreglo(0, 162) = "Denley"
    Arreglo(0, 163) = "Dennis"
    Arreglo(0, 164) = "Derek"
    Arreglo(0, 165) = "Dermot"
    Arreglo(0, 166) = "Derwin"
    Arreglo(0, 167) = "Des"
    Arreglo(0, 168) = "Desmond"
    Arreglo(0, 169) = "Dexter"
    Arreglo(0, 170) = "Dillon"
    Arreglo(0, 171) = "Dion"
    Arreglo(0, 172) = "Dirk"
    Arreglo(0, 173) = "Dixon"
    Arreglo(0, 174) = "Dominic"
    Arreglo(0, 175) = "Donald"
    Arreglo(0, 176) = "Dorian"
    Arreglo(0, 177) = "Douglas"
    Arreglo(0, 178) = "Doyle"
    Arreglo(0, 179) = "Drake"
    Arreglo(0, 180) = "Drew"
    Arreglo(0, 181) = "Driscoll"
    Arreglo(0, 182) = "Dudley"
    Arreglo(0, 183) = "Duncan"
    Arreglo(0, 184) = "Durwin"
    Arreglo(0, 185) = "Dwayne"
    Arreglo(0, 186) = "Dwight"
    Arreglo(0, 187) = "Dylan"
    Arreglo(0, 188) = "Earl"
    Arreglo(0, 189) = "Eaton"
    Arreglo(0, 190) = "Ebenezer"
    Arreglo(0, 191) = "Edan"
    Arreglo(0, 192) = "Edgar"
    Arreglo(0, 193) = "Edric"
    Arreglo(0, 194) = "Edmond"
    Arreglo(0, 195) = "Edmund"
    Arreglo(0, 196) = "Edward"
    Arreglo(0, 197) = "Eddie"
    Arreglo(0, 198) = "Edwin"
    Arreglo(0, 199) = "Efrain"
    Arreglo(0, 200) = "Egan"
    Arreglo(0, 201) = "Egbert"
    Arreglo(0, 202) = "Egerton"
    Arreglo(0, 203) = "Egil"
    Arreglo(0, 204) = "Elbert"
    Arreglo(0, 205) = "Eldon"
    Arreglo(0, 206) = "Eldwin"
    Arreglo(0, 207) = "Eli"
    Arreglo(0, 208) = "Ely"
    Arreglo(0, 209) = "Elijah"
    Arreglo(0, 210) = "Elias"
    Arreglo(0, 211) = "Eliot"
    Arreglo(0, 212) = "Elliott"
    Arreglo(0, 213) = "Ellery"
    Arreglo(0, 214) = "Elmer"
    Arreglo(0, 215) = "Elroy"
    Arreglo(0, 216) = "Elton"
    Arreglo(0, 217) = "Elvis"
    Arreglo(0, 218) = "Emerson"
    Arreglo(0, 219) = "Emery"
    Arreglo(0, 220) = "Emmanuel"
    Arreglo(0, 221) = "Emmett"
    Arreglo(0, 222) = "Emrick"
    Arreglo(0, 223) = "Enoch"
    Arreglo(0, 224) = "Eric"
    Arreglo(0, 225) = "Erik"
    Arreglo(0, 226) = "Ernest"
    Arreglo(0, 227) = "Errol"
    Arreglo(0, 228) = "Erskine"
    Arreglo(0, 229) = "Erwin"
    Arreglo(0, 230) = "Esmond"
    Arreglo(0, 231) = "Ethan"
    Arreglo(0, 232) = "Ethanael"
    Arreglo(0, 233) = "Ethen"
    Arreglo(0, 234) = "Eugene"
    Arreglo(0, 235) = "Evan"
    Arreglo(0, 236) = "Everett"
    Arreglo(0, 237) = "Ezra"
    Arreglo(0, 238) = "Fabian"
    Arreglo(0, 239) = "Fairfax"
    Arreglo(0, 240) = "Falkner"
    Arreglo(0, 241) = "Farley"
    Arreglo(0, 242) = "Farrell"
    Arreglo(0, 243) = "Felix"
    Arreglo(0, 244) = "Fenton"
    Arreglo(0, 245) = "Ferdinand"
    Arreglo(0, 246) = "Fergal"
    Arreglo(0, 247) = "Fergus"
    Arreglo(0, 248) = "Ferguson"
    Arreglo(0, 249) = "Ferris"
    Arreglo(0, 250) = "Finbar"
    Arreglo(0, 251) = "Fitzgerald"
    Arreglo(0, 252) = "Fleming"
    Arreglo(0, 253) = "Fletcher"
    Arreglo(0, 254) = "Floyd"
    Arreglo(0, 255) = "Forbes"
    Arreglo(0, 256) = "Forrest"
    Arreglo(0, 257) = "Foster"
    Arreglo(0, 258) = "Fox"
    Arreglo(0, 259) = "Francis"
    Arreglo(0, 260) = "Frank"
    Arreglo(0, 261) = "Frasier"
    Arreglo(0, 262) = "Frederick"
    Arreglo(0, 263) = "Freeman"
    Arreglo(0, 264) = "Gabriel"
    Arreglo(0, 265) = "Gale"
    Arreglo(0, 266) = "Galvin"
    Arreglo(0, 267) = "Gardner"
    Arreglo(0, 268) = "Garret"
    Arreglo(0, 269) = "Garrick"
    Arreglo(0, 270) = "Garth"
    Arreglo(0, 271) = "Gavin"
    Arreglo(0, 272) = "George"
    Arreglo(0, 273) = "Gerald"
    Arreglo(0, 274) = "Gerard"
    Arreglo(0, 275) = "Gerret"
    Arreglo(0, 276) = "Gideon"
    Arreglo(0, 277) = "Gifford"
    Arreglo(0, 278) = "Gilbert"
    Arreglo(0, 279) = "Giles"
    Arreglo(0, 280) = "Gilroy"
    Arreglo(0, 281) = "Glenn"
    Arreglo(0, 282) = "Goddard"
    Arreglo(0, 283) = "Godfrey"
    Arreglo(0, 284) = "Godwin"
    Arreglo(0, 285) = "Graham"
    Arreglo(0, 286) = "Grant"
    Arreglo(0, 287) = "Grayson"
    Arreglo(0, 288) = "Gregory"
    Arreglo(0, 289) = "Gresham"
    Arreglo(0, 290) = "Griswald"
    Arreglo(0, 291) = "Griswold"
    Arreglo(0, 292) = "Grover"
    Arreglo(0, 293) = "Guy"
    Arreglo(0, 294) = "Hadden"
    Arreglo(0, 295) = "Hadley"
    Arreglo(0, 296) = "Hadwin"
    Arreglo(0, 297) = "Hal"
    Arreglo(0, 298) = "Halbert"
    Arreglo(0, 299) = "Halden"
    Arreglo(0, 300) = "Hale"
    Arreglo(0, 301) = "Hall"
    Arreglo(0, 302) = "Halsey"
    Arreglo(0, 303) = "Hamlin"
    Arreglo(0, 304) = "Hanley"
    Arreglo(0, 305) = "Hardy"
    Arreglo(0, 306) = "Harlan"
    Arreglo(0, 307) = "Harland"
    Arreglo(0, 308) = "Harley"
    Arreglo(0, 309) = "Harold"
    Arreglo(0, 310) = "Harry"
    Arreglo(0, 311) = "Harris"
    Arreglo(0, 312) = "Harrison"
    Arreglo(0, 313) = "Hartley"
    Arreglo(0, 314) = "Heath"
    Arreglo(0, 315) = "Heathcliff"
    Arreglo(0, 316) = "Hector"
    Arreglo(0, 317) = "Henry"
    Arreglo(0, 318) = "Herbert"
    Arreglo(0, 319) = "Herman"
    Arreglo(0, 320) = "Homer"
    Arreglo(0, 321) = "Horace"
    Arreglo(0, 322) = "Horatio"
    Arreglo(0, 323) = "Howard"
    Arreglo(0, 324) = "Hubert"
    Arreglo(0, 325) = "Hugh"
    Arreglo(0, 326) = "Hugo"
    Arreglo(0, 327) = "Humphrey"
    Arreglo(0, 328) = "Hunter"
    Arreglo(0, 329) = "Ian"
    Arreglo(0, 330) = "Igor"
    Arreglo(0, 331) = "Irvin"
    Arreglo(0, 332) = "Irving"
    Arreglo(0, 333) = "Isaac"
    Arreglo(0, 334) = "Isaiah"
    Arreglo(0, 335) = "Ivan"
    Arreglo(0, 336) = "Iver"
    Arreglo(0, 337) = "Ivar"
    Arreglo(0, 338) = "Ives"
    Arreglo(0, 339) = "Jack"
    Arreglo(0, 340) = "Jacob"
    Arreglo(0, 341) = "James"
    Arreglo(0, 342) = "Jimmy"
    Arreglo(0, 343) = "Jarvis"
    Arreglo(0, 344) = "Jason"
    Arreglo(0, 345) = "Jasper"
    Arreglo(0, 346) = "Jed"
    Arreglo(0, 347) = "Jeffrey"
    Arreglo(0, 348) = "Jeremiah"
    Arreglo(0, 349) = "Jeremy"
    Arreglo(0, 350) = "Jerome"
    Arreglo(0, 351) = "Jesse"
    Arreglo(0, 352) = "John"
    Arreglo(0, 353) = "Jonathan"
    Arreglo(0, 354) = "Joseph"
    Arreglo(0, 355) = "Joey"
    Arreglo(0, 356) = "Joe"
    Arreglo(0, 357) = "Joshua"
    Arreglo(0, 358) = "Justin"
    Arreglo(0, 359) = "Kane"
    Arreglo(0, 360) = "Keene"
    Arreglo(0, 361) = "Keegan"
    Arreglo(0, 362) = "Keaton"
    Arreglo(0, 363) = "Keith"
    Arreglo(0, 364) = "Kelsey"
    Arreglo(0, 365) = "Kelvin"
    Arreglo(0, 366) = "Kendall"
    Arreglo(0, 367) = "Kendrick"
    Arreglo(0, 368) = "Kenneth"
    Arreglo(0, 369) = "Ken"
    Arreglo(0, 370) = "Kent"
    Arreglo(0, 371) = "Kenway"
    Arreglo(0, 372) = "Kenyon"
    Arreglo(0, 373) = "Kerry"
    Arreglo(0, 374) = "Kerwin"
    Arreglo(0, 375) = "Kevin"
    Arreglo(0, 376) = "Kiefer"
    Arreglo(0, 377) = "Kilby"
    Arreglo(0, 378) = "Kilian"
    Arreglo(0, 379) = "Kim"
    Arreglo(0, 380) = "Kimball"
    Arreglo(0, 381) = "Kingsley"
    Arreglo(0, 382) = "Kirby"
    Arreglo(0, 383) = "Kirk"
    Arreglo(0, 384) = "Kit"
    Arreglo(0, 385) = "Kody"
    Arreglo(0, 386) = "Konrad"
    Arreglo(0, 387) = "Kurt"
    Arreglo(0, 388) = "Kyle"
    Arreglo(0, 389) = "Lambert"
    Arreglo(0, 390) = "Lamont"
    Arreglo(0, 391) = "Lancelot"
    Arreglo(0, 392) = "Landon"
    Arreglo(0, 393) = "Landry"
    Arreglo(0, 394) = "Lane"
    Arreglo(0, 395) = "Lars"
    Arreglo(0, 396) = "Laurence"
    Arreglo(0, 397) = "Lee"
    Arreglo(0, 398) = "Leith"
    Arreglo(0, 399) = "Leonard"
    Arreglo(0, 400) = "Leo"
    Arreglo(0, 401) = "Leon"
    Arreglo(0, 402) = "Leroy"
    Arreglo(0, 403) = "Leslie"
    Arreglo(0, 404) = "Lester"
    Arreglo(0, 405) = "Lincoln"
    Arreglo(0, 406) = "Lionel"
    Arreglo(0, 407) = "Lloyd"
    Arreglo(0, 408) = "Logan"
    Arreglo(0, 409) = "Lombard"
    Arreglo(0, 410) = "Louis"
    Arreglo(0, 411) = "Lewis"
    Arreglo(0, 412) = "Lowell"
    Arreglo(0, 413) = "Lucas"
    Arreglo(0, 414) = "Luke"
    Arreglo(0, 415) = "Luther"
    Arreglo(0, 416) = "Lyndon"
    Arreglo(0, 417) = "Maddox"
    Arreglo(0, 418) = "Magnus"
    Arreglo(0, 419) = "Malcolm"
    Arreglo(0, 420) = "Melvin"
    Arreglo(0, 421) = "Marcus"
    Arreglo(0, 422) = "Mark"
    Arreglo(0, 423) = "Marc"
    Arreglo(0, 424) = "Marlon"
    Arreglo(0, 425) = "Martin"
    Arreglo(0, 426) = "Marvin"
    Arreglo(0, 427) = "Matthew"
    Arreglo(0, 428) = "Maurice"
    Arreglo(0, 429) = "Max"
    Arreglo(0, 430) = "Maxwell"
    Arreglo(0, 431) = "Medwin"
    Arreglo(0, 432) = "Melville"
    Arreglo(0, 433) = "Merlin"
    Arreglo(0, 434) = "Michael"
    Arreglo(0, 435) = "Milburn"
    Arreglo(0, 436) = "Miles"
    Arreglo(0, 437) = "Monroe"
    Arreglo(0, 438) = "Montague"
    Arreglo(0, 439) = "Montgomery"
    Arreglo(0, 440) = "Morgan"
    Arreglo(0, 441) = "Morris"
    Arreglo(0, 442) = "Morton"
    Arreglo(0, 443) = "Murray"
    Arreglo(0, 444) = "Nathaniel"
    Arreglo(0, 445) = "Nathan"
    Arreglo(0, 446) = "Neal"
    Arreglo(0, 447) = "Neville"
    Arreglo(0, 448) = "Nicholas"
    Arreglo(0, 449) = "Nigel"
    Arreglo(0, 450) = "Noel"
    Arreglo(0, 451) = "Norman"
    Arreglo(0, 452) = "Norris"
    Arreglo(0, 453) = "Olaf"
    Arreglo(0, 454) = "Olin"
    Arreglo(0, 455) = "Oliver"
    Arreglo(0, 456) = "Orson"
    Arreglo(0, 457) = "Oscar"
    Arreglo(0, 458) = "Oswald"
    Arreglo(0, 459) = "Otis"
    Arreglo(0, 460) = "Owen"
    Arreglo(0, 461) = "Paul"
    Arreglo(0, 462) = "Paxton"
    Arreglo(0, 463) = "Percival"
    Arreglo(0, 464) = "Percy"
    Arreglo(0, 465) = "Perry"
    Arreglo(0, 466) = "Peter"
    Arreglo(0, 467) = "Peyton"
    Arreglo(0, 468) = "Philbert"
    Arreglo(0, 469) = "Philip"
    Arreglo(0, 470) = "Phineas"
    Arreglo(0, 471) = "Pierce"
    Arreglo(0, 472) = "Quade"
    Arreglo(0, 473) = "Quenby"
    Arreglo(0, 474) = "Quillan"
    Arreglo(0, 475) = "Quimby"
    Arreglo(0, 476) = "Quentin"
    Arreglo(0, 477) = "Quinby"
    Arreglo(0, 478) = "Quincy"
    Arreglo(0, 479) = "Quinlan"
    Arreglo(0, 480) = "Quinn"
    Arreglo(0, 481) = "Ralph"
    Arreglo(0, 482) = "Ramsey"
    Arreglo(0, 483) = "Randolph"
    Arreglo(0, 484) = "Raymond"
    Arreglo(0, 485) = "Reginald"
    Arreglo(0, 486) = "Renfred"
    Arreglo(0, 487) = "Rex"
    Arreglo(0, 488) = "Rhett"
    Arreglo(0, 489) = "Richard"
    Arreglo(0, 490) = "Ridley"
    Arreglo(0, 491) = "Riley"
    Arreglo(0, 492) = "Robert"
    Arreglo(0, 493) = "Robin"
    Arreglo(0, 494) = "Roderick"
    Arreglo(0, 495) = "Rodney"
    Arreglo(0, 496) = "Roger"
    Arreglo(0, 497) = "Roland"
    Arreglo(0, 498) = "Rolf"
    Arreglo(0, 499) = "Ronald"
    Arreglo(0, 500) = "Rory"
    Arreglo(0, 501) = "Ross"
    Arreglo(0, 502) = "Roswell"
    Arreglo(0, 503) = "Roy"
    Arreglo(0, 504) = "Royce"
    Arreglo(0, 505) = "Rufus"
    Arreglo(0, 506) = "Rupert"
    Arreglo(0, 507) = "Russell"
    Arreglo(0, 508) = "Ryan"
    Arreglo(0, 509) = "Abby"
    Arreglo(0, 510) = "Abigail"
    Arreglo(0, 511) = "Ada"
    Arreglo(0, 512) = "Addison"
    Arreglo(0, 513) = "Adelaide"
    Arreglo(0, 514) = "Adele"
    Arreglo(0, 515) = "Agatha"
    Arreglo(0, 516) = "Agnes"
    Arreglo(0, 517) = "Alaina"
    Arreglo(0, 518) = "Alanna"
    Arreglo(0, 519) = "Alberta"
    Arreglo(0, 520) = "Albina"
    Arreglo(0, 521) = "Alex"
    Arreglo(0, 522) = "Alexandria"
    Arreglo(0, 523) = "Alice"
    Arreglo(0, 524) = "Alicia"
    Arreglo(0, 525) = "Alisha"
    Arreglo(0, 526) = "Alison"
    Arreglo(0, 527) = "Alma"
    Arreglo(0, 528) = "Alvina"
    Arreglo(0, 529) = "Amanda"
    Arreglo(0, 530) = "Amber"
    Arreglo(0, 531) = "Amelia"
    Arreglo(0, 532) = "Amy"
    Arreglo(0, 533) = "Ana"
    Arreglo(0, 534) = "Andrea"
    Arreglo(0, 535) = "Andy"
    Arreglo(0, 536) = "Angel"
    Arreglo(0, 537) = "Angela"
    Arreglo(0, 538) = "Angie"
    Arreglo(0, 539) = "Anna"
    Arreglo(0, 540) = "Annabelle"
    Arreglo(0, 541) = "Annabeth"
    Arreglo(0, 542) = "Anne"
    Arreglo(0, 543) = "Annie"
    Arreglo(0, 544) = "Antonia"
    Arreglo(0, 545) = "April"
    Arreglo(0, 546) = "Arabella"
    Arreglo(0, 547) = "Arda"
    Arreglo(0, 548) = "Ashley"
    Arreglo(0, 549) = "Astrid"
    Arreglo(0, 550) = "Aubrey"
    Arreglo(0, 551) = "Audrey"
    Arreglo(0, 552) = "Aurora"
    Arreglo(0, 553) = "Autumn"
    Arreglo(0, 554) = "Averil"
    Arreglo(0, 555) = "Avis"
    Arreglo(0, 556) = "Aviva"
    Arreglo(0, 557) = "Barbara"
    Arreglo(0, 558) = "Beatrice"
    Arreglo(0, 559) = "Becki"
    Arreglo(0, 560) = "Belinda"
    Arreglo(0, 561) = "Bella"
    Arreglo(0, 562) = "Berenice"
    Arreglo(0, 563) = "Bertha"
    Arreglo(0, 564) = "Betsy"
    Arreglo(0, 565) = "Betty"
    Arreglo(0, 566) = "Blanche"
    Arreglo(0, 567) = "Bobbi"
    Arreglo(0, 568) = "Bobby"
    Arreglo(0, 569) = "Brandy"
    Arreglo(0, 570) = "Brenda"
    Arreglo(0, 571) = "Bridget"
    Arreglo(0, 572) = "Bronwen"
    Arreglo(0, 573) = "Bronwyn"
    Arreglo(0, 574) = "Bryony"
    Arreglo(0, 575) = "Calla"
    Arreglo(0, 576) = "Candy"
    Arreglo(0, 577) = "Cari"
    Arreglo(0, 578) = "Carla"
    Arreglo(0, 579) = "Carlene"
    Arreglo(0, 580) = "Carlie"
    Arreglo(0, 581) = "Carmelita"
    Arreglo(0, 582) = "Carol"
    Arreglo(0, 583) = "Carol Ann"
    Arreglo(0, 584) = "Carol Anne"
    Arreglo(0, 585) = "Carole"
    Arreglo(0, 586) = "Caroline"
    Arreglo(0, 587) = "Carolyn"
    Arreglo(0, 588) = "Carrie Ann"
    Arreglo(0, 589) = "Carrie Anne"
    Arreglo(0, 590) = "Carroll"
    Arreglo(0, 591) = "Carry"
    Arreglo(0, 592) = "Cassandra"
    Arreglo(0, 593) = "Cathleen"
    Arreglo(0, 594) = "Cathy"
    Arreglo(0, 595) = "Cecilia"
    Arreglo(0, 596) = "Cecily"
    Arreglo(0, 597) = "Celestia"
    Arreglo(0, 598) = "Celia"
    Arreglo(0, 599) = "Celinda"
    Arreglo(0, 600) = "Chara"
    Arreglo(0, 601) = "Charis"
    Arreglo(0, 602) = "Charisse"
    Arreglo(0, 603) = "Charity"
    Arreglo(0, 604) = "Charla"
    Arreglo(0, 605) = "Charle"
    Arreglo(0, 606) = "Charlee"
    Arreglo(0, 607) = "Charlene"
    Arreglo(0, 608) = "Charley"
    Arreglo(0, 609) = "Charli"
    Arreglo(0, 610) = "Charlie"
    Arreglo(0, 611) = "Charlotte"
    Arreglo(0, 612) = "Charly"
    Arreglo(0, 613) = "Charlyne"
    Arreglo(0, 614) = "Charmaine"
    Arreglo(0, 615) = "Chas"
    Arreglo(0, 616) = "Chelsea"
    Arreglo(0, 617) = "Cherry"
    Arreglo(0, 618) = "Cheryl"
    Arreglo(0, 619) = "Chloe"
    Arreglo(0, 620) = "Chris"
    Arreglo(0, 621) = "Christabel"
    Arreglo(0, 622) = "Christina"
    Arreglo(0, 623) = "Christine"
    Arreglo(0, 624) = "Christy"
    Arreglo(0, 625) = "Cindy"
    Arreglo(0, 626) = "Claire"
    Arreglo(0, 627) = "Clara"
    Arreglo(0, 628) = "Clare"
    Arreglo(0, 629) = "Claribel"
    Arreglo(0, 630) = "Clarissa"
    Arreglo(0, 631) = "Claudia"
    Arreglo(0, 632) = "Clementine"
    Arreglo(0, 633) = "Cleo"
    Arreglo(0, 634) = "Colette"
    Arreglo(0, 635) = "Colleen"
    Arreglo(0, 636) = "Cordelia"
    Arreglo(0, 637) = "Courtney"
    Arreglo(0, 638) = "Crystal"
    Arreglo(0, 639) = "Cynthia"
    Arreglo(0, 640) = "Daisy"
    Arreglo(0, 641) = "Dana"
    Arreglo(0, 642) = "Danielle"
    Arreglo(0, 643) = "Danna"
    Arreglo(0, 644) = "Daphne"
    Arreglo(0, 645) = "Darlene"
    Arreglo(0, 646) = "Davina"
    Arreglo(0, 647) = "Dawn"
    Arreglo(0, 648) = "Deanna"
    Arreglo(0, 649) = "Deanne"
    Arreglo(0, 650) = "Debbie"
    Arreglo(0, 651) = "Deborah"
    Arreglo(0, 652) = "Dede"
    Arreglo(0, 653) = "Delia"
    Arreglo(0, 654) = "Denise"
    Arreglo(0, 655) = "Destiny"
    Arreglo(0, 656) = "Devon"
    Arreglo(0, 657) = "Donna"
    Arreglo(0, 658) = "Dora"
    Arreglo(0, 659) = "Doreen"
    Arreglo(0, 660) = "Dorothy"
    Arreglo(0, 661) = "Drew"
    Arreglo(0, 662) = "Drusilla"
    Arreglo(0, 663) = "Dulcie"
    Arreglo(0, 664) = "Edith"
    Arreglo(0, 665) = "Edna"
    Arreglo(0, 666) = "Edwina"
    Arreglo(0, 667) = "Effie"
    Arreglo(0, 668) = "Eleanor"
    Arreglo(0, 669) = "Elektra"
    Arreglo(0, 670) = "Eliza"
    Arreglo(0, 671) = "Elizabeth"
    Arreglo(0, 672) = "Ella"
    Arreglo(0, 673) = "Ellen"
    Arreglo(0, 674) = "Ellie"
    Arreglo(0, 675) = "Emily"
    Arreglo(0, 676) = "Emma"
    Arreglo(0, 677) = "Enid"
    Arreglo(0, 678) = "Erika"
    Arreglo(0, 679) = "Erin"
    Arreglo(0, 680) = "Estelle"
    Arreglo(0, 681) = "Esther"
    Arreglo(0, 682) = "Esty"
    Arreglo(0, 683) = "Ethel"
    Arreglo(0, 684) = "Ethelreda"
    Arreglo(0, 685) = "Eudora"
    Arreglo(0, 686) = "Eva"
    Arreglo(0, 687) = "Eve"
    Arreglo(0, 688) = "Evelyn"
    Arreglo(0, 689) = "Faith"
    Arreglo(0, 690) = "Felicity"
    Arreglo(0, 691) = "Fleur"
    Arreglo(0, 692) = "Flora"
    Arreglo(0, 693) = "Florence"
    Arreglo(0, 694) = "Francie"
    Arreglo(0, 695) = "Frida"
    Arreglo(0, 696) = "Gail"
    Arreglo(0, 697) = "Gemma"
    Arreglo(0, 698) = "Genevieve"
    Arreglo(0, 699) = "Georgia"
    Arreglo(0, 700) = "Georgiana"
    Arreglo(0, 701) = "Gertie"
    Arreglo(0, 702) = "Gertrude"
    Arreglo(0, 703) = "Gia"
    Arreglo(0, 704) = "Giselle"
    Arreglo(0, 705) = "Glenda"
    Arreglo(0, 706) = "Glynis"
    Arreglo(0, 707) = "Grace"
    Arreglo(0, 708) = "Gwenda"
    Arreglo(0, 709) = "Gwendolen"
    Arreglo(0, 710) = "Gwendoline"
    Arreglo(0, 711) = "Gwendolyn"
    Arreglo(0, 712) = "Gwyneth"
    Arreglo(0, 713) = "Hannah"
    Arreglo(0, 714) = "Harriet"
    Arreglo(0, 715) = "Heather"
    Arreglo(0, 716) = "Heidi"
    Arreglo(0, 717) = "Helen"
    Arreglo(0, 718) = "Helena"
    Arreglo(0, 719) = "Helene"
    Arreglo(0, 720) = "Henrietta"
    Arreglo(0, 721) = "Hero"
    Arreglo(0, 722) = "Hester"
    Arreglo(0, 723) = "Hilary"
    Arreglo(0, 724) = "Hilda"
    Arreglo(0, 725) = "Hodierna"
    Arreglo(0, 726) = "Holly"
    Arreglo(0, 727) = "Honor"
    Arreglo(0, 728) = "Hope"
    Arreglo(0, 729) = "Hunter"
    Arreglo(0, 730) = "Ida"
    Arreglo(0, 731) = "Imelda"
    Arreglo(0, 732) = "Imogen"
    Arreglo(0, 733) = "Iona"
    Arreglo(0, 734) = "Irene"
    Arreglo(0, 735) = "Iris"
    Arreglo(0, 736) = "Isabella"
    Arreglo(0, 737) = "Isla"
    Arreglo(0, 738) = "Ivy"
    Arreglo(0, 739) = "Jack"
    Arreglo(0, 740) = "Jackie"
    Arreglo(0, 741) = "Jacqueline"
    Arreglo(0, 742) = "Jacqui"
    Arreglo(0, 743) = "Jaime"
    Arreglo(0, 744) = "Jamie"
    Arreglo(0, 745) = "Jan"
    Arreglo(0, 746) = "Jana"
    Arreglo(0, 747) = "Jane"
    Arreglo(0, 748) = "Janee"
    Arreglo(0, 749) = "Janey"
    Arreglo(0, 750) = "Janie"
    Arreglo(0, 751) = "Jasmine"
    Arreglo(0, 752) = "Jay"
    Arreglo(0, 753) = "Jayne"
    Arreglo(0, 754) = "Jaynie"
    Arreglo(0, 755) = "Jemima"
    Arreglo(0, 756) = "Jemma"
    Arreglo(0, 757) = "Jenna"
    Arreglo(0, 758) = "Jennifer"
    Arreglo(0, 759) = "Jenny"
    Arreglo(0, 760) = "Jerry"
    Arreglo(0, 761) = "Jess"
    Arreglo(0, 762) = "Jessica"
    Arreglo(0, 763) = "Jessie"
    Arreglo(0, 764) = "Joan"
    Arreglo(0, 765) = "Joanna"
    Arreglo(0, 766) = "Joanne"
    Arreglo(0, 767) = "Jodie"
    Arreglo(0, 768) = "Joelle"
    Arreglo(0, 769) = "Joey"
    Arreglo(0, 770) = "Johnny"
    Arreglo(0, 771) = "Jolie"
    Arreglo(0, 772) = "Jordan"
    Arreglo(0, 773) = "Josephine"
    Arreglo(0, 774) = "Josie"
    Arreglo(0, 775) = "Joy"
    Arreglo(0, 776) = "Judith"
    Arreglo(0, 777) = "Jules"
    Arreglo(0, 778) = "Julia"
    Arreglo(0, 779) = "Julianne"
    Arreglo(0, 780) = "Julie"
    Arreglo(0, 781) = "Kalla"
    Arreglo(0, 782) = "Karen"
    Arreglo(0, 783) = "Karina"
    Arreglo(0, 784) = "Karlee"
    Arreglo(0, 785) = "Karlene"
    Arreglo(0, 786) = "Karli"
    Arreglo(0, 787) = "Karlie"
    Arreglo(0, 788) = "Karly"
    Arreglo(0, 789) = "Karolyn"
    Arreglo(0, 790) = "Karrie"
    Arreglo(0, 791) = "Katey"
    Arreglo(0, 792) = "Kathleen"
    Arreglo(0, 793) = "Kathy"
    Arreglo(0, 794) = "Katie"
    Arreglo(0, 795) = "Katrina"
    Arreglo(0, 796) = "Kay"
    Arreglo(0, 797) = "Kaylee"
    Arreglo(0, 798) = "Kelsey"
    Arreglo(0, 799) = "Kierra"
    Arreglo(0, 800) = "Kim"
    Arreglo(0, 801) = "Kirsten"
    Arreglo(0, 802) = "Kirstin"
    Arreglo(0, 803) = "Kristen"
    Arreglo(0, 804) = "Kristi"
    Arreglo(0, 805) = "Kristin"
    Arreglo(0, 806) = "Lana"
    Arreglo(0, 807) = "Lanna"
    Arreglo(0, 808) = "Lara"
    Arreglo(0, 809) = "Laura"
    Arreglo(0, 810) = "Lauren"
    Arreglo(0, 811) = "Laurence"
    Arreglo(0, 812) = "Lauretta"
    Arreglo(0, 813) = "Laurie"
    Arreglo(0, 814) = "Leah"
    Arreglo(0, 815) = "Leanne"
    Arreglo(0, 816) = "Lee"
    Arreglo(0, 817) = "Leila"
    Arreglo(0, 818) = "Leisha"
    Arreglo(0, 819) = "Lena"
    Arreglo(0, 820) = "Lenna"
    Arreglo(0, 821) = "Leonora"
    Arreglo(0, 822) = "Leslie"
    Arreglo(0, 823) = "Lettice"
    Arreglo(0, 824) = "Liana"
    Arreglo(0, 825) = "Lila"
    Arreglo(0, 826) = "Lilla"
    Arreglo(0, 827) = "Lillian"
    Arreglo(0, 828) = "Lily"
    Arreglo(0, 829) = "Linda"
    Arreglo(0, 830) = "Lindsay"
    Arreglo(0, 831) = "Lisa"
    Arreglo(0, 832) = "Liza"
    Arreglo(0, 833) = "Lois"
    Arreglo(0, 834) = "Loraine"
    Arreglo(0, 835) = "Lorelei"
    Arreglo(0, 836) = "Lorena"
    Arreglo(0, 837) = "Loretta"
    Arreglo(0, 838) = "Lorinda"
    Arreglo(0, 839) = "Lorna"
    Arreglo(0, 840) = "Lorraine"
    Arreglo(0, 841) = "Lottie"
    Arreglo(0, 842) = "Lotty"
    Arreglo(0, 843) = "Louella"
    Arreglo(0, 844) = "Louisa"
    Arreglo(0, 845) = "Louise"
    Arreglo(0, 846) = "Lucia"
    Arreglo(0, 847) = "Lucinda"
    Arreglo(0, 848) = "Lucy"
    Arreglo(0, 849) = "Lynnette"
    Arreglo(0, 850) = "Lysette"
    Arreglo(0, 851) = "Mabel"
    Arreglo(0, 852) = "Madelaine"
    Arreglo(0, 853) = "Madge"
    Arreglo(0, 854) = "Maggie"
    Arreglo(0, 855) = "Mandy"
    Arreglo(0, 856) = "Marcia"
    Arreglo(0, 857) = "Marcie"
    Arreglo(0, 858) = "Margaret"
    Arreglo(0, 859) = "Mariah"
    Arreglo(0, 860) = "Marian"
    Arreglo(0, 861) = "Marianne"
    Arreglo(0, 862) = "Marie"
    Arreglo(0, 863) = "Marilyn"
    Arreglo(0, 864) = "Marina"
    Arreglo(0, 865) = "Marissa"
    Arreglo(0, 866) = "Marjorie"
    Arreglo(0, 867) = "Marsha"
    Arreglo(0, 868) = "Marta"
    Arreglo(0, 869) = "Mary"
    Arreglo(0, 870) = "Mason"
    Arreglo(0, 871) = "Matilda"
    Arreglo(0, 872) = "Maud"
    Arreglo(0, 873) = "Maude"
    Arreglo(0, 874) = "Maureen"
    Arreglo(0, 875) = "Mavis"
    Arreglo(0, 876) = "May"
    Arreglo(0, 877) = "Maya"
    Arreglo(0, 878) = "Mayola"
    Arreglo(0, 879) = "Medea"
    Arreglo(0, 880) = "Megan"
    Arreglo(0, 881) = "Mehitable"
    Arreglo(0, 882) = "Melanie"
    Arreglo(0, 883) = "Melissa"
    Arreglo(0, 884) = "Mercedes"
    Arreglo(0, 885) = "Merle"
    Arreglo(0, 886) = "Michele"
    Arreglo(0, 887) = "Michelle"
    Arreglo(0, 888) = "Mildred"
    Arreglo(0, 889) = "Millicent"
    Arreglo(0, 890) = "Minna"
    Arreglo(0, 891) = "Minnie"
    Arreglo(0, 892) = "Miranda"
    Arreglo(0, 893) = "Moira"
    Arreglo(0, 894) = "Morgan"
    Arreglo(0, 895) = "Myra"
    Arreglo(0, 896) = "Myrna"
    Arreglo(0, 897) = "Myrtle"
    Arreglo(0, 898) = "Nadine"
    Arreglo(0, 899) = "Naila"
    Arreglo(0, 900) = "Nancy"
    Arreglo(0, 901) = "Narcissa"
    Arreglo(0, 902) = "Natalie"
    Arreglo(0, 903) = "Nena"
    Arreglo(0, 904) = "Nettie"
    Arreglo(0, 905) = "Nia"
    Arreglo(0, 906) = "Nicola"
    Arreglo(0, 907) = "Nicole"
    Arreglo(0, 908) = "Nina"
    Arreglo(0, 909) = "Nora"
    Arreglo(0, 910) = "Odette"
    Arreglo(0, 911) = "Olivia"
    Arreglo(0, 912) = "Opal"
    Arreglo(0, 913) = "Patience"
    Arreglo(0, 914) = "Patrice"
    Arreglo(0, 915) = "Patsy"
    Arreglo(0, 916) = "Patty"
    Arreglo(0, 917) = "Paula"
    Arreglo(0, 918) = "Paulina"
    Arreglo(0, 919) = "Pearl"
    Arreglo(0, 920) = "Peggy"
    Arreglo(0, 921) = "Penelope"
    Arreglo(0, 922) = "Penny"
    Arreglo(0, 923) = "Persis"
    Arreglo(0, 924) = "Petunia"
    Arreglo(0, 925) = "Philippa"
    Arreglo(0, 926) = "Poppy"
    Arreglo(0, 927) = "Precious"
    Arreglo(0, 928) = "Priscilla"
    Arreglo(0, 929) = "Rachel"
    Arreglo(0, 930) = "Reba"
    Arreglo(0, 931) = "Rhiannon"
    Arreglo(0, 932) = "Rhoda"
    Arreglo(0, 933) = "Rhonda"
    Arreglo(0, 934) = "Richeldis"
    Arreglo(0, 935) = "Rita"
    Arreglo(0, 936) = "Roberta"
    Arreglo(0, 937) = "Robin"
    Arreglo(0, 938) = "Ronnie"
    Arreglo(0, 939) = "Rosamund"
    Arreglo(0, 940) = "Rose"
    Arreglo(0, 941) = "Rosemary"
    Arreglo(0, 942) = "Ruth"
    Arreglo(0, 943) = "Sabrina"
    Arreglo(0, 944) = "Sadie"
    Arreglo(0, 945) = "Salma"
    Arreglo(0, 946) = "Sam"
    Arreglo(0, 947) = "Samantha"
    Arreglo(0, 948) = "Sandra"
    Arreglo(0, 949) = "Sarah"
    Arreglo(0, 950) = "Selma"
    Arreglo(0, 951) = "Serena"
    Arreglo(0, 952) = "Shania"
    Arreglo(0, 953) = "Shannon"
    Arreglo(0, 954) = "Sharla"
    Arreglo(0, 955) = "Sharleen"
    Arreglo(0, 956) = "Sharlene"
    Arreglo(0, 957) = "Sharon"
    Arreglo(0, 958) = "Shawna"
    Arreglo(0, 959) = "Sheryl"
    Arreglo(0, 960) = "Sibyl"
    Arreglo(0, 961) = "Simone"
    Arreglo(0, 962) = "Skyler"
    Arreglo(0, 963) = "Sophia"
    Arreglo(0, 964) = "Sophie"
    Arreglo(0, 965) = "Sorrel"
    Arreglo(0, 966) = "Stella"
    Arreglo(0, 967) = "Stevie"
    Arreglo(0, 968) = "Summer"
    Arreglo(0, 969) = "Susan"
    Arreglo(0, 970) = "Susanna"
    Arreglo(0, 971) = "Susanne"
    Arreglo(0, 972) = "Suzanne"
    Arreglo(0, 973) = "Sylvia"
    Arreglo(0, 974) = "Talitha"
    Arreglo(0, 975) = "Tallulah"
    Arreglo(0, 976) = "Tamara"
    Arreglo(0, 977) = "Tammy"
    Arreglo(0, 978) = "Tara"
    Arreglo(0, 979) = "Teresa"
    Arreglo(0, 980) = "Terry"
    Arreglo(0, 981) = "Thelma"
    Arreglo(0, 982) = "Thomasina"
    Arreglo(0, 983) = "Thurza"
    Arreglo(0, 984) = "Tiffany"
    Arreglo(0, 985) = "Tina"
    Arreglo(0, 986) = "Tonja"
    Arreglo(0, 987) = "Tonya"
    Arreglo(0, 988) = "Tracy"
    Arreglo(0, 989) = "Trisha"
    Arreglo(0, 990) = "Tyler"
    Arreglo(0, 991) = "Tyra"
    Arreglo(0, 992) = "Urith"
    Arreglo(0, 993) = "Valerie"
    Arreglo(0, 994) = "Vanessa"
    Arreglo(0, 995) = "Venetia"
    Arreglo(0, 996) = "Vera"
    Arreglo(0, 997) = "Victoria"
    Arreglo(0, 998) = "Vilma"
    Arreglo(0, 999) = "Viola"
    Arreglo(0, 1000) = "Violette"
    Arreglo(0, 1001) = "Virginia"
    Arreglo(0, 1002) = "Wanda"
    Arreglo(0, 1003) = "Wendy"
    Arreglo(0, 1004) = "Whitney"
    Arreglo(0, 1005) = "Wilma"
    Arreglo(0, 1006) = "Winifred"
    Arreglo(0, 1007) = "Winnie"
    Arreglo(0, 1008) = "Winnifred"
    Arreglo(0, 1009) = "Yasmin"
    Arreglo(0, 1010) = "Yvette"
    Arreglo(0, 1011) = "Yvonne"
    Arreglo(0, 1012) = "Zelda"


    Set Myrange = Range("A2:B1014")
    i = 0
    j = 0
    
    For Each Cell In Myrange

        If j Mod 2 = 0 Then

            Cell.Value = i + 1

            i = i + 1
            j = j + 1
        Else

            Cell.Value = Arreglo(0, i - 1)

            j = j + 1
        End If

    Next Cell
        
        
    Range("A1").Value = "Order"
    Range("A1").Font.Bold = True

    Range("B1").Value = "Name"
    Range("B1").Font.Bold = True

    Columns("A:A").EntireColumn.AutoFit
    Columns("A:A").HorizontalAlignment = xlCenter

    Columns("B:B").EntireColumn.AutoFit

    Range("A2").Select
    ActiveWindow.FreezePanes = True


End Sub

Sub MacroN_menos_1_Fill_LastNames()

    Sheets("Last Names").Select

    Dim Myrange As Range
    Dim Cell As Range
    Dim i As Double
    Dim j As Double
    
    Dim Arreglo(0, 2011) As String
    
    
    Arreglo(0, 0) = "Abbott"
    Arreglo(0, 1) = "Abeita"
    Arreglo(0, 2) = "Abel"
    Arreglo(0, 3) = "Abeyta"
    Arreglo(0, 4) = "Abraham"
    Arreglo(0, 5) = "Abrahamson"
    Arreglo(0, 6) = "Abrams"
    Arreglo(0, 7) = "Ackerman"
    Arreglo(0, 8) = "Acosta"
    Arreglo(0, 9) = "Adair"
    Arreglo(0, 10) = "Adakai"
    Arreglo(0, 11) = "Adams"
    Arreglo(0, 12) = "Addison"
    Arreglo(0, 13) = "Adkins"
    Arreglo(0, 14) = "Aguilar"
    Arreglo(0, 15) = "Aguirre"
    Arreglo(0, 16) = "Akers"
    Arreglo(0, 17) = "Albert"
    Arreglo(0, 18) = "Aldrich"
    Arreglo(0, 19) = "Alexander"
    Arreglo(0, 20) = "Alexie"
    Arreglo(0, 21) = "Alford"
    Arreglo(0, 22) = "Ali"
    Arreglo(0, 23) = "Allard"
    Arreglo(0, 24) = "Allen"
    Arreglo(0, 25) = "Allery"
    Arreglo(0, 26) = "Allison"
    Arreglo(0, 27) = "Alonzo"
    Arreglo(0, 28) = "Altaha"
    Arreglo(0, 29) = "Alvarado"
    Arreglo(0, 30) = "Alvarez"
    Arreglo(0, 31) = "Ambrose"
    Arreglo(0, 32) = "Americanhorse"
    Arreglo(0, 33) = "Ames"
    Arreglo(0, 34) = "Ammons"
    Arreglo(0, 35) = "Amos"
    Arreglo(0, 36) = "Andersen"
    Arreglo(0, 37) = "Anderson"
    Arreglo(0, 38) = "Andrew"
    Arreglo(0, 39) = "Andrews"
    Arreglo(0, 40) = "Antelope"
    Arreglo(0, 41) = "Anthony"
    Arreglo(0, 42) = "Antoine"
    Arreglo(0, 43) = "Antone"
    Arreglo(0, 44) = "Antonio"
    Arreglo(0, 45) = "Apache"
    Arreglo(0, 46) = "Apachito"
    Arreglo(0, 47) = "Apodaca"
    Arreglo(0, 48) = "Apple"
    Arreglo(0, 49) = "Aragon"
    Arreglo(0, 50) = "Archambault"
    Arreglo(0, 51) = "Archambeau"
    Arreglo(0, 52) = "Archer"
    Arreglo(0, 53) = "Archuleta"
    Arreglo(0, 54) = "Armijo"
    Arreglo(0, 55) = "Armstrong"
    Arreglo(0, 56) = "Arnold"
    Arreglo(0, 57) = "Arthur"
    Arreglo(0, 58) = "Arviso"
    Arreglo(0, 59) = "Ashley"
    Arreglo(0, 60) = "Atcitty"
    Arreglo(0, 61) = "Atencio"
    Arreglo(0, 62) = "Atene"
    Arreglo(0, 63) = "Atkins"
    Arreglo(0, 64) = "Atkinson"
    Arreglo(0, 65) = "Attakai"
    Arreglo(0, 66) = "Augustine"
    Arreglo(0, 67) = "Austin"
    Arreglo(0, 68) = "Avery"
    Arreglo(0, 69) = "Avila"
    Arreglo(0, 70) = "Ayala"
    Arreglo(0, 71) = "Ayers"
    Arreglo(0, 72) = "Azure"
    Arreglo(0, 73) = "Baca"
    Arreglo(0, 74) = "Bacon"
    Arreglo(0, 75) = "Bahe"
    Arreglo(0, 76) = "Bailey"
    Arreglo(0, 77) = "Baird"
    Arreglo(0, 78) = "Baker"
    Arreglo(0, 79) = "Baldwin"
    Arreglo(0, 80) = "Bales"
    Arreglo(0, 81) = "Ball"
    Arreglo(0, 82) = "Ballard"
    Arreglo(0, 83) = "Banks"
    Arreglo(0, 84) = "Barber"
    Arreglo(0, 85) = "Barbone"
    Arreglo(0, 86) = "Barger"
    Arreglo(0, 87) = "Barker"
    Arreglo(0, 88) = "Barlow"
    Arreglo(0, 89) = "Barnes"
    Arreglo(0, 90) = "Barnett"
    Arreglo(0, 91) = "Barney"
    Arreglo(0, 92) = "Barr"
    Arreglo(0, 93) = "Barrett"
    Arreglo(0, 94) = "Barron"
    Arreglo(0, 95) = "Barry"
    Arreglo(0, 96) = "Bartlett"
    Arreglo(0, 97) = "Bartley"
    Arreglo(0, 98) = "Barton"
    Arreglo(0, 99) = "Bass"
    Arreglo(0, 100) = "Bates"
    Arreglo(0, 101) = "Battiest"
    Arreglo(0, 102) = "Battise"
    Arreglo(0, 103) = "Bauer"
    Arreglo(0, 104) = "Bautista"
    Arreglo(0, 105) = "Baxter"
    Arreglo(0, 106) = "Beach"
    Arreglo(0, 107) = "Bean"
    Arreglo(0, 108) = "Bear"
    Arreglo(0, 109) = "Beard"
    Arreglo(0, 110) = "Beasley"
    Arreglo(0, 111) = "Beatty"
    Arreglo(0, 112) = "Beaulieu"
    Arreglo(0, 113) = "Beaver"
    Arreglo(0, 114) = "Beavers"
    Arreglo(0, 115) = "Becenti"
    Arreglo(0, 116) = "Beck"
    Arreglo(0, 117) = "Becker"
    Arreglo(0, 118) = "Bedoni"
    Arreglo(0, 119) = "Bedonie"
    Arreglo(0, 120) = "Begay"
    Arreglo(0, 121) = "Begaye"
    Arreglo(0, 122) = "Belcher"
    Arreglo(0, 123) = "Belcourt"
    Arreglo(0, 124) = "Belgarde"
    Arreglo(0, 125) = "Belin"
    Arreglo(0, 126) = "Bell"
    Arreglo(0, 127) = "Bellanger"
    Arreglo(0, 128) = "Belone"
    Arreglo(0, 129) = "Ben"
    Arreglo(0, 130) = "Benallie"
    Arreglo(0, 131) = "Benally"
    Arreglo(0, 132) = "Bender"
    Arreglo(0, 133) = "Benedict"
    Arreglo(0, 134) = "Benjamin"
    Arreglo(0, 135) = "Bennett"
    Arreglo(0, 136) = "Benson"
    Arreglo(0, 137) = "Bentley"
    Arreglo(0, 138) = "Benton"
    Arreglo(0, 139) = "Bercier"
    Arreglo(0, 140) = "Berg"
    Arreglo(0, 141) = "Berger"
    Arreglo(0, 142) = "Bernard"
    Arreglo(0, 143) = "Berry"
    Arreglo(0, 144) = "Berryhill"
    Arreglo(0, 145) = "Best"
    Arreglo(0, 146) = "Bettelyoun"
    Arreglo(0, 147) = "Beyale"
    Arreglo(0, 148) = "Bia"
    Arreglo(0, 149) = "Bible"
    Arreglo(0, 150) = "Bigcrow"
    Arreglo(0, 151) = "Bigeagle"
    Arreglo(0, 152) = "Biggs"
    Arreglo(0, 153) = "Bighorse"
    Arreglo(0, 154) = "Bigman"
    Arreglo(0, 155) = "Bill"
    Arreglo(0, 156) = "Billie"
    Arreglo(0, 157) = "Billings"
    Arreglo(0, 158) = "Billiot"
    Arreglo(0, 159) = "Billy"
    Arreglo(0, 160) = "Bingham"
    Arreglo(0, 161) = "Bird"
    Arreglo(0, 162) = "Bishop"
    Arreglo(0, 163) = "Bissonette"
    Arreglo(0, 164) = "Bitsie"
    Arreglo(0, 165) = "Bitsilly"
    Arreglo(0, 166) = "Bitsui"
    Arreglo(0, 167) = "Bitsuie"
    Arreglo(0, 168) = "Black"
    Arreglo(0, 169) = "Blackbear"
    Arreglo(0, 170) = "Blackbird"
    Arreglo(0, 171) = "Blackburn"
    Arreglo(0, 172) = "Blackgoat"
    Arreglo(0, 173) = "Blackhorse"
    Arreglo(0, 174) = "Blacksmith"
    Arreglo(0, 175) = "Blackwater"
    Arreglo(0, 176) = "Blackwell"
    Arreglo(0, 177) = "Blackwolf"
    Arreglo(0, 178) = "Blaine"
    Arreglo(0, 179) = "Blair"
    Arreglo(0, 180) = "Blake"
    Arreglo(0, 181) = "Blanchard"
    Arreglo(0, 182) = "Blankenship"
    Arreglo(0, 183) = "Blanks"
    Arreglo(0, 184) = "Blaylock"
    Arreglo(0, 185) = "Blevins"
    Arreglo(0, 186) = "Blossom"
    Arreglo(0, 187) = "Blue"
    Arreglo(0, 188) = "Bluebird"
    Arreglo(0, 189) = "Blueeyes"
    Arreglo(0, 190) = "Bob"
    Arreglo(0, 191) = "Boggs"
    Arreglo(0, 192) = "Bolin"
    Arreglo(0, 193) = "Bolton"
    Arreglo(0, 194) = "Bond"
    Arreglo(0, 195) = "Boone"
    Arreglo(0, 196) = "Booth"
    Arreglo(0, 197) = "Bordeaux"
    Arreglo(0, 198) = "Boswell"
    Arreglo(0, 199) = "Bowen"
    Arreglo(0, 200) = "Bowers"
    Arreglo(0, 201) = "Bowman"
    Arreglo(0, 202) = "Boyd"
    Arreglo(0, 203) = "Boyer"
    Arreglo(0, 204) = "Brackett"
    Arreglo(0, 205) = "Bradford"
    Arreglo(0, 206) = "Bradley"
    Arreglo(0, 207) = "Bradshaw"
    Arreglo(0, 208) = "Brady"
    Arreglo(0, 209) = "Branch"
    Arreglo(0, 210) = "Brandon"
    Arreglo(0, 211) = "Branham"
    Arreglo(0, 212) = "Brant"
    Arreglo(0, 213) = "Brave"
    Arreglo(0, 214) = "Bray"
    Arreglo(0, 215) = "Brayboy"
    Arreglo(0, 216) = "Brennan"
    Arreglo(0, 217) = "Brewer"
    Arreglo(0, 218) = "Brewington"
    Arreglo(0, 219) = "Brewster"
    Arreglo(0, 220) = "Bridges"
    Arreglo(0, 221) = "Brien"
    Arreglo(0, 222) = "Briggs"
    Arreglo(0, 223) = "Bright"
    Arreglo(0, 224) = "Brink"
    Arreglo(0, 225) = "Britt"
    Arreglo(0, 226) = "Britton"
    Arreglo(0, 227) = "Brock"
    Arreglo(0, 228) = "Brooks"
    Arreglo(0, 229) = "Brower"
    Arreglo(0, 230) = "Brown"
    Arreglo(0, 231) = "Browning"
    Arreglo(0, 232) = "Bruce"
    Arreglo(0, 233) = "Bruner"
    Arreglo(0, 234) = "Bruno"
    Arreglo(0, 235) = "Bryan"
    Arreglo(0, 236) = "Bryant"
    Arreglo(0, 237) = "Buchanan"
    Arreglo(0, 238) = "Buck"
    Arreglo(0, 239) = "Buckley"
    Arreglo(0, 240) = "Buckman"
    Arreglo(0, 241) = "Buckner"
    Arreglo(0, 242) = "Buffalo"
    Arreglo(0, 243) = "Bull"
    Arreglo(0, 244) = "Bullard"
    Arreglo(0, 245) = "Bullock"
    Arreglo(0, 246) = "Bunch"
    Arreglo(0, 247) = "Burbank"
    Arreglo(0, 248) = "Burch"
    Arreglo(0, 249) = "Burgess"
    Arreglo(0, 250) = "Burke"
    Arreglo(0, 251) = "Burnett"
    Arreglo(0, 252) = "Burnette"
    Arreglo(0, 253) = "Burns"
    Arreglo(0, 254) = "Burnside"
    Arreglo(0, 255) = "Burr"
    Arreglo(0, 256) = "Burris"
    Arreglo(0, 257) = "Burrows"
    Arreglo(0, 258) = "Burton"
    Arreglo(0, 259) = "Bush"
    Arreglo(0, 260) = "Butcher"
    Arreglo(0, 261) = "Butler"
    Arreglo(0, 262) = "Buzzard"
    Arreglo(0, 263) = "Byers"
    Arreglo(0, 264) = "Byrd"
    Arreglo(0, 265) = "Cadman"
    Arreglo(0, 266) = "Cadotte"
    Arreglo(0, 267) = "Cagle"
    Arreglo(0, 268) = "Cain"
    Arreglo(0, 269) = "Calabaza"
    Arreglo(0, 270) = "Caldwell"
    Arreglo(0, 271) = "Calhoun"
    Arreglo(0, 272) = "Callahan"
    Arreglo(0, 273) = "Cameron"
    Arreglo(0, 274) = "Camp"
    Arreglo(0, 275) = "Campbell"
    Arreglo(0, 276) = "Campos"
    Arreglo(0, 277) = "Candelaria"
    Arreglo(0, 278) = "Cannon"
    Arreglo(0, 279) = "Cantrell"
    Arreglo(0, 280) = "Capps"
    Arreglo(0, 281) = "Carey"
    Arreglo(0, 282) = "Carl"
    Arreglo(0, 283) = "Carlson"
    Arreglo(0, 284) = "Carman"
    Arreglo(0, 285) = "Carney"
    Arreglo(0, 286) = "Carpenter"
    Arreglo(0, 287) = "Carr"
    Arreglo(0, 288) = "Carrillo"
    Arreglo(0, 289) = "Carroll"
    Arreglo(0, 290) = "Carson"
    Arreglo(0, 291) = "Carter"
    Arreglo(0, 292) = "Cartwright"
    Arreglo(0, 293) = "Carver"
    Arreglo(0, 294) = "Case"
    Arreglo(0, 295) = "Casey"
    Arreglo(0, 296) = "Cash"
    Arreglo(0, 297) = "Castaneda"
    Arreglo(0, 298) = "Castillo"
    Arreglo(0, 299) = "Castro"
    Arreglo(0, 300) = "Cates"
    Arreglo(0, 301) = "Catron"
    Arreglo(0, 302) = "Cavanaugh"
    Arreglo(0, 303) = "Cayaditto"
    Arreglo(0, 304) = "Cervantes"
    Arreglo(0, 305) = "Chacon"
    Arreglo(0, 306) = "Chamberlain"
    Arreglo(0, 307) = "Chambers"
    Arreglo(0, 308) = "Champagne"
    Arreglo(0, 309) = "Chance"
    Arreglo(0, 310) = "Chandler"
    Arreglo(0, 311) = "Chaney"
    Arreglo(0, 312) = "Chapman"
    Arreglo(0, 313) = "Charette"
    Arreglo(0, 314) = "Charles"
    Arreglo(0, 315) = "Charley"
    Arreglo(0, 316) = "Charlie"
    Arreglo(0, 317) = "Chase"
    Arreglo(0, 318) = "Chasinghawk"
    Arreglo(0, 319) = "Chavez"
    Arreglo(0, 320) = "Chavis"
    Arreglo(0, 321) = "Chee"
    Arreglo(0, 322) = "Cheek"
    Arreglo(0, 323) = "Cheney"
    Arreglo(0, 324) = "Cheromiah"
    Arreglo(0, 325) = "Cherry"
    Arreglo(0, 326) = "Chester"
    Arreglo(0, 327) = "Chico"
    Arreglo(0, 328) = "Chief"
    Arreglo(0, 329) = "Childers"
    Arreglo(0, 330) = "Childress"
    Arreglo(0, 331) = "Childs"
    Arreglo(0, 332) = "Chinana"
    Arreglo(0, 333) = "Chino"
    Arreglo(0, 334) = "Chiquito"
    Arreglo(0, 335) = "Chischilly"
    Arreglo(0, 336) = "Choate"
    Arreglo(0, 337) = "Chosa"
    Arreglo(0, 338) = "Christensen"
    Arreglo(0, 339) = "Christian"
    Arreglo(0, 340) = "Christiansen"
    Arreglo(0, 341) = "Christie"
    Arreglo(0, 342) = "Chuculate"
    Arreglo(0, 343) = "Church"
    Arreglo(0, 344) = "Cisco"
    Arreglo(0, 345) = "Clah"
    Arreglo(0, 346) = "Clairmont"
    Arreglo(0, 347) = "Clark"
    Arreglo(0, 348) = "Clarke"
    Arreglo(0, 349) = "Claw"
    Arreglo(0, 350) = "Clay"
    Arreglo(0, 351) = "Claymore"
    Arreglo(0, 352) = "Clayton"
    Arreglo(0, 353) = "Clement"
    Arreglo(0, 354) = "Clements"
    Arreglo(0, 355) = "Clemons"
    Arreglo(0, 356) = "Cleveland"
    Arreglo(0, 357) = "Clifford"
    Arreglo(0, 358) = "Clifton"
    Arreglo(0, 359) = "Cline"
    Arreglo(0, 360) = "Clinton"
    Arreglo(0, 361) = "Clitso"
    Arreglo(0, 362) = "Cloud"
    Arreglo(0, 363) = "Cly"
    Arreglo(0, 364) = "Clyde"
    Arreglo(0, 365) = "Coats"
    Arreglo(0, 366) = "Cobb"
    Arreglo(0, 367) = "Cochran"
    Arreglo(0, 368) = "Cody"
    Arreglo(0, 369) = "Coffey"
    Arreglo(0, 370) = "Coffman"
    Arreglo(0, 371) = "Coho"
    Arreglo(0, 372) = "Coker"
    Arreglo(0, 373) = "Colbert"
    Arreglo(0, 374) = "Cole"
    Arreglo(0, 375) = "Coleman"
    Arreglo(0, 376) = "Coley"
    Arreglo(0, 377) = "Collier"
    Arreglo(0, 378) = "Collins"
    Arreglo(0, 379) = "Combs"
    Arreglo(0, 380) = "Compton"
    Arreglo(0, 381) = "Concha"
    Arreglo(0, 382) = "Condon"
    Arreglo(0, 383) = "Conklin"
    Arreglo(0, 384) = "Conley"
    Arreglo(0, 385) = "Conn"
    Arreglo(0, 386) = "Conner"
    Arreglo(0, 387) = "Connor"
    Arreglo(0, 388) = "Connors"
    Arreglo(0, 389) = "Conrad"
    Arreglo(0, 390) = "Contreras"
    Arreglo(0, 391) = "Conway"
    Arreglo(0, 392) = "Cook"
    Arreglo(0, 393) = "Cooke"
    Arreglo(0, 394) = "Cooley"
    Arreglo(0, 395) = "Coon"
    Arreglo(0, 396) = "Cooper"
    Arreglo(0, 397) = "Copeland"
    Arreglo(0, 398) = "Corbett"
    Arreglo(0, 399) = "Cordova"
    Arreglo(0, 400) = "Coriz"
    Arreglo(0, 401) = "Corn"
    Arreglo(0, 402) = "Cornelius"
    Arreglo(0, 403) = "Cortez"
    Arreglo(0, 404) = "Cosay"
    Arreglo(0, 405) = "Cote"
    Arreglo(0, 406) = "Cotton"
    Arreglo(0, 407) = "Couch"
    Arreglo(0, 408) = "Counts"
    Arreglo(0, 409) = "Cournoyer"
    Arreglo(0, 410) = "Courtney"
    Arreglo(0, 411) = "Couture"
    Arreglo(0, 412) = "Covington"
    Arreglo(0, 413) = "Cowan"
    Arreglo(0, 414) = "Cowboy"
    Arreglo(0, 415) = "Cox"
    Arreglo(0, 416) = "Crabtree"
    Arreglo(0, 417) = "Craft"
    Arreglo(0, 418) = "Craig"
    Arreglo(0, 419) = "Crain"
    Arreglo(0, 420) = "Cramer"
    Arreglo(0, 421) = "Crane"
    Arreglo(0, 422) = "Crank"
    Arreglo(0, 423) = "Crawford"
    Arreglo(0, 424) = "Cree"
    Arreglo(0, 425) = "Creel"
    Arreglo(0, 426) = "Crespin"
    Arreglo(0, 427) = "Crittenden"
    Arreglo(0, 428) = "Cromwell"
    Arreglo(0, 429) = "Crosby"
    Arreglo(0, 430) = "Cross"
    Arreglo(0, 431) = "Crow"
    Arreglo(0, 432) = "Crowe"
    Arreglo(0, 433) = "Crutcher"
    Arreglo(0, 434) = "Cruz"
    Arreglo(0, 435) = "Cuch"
    Arreglo(0, 436) = "Cummings"
    Arreglo(0, 437) = "Cummins"
    Arreglo(0, 438) = "Cunningham"
    Arreglo(0, 439) = "Curley"
    Arreglo(0, 440) = "Curry"
    Arreglo(0, 441) = "Curtis"
    Arreglo(0, 442) = "Dahl"
    Arreglo(0, 443) = "Dailey"
    Arreglo(0, 444) = "Dale"
    Arreglo(0, 445) = "Dallas"
    Arreglo(0, 446) = "Dalton"
    Arreglo(0, 447) = "Damon"
    Arreglo(0, 448) = "Dan"
    Arreglo(0, 449) = "Danforth"
    Arreglo(0, 450) = "Daniel"
    Arreglo(0, 451) = "Daniels"
    Arreglo(0, 452) = "Dardar"
    Arreglo(0, 453) = "Daugherty"
    Arreglo(0, 454) = "Davenport"
    Arreglo(0, 455) = "David"
    Arreglo(0, 456) = "Davidson"
    Arreglo(0, 457) = "Davis"
    Arreglo(0, 458) = "Davison"
    Arreglo(0, 459) = "Dawes"
    Arreglo(0, 460) = "Dawson"
    Arreglo(0, 461) = "Day"
    Arreglo(0, 462) = "Deal"
    Arreglo(0, 463) = "Dean"
    Arreglo(0, 464) = "Decker"
    Arreglo(0, 465) = "Declay"
    Arreglo(0, 466) = "Decorah"
    Arreglo(0, 467) = "Decoteau"
    Arreglo(0, 468) = "Dedman"
    Arreglo(0, 469) = "Dee"
    Arreglo(0, 470) = "Deer"
    Arreglo(0, 471) = "Deere"
    Arreglo(0, 472) = "Deese"
    Arreglo(0, 473) = "Defoe"
    Arreglo(0, 474) = "Degroat"
    Arreglo(0, 475) = "Deleon"
    Arreglo(0, 476) = "Delgado"
    Arreglo(0, 477) = "Delgarito"
    Arreglo(0, 478) = "Delong"
    Arreglo(0, 479) = "Delorme"
    Arreglo(0, 480) = "Demarrias"
    Arreglo(0, 481) = "Demery"
    Arreglo(0, 482) = "Demientieff"
    Arreglo(0, 483) = "Dempsey"
    Arreglo(0, 484) = "Dennis"
    Arreglo(0, 485) = "Dennison"
    Arreglo(0, 486) = "Denny"
    Arreglo(0, 487) = "Denson"
    Arreglo(0, 488) = "Denton"
    Arreglo(0, 489) = "Desjarlais"
    Arreglo(0, 490) = "Devine"
    Arreglo(0, 491) = "Devore"
    Arreglo(0, 492) = "Dewey"
    Arreglo(0, 493) = "Dewitt"
    Arreglo(0, 494) = "Dial"
    Arreglo(0, 495) = "Diaz"
    Arreglo(0, 496) = "Dick"
    Arreglo(0, 497) = "Dickens"
    Arreglo(0, 498) = "Dickerson"
    Arreglo(0, 499) = "Dickson"
    Arreglo(0, 500) = "Dill"
    Arreglo(0, 501) = "Dillard"
    Arreglo(0, 502) = "Dillon"
    Arreglo(0, 503) = "Dion"
    Arreglo(0, 504) = "Dionne"
    Arreglo(0, 505) = "Diver"
    Arreglo(0, 506) = "Dixon"
    Arreglo(0, 507) = "Dobbs"
    Arreglo(0, 508) = "Doctor"
    Arreglo(0, 509) = "Dodd"
    Arreglo(0, 510) = "Dodge"
    Arreglo(0, 511) = "Dodson"
    Arreglo(0, 512) = "Domingo"
    Arreglo(0, 513) = "Dominguez"
    Arreglo(0, 514) = "Donahue"
    Arreglo(0, 515) = "Donaldson"
    Arreglo(0, 516) = "Doney"
    Arreglo(0, 517) = "Dooley"
    Arreglo(0, 518) = "Dorsey"
    Arreglo(0, 519) = "Dosela"
    Arreglo(0, 520) = "Dotson"
    Arreglo(0, 521) = "Douglas"
    Arreglo(0, 522) = "Dow"
    Arreglo(0, 523) = "Downey"
    Arreglo(0, 524) = "Downing"
    Arreglo(0, 525) = "Downs"
    Arreglo(0, 526) = "Doxtator"
    Arreglo(0, 527) = "Doyle"
    Arreglo(0, 528) = "Drake"
    Arreglo(0, 529) = "Drapeau"
    Arreglo(0, 530) = "Draper"
    Arreglo(0, 531) = "Drew"
    Arreglo(0, 532) = "Driver"
    Arreglo(0, 533) = "Dry"
    Arreglo(0, 534) = "Drywater"
    Arreglo(0, 535) = "Dubois"
    Arreglo(0, 536) = "Dubray"
    Arreglo(0, 537) = "Ducheneaux"
    Arreglo(0, 538) = "Dudley"
    Arreglo(0, 539) = "Duffy"
    Arreglo(0, 540) = "Duke"
    Arreglo(0, 541) = "Dumarce"
    Arreglo(0, 542) = "Dumont"
    Arreglo(0, 543) = "Duncan"
    Arreglo(0, 544) = "Dunham"
    Arreglo(0, 545) = "Dunlap"
    Arreglo(0, 546) = "Dunn"
    Arreglo(0, 547) = "Dupree"
    Arreglo(0, 548) = "Dupris"
    Arreglo(0, 549) = "Duran"
    Arreglo(0, 550) = "Durant"
    Arreglo(0, 551) = "Durham"
    Arreglo(0, 552) = "Duvall"
    Arreglo(0, 553) = "Dye"
    Arreglo(0, 554) = "Dyer"
    Arreglo(0, 555) = "Eagle"
    Arreglo(0, 556) = "Eagleman"
    Arreglo(0, 557) = "Easley"
    Arreglo(0, 558) = "Eastman"
    Arreglo(0, 559) = "Eaton"
    Arreglo(0, 560) = "Ebarb"
    Arreglo(0, 561) = "Eddy"
    Arreglo(0, 562) = "Edison"
    Arreglo(0, 563) = "Edmo"
    Arreglo(0, 564) = "Edmonds"
    Arreglo(0, 565) = "Edwards"
    Arreglo(0, 566) = "Eldridge"
    Arreglo(0, 567) = "Elkins"
    Arreglo(0, 568) = "Elliott"
    Arreglo(0, 569) = "Ellis"
    Arreglo(0, 570) = "Ellison"
    Arreglo(0, 571) = "Ellsworth"
    Arreglo(0, 572) = "Emanuel"
    Arreglo(0, 573) = "Emerson"
    Arreglo(0, 574) = "Emery"
    Arreglo(0, 575) = "England"
    Arreglo(0, 576) = "English"
    Arreglo(0, 577) = "Enno"
    Arreglo(0, 578) = "Enos"
    Arreglo(0, 579) = "Epperson"
    Arreglo(0, 580) = "Epps"
    Arreglo(0, 581) = "Eriacho"
    Arreglo(0, 582) = "Erickson"
    Arreglo(0, 583) = "Ervin"
    Arreglo(0, 584) = "Erwin"
    Arreglo(0, 585) = "Escalante"
    Arreglo(0, 586) = "Espinoza"
    Arreglo(0, 587) = "Esquibel"
    Arreglo(0, 588) = "Estes"
    Arreglo(0, 589) = "Estrada"
    Arreglo(0, 590) = "Etcitty"
    Arreglo(0, 591) = "Ethelbah"
    Arreglo(0, 592) = "Etsitty"
    Arreglo(0, 593) = "Eubanks"
    Arreglo(0, 594) = "Evan"
    Arreglo(0, 595) = "Evans"
    Arreglo(0, 596) = "Everett"
    Arreglo(0, 597) = "Ewing"
    Arreglo(0, 598) = "Factor"
    Arreglo(0, 599) = "Fairbanks"
    Arreglo(0, 600) = "Falcon"
    Arreglo(0, 601) = "Farley"
    Arreglo(0, 602) = "Farmer"
    Arreglo(0, 603) = "Farrell"
    Arreglo(0, 604) = "Farris"
    Arreglo(0, 605) = "Fasthorse"
    Arreglo(0, 606) = "Faulkner"
    Arreglo(0, 607) = "Feather"
    Arreglo(0, 608) = "Feathers"
    Arreglo(0, 609) = "Felix"
    Arreglo(0, 610) = "Ferguson"
    Arreglo(0, 611) = "Fernandez"
    Arreglo(0, 612) = "Ferrell"
    Arreglo(0, 613) = "Ferris"
    Arreglo(0, 614) = "Fields"
    Arreglo(0, 615) = "Figueroa"
    Arreglo(0, 616) = "Finley"
    Arreglo(0, 617) = "Fischer"
    Arreglo(0, 618) = "Fish"
    Arreglo(0, 619) = "Fisher"
    Arreglo(0, 620) = "Fitch"
    Arreglo(0, 621) = "Fitzgerald"
    Arreglo(0, 622) = "Fitzpatrick"
    Arreglo(0, 623) = "Fixico"
    Arreglo(0, 624) = "Fleming"
    Arreglo(0, 625) = "Fletcher"
    Arreglo(0, 626) = "Flood"
    Arreglo(0, 627) = "Flores"
    Arreglo(0, 628) = "Flowers"
    Arreglo(0, 629) = "Floyd"
    Arreglo(0, 630) = "Flute"
    Arreglo(0, 631) = "Flynn"
    Arreglo(0, 632) = "Foley"
    Arreglo(0, 633) = "Folsom"
    Arreglo(0, 634) = "Foote"
    Arreglo(0, 635) = "Forbes"
    Arreglo(0, 636) = "Ford"
    Arreglo(0, 637) = "Foreman"
    Arreglo(0, 638) = "Forrest"
    Arreglo(0, 639) = "Foster"
    Arreglo(0, 640) = "Fourkiller"
    Arreglo(0, 641) = "Fowler"
    Arreglo(0, 642) = "Fox"
    Arreglo(0, 643) = "Fragua"
    Arreglo(0, 644) = "Francis"
    Arreglo(0, 645) = "Francisco"
    Arreglo(0, 646) = "Franco"
    Arreglo(0, 647) = "Frank"
    Arreglo(0, 648) = "Franklin"
    Arreglo(0, 649) = "Franks"
    Arreglo(0, 650) = "Fraser"
    Arreglo(0, 651) = "Frazier"
    Arreglo(0, 652) = "Fred"
    Arreglo(0, 653) = "Frederick"
    Arreglo(0, 654) = "Fredericks"
    Arreglo(0, 655) = "Free"
    Arreglo(0, 656) = "Freeland"
    Arreglo(0, 657) = "Freeman"
    Arreglo(0, 658) = "Freemont"
    Arreglo(0, 659) = "French"
    Arreglo(0, 660) = "Friday"
    Arreglo(0, 661) = "Friend"
    Arreglo(0, 662) = "Fritz"
    Arreglo(0, 663) = "Frost"
    Arreglo(0, 664) = "Fry"
    Arreglo(0, 665) = "Frye"
    Arreglo(0, 666) = "Fuentes"
    Arreglo(0, 667) = "Fuller"
    Arreglo(0, 668) = "Fulton"
    Arreglo(0, 669) = "Funmaker"
    Arreglo(0, 670) = "Gachupin"
    Arreglo(0, 671) = "Gaddy"
    Arreglo(0, 672) = "Gagnon"
    Arreglo(0, 673) = "Gaines"
    Arreglo(0, 674) = "Gallagher"
    Arreglo(0, 675) = "Gallegos"
    Arreglo(0, 676) = "Galloway"
    Arreglo(0, 677) = "Galvan"
    Arreglo(0, 678) = "Gamble"
    Arreglo(0, 679) = "Gann"
    Arreglo(0, 680) = "Garcia"
    Arreglo(0, 681) = "Gardipee"
    Arreglo(0, 682) = "Gardner"
    Arreglo(0, 683) = "Garfield"
    Arreglo(0, 684) = "Garner"
    Arreglo(0, 685) = "Garrett"
    Arreglo(0, 686) = "Garrison"
    Arreglo(0, 687) = "Garrow"
    Arreglo(0, 688) = "Garza"
    Arreglo(0, 689) = "Gates"
    Arreglo(0, 690) = "Gatewood"
    Arreglo(0, 691) = "Gauthier"
    Arreglo(0, 692) = "Gay"
    Arreglo(0, 693) = "Gee"
    Arreglo(0, 694) = "Gene"
    Arreglo(0, 695) = "Gentry"
    Arreglo(0, 696) = "George"
    Arreglo(0, 697) = "Gibbons"
    Arreglo(0, 698) = "Gibbs"
    Arreglo(0, 699) = "Gibson"
    Arreglo(0, 700) = "Gilbert"
    Arreglo(0, 701) = "Giles"
    Arreglo(0, 702) = "Gill"
    Arreglo(0, 703) = "Gillespie"
    Arreglo(0, 704) = "Gillis"
    Arreglo(0, 705) = "Gilmore"
    Arreglo(0, 706) = "Gipson"
    Arreglo(0, 707) = "Gishie"
    Arreglo(0, 708) = "Givens"
    Arreglo(0, 709) = "Glass"
    Arreglo(0, 710) = "Gleason"
    Arreglo(0, 711) = "Glenn"
    Arreglo(0, 712) = "Glover"
    Arreglo(0, 713) = "Goad"
    Arreglo(0, 714) = "Godfrey"
    Arreglo(0, 715) = "Godwin"
    Arreglo(0, 716) = "Goff"
    Arreglo(0, 717) = "Goings"
    Arreglo(0, 718) = "Goins"
    Arreglo(0, 719) = "Golden"
    Arreglo(0, 720) = "Goldtooth"
    Arreglo(0, 721) = "Gomez"
    Arreglo(0, 722) = "Gonzales"
    Arreglo(0, 723) = "Gonzalez"
    Arreglo(0, 724) = "Good"
    Arreglo(0, 725) = "Goode"
    Arreglo(0, 726) = "Goodluck"
    Arreglo(0, 727) = "Goodman"
    Arreglo(0, 728) = "Goodrich"
    Arreglo(0, 729) = "Goodwin"
    Arreglo(0, 730) = "Gordon"
    Arreglo(0, 731) = "Gore"
    Arreglo(0, 732) = "Gorman"
    Arreglo(0, 733) = "Goseyun"
    Arreglo(0, 734) = "Goss"
    Arreglo(0, 735) = "Gouge"
    Arreglo(0, 736) = "Gould"
    Arreglo(0, 737) = "Gourd"
    Arreglo(0, 738) = "Gourneau"
    Arreglo(0, 739) = "Grace"
    Arreglo(0, 740) = "Graham"
    Arreglo(0, 741) = "Grant"
    Arreglo(0, 742) = "Grass"
    Arreglo(0, 743) = "Graves"
    Arreglo(0, 744) = "Gray"
    Arreglo(0, 745) = "Grayson"
    Arreglo(0, 746) = "Green"
    Arreglo(0, 747) = "Greene"
    Arreglo(0, 748) = "Greenwood"
    Arreglo(0, 749) = "Greer"
    Arreglo(0, 750) = "Gregg"
    Arreglo(0, 751) = "Gregory"
    Arreglo(0, 752) = "Grey"
    Arreglo(0, 753) = "Greyeyes"
    Arreglo(0, 754) = "Griffin"
    Arreglo(0, 755) = "Griffith"
    Arreglo(0, 756) = "Griggs"
    Arreglo(0, 757) = "Grimes"
    Arreglo(0, 758) = "Gross"
    Arreglo(0, 759) = "Grover"
    Arreglo(0, 760) = "Groves"
    Arreglo(0, 761) = "Grubbs"
    Arreglo(0, 762) = "Guardipee"
    Arreglo(0, 763) = "Guerra"
    Arreglo(0, 764) = "Guerrero"
    Arreglo(0, 765) = "Guidry"
    Arreglo(0, 766) = "Guinn"
    Arreglo(0, 767) = "Gunter"
    Arreglo(0, 768) = "Guthrie"
    Arreglo(0, 769) = "Gutierrez"
    Arreglo(0, 770) = "Guy"
    Arreglo(0, 771) = "Guzman"
    Arreglo(0, 772) = "Hadley"
    Arreglo(0, 773) = "Haines"
    Arreglo(0, 774) = "Hair"
    Arreglo(0, 775) = "Hale"
    Arreglo(0, 776) = "Haley"
    Arreglo(0, 777) = "Hall"
    Arreglo(0, 778) = "Hamilton"
    Arreglo(0, 779) = "Hamm"
    Arreglo(0, 780) = "Hammer"
    Arreglo(0, 781) = "Hammond"
    Arreglo(0, 782) = "Hammonds"
    Arreglo(0, 783) = "Hammons"
    Arreglo(0, 784) = "Hampton"
    Arreglo(0, 785) = "Hancock"
    Arreglo(0, 786) = "Hand"
    Arreglo(0, 787) = "Haney"
    Arreglo(0, 788) = "Hankins"
    Arreglo(0, 789) = "Hanks"
    Arreglo(0, 790) = "Hanley"
    Arreglo(0, 791) = "Hanna"
    Arreglo(0, 792) = "Hansen"
    Arreglo(0, 793) = "Hanson"
    Arreglo(0, 794) = "Harden"
    Arreglo(0, 795) = "Hardin"
    Arreglo(0, 796) = "Harding"
    Arreglo(0, 797) = "Hardy"
    Arreglo(0, 798) = "Hare"
    Arreglo(0, 799) = "Harjo"
    Arreglo(0, 800) = "Harlan"
    Arreglo(0, 801) = "Harley"
    Arreglo(0, 802) = "Harmon"
    Arreglo(0, 803) = "Harp"
    Arreglo(0, 804) = "Harper"
    Arreglo(0, 805) = "Harrell"
    Arreglo(0, 806) = "Harrington"
    Arreglo(0, 807) = "Harris"
    Arreglo(0, 808) = "Harrison"
    Arreglo(0, 809) = "Harry"
    Arreglo(0, 810) = "Hart"
    Arreglo(0, 811) = "Hartman"
    Arreglo(0, 812) = "Harvey"
    Arreglo(0, 813) = "Harwood"
    Arreglo(0, 814) = "Haskie"
    Arreglo(0, 815) = "Hastings"
    Arreglo(0, 816) = "Hatch"
    Arreglo(0, 817) = "Hatcher"
    Arreglo(0, 818) = "Hatfield"
    Arreglo(0, 819) = "Hatton"
    Arreglo(0, 820) = "Hawk"
    Arreglo(0, 821) = "Hawkins"
    Arreglo(0, 822) = "Hawley"
    Arreglo(0, 823) = "Hayden"
    Arreglo(0, 824) = "Hayes"
    Arreglo(0, 825) = "Haynes"
    Arreglo(0, 826) = "Hays"
    Arreglo(0, 827) = "Hayward"
    Arreglo(0, 828) = "Hazard"
    Arreglo(0, 829) = "Head"
    Arreglo(0, 830) = "Healy"
    Arreglo(0, 831) = "Heath"
    Arreglo(0, 832) = "Hebert"
    Arreglo(0, 833) = "Hedrick"
    Arreglo(0, 834) = "Helms"
    Arreglo(0, 835) = "Helton"
    Arreglo(0, 836) = "Henderson"
    Arreglo(0, 837) = "Hendricks"
    Arreglo(0, 838) = "Hendrickson"
    Arreglo(0, 839) = "Hendrix"
    Arreglo(0, 840) = "Henio"
    Arreglo(0, 841) = "Henry"
    Arreglo(0, 842) = "Hensley"
    Arreglo(0, 843) = "Henson"
    Arreglo(0, 844) = "Herbert"
    Arreglo(0, 845) = "Herman"
    Arreglo(0, 846) = "Hernandez"
    Arreglo(0, 847) = "Herne"
    Arreglo(0, 848) = "Herrera"
    Arreglo(0, 849) = "Herring"
    Arreglo(0, 850) = "Herron"
    Arreglo(0, 851) = "Hess"
    Arreglo(0, 852) = "Hewitt"
    Arreglo(0, 853) = "Hickman"
    Arreglo(0, 854) = "Hicks"
    Arreglo(0, 855) = "Higgins"
    Arreglo(0, 856) = "Hill"
    Arreglo(0, 857) = "Hilton"
    Arreglo(0, 858) = "Hines"
    Arreglo(0, 859) = "Hinton"
    Arreglo(0, 860) = "Hobbs"
    Arreglo(0, 861) = "Hobson"
    Arreglo(0, 862) = "Hodge"
    Arreglo(0, 863) = "Hodges"
    Arreglo(0, 864) = "Hoffman"
    Arreglo(0, 865) = "Hogan"
    Arreglo(0, 866) = "Hogue"
    Arreglo(0, 867) = "Holcomb"
    Arreglo(0, 868) = "Holden"
    Arreglo(0, 869) = "Holder"
    Arreglo(0, 870) = "Holiday"
    Arreglo(0, 871) = "Holland"
    Arreglo(0, 872) = "Holley"
    Arreglo(0, 873) = "Holliday"
    Arreglo(0, 874) = "Holloway"
    Arreglo(0, 875) = "Holman"
    Arreglo(0, 876) = "Holmes"
    Arreglo(0, 877) = "Holt"
    Arreglo(0, 878) = "Homer"
    Arreglo(0, 879) = "Honeycutt"
    Arreglo(0, 880) = "Hood"
    Arreglo(0, 881) = "Hooper"
    Arreglo(0, 882) = "Hoover"
    Arreglo(0, 883) = "Hopkins"
    Arreglo(0, 884) = "Hopper"
    Arreglo(0, 885) = "Hopson"
    Arreglo(0, 886) = "Horn"
    Arreglo(0, 887) = "Horne"
    Arreglo(0, 888) = "Horner"
    Arreglo(0, 889) = "Horse"
    Arreglo(0, 890) = "Horton"
    Arreglo(0, 891) = "Hoskie"
    Arreglo(0, 892) = "Hosteen"
    Arreglo(0, 893) = "Houle"
    Arreglo(0, 894) = "House"
    Arreglo(0, 895) = "Houston"
    Arreglo(0, 896) = "Howard"
    Arreglo(0, 897) = "Howe"
    Arreglo(0, 898) = "Howell"
    Arreglo(0, 899) = "Hubbard"
    Arreglo(0, 900) = "Huddleston"
    Arreglo(0, 901) = "Hudson"
    Arreglo(0, 902) = "Huff"
    Arreglo(0, 903) = "Huffman"
    Arreglo(0, 904) = "Huggins"
    Arreglo(0, 905) = "Hughes"
    Arreglo(0, 906) = "Hull"
    Arreglo(0, 907) = "Hummingbird"
    Arreglo(0, 908) = "Humphrey"
    Arreglo(0, 909) = "Hunt"
    Arreglo(0, 910) = "Hunter"
    Arreglo(0, 911) = "Hurley"
    Arreglo(0, 912) = "Hurst"
    Arreglo(0, 913) = "Hurt"
    Arreglo(0, 914) = "Hutchins"
    Arreglo(0, 915) = "Hutchinson"
    Arreglo(0, 916) = "Hutchison"
    Arreglo(0, 917) = "Hyatt"
    Arreglo(0, 918) = "Hyde"
    Arreglo(0, 919) = "Ignacio"
    Arreglo(0, 920) = "Ingram"
    Arreglo(0, 921) = "Inman"
    Arreglo(0, 922) = "Iron"
    Arreglo(0, 923) = "Ironcloud"
    Arreglo(0, 924) = "Irving"
    Arreglo(0, 925) = "Irwin"
    Arreglo(0, 926) = "Isaac"
    Arreglo(0, 927) = "Isaacs"
    Arreglo(0, 928) = "Isham"
    Arreglo(0, 929) = "Ivanoff"
    Arreglo(0, 930) = "Ivey"
    Arreglo(0, 931) = "Jack"
    Arreglo(0, 932) = "Jackson"
    Arreglo(0, 933) = "Jacob"
    Arreglo(0, 934) = "Jacobs"
    Arreglo(0, 935) = "Jacobson"
    Arreglo(0, 936) = "Jake"
    Arreglo(0, 937) = "James"
    Arreglo(0, 938) = "Jameson"
    Arreglo(0, 939) = "Jamison"
    Arreglo(0, 940) = "Janis"
    Arreglo(0, 941) = "Jaramillo"
    Arreglo(0, 942) = "Jarvis"
    Arreglo(0, 943) = "Jay"
    Arreglo(0, 944) = "Jeff"
    Arreglo(0, 945) = "Jefferson"
    Arreglo(0, 946) = "Jeffries"
    Arreglo(0, 947) = "Jenkins"
    Arreglo(0, 948) = "Jennings"
    Arreglo(0, 949) = "Jensen"
    Arreglo(0, 950) = "Jerome"
    Arreglo(0, 951) = "Jewell"
    Arreglo(0, 952) = "Jewett"
    Arreglo(0, 953) = "Jim"
    Arreglo(0, 954) = "Jimenez"
    Arreglo(0, 955) = "Jimerson"
    Arreglo(0, 956) = "Jimmie"
    Arreglo(0, 957) = "Jimmy"
    Arreglo(0, 958) = "Jiron"
    Arreglo(0, 959) = "Joaquin"
    Arreglo(0, 960) = "Joe"
    Arreglo(0, 961) = "John"
    Arreglo(0, 962) = "Johns"
    Arreglo(0, 963) = "Johnson"
    Arreglo(0, 964) = "Johnston"
    Arreglo(0, 965) = "Jojola"
    Arreglo(0, 966) = "Jones"
    Arreglo(0, 967) = "Jordan"
    Arreglo(0, 968) = "Jose"
    Arreglo(0, 969) = "Joseph"
    Arreglo(0, 970) = "Jourdain"
    Arreglo(0, 971) = "Juan"
    Arreglo(0, 972) = "Juarez"
    Arreglo(0, 973) = "Julian"
    Arreglo(0, 974) = "Jumbo"
    Arreglo(0, 975) = "Jumper"
    Arreglo(0, 976) = "June"
    Arreglo(0, 977) = "Justice"
    Arreglo(0, 978) = "Kaiser"
    Arreglo(0, 979) = "Kameroff"
    Arreglo(0, 980) = "Kane"
    Arreglo(0, 981) = "Kanuho"
    Arreglo(0, 982) = "Kaulaity"
    Arreglo(0, 983) = "Kay"
    Arreglo(0, 984) = "Kaye"
    Arreglo(0, 985) = "Keams"
    Arreglo(0, 986) = "Kee"
    Arreglo(0, 987) = "Keene"
    Arreglo(0, 988) = "Keener"
    Arreglo(0, 989) = "Keith"
    Arreglo(0, 990) = "Keller"
    Arreglo(0, 991) = "Kelley"
    Arreglo(0, 992) = "Kelly"
    Arreglo(0, 993) = "Kelsey"
    Arreglo(0, 994) = "Kemp"
    Arreglo(0, 995) = "Kendall"
    Arreglo(0, 996) = "Kendrick"
    Arreglo(0, 997) = "Kennedy"
    Arreglo(0, 998) = "Kent"
    Arreglo(0, 999) = "Kenton"
    Arreglo(0, 1000) = "Keplin"
    Arreglo(0, 1001) = "Kerns"
    Arreglo(0, 1002) = "Kerr"
    Arreglo(0, 1003) = "Ketcher"
    Arreglo(0, 1004) = "Ketchum"
    Arreglo(0, 1005) = "Key"
    Arreglo(0, 1006) = "Keyonnie"
    Arreglo(0, 1007) = "Keys"
    Arreglo(0, 1008) = "Khan"
    Arreglo(0, 1009) = "Kidd"
    Arreglo(0, 1010) = "King"
    Arreglo(0, 1011) = "Kingbird"
    Arreglo(0, 1012) = "Kingfisher"
    Arreglo(0, 1013) = "Kinney"
    Arreglo(0, 1014) = "Kinsel"
    Arreglo(0, 1015) = "Kinsey"
    Arreglo(0, 1016) = "Kipp"
    Arreglo(0, 1017) = "Kirby"
    Arreglo(0, 1018) = "Kirk"
    Arreglo(0, 1019) = "Kirkland"
    Arreglo(0, 1020) = "Kisto"
    Arreglo(0, 1021) = "Klein"
    Arreglo(0, 1022) = "Kline"
    Arreglo(0, 1023) = "Knapp"
    Arreglo(0, 1024) = "Knight"
    Arreglo(0, 1025) = "Knox"
    Arreglo(0, 1026) = "Koenig"
    Arreglo(0, 1027) = "Kraft"
    Arreglo(0, 1028) = "Kramer"
    Arreglo(0, 1029) = "Lackey"
    Arreglo(0, 1030) = "Lacroix"
    Arreglo(0, 1031) = "Lacy"
    Arreglo(0, 1032) = "Ladd"
    Arreglo(0, 1033) = "Laducer"
    Arreglo(0, 1034) = "Lafferty"
    Arreglo(0, 1035) = "Lafontaine"
    Arreglo(0, 1036) = "Lafountain"
    Arreglo(0, 1037) = "Lafromboise"
    Arreglo(0, 1038) = "Lake"
    Arreglo(0, 1039) = "Lamb"
    Arreglo(0, 1040) = "Lambert"
    Arreglo(0, 1041) = "Lamere"
    Arreglo(0, 1042) = "Lamont"
    Arreglo(0, 1043) = "Lancaster"
    Arreglo(0, 1044) = "Landry"
    Arreglo(0, 1045) = "Lane"
    Arreglo(0, 1046) = "Lang"
    Arreglo(0, 1047) = "Langley"
    Arreglo(0, 1048) = "Langston"
    Arreglo(0, 1049) = "Lansing"
    Arreglo(0, 1050) = "Laplante"
    Arreglo(0, 1051) = "Lapointe"
    Arreglo(0, 1052) = "Lara"
    Arreglo(0, 1053) = "Largo"
    Arreglo(0, 1054) = "Larney"
    Arreglo(0, 1055) = "Laroche"
    Arreglo(0, 1056) = "Larocque"
    Arreglo(0, 1057) = "Laroque"
    Arreglo(0, 1058) = "Larose"
    Arreglo(0, 1059) = "Larsen"
    Arreglo(0, 1060) = "Larson"
    Arreglo(0, 1061) = "Larvie"
    Arreglo(0, 1062) = "Lasley"
    Arreglo(0, 1063) = "Latham"
    Arreglo(0, 1064) = "Laughing"
    Arreglo(0, 1065) = "Laughlin"
    Arreglo(0, 1066) = "Laughter"
    Arreglo(0, 1067) = "Lavallie"
    Arreglo(0, 1068) = "Laverdure"
    Arreglo(0, 1069) = "Lawrence"
    Arreglo(0, 1070) = "Lawson"
    Arreglo(0, 1071) = "Lay"
    Arreglo(0, 1072) = "Lazore"
    Arreglo(0, 1073) = "Leach"
    Arreglo(0, 1074) = "Leavitt"
    Arreglo(0, 1075) = "Lebeau"
    Arreglo(0, 1076) = "Leblanc"
    Arreglo(0, 1077) = "Leclair"
    Arreglo(0, 1078) = "Leclaire"
    Arreglo(0, 1079) = "Ledford"
    Arreglo(0, 1080) = "Lee"
    Arreglo(0, 1081) = "Leflore"
    Arreglo(0, 1082) = "Lefthand"
    Arreglo(0, 1083) = "Lemieux"
    Arreglo(0, 1084) = "Lente"
    Arreglo(0, 1085) = "Leon"
    Arreglo(0, 1086) = "Leonard"
    Arreglo(0, 1087) = "Leroy"
    Arreglo(0, 1088) = "Leslie"
    Arreglo(0, 1089) = "Lester"
    Arreglo(0, 1090) = "Levi"
    Arreglo(0, 1091) = "Lewis"
    Arreglo(0, 1092) = "Lightfoot"
    Arreglo(0, 1093) = "Lilly"
    Arreglo(0, 1094) = "Lincoln"
    Arreglo(0, 1095) = "Lind"
    Arreglo(0, 1096) = "Lindsay"
    Arreglo(0, 1097) = "Lindsey"
    Arreglo(0, 1098) = "Linton"
    Arreglo(0, 1099) = "Little"
    Arreglo(0, 1100) = "Littlebear"
    Arreglo(0, 1101) = "Littledog"
    Arreglo(0, 1102) = "Littlefield"
    Arreglo(0, 1103) = "Littlejohn"
    Arreglo(0, 1104) = "Littlelight"
    Arreglo(0, 1105) = "Littleman"
    Arreglo(0, 1106) = "Littlethunder"
    Arreglo(0, 1107) = "Littlewolf"
    Arreglo(0, 1108) = "Livingston"
    Arreglo(0, 1109) = "Lloyd"
    Arreglo(0, 1110) = "Locke"
    Arreglo(0, 1111) = "Locklear"
    Arreglo(0, 1112) = "Lockwood"
    Arreglo(0, 1113) = "Locust"
    Arreglo(0, 1114) = "Lofton"
    Arreglo(0, 1115) = "Logan"
    Arreglo(0, 1116) = "Lonebear"
    Arreglo(0, 1117) = "Long"
    Arreglo(0, 1118) = "Longie"
    Arreglo(0, 1119) = "Looney"
    Arreglo(0, 1120) = "Lopez"
    Arreglo(0, 1121) = "Lorenzo"
    Arreglo(0, 1122) = "Loretto"
    Arreglo(0, 1123) = "Lott"
    Arreglo(0, 1124) = "Louis"
    Arreglo(0, 1125) = "Lovato"
    Arreglo(0, 1126) = "Love"
    Arreglo(0, 1127) = "Lovejoy"
    Arreglo(0, 1128) = "Lovell"
    Arreglo(0, 1129) = "Lowe"
    Arreglo(0, 1130) = "Lowery"
    Arreglo(0, 1131) = "Lowry"
    Arreglo(0, 1132) = "Lucas"
    Arreglo(0, 1133) = "Lucero"
    Arreglo(0, 1134) = "Ludlow"
    Arreglo(0, 1135) = "Lujan"
    Arreglo(0, 1136) = "Luke"
    Arreglo(0, 1137) = "Luna"
    Arreglo(0, 1138) = "Lupe"
    Arreglo(0, 1139) = "Lussier"
    Arreglo(0, 1140) = "Luther"
    Arreglo(0, 1141) = "Lynch"
    Arreglo(0, 1142) = "Lynn"
    Arreglo(0, 1143) = "Lyons"
    Arreglo(0, 1144) = "Macdonald"
    Arreglo(0, 1145) = "Mace"
    Arreglo(0, 1146) = "Mack"
    Arreglo(0, 1147) = "Mackey"
    Arreglo(0, 1148) = "Madden"
    Arreglo(0, 1149) = "Maddox"
    Arreglo(0, 1150) = "Madison"
    Arreglo(0, 1151) = "Madrid"
    Arreglo(0, 1152) = "Maestas"
    Arreglo(0, 1153) = "Magee"
    Arreglo(0, 1154) = "Main"
    Arreglo(0, 1155) = "Maldonado"
    Arreglo(0, 1156) = "Mallory"
    Arreglo(0, 1157) = "Malone"
    Arreglo(0, 1158) = "Maloney"
    Arreglo(0, 1159) = "Manley"
    Arreglo(0, 1160) = "Mann"
    Arreglo(0, 1161) = "Manning"
    Arreglo(0, 1162) = "Manson"
    Arreglo(0, 1163) = "Manuel"
    Arreglo(0, 1164) = "Manuelito"
    Arreglo(0, 1165) = "Manygoats"
    Arreglo(0, 1166) = "Maracle"
    Arreglo(0, 1167) = "Marchand"
    Arreglo(0, 1168) = "Maria"
    Arreglo(0, 1169) = "Mariano"
    Arreglo(0, 1170) = "Marion"
    Arreglo(0, 1171) = "Mark"
    Arreglo(0, 1172) = "Marks"
    Arreglo(0, 1173) = "Marquez"
    Arreglo(0, 1174) = "Marsh"
    Arreglo(0, 1175) = "Marshall"
    Arreglo(0, 1176) = "Martell"
    Arreglo(0, 1177) = "Martin"
    Arreglo(0, 1178) = "Martine"
    Arreglo(0, 1179) = "Martinez"
    Arreglo(0, 1180) = "Mason"
    Arreglo(0, 1181) = "Massey"
    Arreglo(0, 1182) = "Masters"
    Arreglo(0, 1183) = "Mathews"
    Arreglo(0, 1184) = "Mathis"
    Arreglo(0, 1185) = "Matt"
    Arreglo(0, 1186) = "Matthews"
    Arreglo(0, 1187) = "Matus"
    Arreglo(0, 1188) = "Maxwell"
    Arreglo(0, 1189) = "May"
    Arreglo(0, 1190) = "Mayes"
    Arreglo(0, 1191) = "Mayfield"
    Arreglo(0, 1192) = "Mayle"
    Arreglo(0, 1193) = "Maynard"
    Arreglo(0, 1194) = "Maynor"
    Arreglo(0, 1195) = "Mayo"
    Arreglo(0, 1196) = "Mays"
    Arreglo(0, 1197) = "Mcbride"
    Arreglo(0, 1198) = "Mccabe"
    Arreglo(0, 1199) = "Mccall"
    Arreglo(0, 1200) = "Mccann"
    Arreglo(0, 1201) = "Mccarthy"
    Arreglo(0, 1202) = "Mccarty"
    Arreglo(0, 1203) = "Mccauley"
    Arreglo(0, 1204) = "Mcclain"
    Arreglo(0, 1205) = "Mcclellan"
    Arreglo(0, 1206) = "Mccloud"
    Arreglo(0, 1207) = "Mcclure"
    Arreglo(0, 1208) = "Mcconnell"
    Arreglo(0, 1209) = "Mccord"
    Arreglo(0, 1210) = "Mccormick"
    Arreglo(0, 1211) = "Mccovey"
    Arreglo(0, 1212) = "Mccoy"
    Arreglo(0, 1213) = "Mccray"
    Arreglo(0, 1214) = "Mccullough"
    Arreglo(0, 1215) = "Mccurtain"
    Arreglo(0, 1216) = "Mcdaniel"
    Arreglo(0, 1217) = "Mcdonald"
    Arreglo(0, 1218) = "Mcdowell"
    Arreglo(0, 1219) = "Mcfarland"
    Arreglo(0, 1220) = "Mcgee"
    Arreglo(0, 1221) = "Mcgeshick"
    Arreglo(0, 1222) = "Mcghee"
    Arreglo(0, 1223) = "Mcgill"
    Arreglo(0, 1224) = "Mcgirt"
    Arreglo(0, 1225) = "Mcguire"
    Arreglo(0, 1226) = "Mcintosh"
    Arreglo(0, 1227) = "Mcintyre"
    Arreglo(0, 1228) = "Mckay"
    Arreglo(0, 1229) = "Mckee"
    Arreglo(0, 1230) = "Mckenzie"
    Arreglo(0, 1231) = "Mckinley"
    Arreglo(0, 1232) = "Mckinney"
    Arreglo(0, 1233) = "Mcknight"
    Arreglo(0, 1234) = "Mclain"
    Arreglo(0, 1235) = "Mclaughlin"
    Arreglo(0, 1236) = "Mclean"
    Arreglo(0, 1237) = "Mclemore"
    Arreglo(0, 1238) = "Mcleod"
    Arreglo(0, 1239) = "Mcmillan"
    Arreglo(0, 1240) = "Mcmillian"
    Arreglo(0, 1241) = "Mcneal"
    Arreglo(0, 1242) = "Mcneil"
    Arreglo(0, 1243) = "Mcneill"
    Arreglo(0, 1244) = "Mcpherson"
    Arreglo(0, 1245) = "Mead"
    Arreglo(0, 1246) = "Meade"
    Arreglo(0, 1247) = "Meadows"
    Arreglo(0, 1248) = "Means"
    Arreglo(0, 1249) = "Medina"
    Arreglo(0, 1250) = "Meeks"
    Arreglo(0, 1251) = "Mejia"
    Arreglo(0, 1252) = "Melton"
    Arreglo(0, 1253) = "Menard"
    Arreglo(0, 1254) = "Mendez"
    Arreglo(0, 1255) = "Mendoza"
    Arreglo(0, 1256) = "Mercer"
    Arreglo(0, 1257) = "Merculief"
    Arreglo(0, 1258) = "Merrill"
    Arreglo(0, 1259) = "Merritt"
    Arreglo(0, 1260) = "Meshell"
    Arreglo(0, 1261) = "Messer"
    Arreglo(0, 1262) = "Mesteth"
    Arreglo(0, 1263) = "Metcalf"
    Arreglo(0, 1264) = "Metoxen"
    Arreglo(0, 1265) = "Meyer"
    Arreglo(0, 1266) = "Meyers"
    Arreglo(0, 1267) = "Michael"
    Arreglo(0, 1268) = "Michel"
    Arreglo(0, 1269) = "Middleton"
    Arreglo(0, 1270) = "Miguel"
    Arreglo(0, 1271) = "Mike"
    Arreglo(0, 1272) = "Miles"
    Arreglo(0, 1273) = "Miller"
    Arreglo(0, 1274) = "Milligan"
    Arreglo(0, 1275) = "Mills"
    Arreglo(0, 1276) = "Miner"
    Arreglo(0, 1277) = "Miranda"
    Arreglo(0, 1278) = "Mitchell"
    Arreglo(0, 1279) = "Mix"
    Arreglo(0, 1280) = "Molina"
    Arreglo(0, 1281) = "Monroe"
    Arreglo(0, 1282) = "Montana"
    Arreglo(0, 1283) = "Montano"
    Arreglo(0, 1284) = "Monte"
    Arreglo(0, 1285) = "Montgomery"
    Arreglo(0, 1286) = "Montoya"
    Arreglo(0, 1287) = "Moody"
    Arreglo(0, 1288) = "Moon"
    Arreglo(0, 1289) = "Mooney"
    Arreglo(0, 1290) = "Moore"
    Arreglo(0, 1291) = "Moose"
    Arreglo(0, 1292) = "Moquino"
    Arreglo(0, 1293) = "Mora"
    Arreglo(0, 1294) = "Morales"
    Arreglo(0, 1295) = "Moran"
    Arreglo(0, 1296) = "Moreland"
    Arreglo(0, 1297) = "Moreno"
    Arreglo(0, 1298) = "Morgan"
    Arreglo(0, 1299) = "Morin"
    Arreglo(0, 1300) = "Morris"
    Arreglo(0, 1301) = "Morrison"
    Arreglo(0, 1302) = "Morrow"
    Arreglo(0, 1303) = "Morse"
    Arreglo(0, 1304) = "Morsette"
    Arreglo(0, 1305) = "Morton"
    Arreglo(0, 1306) = "Mose"
    Arreglo(0, 1307) = "Moses"
    Arreglo(0, 1308) = "Mosley"
    Arreglo(0, 1309) = "Moss"
    Arreglo(0, 1310) = "Mountain"
    Arreglo(0, 1311) = "Mouse"
    Arreglo(0, 1312) = "Mullen"
    Arreglo(0, 1313) = "Mullins"
    Arreglo(0, 1314) = "Munoz"
    Arreglo(0, 1315) = "Munson"
    Arreglo(0, 1316) = "Murdock"
    Arreglo(0, 1317) = "Murphy"
    Arreglo(0, 1318) = "Murray"
    Arreglo(0, 1319) = "Myers"
    Arreglo(0, 1320) = "Nadeau"
    Arreglo(0, 1321) = "Nail"
    Arreglo(0, 1322) = "Nakai"
    Arreglo(0, 1323) = "Nance"
    Arreglo(0, 1324) = "Naquin"
    Arreglo(0, 1325) = "Naranjo"
    Arreglo(0, 1326) = "Nash"
    Arreglo(0, 1327) = "Nation"
    Arreglo(0, 1328) = "Navarro"
    Arreglo(0, 1329) = "Neal"
    Arreglo(0, 1330) = "Needham"
    Arreglo(0, 1331) = "Neff"
    Arreglo(0, 1332) = "Nelson"
    Arreglo(0, 1333) = "Nephew"
    Arreglo(0, 1334) = "Newell"
    Arreglo(0, 1335) = "Newman"
    Arreglo(0, 1336) = "Newton"
    Arreglo(0, 1337) = "Nez"
    Arreglo(0, 1338) = "Nicholas"
    Arreglo(0, 1339) = "Nichols"
    Arreglo(0, 1340) = "Nicholson"
    Arreglo(0, 1341) = "Nick"
    Arreglo(0, 1342) = "Nielsen"
    Arreglo(0, 1343) = "Nieto"
    Arreglo(0, 1344) = "Ninham"
    Arreglo(0, 1345) = "Nix"
    Arreglo(0, 1346) = "Nixon"
    Arreglo(0, 1347) = "Noah"
    Arreglo(0, 1348) = "Noble"
    Arreglo(0, 1349) = "Noel"
    Arreglo(0, 1350) = "Nolan"
    Arreglo(0, 1351) = "Norman"
    Arreglo(0, 1352) = "Norris"
    Arreglo(0, 1353) = "North"
    Arreglo(0, 1354) = "Norton"
    Arreglo(0, 1355) = "Norwood"
    Arreglo(0, 1356) = "Nosie"
    Arreglo(0, 1357) = "Notah"
    Arreglo(0, 1358) = "Nunez"
    Arreglo(0, 1359) = "Oakes"
    Arreglo(0, 1360) = "Obrien"
    Arreglo(0, 1361) = "Ochoa"
    Arreglo(0, 1362) = "Oconnor"
    Arreglo(0, 1363) = "Odell"
    Arreglo(0, 1364) = "Odom"
    Arreglo(0, 1365) = "Ogden"
    Arreglo(0, 1366) = "Ogle"
    Arreglo(0, 1367) = "Oldman"
    Arreglo(0, 1368) = "Olguin"
    Arreglo(0, 1369) = "Oliver"
    Arreglo(0, 1370) = "Olney"
    Arreglo(0, 1371) = "Olsen"
    Arreglo(0, 1372) = "Olson"
    Arreglo(0, 1373) = "Oneal"
    Arreglo(0, 1374) = "Orr"
    Arreglo(0, 1375) = "Ortega"
    Arreglo(0, 1376) = "Ortiz"
    Arreglo(0, 1377) = "Osborn"
    Arreglo(0, 1378) = "Osborne"
    Arreglo(0, 1379) = "Osceola"
    Arreglo(0, 1380) = "Osife"
    Arreglo(0, 1381) = "Otero"
    Arreglo(0, 1382) = "Ott"
    Arreglo(0, 1383) = "Owen"
    Arreglo(0, 1384) = "Owens"
    Arreglo(0, 1385) = "Owle"
    Arreglo(0, 1386) = "Oxendine"
    Arreglo(0, 1387) = "Pablo"
    Arreglo(0, 1388) = "Pace"
    Arreglo(0, 1389) = "Pacheco"
    Arreglo(0, 1390) = "Pack"
    Arreglo(0, 1391) = "Paddock"
    Arreglo(0, 1392) = "Padgett"
    Arreglo(0, 1393) = "Padilla"
    Arreglo(0, 1394) = "Page"
    Arreglo(0, 1395) = "Painter"
    Arreglo(0, 1396) = "Palmer"
    Arreglo(0, 1397) = "Panther"
    Arreglo(0, 1398) = "Pappan"
    Arreglo(0, 1399) = "Paquin"
    Arreglo(0, 1400) = "Parfait"
    Arreglo(0, 1401) = "Parish"
    Arreglo(0, 1402) = "Parisien"
    Arreglo(0, 1403) = "Park"
    Arreglo(0, 1404) = "Parker"
    Arreglo(0, 1405) = "Parks"
    Arreglo(0, 1406) = "Parrish"
    Arreglo(0, 1407) = "Parsons"
    Arreglo(0, 1408) = "Pate"
    Arreglo(0, 1409) = "Patel"
    Arreglo(0, 1410) = "Patrick"
    Arreglo(0, 1411) = "Patten"
    Arreglo(0, 1412) = "Patterson"
    Arreglo(0, 1413) = "Patton"
    Arreglo(0, 1414) = "Paul"
    Arreglo(0, 1415) = "Payne"
    Arreglo(0, 1416) = "Payton"
    Arreglo(0, 1417) = "Peacock"
    Arreglo(0, 1418) = "Pearce"
    Arreglo(0, 1419) = "Pearson"
    Arreglo(0, 1420) = "Pease"
    Arreglo(0, 1421) = "Peck"
    Arreglo(0, 1422) = "Pedro"
    Arreglo(0, 1423) = "Peltier"
    Arreglo(0, 1424) = "Pemberton"
    Arreglo(0, 1425) = "Pena"
    Arreglo(0, 1426) = "Pendleton"
    Arreglo(0, 1427) = "Penn"
    Arreglo(0, 1428) = "Pennington"
    Arreglo(0, 1429) = "Perez"
    Arreglo(0, 1430) = "Perkins"
    Arreglo(0, 1431) = "Perry"
    Arreglo(0, 1432) = "Persaud"
    Arreglo(0, 1433) = "Person"
    Arreglo(0, 1434) = "Peshlakai"
    Arreglo(0, 1435) = "Pete"
    Arreglo(0, 1436) = "Peter"
    Arreglo(0, 1437) = "Peters"
    Arreglo(0, 1438) = "Petersen"
    Arreglo(0, 1439) = "Peterson"
    Arreglo(0, 1440) = "Pettigrew"
    Arreglo(0, 1441) = "Pettit"
    Arreglo(0, 1442) = "Petty"
    Arreglo(0, 1443) = "Phelps"
    Arreglo(0, 1444) = "Phillip"
    Arreglo(0, 1445) = "Phillips"
    Arreglo(0, 1446) = "Picard"
    Arreglo(0, 1447) = "Pickett"
    Arreglo(0, 1448) = "Picotte"
    Arreglo(0, 1449) = "Pierce"
    Arreglo(0, 1450) = "Pierre"
    Arreglo(0, 1451) = "Pigeon"
    Arreglo(0, 1452) = "Pike"
    Arreglo(0, 1453) = "Pina"
    Arreglo(0, 1454) = "Pine"
    Arreglo(0, 1455) = "Pino"
    Arreglo(0, 1456) = "Pinto"
    Arreglo(0, 1457) = "Piper"
    Arreglo(0, 1458) = "Pitka"
    Arreglo(0, 1459) = "Pittman"
    Arreglo(0, 1460) = "Pitts"
    Arreglo(0, 1461) = "Platero"
    Arreglo(0, 1462) = "Plummer"
    Arreglo(0, 1463) = "Poe"
    Arreglo(0, 1464) = "Poitra"
    Arreglo(0, 1465) = "Polk"
    Arreglo(0, 1466) = "Pollard"
    Arreglo(0, 1467) = "Poncho"
    Arreglo(0, 1468) = "Poole"
    Arreglo(0, 1469) = "Poorbear"
    Arreglo(0, 1470) = "Pope"
    Arreglo(0, 1471) = "Porter"
    Arreglo(0, 1472) = "Posey"
    Arreglo(0, 1473) = "Post"
    Arreglo(0, 1474) = "Postoak"
    Arreglo(0, 1475) = "Potter"
    Arreglo(0, 1476) = "Potts"
    Arreglo(0, 1477) = "Pourier"
    Arreglo(0, 1478) = "Powell"
    Arreglo(0, 1479) = "Powers"
    Arreglo(0, 1480) = "Powless"
    Arreglo(0, 1481) = "Poyer"
    Arreglo(0, 1482) = "Prater"
    Arreglo(0, 1483) = "Pratt"
    Arreglo(0, 1484) = "Prescott"
    Arreglo(0, 1485) = "Presley"
    Arreglo(0, 1486) = "Preston"
    Arreglo(0, 1487) = "Price"
    Arreglo(0, 1488) = "Primeaux"
    Arreglo(0, 1489) = "Prince"
    Arreglo(0, 1490) = "Printup"
    Arreglo(0, 1491) = "Pritchett"
    Arreglo(0, 1492) = "Proctor"
    Arreglo(0, 1493) = "Provost"
    Arreglo(0, 1494) = "Pruitt"
    Arreglo(0, 1495) = "Puckett"
    Arreglo(0, 1496) = "Pugh"
    Arreglo(0, 1497) = "Qualls"
    Arreglo(0, 1498) = "Quam"
    Arreglo(0, 1499) = "Queen"
    Arreglo(0, 1500) = "Quick"
    Arreglo(0, 1501) = "Quinn"
    Arreglo(0, 1502) = "Quintana"
    Arreglo(0, 1503) = "Quintero"
    Arreglo(0, 1504) = "Racine"
    Arreglo(0, 1505) = "Ragsdale"
    Arreglo(0, 1506) = "Raines"
    Arreglo(0, 1507) = "Rains"
    Arreglo(0, 1508) = "Rainwater"
    Arreglo(0, 1509) = "Ramey"
    Arreglo(0, 1510) = "Ramirez"
    Arreglo(0, 1511) = "Ramon"
    Arreglo(0, 1512) = "Ramone"
    Arreglo(0, 1513) = "Ramos"
    Arreglo(0, 1514) = "Ramsey"
    Arreglo(0, 1515) = "Randall"
    Arreglo(0, 1516) = "Randolph"
    Arreglo(0, 1517) = "Ransom"
    Arreglo(0, 1518) = "Ratliff"
    Arreglo(0, 1519) = "Rave"
    Arreglo(0, 1520) = "Ray"
    Arreglo(0, 1521) = "Raymond"
    Arreglo(0, 1522) = "Redbear"
    Arreglo(0, 1523) = "Redbird"
    Arreglo(0, 1524) = "Redcloud"
    Arreglo(0, 1525) = "Redeagle"
    Arreglo(0, 1526) = "Redelk"
    Arreglo(0, 1527) = "Redfox"
    Arreglo(0, 1528) = "Redhouse"
    Arreglo(0, 1529) = "Reece"
    Arreglo(0, 1530) = "Reed"
    Arreglo(0, 1531) = "Reeder"
    Arreglo(0, 1532) = "Reese"
    Arreglo(0, 1533) = "Reeves"
    Arreglo(0, 1534) = "Reid"
    Arreglo(0, 1535) = "Renville"
    Arreglo(0, 1536) = "Revels"
    Arreglo(0, 1537) = "Reyes"
    Arreglo(0, 1538) = "Reyna"
    Arreglo(0, 1539) = "Reynolds"
    Arreglo(0, 1540) = "Rhoades"
    Arreglo(0, 1541) = "Rhodes"
    Arreglo(0, 1542) = "Rice"
    Arreglo(0, 1543) = "Rich"
    Arreglo(0, 1544) = "Richard"
    Arreglo(0, 1545) = "Richards"
    Arreglo(0, 1546) = "Richardson"
    Arreglo(0, 1547) = "Richmond"
    Arreglo(0, 1548) = "Riddle"
    Arreglo(0, 1549) = "Rider"
    Arreglo(0, 1550) = "Ridley"
    Arreglo(0, 1551) = "Riggs"
    Arreglo(0, 1552) = "Riley"
    Arreglo(0, 1553) = "Ring"
    Arreglo(0, 1554) = "Rios"
    Arreglo(0, 1555) = "Ritchie"
    Arreglo(0, 1556) = "Rivas"
    Arreglo(0, 1557) = "Rivera"
    Arreglo(0, 1558) = "Rivers"
    Arreglo(0, 1559) = "Roach"
    Arreglo(0, 1560) = "Roan"
    Arreglo(0, 1561) = "Roanhorse"
    Arreglo(0, 1562) = "Robbins"
    Arreglo(0, 1563) = "Roberson"
    Arreglo(0, 1564) = "Roberts"
    Arreglo(0, 1565) = "Robertson"
    Arreglo(0, 1566) = "Robinson"
    Arreglo(0, 1567) = "Robison"
    Arreglo(0, 1568) = "Robles"
    Arreglo(0, 1569) = "Rocha"
    Arreglo(0, 1570) = "Rock"
    Arreglo(0, 1571) = "Rodgers"
    Arreglo(0, 1572) = "Rodriguez"
    Arreglo(0, 1573) = "Rodriquez"
    Arreglo(0, 1574) = "Roe"
    Arreglo(0, 1575) = "Rogers"
    Arreglo(0, 1576) = "Rojas"
    Arreglo(0, 1577) = "Rollins"
    Arreglo(0, 1578) = "Romero"
    Arreglo(0, 1579) = "Root"
    Arreglo(0, 1580) = "Rose"
    Arreglo(0, 1581) = "Ross"
    Arreglo(0, 1582) = "Roth"
    Arreglo(0, 1583) = "Roubideaux"
    Arreglo(0, 1584) = "Rouse"
    Arreglo(0, 1585) = "Rowe"
    Arreglo(0, 1586) = "Rowland"
    Arreglo(0, 1587) = "Roy"
    Arreglo(0, 1588) = "Roybal"
    Arreglo(0, 1589) = "Rubio"
    Arreglo(0, 1590) = "Rudd"
    Arreglo(0, 1591) = "Ruiz"
    Arreglo(0, 1592) = "Rush"
    Arreglo(0, 1593) = "Russell"
    Arreglo(0, 1594) = "Rutherford"
    Arreglo(0, 1595) = "Rutledge"
    Arreglo(0, 1596) = "Ryan"
    Arreglo(0, 1597) = "Sage"
    Arreglo(0, 1598) = "Salas"
    Arreglo(0, 1599) = "Salazar"
    Arreglo(0, 1600) = "Salgado"
    Arreglo(0, 1601) = "Salinas"
    Arreglo(0, 1602) = "Salt"
    Arreglo(0, 1603) = "Salway"
    Arreglo(0, 1604) = "Sam"
    Arreglo(0, 1605) = "Sampson"
    Arreglo(0, 1606) = "Samuel"
    Arreglo(0, 1607) = "Samuels"
    Arreglo(0, 1608) = "Sanchez"
    Arreglo(0, 1609) = "Sanders"
    Arreglo(0, 1610) = "Sanderson"
    Arreglo(0, 1611) = "Sandoval"
    Arreglo(0, 1612) = "Sands"
    Arreglo(0, 1613) = "Sanford"
    Arreglo(0, 1614) = "Sangster"
    Arreglo(0, 1615) = "Santiago"
    Arreglo(0, 1616) = "Santos"
    Arreglo(0, 1617) = "Sargent"
    Arreglo(0, 1618) = "Sarracino"
    Arreglo(0, 1619) = "Saunders"
    Arreglo(0, 1620) = "Savage"
    Arreglo(0, 1621) = "Sawyer"
    Arreglo(0, 1622) = "Sayers"
    Arreglo(0, 1623) = "Schaeffer"
    Arreglo(0, 1624) = "Schmidt"
    Arreglo(0, 1625) = "Schneider"
    Arreglo(0, 1626) = "Schroeder"
    Arreglo(0, 1627) = "Schultz"
    Arreglo(0, 1628) = "Schwartz"
    Arreglo(0, 1629) = "Scott"
    Arreglo(0, 1630) = "Seals"
    Arreglo(0, 1631) = "Sears"
    Arreglo(0, 1632) = "Seaton"
    Arreglo(0, 1633) = "Sebastian"
    Arreglo(0, 1634) = "Secatero"
    Arreglo(0, 1635) = "Secody"
    Arreglo(0, 1636) = "Self"
    Arreglo(0, 1637) = "Sellers"
    Arreglo(0, 1638) = "Sells"
    Arreglo(0, 1639) = "Sepulvado"
    Arreglo(0, 1640) = "Sewell"
    Arreglo(0, 1641) = "Sexton"
    Arreglo(0, 1642) = "Seymour"
    Arreglo(0, 1643) = "Shade"
    Arreglo(0, 1644) = "Shaffer"
    Arreglo(0, 1645) = "Shannon"
    Arreglo(0, 1646) = "Sharp"
    Arreglo(0, 1647) = "Shaw"
    Arreglo(0, 1648) = "Shay"
    Arreglo(0, 1649) = "Sheldon"
    Arreglo(0, 1650) = "Shell"
    Arreglo(0, 1651) = "Shelton"
    Arreglo(0, 1652) = "Shepard"
    Arreglo(0, 1653) = "Shepherd"
    Arreglo(0, 1654) = "Sheppard"
    Arreglo(0, 1655) = "Sheridan"
    Arreglo(0, 1656) = "Sherman"
    Arreglo(0, 1657) = "Sherwood"
    Arreglo(0, 1658) = "Shields"
    Arreglo(0, 1659) = "Shije"
    Arreglo(0, 1660) = "Shipley"
    Arreglo(0, 1661) = "Shirley"
    Arreglo(0, 1662) = "Shoemaker"
    Arreglo(0, 1663) = "Short"
    Arreglo(0, 1664) = "Shortman"
    Arreglo(0, 1665) = "Shorty"
    Arreglo(0, 1666) = "Sierra"
    Arreglo(0, 1667) = "Silas"
    Arreglo(0, 1668) = "Silk"
    Arreglo(0, 1669) = "Silva"
    Arreglo(0, 1670) = "Silvas"
    Arreglo(0, 1671) = "Silver"
    Arreglo(0, 1672) = "Silversmith"
    Arreglo(0, 1673) = "Simmons"
    Arreglo(0, 1674) = "Simms"
    Arreglo(0, 1675) = "Simon"
    Arreglo(0, 1676) = "Simpson"
    Arreglo(0, 1677) = "Sims"
    Arreglo(0, 1678) = "Sinclair"
    Arreglo(0, 1679) = "Singer"
    Arreglo(0, 1680) = "Singh"
    Arreglo(0, 1681) = "Singleton"
    Arreglo(0, 1682) = "Six"
    Arreglo(0, 1683) = "Sixkiller"
    Arreglo(0, 1684) = "Sizemore"
    Arreglo(0, 1685) = "Skeet"
    Arreglo(0, 1686) = "Skeets"
    Arreglo(0, 1687) = "Skenandore"
    Arreglo(0, 1688) = "Skidmore"
    Arreglo(0, 1689) = "Skinner"
    Arreglo(0, 1690) = "Slater"
    Arreglo(0, 1691) = "Slim"
    Arreglo(0, 1692) = "Sloan"
    Arreglo(0, 1693) = "Small"
    Arreglo(0, 1694) = "Smallcanyon"
    Arreglo(0, 1695) = "Smallwood"
    Arreglo(0, 1696) = "Smart"
    Arreglo(0, 1697) = "Smiley"
    Arreglo(0, 1698) = "Smiling"
    Arreglo(0, 1699) = "Smith"
    Arreglo(0, 1700) = "Smoke"
    Arreglo(0, 1701) = "Sneed"
    Arreglo(0, 1702) = "Snell"
    Arreglo(0, 1703) = "Snider"
    Arreglo(0, 1704) = "Snow"
    Arreglo(0, 1705) = "Snyder"
    Arreglo(0, 1706) = "Soap"
    Arreglo(0, 1707) = "Solomon"
    Arreglo(0, 1708) = "Sorrell"
    Arreglo(0, 1709) = "Soto"
    Arreglo(0, 1710) = "Spang"
    Arreglo(0, 1711) = "Sparks"
    Arreglo(0, 1712) = "Spaulding"
    Arreglo(0, 1713) = "Spears"
    Arreglo(0, 1714) = "Spence"
    Arreglo(0, 1715) = "Spencer"
    Arreglo(0, 1716) = "Spoonhunter"
    Arreglo(0, 1717) = "Spottedbear"
    Arreglo(0, 1718) = "Spottedelk"
    Arreglo(0, 1719) = "Sprague"
    Arreglo(0, 1720) = "Springer"
    Arreglo(0, 1721) = "Stacy"
    Arreglo(0, 1722) = "Stafford"
    Arreglo(0, 1723) = "Standingbear"
    Arreglo(0, 1724) = "Stands"
    Arreglo(0, 1725) = "Stanley"
    Arreglo(0, 1726) = "Stanton"
    Arreglo(0, 1727) = "Staples"
    Arreglo(0, 1728) = "Stark"
    Arreglo(0, 1729) = "Starr"
    Arreglo(0, 1730) = "Stately"
    Arreglo(0, 1731) = "Stclair"
    Arreglo(0, 1732) = "Stclaire"
    Arreglo(0, 1733) = "Steele"
    Arreglo(0, 1734) = "Stephens"
    Arreglo(0, 1735) = "Stephenson"
    Arreglo(0, 1736) = "Steve"
    Arreglo(0, 1737) = "Stevens"
    Arreglo(0, 1738) = "Stevenson"
    Arreglo(0, 1739) = "Stewart"
    Arreglo(0, 1740) = "Stgermaine"
    Arreglo(0, 1741) = "Stiffarm"
    Arreglo(0, 1742) = "Stiles"
    Arreglo(0, 1743) = "Still"
    Arreglo(0, 1744) = "Stinson"
    Arreglo(0, 1745) = "Stjohn"
    Arreglo(0, 1746) = "Stokes"
    Arreglo(0, 1747) = "Stone"
    Arreglo(0, 1748) = "Stout"
    Arreglo(0, 1749) = "Stover"
    Arreglo(0, 1750) = "Stpierre"
    Arreglo(0, 1751) = "Street"
    Arreglo(0, 1752) = "Strickland"
    Arreglo(0, 1753) = "Strong"
    Arreglo(0, 1754) = "Stroud"
    Arreglo(0, 1755) = "Stuart"
    Arreglo(0, 1756) = "Stump"
    Arreglo(0, 1757) = "Suazo"
    Arreglo(0, 1758) = "Sullivan"
    Arreglo(0, 1759) = "Summers"
    Arreglo(0, 1760) = "Sumner"
    Arreglo(0, 1761) = "Sunday"
    Arreglo(0, 1762) = "Sutherland"
    Arreglo(0, 1763) = "Sutton"
    Arreglo(0, 1764) = "Swain"
    Arreglo(0, 1765) = "Swallow"
    Arreglo(0, 1766) = "Swan"
    Arreglo(0, 1767) = "Swann"
    Arreglo(0, 1768) = "Swanson"
    Arreglo(0, 1769) = "Sweat"
    Arreglo(0, 1770) = "Sweeney"
    Arreglo(0, 1771) = "Sweet"
    Arreglo(0, 1772) = "Swift"
    Arreglo(0, 1773) = "Swimmer"
    Arreglo(0, 1774) = "Tabaha"
    Arreglo(0, 1775) = "Tackett"
    Arreglo(0, 1776) = "Tafoya"
    Arreglo(0, 1777) = "Talley"
    Arreglo(0, 1778) = "Tallman"
    Arreglo(0, 1779) = "Tanner"
    Arreglo(0, 1780) = "Tapaha"
    Arreglo(0, 1781) = "Tapia"
    Arreglo(0, 1782) = "Tarbell"
    Arreglo(0, 1783) = "Tate"
    Arreglo(0, 1784) = "Tatum"
    Arreglo(0, 1785) = "Taylor"
    Arreglo(0, 1786) = "Teague"
    Arreglo(0, 1787) = "Teehee"
    Arreglo(0, 1788) = "Teller"
    Arreglo(0, 1789) = "Tenorio"
    Arreglo(0, 1790) = "Terrance"
    Arreglo(0, 1791) = "Terrell"
    Arreglo(0, 1792) = "Terry"
    Arreglo(0, 1793) = "Tessay"
    Arreglo(0, 1794) = "Thayer"
    Arreglo(0, 1795) = "Thomas"
    Arreglo(0, 1796) = "Thomason"
    Arreglo(0, 1797) = "Thompson"
    Arreglo(0, 1798) = "Thorne"
    Arreglo(0, 1799) = "Thornton"
    Arreglo(0, 1800) = "Thorpe"
    Arreglo(0, 1801) = "Thunder"
    Arreglo(0, 1802) = "Thunderhawk"
    Arreglo(0, 1803) = "Thurman"
    Arreglo(0, 1804) = "Tibbetts"
    Arreglo(0, 1805) = "Tidwell"
    Arreglo(0, 1806) = "Tiger"
    Arreglo(0, 1807) = "Tilley"
    Arreglo(0, 1808) = "Tillman"
    Arreglo(0, 1809) = "Tinker"
    Arreglo(0, 1810) = "Tipton"
    Arreglo(0, 1811) = "Titus"
    Arreglo(0, 1812) = "Todacheenie"
    Arreglo(0, 1813) = "Todd"
    Arreglo(0, 1814) = "Toledo"
    Arreglo(0, 1815) = "Tom"
    Arreglo(0, 1816) = "Tompkins"
    Arreglo(0, 1817) = "Toney"
    Arreglo(0, 1818) = "Torivio"
    Arreglo(0, 1819) = "Torres"
    Arreglo(0, 1820) = "Towne"
    Arreglo(0, 1821) = "Townsend"
    Arreglo(0, 1822) = "Toya"
    Arreglo(0, 1823) = "Tracey"
    Arreglo(0, 1824) = "Tracy"
    Arreglo(0, 1825) = "Trammell"
    Arreglo(0, 1826) = "Traversie"
    Arreglo(0, 1827) = "Travis"
    Arreglo(0, 1828) = "Trevino"
    Arreglo(0, 1829) = "Tripp"
    Arreglo(0, 1830) = "Trottier"
    Arreglo(0, 1831) = "Trout"
    Arreglo(0, 1832) = "Truax"
    Arreglo(0, 1833) = "Trujillo"
    Arreglo(0, 1834) = "Tsethlikai"
    Arreglo(0, 1835) = "Tsinnie"
    Arreglo(0, 1836) = "Tsinnijinnie"
    Arreglo(0, 1837) = "Tso"
    Arreglo(0, 1838) = "Tsosie"
    Arreglo(0, 1839) = "Tubby"
    Arreglo(0, 1840) = "Tucker"
    Arreglo(0, 1841) = "Turner"
    Arreglo(0, 1842) = "Turney"
    Arreglo(0, 1843) = "Tuttle"
    Arreglo(0, 1844) = "Twiss"
    Arreglo(0, 1845) = "Twobulls"
    Arreglo(0, 1846) = "Twocrow"
    Arreglo(0, 1847) = "Tyler"
    Arreglo(0, 1848) = "Tyndall"
    Arreglo(0, 1849) = "Tyner"
    Arreglo(0, 1850) = "Tyson"
    Arreglo(0, 1851) = "Underwood"
    Arreglo(0, 1852) = "Upshaw"
    Arreglo(0, 1853) = "Valdez"
    Arreglo(0, 1854) = "Valencia"
    Arreglo(0, 1855) = "Valentine"
    Arreglo(0, 1856) = "Valenzuela"
    Arreglo(0, 1857) = "Vallo"
    Arreglo(0, 1858) = "Vance"
    Arreglo(0, 1859) = "Vandever"
    Arreglo(0, 1860) = "Vandunk"
    Arreglo(0, 1861) = "Vann"
    Arreglo(0, 1862) = "Vanwinkle"
    Arreglo(0, 1863) = "Vargas"
    Arreglo(0, 1864) = "Vasquez"
    Arreglo(0, 1865) = "Vaughan"
    Arreglo(0, 1866) = "Vaughn"
    Arreglo(0, 1867) = "Vega"
    Arreglo(0, 1868) = "Velarde"
    Arreglo(0, 1869) = "Velasquez"
    Arreglo(0, 1870) = "Ventura"
    Arreglo(0, 1871) = "Verdin"
    Arreglo(0, 1872) = "Verret"
    Arreglo(0, 1873) = "Vicente"
    Arreglo(0, 1874) = "Vicenti"
    Arreglo(0, 1875) = "Victor"
    Arreglo(0, 1876) = "Vigil"
    Arreglo(0, 1877) = "Villa"
    Arreglo(0, 1878) = "Villegas"
    Arreglo(0, 1879) = "Vincent"
    Arreglo(0, 1880) = "Vinson"
    Arreglo(0, 1881) = "Vivier"
    Arreglo(0, 1882) = "Wade"
    Arreglo(0, 1883) = "Wadsworth"
    Arreglo(0, 1884) = "Wagner"
    Arreglo(0, 1885) = "Walden"
    Arreglo(0, 1886) = "Walker"
    Arreglo(0, 1887) = "Walkingeagle"
    Arreglo(0, 1888) = "Walkingstick"
    Arreglo(0, 1889) = "Wall"
    Arreglo(0, 1890) = "Wallace"
    Arreglo(0, 1891) = "Waller"
    Arreglo(0, 1892) = "Walls"
    Arreglo(0, 1893) = "Walsh"
    Arreglo(0, 1894) = "Walter"
    Arreglo(0, 1895) = "Walters"
    Arreglo(0, 1896) = "Walton"
    Arreglo(0, 1897) = "Ward"
    Arreglo(0, 1898) = "Ware"
    Arreglo(0, 1899) = "Warner"
    Arreglo(0, 1900) = "Warren"
    Arreglo(0, 1901) = "Warrior"
    Arreglo(0, 1902) = "Washburn"
    Arreglo(0, 1903) = "Washington"
    Arreglo(0, 1904) = "Wassillie"
    Arreglo(0, 1905) = "Watchman"
    Arreglo(0, 1906) = "Waterman"
    Arreglo(0, 1907) = "Waters"
    Arreglo(0, 1908) = "Watkins"
    Arreglo(0, 1909) = "Watson"
    Arreglo(0, 1910) = "Watt"
    Arreglo(0, 1911) = "Watts"
    Arreglo(0, 1912) = "Wauneka"
    Arreglo(0, 1913) = "Waupoose"
    Arreglo(0, 1914) = "Weaver"
    Arreglo(0, 1915) = "Webb"
    Arreglo(0, 1916) = "Webber"
    Arreglo(0, 1917) = "Weber"
    Arreglo(0, 1918) = "Webster"
    Arreglo(0, 1919) = "Weeks"
    Arreglo(0, 1920) = "Welch"
    Arreglo(0, 1921) = "Wells"
    Arreglo(0, 1922) = "Welsh"
    Arreglo(0, 1923) = "Werito"
    Arreglo(0, 1924) = "Wesley"
    Arreglo(0, 1925) = "West"
    Arreglo(0, 1926) = "Westbrook"
    Arreglo(0, 1927) = "Weston"
    Arreglo(0, 1928) = "Wheeler"
    Arreglo(0, 1929) = "Whipple"
    Arreglo(0, 1930) = "Whitaker"
    Arreglo(0, 1931) = "White"
    Arreglo(0, 1932) = "Whitebear"
    Arreglo(0, 1933) = "Whitebird"
    Arreglo(0, 1934) = "Whitebull"
    Arreglo(0, 1935) = "Whiteeagle"
    Arreglo(0, 1936) = "Whitefeather"
    Arreglo(0, 1937) = "Whitehair"
    Arreglo(0, 1938) = "Whitehat"
    Arreglo(0, 1939) = "Whitehead"
    Arreglo(0, 1940) = "Whitehorse"
    Arreglo(0, 1941) = "Whitehouse"
    Arreglo(0, 1942) = "Whiteman"
    Arreglo(0, 1943) = "Whiterock"
    Arreglo(0, 1944) = "Whitfield"
    Arreglo(0, 1945) = "Whitford"
    Arreglo(0, 1946) = "Whiting"
    Arreglo(0, 1947) = "Whitley"
    Arreglo(0, 1948) = "Whitman"
    Arreglo(0, 1949) = "Whitney"
    Arreglo(0, 1950) = "Wiggins"
    Arreglo(0, 1951) = "Wilber"
    Arreglo(0, 1952) = "Wilbur"
    Arreglo(0, 1953) = "Wilcox"
    Arreglo(0, 1954) = "Wildcat"
    Arreglo(0, 1955) = "Wilder"
    Arreglo(0, 1956) = "Wiley"
    Arreglo(0, 1957) = "Wilkerson"
    Arreglo(0, 1958) = "Wilkie"
    Arreglo(0, 1959) = "Wilkins"
    Arreglo(0, 1960) = "Wilkinson"
    Arreglo(0, 1961) = "Willard"
    Arreglo(0, 1962) = "Williams"
    Arreglo(0, 1963) = "Williamson"
    Arreglo(0, 1964) = "Willie"
    Arreglo(0, 1965) = "Willis"
    Arreglo(0, 1966) = "Willoughby"
    Arreglo(0, 1967) = "Wills"
    Arreglo(0, 1968) = "Wilson"
    Arreglo(0, 1969) = "Wind"
    Arreglo(0, 1970) = "Wing"
    Arreglo(0, 1971) = "Winn"
    Arreglo(0, 1972) = "Winters"
    Arreglo(0, 1973) = "Wise"
    Arreglo(0, 1974) = "Witt"
    Arreglo(0, 1975) = "Wofford"
    Arreglo(0, 1976) = "Wolf"
    Arreglo(0, 1977) = "Wolfe"
    Arreglo(0, 1978) = "Womack"
    Arreglo(0, 1979) = "Wood"
    Arreglo(0, 1980) = "Woodall"
    Arreglo(0, 1981) = "Woodard"
    Arreglo(0, 1982) = "Woodruff"
    Arreglo(0, 1983) = "Woods"
    Arreglo(0, 1984) = "Woodward"
    Arreglo(0, 1985) = "Woody"
    Arreglo(0, 1986) = "Wooten"
    Arreglo(0, 1987) = "Workman"
    Arreglo(0, 1988) = "Worley"
    Arreglo(0, 1989) = "Wright"
    Arreglo(0, 1990) = "Wyatt"
    Arreglo(0, 1991) = "Wynn"
    Arreglo(0, 1992) = "Yahola"
    Arreglo(0, 1993) = "Yankton"
    Arreglo(0, 1994) = "Yarbrough"
    Arreglo(0, 1995) = "Yates"
    Arreglo(0, 1996) = "Yazzie"
    Arreglo(0, 1997) = "Yellow"
    Arreglo(0, 1998) = "Yellowhair"
    Arreglo(0, 1999) = "Yellowhorse"
    Arreglo(0, 2000) = "Yellowman"
    Arreglo(0, 2001) = "Yepa"
    Arreglo(0, 2002) = "York"
    Arreglo(0, 2003) = "Young"
    Arreglo(0, 2004) = "Youngbear"
    Arreglo(0, 2005) = "Youngbird"
    Arreglo(0, 2006) = "Youngblood"
    Arreglo(0, 2007) = "Youngman"
    Arreglo(0, 2008) = "Zamora"
    Arreglo(0, 2009) = "Zephier"
    Arreglo(0, 2010) = "Zimmerman"
    Arreglo(0, 2011) = "Zuni"

   
    Set Myrange = Range("A2:B2013")
    i = 0
    
    j = 0
            
    For Each Cell In Myrange
        
        If j Mod 2 = 0 Then
        
            Cell.Value = i + 1
        
            i = i + 1
            j = j + 1
        Else
        
            Cell.Value = Arreglo(0, i - 1)
            
            j = j + 1
        End If
        
            
    Next Cell

    Range("A1").Value = "Order"
    Range("A1").Font.Bold = True
    
    Range("B1").Value = "Last name / Surname"
    Range("B1").Font.Bold = True
    
    Columns("A:A").EntireColumn.AutoFit
    Columns("A:A").HorizontalAlignment = xlCenter
    
    Columns("B:B").EntireColumn.AutoFit
        
    Range("A2").Select
    ActiveWindow.FreezePanes = True
    
End Sub

Sub MacroN_All()

    Call Macro00_NewFile
    Call Macro01_Headers
    
    Call Macro03_InsertCitiesSheet
    Call Macro04_InsertSheets
    Call Macro05_Fill_Category
    Call Macro06_Fill_Sub_Category
           
    Call Macro08_Fill_Accesories
    Call Macro09_Fill_Appliances
    Call Macro10_Fill_Binders
    Call Macro11_Fill_Art
    Call Macro12_Fill_Bookcases
    Call Macro13_Fill_Chairs
    Call Macro14_Fill_Copiers
    Call Macro15_Fill_Envelopes
    Call Macro16_Fill_Fasteners
    Call Macro17_Fill_Furnishings
    Call Macro18_Fill_Labels
    Call Macro19_Fill_Gym_Machines
    Call Macro20_Fill_Papers
    Call Macro21_Fill_Storage
    Call Macro22_Fill_Supplies
    Call Macro23_Fill_Tables
    
    Call Macro24_Fill_Products
            
        Call MacroN_menos_2_Fill_Names
        Call MacroN_menos_1_Fill_LastNames
        
    Call Macro02_OrdersWithRANDQty
    
    Sheets("Sales SuperStore").Select
    
End Sub
