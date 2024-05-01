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

    Dim My_range As Range
    
    Set My_range = Range("A1:Y1")
        
    Range("A1").FormulaR1C1 = "Category"
    Columns("A:A").EntireColumn.AutoFit
    
    Range("B1").FormulaR1C1 = "Customer Name"
    Columns("B:B").EntireColumn.AutoFit
    
    Range("C1").FormulaR1C1 = "Order Date"
    Columns("C:C").EntireColumn.AutoFit
    
    Range("D1").FormulaR1C1 = "Order ID"
    Columns("D:D").EntireColumn.AutoFit
    
    Range("E1").FormulaR1C1 = "Product Name"
    Columns("E:E").EntireColumn.AutoFit
    
    Range("F1").FormulaR1C1 = "Unit Price"
    Columns("F:F").EntireColumn.AutoFit
    
    Range("G1").FormulaR1C1 = "Segment"
    Columns("G:G").EntireColumn.AutoFit
    
    Range("H1").FormulaR1C1 = "Ship Date"
    Columns("H:H").EntireColumn.AutoFit
    
    Range("I1").FormulaR1C1 = "Ship Mode"
    Columns("I:I").EntireColumn.AutoFit
    
    Range("J1").FormulaR1C1 = "Country"
    Columns("J:J").EntireColumn.AutoFit
    
    Range("K1").FormulaR1C1 = "Region"
    Columns("K:K").EntireColumn.AutoFit
    
    Range("L1").FormulaR1C1 = "State"
    Columns("L:L").EntireColumn.AutoFit
    
    Range("M1").FormulaR1C1 = "City order"
    Columns("M:M").EntireColumn.AutoFit
    
    Range("N1").FormulaR1C1 = "City"
    Columns("N:N").EntireColumn.AutoFit
    
    Range("O1").FormulaR1C1 = "Postal Code"
    Columns("O:O").EntireColumn.AutoFit
    
    Range("P1").FormulaR1C1 = "Sub-Category"
    Columns("P:P").EntireColumn.AutoFit
    
    Range("Q1").FormulaR1C1 = "Maesure Names"
    Columns("Q:Q").EntireColumn.AutoFit
    
    Range("R1").FormulaR1C1 = "Discount"
    Columns("R:R").EntireColumn.AutoFit
    
    Range("S1").FormulaR1C1 = "Profit"
    Columns("S:S").EntireColumn.AutoFit
    
    Range("T1").FormulaR1C1 = "Quantity"
    Columns("T:T").EntireColumn.AutoFit
    
    Range("U1").FormulaR1C1 = "Total Price"
    Columns("U:U").EntireColumn.AutoFit
    
    Range("V1").FormulaR1C1 = "Latitude"
    Columns("V:V").EntireColumn.AutoFit
    
    Range("W1").FormulaR1C1 = "Longitude"
    Columns("W:W").EntireColumn.AutoFit
    
    Range("X1").FormulaR1C1 = "Number of records"
    Columns("X:X").EntireColumn.AutoFit
    
    Range("Y1").FormulaR1C1 = "Sub-Region"
    Columns("Y:Y").EntireColumn.AutoFit
    
    My_range.Font.Bold = True
        
End Sub