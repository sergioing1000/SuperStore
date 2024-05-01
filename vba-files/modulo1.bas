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