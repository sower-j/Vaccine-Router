Attribute VB_Name = "Setup"
Sub Setup()
Attribute Setup.VB_Description = "Format Rows and Columns to be more readable"
Attribute Setup.VB_ProcData.VB_Invoke_Func = "s\n14"

' Setup Macro to:
' reformat rows and columns to be more readable
' Assign routes based on zip code
' Sort by route number then by zip code

    DocClean
    Call DynamicRoute.DynamicRoute
End Sub

Sub DocClean()

Attribute DocClean.VB_Description = "Execute VaccineSetup and Vaccine Router with one command"
Attribute DocClean.VB_ProcData.VB_Invoke_Func = " \n14"

' Formats Rows and Columns to be more readable

Dim ColumnHead As Variant ' Array used to search and reorder columns
Dim search As Range
Dim indx, count As Integer
    
' Sort Columns and delete extras
ColumnHead = Array( _
            "Route", _
            "Seq", _
            "Airbill", _
            "Address", _
            "Zip", _
            "Commit Time", _
            "Cmt")
count = 1

Application.ScreenUpdating = False

For indx = LBound(ColumnHead) To UBound(ColumnHead)
    Set search = Rows(1).Find(ColumnHead(indx), _
                LookIn:=xlValues, _
                LookAt:=xlPart, _
                SearchOrder:=xlByColumns, _
                SearchDirection:=xlNext, _
                MatchCase:=False)

    If Not search Is Nothing Then
        If search.column <> count Then
            search.EntireColumn.Cut
            Columns(count).Insert Shift:=xlToRight
            Application.CutCopyMode = False
        End If
        
    count = count + 1
    
    End If
    
Next indx

ActiveSheet.UsedRange.Offset(, indx - 2).ClearContents

' Ensure column B uses standard number format
    Columns("B:B").Select
    Selection.NumberFormat = "0"
    
' Auto expand columns to fit text and make them sortable
    Columns("A:E").Select
    Columns("A:E").EntireColumn.AutoFit
    Selection.AutoFilter
    
' Center text and disable text wrapping
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With


End Sub