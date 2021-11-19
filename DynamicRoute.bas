Attribute VB_Name = "DynamicRoute"
Option Explicit

Sub DynamicRoute()
Dim index, Route As Integer
Dim lastRow, refZip, zip, rzcell, zcell As Range

Application.ScreenUpdating = False

Set zip = Workbooks(2).Worksheets(1).Range(Cells(2, 4), Cells(2, 4).End(xlDown))

    For Each zcell In zip
        Trim (zcell)
        If Len(zcell) > 5 Then
            zcell = Left(zcell, Len(zcell) - 4)
        End If
    Next zcell

index = 3
    
    Do While Cells(1, index).Value <> ""
        Workbooks(1).Activate
        Route = Cells(1, index).Value
        Set refZip = Workbooks(1).Worksheets(1).Range(Cells(3, index), Cells(3, index).End(xlDown))
        Workbooks(2).Activate
        
        For Each rzcell In refZip
            For Each zcell In zip
                If zcell.Value = rzcell.Value Then
                    Cells(zcell.row, 1).Value = Route
                End If
            Next zcell
        Next rzcell
        index = index + 1
    Loop
    
Application.ScreenUpdating = True

End Sub
