VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Route_Click()
'This button runs the Setup module after opening another file

Dim fileName As String

' open up a selection window and assigns the file name as a string
fileName = Application.GetOpenFilename

' open the previously selected file and set it as the active Workbook
Workbooks.Open (fileName)

Workbooks(2).Activate
' run the Setup module

Call Setup.Setup

End Sub


