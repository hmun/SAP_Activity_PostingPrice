VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DateFormatString"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Function getString() As String
    Dim aDateOrder As Long

    aDateOrder = Application.International(xlDateOrder)

    Select Case aDateOrder
        Case 0
            getString = "mm/dd/yyyy"
        Case 1
            getString = "dd/mm/yyyy"
        Case 2
            getString = "yyyy/mm/dd"
    End Select
End Function

