Attribute VB_Name = "SAPMakro"
Public sActivity As Integer

Sub SAP_Activity_post()
    SAP_Activity_exec ("post")
End Sub

Sub SAP_Activity_check()
    SAP_Activity_exec ("check")
End Sub

Sub SAP_Activity_exec(p_mode As String)
    Dim aSAPAcctngActivityAlloc As New SAPAcctngActivityAlloc
    Dim aSAPDocItem As New SAPDocItem
    Dim aDateFormatString As New DateFormatString
    Dim aData As New Collection
    Dim aRetStr As String

    Dim bRetStr As String

    Dim aKOKRS As String
    Dim aEB As String
    Dim aFromLine As Integer
    Dim aToLine As Integer

    Dim aBLDAT As String
    Dim aBUDAT As String
    Dim aMENGE As String
    Dim aEPSP As String
    Dim aSKOSTL As String
    Dim aLEART As String

    Worksheets("Parameter").Activate
    aKOKRS = Format(Cells(2, 2), "0000")
    aEB = Cells(3, 2)
    If IsNull(aKOKRS) Or aKOKRS = "" Then
        MsgBox "Bitte alle Mussfelder der Parameter füllen!", vbCritical + vbOKOnly
        Exit Sub
    End If
    aRet = SAPCheck()
    If Not aRet Then
        MsgBox "Connection to SAP failed!", vbCritical + vbOKOnly
        Exit Sub
    End If

    Worksheets("Data").Activate
    i = 2
    Do
        If InStr(Cells(i, 13), "Beleg wird unter der Nummer") = 0 And InStr(Cells(i, 13), "Document is posted under number") = 0 Then
            If aBUDAT = "" Or aEB = "J" Then
                aBUDAT = Format(Cells(i, 1), aDateFormatString.getString)
                aBLDAT = Format(Cells(i, 2), aDateFormatString.getString)
            End If
            Set aSAPDocItem = New SAPDocItem
            aSAPDocItem.create Cells(i, 3).Value, Cells(i, 4).Value, Cells(i, 5).Value, CDbl(Cells(i, 6).Value), _
            Cells(i, 7).Value, Cells(i, 8).Value, Cells(i, 9).Value, Cells(i, 10).Value, _
            Cells(i, 11).Value, Cells(i, 12).Value, _
            CDbl(Cells(i, 13).Value), CDbl(Cells(i, 14).Value), CDbl(Cells(i, 15).Value), CInt(Cells(i, 16).Value), Cells(i, 17).Value
            aData.Add aSAPDocItem
            If aEB = "J" Or aEB = "Y" Then
                If p_mode = "post" Then
                    aRetStr = aSAPAcctngActivityAlloc.post(aKOKRS, aBUDAT, aBLDAT, aData)
                Else
                    aRetStr = aSAPAcctngActivityAlloc.check(aKOKRS, aBUDAT, aBLDAT, aData)
                End If
                Cells(i, 18) = aRetStr
                Set aData = New Collection
            End If
        End If
        i = i + 1
    Loop While Not IsNull(Cells(i, 1)) And Cells(i, 1) <> ""
    If aEB <> "J" And aEB <> "Y" Then
        If p_mode = "post" Then
            aRetStr = aSAPAcctngActivityAlloc.post(aKOKRS, aBUDAT, aBLDAT, aData)
        Else
            aRetStr = aSAPAcctngActivityAlloc.check(aKOKRS, aBUDAT, aBLDAT, aData)
        End If
        Cells(i, 18) = aRetStr
    End If
End Sub

