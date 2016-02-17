VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SAPAcctngActivityAlloc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public IsConnected As Integer

Private Sub Class_Initialize()
    IsConnected = SAPCheck()
End Sub

Public Function post(pKokrs As String, pBuDat As String, pBldat As String, pData As Collection) As String
    On Error GoTo post_Error
    Dim boAcctngActivityAlloc As Object
    Dim oDocHeader As Object
    Dim oDocItems As Object
    Dim oRETURN As Object
    Dim oDocNo As String
    Dim lSAPFormat As New SAPFormat

    Set boAcctngActivityAlloc = MySAPCon.SAPBapiControl.GetSAPObject("AcctngActivityAlloc")
    Set oRETURN = MySAPCon.SAPBapiControl.DimAs(boAcctngActivityAlloc, "Post", "Return")
    Set oDocHeader = MySAPCon.SAPBapiControl.DimAs(boAcctngActivityAlloc, "Post", "DocHeader")
    Set oDocItems = MySAPCon.SAPBapiControl.DimAs(boAcctngActivityAlloc, "Post", "DocItems")
    '  Set oDocNo = MySAPCon.SAPBapiControl.DimAs(boAcctngActivityAlloc, "Post", "DocNo")
    oDocHeader("CO_AREA") = pKokrs
    oDocHeader("DOCDATE") = pBldat
    oDocHeader("POSTGDATE") = pBuDat
    oDocHeader("USERNAME") = MySAPCon.SAPCon.User
    Dim aItemRow As Object
    Dim lRow As Object
    For Each lRow In pData
        Set aItemRow = oDocItems.Rows.Add
        aItemRow.Value("SEND_CCTR") = lSAPFormat.unpack(lRow.SEND_CCTR, 10)
        aItemRow.Value("PERSON_NO") = lRow.PERSON_NO
        aItemRow.Value("ACTTYPE") = lRow.ACTTYPE
        aItemRow.Value("ACTVTY_QTY") = Format$(lRow.ACTVTY_QTY, "0.000")
        aItemRow.Value("SEG_TEXT") = lRow.SEG_TEXT
        aItemRow.Value("REC_WBS_EL") = lRow.REC_WBS_EL
        aItemRow.Value("REC_NETWRK") = lSAPFormat.unpack(lRow.REC_NETWRK, 12)
        aItemRow.Value("RECOPERATN") = lSAPFormat.unpack(lRow.RECOPERATN, 4)
        aItemRow.Value("REC_ORDER") = lSAPFormat.unpack(lRow.REC_ORDER, 12)
        aItemRow.Value("REC_CCTR") = lSAPFormat.unpack(lRow.REC_CCTR, 10)
        If lRow.PRICE <> 0 Then
            aItemRow.Value("PRICE") = lRow.PRICE
        End If
        If lRow.PRICE_FIX <> 0 Then
            aItemRow.Value("PRICE_FIX") = lRow.PRICE_FIX
        End If
        If lRow.PRICE_VAR <> 0 Then
            aItemRow.Value("PRICE_VAR") = lRow.PRICE_VAR
        End If
        If lRow.PRICE_UNIT <> 0 Then
            aItemRow.Value("PRICE_UNIT") = lRow.PRICE_UNIT
        End If
        If lRow.CURR <> "" Then
            aItemRow.Value("CURRENCY") = lRow.CURR
        End If
    Next lRow

    boAcctngActivityAlloc.post DocHeader:=oDocHeader, _
    IgnoreWarnings:="X", _
    DocItems:=oDocItems, _
    DocNo:=oDocNo, _
    RETURN:=oRETURN
    Dim RetRow As Object
    Dim aErr As Boolean
    aErr = False
    For Each RetRow In oRETURN.Rows
        post = post & ";" & RetRow("TYPE") & "-" & RetRow("MESSAGE")
        If RetRow("TYPE") <> "S" And RetRow("TYPE") <> "I" And RetRow("TYPE") <> "W" Then
            aErr = True
        End If
    Next RetRow
    If aErr = False Then
        Dim aSAPBapiTranctionCommit As New SAPBapiTranctionCommit
        aSAPBapiTranctionCommit.commit
    End If
    Set boAcctngDocument = Nothing
    Exit Function
post_Error:
    post = "Error"
    MySAPErr.MSGProt "SAPAcctngActivityAlloc", "post", "", Err.Number, Err.Description
    Exit Function
End Function

Public Function check(pKokrs As String, pBuDat As String, pBldat As String, pData As Collection) As String
    On Error GoTo check_Error
    Dim boAcctngActivityAlloc As Object
    Dim oDocHeader As Object
    Dim oDocItems As Object
    Dim oRETURN As Object
    Dim oDocNo As String
    Dim lSAPFormat As New SAPFormat

    Set boAcctngActivityAlloc = MySAPCon.SAPBapiControl.GetSAPObject("AcctngActivityAlloc")
    Set oRETURN = MySAPCon.SAPBapiControl.DimAs(boAcctngActivityAlloc, "Check", "Return")
    Set oDocHeader = MySAPCon.SAPBapiControl.DimAs(boAcctngActivityAlloc, "Check", "DocHeader")
    Set oDocItems = MySAPCon.SAPBapiControl.DimAs(boAcctngActivityAlloc, "Check", "DocItems")
    '  Set oDocNo = MySAPCon.SAPBapiControl.DimAs(boAcctngActivityAlloc, "Post", "DocNo")
    oDocHeader("CO_AREA") = pKokrs
    oDocHeader("DOCDATE") = pBldat
    oDocHeader("POSTGDATE") = pBuDat
    oDocHeader("USERNAME") = MySAPCon.SAPCon.User
    Dim aItemRow As Object
    Dim lRow As Object
    For Each lRow In pData
        Set aItemRow = oDocItems.Rows.Add
        aItemRow.Value("SEND_CCTR") = lSAPFormat.unpack(lRow.SEND_CCTR, 10)
        aItemRow.Value("PERSON_NO") = lRow.PERSON_NO
        aItemRow.Value("ACTTYPE") = lRow.ACTTYPE
        aItemRow.Value("ACTVTY_QTY") = Format$(lRow.ACTVTY_QTY, "0.000")
        aItemRow.Value("SEG_TEXT") = lRow.SEG_TEXT
        aItemRow.Value("REC_WBS_EL") = lRow.REC_WBS_EL
        aItemRow.Value("REC_NETWRK") = lSAPFormat.unpack(lRow.REC_NETWRK, 12)
        aItemRow.Value("RECOPERATN") = lSAPFormat.unpack(lRow.RECOPERATN, 4)
        aItemRow.Value("REC_ORDER") = lSAPFormat.unpack(lRow.REC_ORDER, 12)
        aItemRow.Value("REC_CCTR") = lSAPFormat.unpack(lRow.REC_CCTR, 10)
    Next lRow

    boAcctngActivityAlloc.check DocHeader:=oDocHeader, _
    IgnoreWarnings:="X", _
    DocItems:=oDocItems, _
    RETURN:=oRETURN
    Dim RetRow As Object
    Dim aErr As Boolean
    aErr = False
    For Each RetRow In oRETURN.Rows
        check = check & ";" & RetRow("TYPE") & "-" & RetRow("MESSAGE")
        If RetRow("TYPE") <> "S" And RetRow("TYPE") <> "I" And RetRow("TYPE") <> "W" Then
            aErr = True
        End If
    Next RetRow
    If aErr = False Then
        Dim aSAPBapiTranctionCommit As New SAPBapiTranctionCommit
        aSAPBapiTranctionCommit.commit
    End If
    Set boAcctngDocument = Nothing
    Exit Function
check_Error:
    check = "Error"
    MySAPErr.MSGProt "SAPAcctngActivityAlloc", "check", "", Err.Number, Err.Description
    Exit Function
End Function





