VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SAPAcctngManCostAlloc"
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
    Dim boAcctngManCostAlloc As Object
    Dim oDocHeader As Object
    Dim oDocItems As Object
    Dim oRETURN As Object
    Dim oDocNo As String
    Dim lSAPFormat As New SAPFormat

    Set boAcctngManCostAlloc = MySAPCon.SAPBapiControl.GetSAPObject("AcctngManCostAlloc")
    Set oRETURN = MySAPCon.SAPBapiControl.DimAs(boAcctngManCostAlloc, "Post", "Return")
    Set oDocHeader = MySAPCon.SAPBapiControl.DimAs(boAcctngManCostAlloc, "Post", "DocHeader")
    Set oDocItems = MySAPCon.SAPBapiControl.DimAs(boAcctngManCostAlloc, "Post", "DocItems")
    oDocHeader("CO_AREA") = pKokrs
    oDocHeader("DOCDATE") = pBldat
    oDocHeader("POSTGDATE") = pBuDat
    oDocHeader("USERNAME") = MySAPCon.SAPCon.User
    Dim aItemRow As Object
    Dim lRow As Object
    For Each lRow In pData
        Set aItemRow = oDocItems.Rows.Add
        aItemRow.Value("SEND_CCTR") = lSAPFormat.unpack(lRow.SEND_CCTR, 10)
        aItemRow.Value("SEN_ORDER") = lSAPFormat.unpack(lRow.SEN_ORDER, 12)
        aItemRow.Value("SEN_WBS_EL") = lRow.SEN_WBS_EL
        aItemRow.Value("SEN_NETWRK") = lSAPFormat.unpack(lRow.SEN_NETWRK, 12)
        aItemRow.Value("SENOPERATN") = lSAPFormat.unpack(lRow.SENOPERATN, 4)
        aItemRow.Value("SEND_FUNCTION") = lRow.SEND_FUNCTION
        aItemRow.Value("PERSON_NO") = lSAPFormat.unpack(lRow.PERSON_NO, 8)
        aItemRow.Value("COST_ELEM") = lSAPFormat.unpack(lRow.COST_ELEM, 10)
        aItemRow.Value("VALUE_TCUR") = Format$(lRow.VALUE_TCUR, "0.00")
        aItemRow.Value("SEG_TEXT") = lRow.SEG_TEXT
        aItemRow.Value("REC_CCTR") = lSAPFormat.unpack(lRow.REC_CCTR, 10)
        aItemRow.Value("REC_ORDER") = lSAPFormat.unpack(lRow.REC_ORDER, 12)
        aItemRow.Value("REC_WBS_EL") = lRow.REC_WBS_EL
        aItemRow.Value("REC_NETWRK") = lSAPFormat.unpack(lRow.REC_NETWRK, 12)
        aItemRow.Value("RECOPERATN") = lSAPFormat.unpack(lRow.RECOPERATN, 4)
        aItemRow.Value("REC_FUNCTION") = lRow.REC_FUNCTION
        aItemRow.Value("TRANS_CURR") = lRow.TRANS_CURR
    Next lRow

    boAcctngManCostAlloc.post DocHeader:=oDocHeader, _
    IgnoreWarnings:="X", _
    DocItems:=oDocItems, _
    DocNo:=oDocNo, _
    RETURN:=oRETURN
    Dim RetRow As Object
    Dim aErr As Boolean
    aErr = False
    For Each RetRow In oRETURN.Rows
        If RetRow("TYPE") <> "I" And RetRow("TYPE") <> "W" Then
            post = post & ";" & RetRow("TYPE") & "-" & RetRow("MESSAGE")
            If RetRow("TYPE") <> "S" And RetRow("TYPE") <> "W" Then
                aErr = True
            End If
        End If
    Next RetRow
    If aErr = False Then
        Dim aSAPBapiTranctionCommit As New SAPBapiTranctionCommit
        aSAPBapiTranctionCommit.commit
    End If
    Set boAcctngManCostAlloc = Nothing
    Exit Function
post_Error:
    post = "Fehler"
    MySAPErr.MSGProt "SAPAcctngManCostAlloc", "post", "", Err.Number, Err.Description
    Exit Function
End Function

Public Function check(pKokrs As String, pBuDat As String, pBldat As String, pData As Collection) As String
    On Error GoTo check_Error
    Dim boAcctngManCostAlloc As Object
    Dim oDocHeader As Object
    Dim oDocItems As Object
    Dim oRETURN As Object
    Dim oDocNo As String
    Dim lSAPFormat As New SAPFormat

    Set boAcctngManCostAlloc = MySAPCon.SAPBapiControl.GetSAPObject("AcctngManCostAlloc")
    Set oRETURN = MySAPCon.SAPBapiControl.DimAs(boAcctngManCostAlloc, "Check", "Return")
    Set oDocHeader = MySAPCon.SAPBapiControl.DimAs(boAcctngManCostAlloc, "Check", "DocHeader")
    Set oDocItems = MySAPCon.SAPBapiControl.DimAs(boAcctngManCostAlloc, "Check", "DocItems")
    oDocHeader("CO_AREA") = pKokrs
    oDocHeader("DOCDATE") = pBldat
    oDocHeader("POSTGDATE") = pBuDat
    oDocHeader("USERNAME") = MySAPCon.SAPCon.User
    Dim aItemRow As Object
    Dim lRow As Object
    For Each lRow In pData
        Set aItemRow = oDocItems.Rows.Add
        aItemRow.Value("SEND_CCTR") = lSAPFormat.unpack(lRow.SEND_CCTR, 10)
        aItemRow.Value("SEN_ORDER") = lSAPFormat.unpack(lRow.SEN_ORDER, 12)
        aItemRow.Value("SEN_WBS_EL") = lRow.SEN_WBS_EL
        aItemRow.Value("SEN_NETWRK") = lSAPFormat.unpack(lRow.SEN_NETWRK, 12)
        aItemRow.Value("SENOPERATN") = lSAPFormat.unpack(lRow.SENOPERATN, 4)
        aItemRow.Value("SEND_FUNCTION") = lRow.SEND_FUNCTION
        aItemRow.Value("PERSON_NO") = lSAPFormat.unpack(lRow.PERSON_NO, 8)
        aItemRow.Value("COST_ELEM") = lSAPFormat.unpack(lRow.COST_ELEM, 10)
        aItemRow.Value("VALUE_TCUR") = Format$(lRow.VALUE_TCUR, "0.00")
        aItemRow.Value("SEG_TEXT") = lRow.SEG_TEXT
        aItemRow.Value("REC_CCTR") = lSAPFormat.unpack(lRow.REC_CCTR, 10)
        aItemRow.Value("REC_ORDER") = lSAPFormat.unpack(lRow.REC_ORDER, 12)
        aItemRow.Value("REC_WBS_EL") = lRow.REC_WBS_EL
        aItemRow.Value("REC_NETWRK") = lSAPFormat.unpack(lRow.REC_NETWRK, 12)
        aItemRow.Value("RECOPERATN") = lSAPFormat.unpack(lRow.RECOPERATN, 4)
        aItemRow.Value("REC_FUNCTION") = lRow.REC_FUNCTION
        aItemRow.Value("TRANS_CURR") = lRow.TRANS_CURR
    Next lRow

    boAcctngManCostAlloc.check DocHeader:=oDocHeader, _
    IgnoreWarnings:="X", _
    DocItems:=oDocItems, _
    RETURN:=oRETURN
    Dim RetRow As Object
    Dim aErr As Boolean
    aErr = False
    For Each RetRow In oRETURN.Rows
        check = check & ";" & RetRow("TYPE") & "-" & RetRow("MESSAGE")
        If RetRow("TYPE") <> "S" And RetRow("TYPE") <> "W" Then
            aErr = True
        End If
    Next RetRow
    If aErr = False Then
        Dim aSAPBapiTranctionCommit As New SAPBapiTranctionCommit
        aSAPBapiTranctionCommit.commit
    End If
    Set boAcctngManCostAlloc = Nothing
    Exit Function
check_Error:
    check = "Fehler"
    MySAPErr.MSGProt "SAPAcctngManCostAlloc", "check", "", Err.Number, Err.Description
    Exit Function
End Function




