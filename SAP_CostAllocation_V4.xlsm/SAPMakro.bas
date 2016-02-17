Attribute VB_Name = "SAPMakro"

Sub SAP_ManCostAlloc_post()
    SAP_ManCostAlloc_exec ("post")
End Sub

Sub SAP_ManCostAlloc_check()
    SAP_ManCostAlloc_exec ("check")
End Sub

Sub SAP_ManCostAlloc_exec(p_mode As String)
    Dim aSAPAcctngManCostAlloc As New SAPAcctngManCostAlloc
    Dim aDateFormatString As New DateFormatString
    Dim aSAPDocItem As New SAPDocItem
    Dim aData As New Collection
    Dim aRetStr As String

    Dim bRetStr As String

    Dim aKOKRS As String
    Dim aEB As String
    Dim aFromLine As Integer
    Dim aToLine As Integer

    Dim aBLDAT As String
    Dim aBUDAT As String
    Dim aNextBUDAT As String
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
    i = 3
    Do
        If InStr(Cells(i, 20), "Beleg wird unter der Nummer") = 0 And InStr(Cells(i, 20), "Document is posted under number") = 0 Then
            aBUDAT = Format(Cells(i, 1), aDateFormatString.getString)
            aBLDAT = Format(Cells(i, 2), aDateFormatString.getString)
            aNextBUDAT = Format(Cells(i + 1, 1), aDateFormatString.getString)
            Set aSAPDocItem = New SAPDocItem
            aSAPDocItem.create Cells(i, 3).Value, Cells(i, 4).Value, Cells(i, 5).Value, Cells(i, 6).Value, _
            Cells(i, 7).Value, Cells(i, 8).Value, Cells(i, 9).Value, _
            Cells(i, 10).Value, CDbl(Cells(i, 11).Value), Cells(i, 12).Value, _
            Cells(i, 13).Value, Cells(i, 14).Value, Cells(i, 15).Value, Cells(i, 16).Value, _
            Cells(i, 17).Value, Cells(i, 18).Value, Cells(i, 19).Value
            aData.Add aSAPDocItem
            If aEB = "J" Or aEB = "Y" Or aBUDAT <> aNextBUDAT Then
                If p_mode = "post" Then
                    aRetStr = aSAPAcctngManCostAlloc.post(aKOKRS, aBUDAT, aBLDAT, aData)
                Else
                    aRetStr = aSAPAcctngManCostAlloc.check(aKOKRS, aBUDAT, aBLDAT, aData)
                End If
                Cells(i, 20) = aRetStr
                Set aData = New Collection
            End If
        End If
        i = i + 1
    Loop While Not IsNull(Cells(i, 1)) And Cells(i, 1) <> ""
End Sub
