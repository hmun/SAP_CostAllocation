VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SAPDocItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public SEND_CCTR As String
Public SEN_ORDER As String
Public SEN_WBS_EL As String
Public SEN_NETWRK As String
Public SENOPERATN As String
Public SEND_FUNCTION As String
Public PERSON_NO As String
Public COST_ELEM As String
Public VALUE_TCUR As Double
Public SEG_TEXT As String
Public REC_CCTR As String
Public REC_ORDER As String
Public REC_WBS_EL As String
Public REC_NETWRK As String
Public RECOPERATN As String
Public REC_FUNCTION As String
Public TRANS_CURR As String

Public Function create(pSEND_CCTR As String, pSEN_ORDER As String, pSEN_WBS_EL As String, pSEN_NETWRK As String, _
    pSENOPERATN As String, pSEND_FUNCTION As String, _
    pPERSON_NO As String, pCOST_ELEM As String, pVALUE_TCUR As Double, pSEG_TEXT As String, _
    pREC_CCTR As String, pREC_ORDER As String, pREC_WBS_EL As String, pREC_NETWRK As String, _
    pRECOPERATN As String, pREC_FUNCTION As String, _
    pTRANS_CURR As String)
    
    SEND_CCTR = pSEND_CCTR
    SEN_ORDER = pSEN_ORDER
    SEN_WBS_EL = pSEN_WBS_EL
    SEN_NETWRK = pSEN_NETWRK
    SENOPERATN = pSENOPERATN
    SEND_FUNCTION = pSEND_FUNCTION
    PERSON_NO = pPERSON_NO
    COST_ELEM = pCOST_ELEM
    VALUE_TCUR = pVALUE_TCUR
    SEG_TEXT = pSEG_TEXT
    REC_CCTR = pREC_CCTR
    REC_ORDER = pREC_ORDER
    REC_WBS_EL = pREC_WBS_EL
    REC_NETWRK = pREC_NETWRK
    RECOPERATN = pRECOPERATN
    REC_FUNCTION = pREC_FUNCTION
    TRANS_CURR = pTRANS_CURR
End Function
