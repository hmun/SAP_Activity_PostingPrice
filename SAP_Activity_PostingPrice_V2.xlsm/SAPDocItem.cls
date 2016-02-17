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
Public PERSON_NO As String
Public ACTTYPE As String
Public ACTVTY_QTY As Double
Public SEG_TEXT As String
Public REC_WBS_EL As String
Public REC_NETWRK As String
Public RECOPERATN As String
Public REC_ORDER As String
Public REC_CCTR As String
Public PRICE As Double
Public PRICE_FIX As Double
Public PRICE_VAR As Double
Public PRICE_UNIT As Integer
Public CURR As String



Public Function create(pSEND_CCTR As String, pPERSON_NO As String, pACTTYPE As String, pACTVTY_QTY As Double, pSEG_TEXT As String, _
    pREC_WBS_EL As String, pREC_NETWRK As String, pRECOPERATN As String, pREC_ORDER As String, pREC_CCTR As String, _
    pPRICE As Double, pPRICE_FIX As Double, pPRICE_VAR As Double, PRICE_UNIT As Integer, pCURR As String)
    SEND_CCTR = pSEND_CCTR
    PERSON_NO = pPERSON_NO
    ACTTYPE = pACTTYPE
    ACTVTY_QTY = pACTVTY_QTY
    SEG_TEXT = pSEG_TEXT
    REC_WBS_EL = pREC_WBS_EL
    REC_NETWRK = pREC_NETWRK
    RECOPERATN = pRECOPERATN
    REC_ORDER = pREC_ORDER
    REC_CCTR = pREC_CCTR
    PRICE = pPRICE
    PRICE_VAR = pPRICE_VAR
    PRICE_FIX = pPRICE_FIX
    PRICE_UNIT = pPRICE_UNIT
    CURR = pCURR
End Function
