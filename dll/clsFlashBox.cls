VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFlashBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Enum FlashBoxType
    fbOK_CANCEL = 1
    fbYES_NO = 6
    fbOKONLY = 11
    fbINPUT = 18
End Enum

Public Function FlashMsg(FlashMessage As String, Optional FormCaption As String = "FlashBox", Optional BoxType As FlashBoxType = fbOKONLY, Optional InputDefault As String) As String

Load frmFlashBox

With frmFlashBox
    .Caption = FormCaption
    .flsMsgbox.Movie = App.Path & "\flashmsg.swf"
    .flsMsgbox.GotoFrame BoxType
    If BoxType = fbINPUT Then
        .flsMsgbox.SetVariable "txtInput", InputDefault
    End If
    .flsMsgbox.SetVariable "txtMsg", FlashMessage
    .Show vbModal
End With

FlashMsg = ReturnCommand

End Function
