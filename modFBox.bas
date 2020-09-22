Attribute VB_Name = "modFBox"
Public ReturnCommand As String
Public ReturnInput As String

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
