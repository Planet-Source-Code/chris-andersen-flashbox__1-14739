VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2010
   ClientLeft      =   3405
   ClientTop       =   2550
   ClientWidth     =   3765
   LinkTopic       =   "Form1"
   ScaleHeight     =   2010
   ScaleWidth      =   3765
   Begin VB.CommandButton Command1 
      Caption         =   "My flash message box"
      Height          =   555
      Left            =   240
      TabIndex        =   0
      Top             =   180
      Width           =   2115
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim FBox As New clsFlashBox
Dim Returned As String

Private Sub Command1_Click()

'Possible Box Types are OK Only, Yes/No, Ok/Cancel, and Input Box
'add a third arguement to the below function to
'see the Enumerated list of possible Box types
'The fourth option is only used when you
'select fbINPUT as the Box Type and is the
'default value you want printed in the Input
'Box
Returned = FBox.FlashMsg("Done", "Action Completion")

'Ok buttons return fbrOK
'Cancel button returne fbrCANCEL
'Yes returns fbrYES
'No returns fbrNO
'If using Input Box, Ok returns typed value, Cancel returns fbrCANCEL

Select Case Returned
    Case "fbrOK"
        MsgBox "You clicked Ok!"
    Case "fbrCANCEL"
        MsgBox "You clicked Cancel!"
    Case "fbrYES"
        MsgBox "You clicked Yes!"
    Case "fbrNO"
        MsgBox "You clicked No!"
    Case Else
        MsgBox Returned
End Select


End Sub

Private Sub Form_Unload(Cancel As Integer)

Set FBox = Nothing

End Sub
