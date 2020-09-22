VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "swflash.ocx"
Begin VB.Form frmFlashBox 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form2"
   ClientHeight    =   2595
   ClientLeft      =   4305
   ClientTop       =   3525
   ClientWidth     =   5625
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2595
   ScaleWidth      =   5625
   ShowInTaskbar   =   0   'False
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash flsMsgbox 
      Height          =   2595
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5655
      _cx             =   22882039
      _cy             =   22876641
      Movie           =   ""
      Src             =   ""
      WMode           =   "Window"
      Play            =   -1  'True
      Loop            =   -1  'True
      Quality         =   "High"
      SAlign          =   ""
      Menu            =   -1  'True
      Base            =   ""
      Scale           =   "ShowAll"
      DeviceFont      =   0   'False
      EmbedMovie      =   0   'False
      BGColor         =   ""
      SWRemote        =   ""
   End
End
Attribute VB_Name = "frmFlashBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub flsMsgbox_FSCommand(ByVal command As String, ByVal args As String)

If command = "btn1" Or command = "btn2" Then
    If args = "fbrOKINPUT" Then
        ReturnCommand = flsMsgbox.GetVariable("txtInput")
    Else:
        ReturnCommand = args
    End If
    
   
    Me.Hide
    Set frmFlashBox = Nothing
End If

End Sub

