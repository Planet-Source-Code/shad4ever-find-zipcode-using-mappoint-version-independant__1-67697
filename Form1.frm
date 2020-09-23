VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Test"
      Height          =   375
      Left            =   600
      TabIndex        =   0
      Top             =   480
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ClsMap As clsMapPoint

Private Sub Command1_Click()
  ClsMap.MapZipCodes "J4B 7L9", "Boucherville - Industrial", "Found Code"
End Sub

Private Sub Form_Load()
On Error GoTo ErrChk:
  Set ClsMap = New clsMapPoint
  Exit Sub
ErrChk:
  MsgBox "MapPoint cannot initialise, Wrong COM or Not Installed"
  Command1.Enabled = False
End Sub
