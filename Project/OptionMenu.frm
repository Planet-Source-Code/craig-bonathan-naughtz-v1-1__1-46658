VERSION 5.00
Begin VB.Form OptionMenu 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Naughtz - Options"
   ClientHeight    =   855
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   2055
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   855
   ScaleWidth      =   2055
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   1815
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Allow UFOs"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "OptionMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
AllowUFO = False
If Check1.Value = 1 Then AllowUFO = True
PlayScreen.Enabled = True
Unload Me
End Sub

Private Sub Form_Load()
If AllowUFO = True Then Check1.Value = 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
PlayScreen.Enabled = True
End Sub
