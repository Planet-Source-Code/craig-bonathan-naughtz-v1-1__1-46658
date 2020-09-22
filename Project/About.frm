VERSION 5.00
Begin VB.Form About 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "About Naughtz"
   ClientHeight    =   2790
   ClientLeft      =   2340
   ClientTop       =   1890
   ClientWidth     =   4515
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "About.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "About.frx":08CA
   ScaleHeight     =   1925.708
   ScaleMode       =   0  'User
   ScaleWidth      =   4239.818
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00000000&
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   1680
      TabIndex        =   0
      Top             =   1800
      Width           =   1260
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "E-mail: cb3software@fire-bug.co.uk"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   2520
      Width           =   4455
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Web: www.cb3software.fire-bug.co.uk"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   2280
      Width           =   4455
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Naughtz"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1680
      TabIndex        =   2
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Version 1.1"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1680
      TabIndex        =   1
      Top             =   1440
      Width           =   1215
   End
End
Attribute VB_Name = "About"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
Unload About
PlayScreen.Show
PlayScreen.Enabled = True
End Sub
