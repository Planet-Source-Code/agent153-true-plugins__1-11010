VERSION 5.00
Begin VB.Form frm_ui 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fr_plg 
      Caption         =   "Plugin Interface"
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3135
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   840
         Width           =   1695
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Send info to Host"
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   1695
      End
   End
End
Attribute VB_Name = "frm_ui"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Coded by Agent153 on 8-26-2000
Option Explicit

Public HostObject As Object

Private Sub Command1_Click()
    HostObject.Testing Text1.Text
End Sub
