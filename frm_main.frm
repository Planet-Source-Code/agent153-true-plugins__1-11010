VERSION 5.00
Begin VB.Form frm_main 
   Caption         =   "Plugin Host"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Load Plugin"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "frm_main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Coded by Agent153 on 8-26-2000
Option Explicit
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long

Private Sub Command1_Click()
    Dim plugin As Object
    Set plugin = CreateObject("SvrApp.plugin")
    
    With plugin
        Set .HostAppObject = New cls_app
        Call SetParent(.interface.hWnd, Me.hWnd)
        MsgBox "Plugin name: """ & .PIName & """."
    End With
End Sub
