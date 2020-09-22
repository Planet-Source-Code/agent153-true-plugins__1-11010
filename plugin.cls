VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "plugin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Coded by Agent153 on 8-26-2000
Option Explicit

Public Function PIName() As String
    PIName = "Plugin Tester"
End Function

Public Function Interface() As Object
    Set Interface = frm_ui.fr_plg
End Function

Property Set HostAppObject(newref As Object)
    Set frm_ui.HostObject = newref
End Property

