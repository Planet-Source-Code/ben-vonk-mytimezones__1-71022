VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMouseWheel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Public Event
Public Event Wheel(ScrollLines As Integer)

Public Sub Hook(ByVal hWnd As Long)

   Call SubclassMouseWheel(Me, hWnd)

End Sub

Public Sub Unhook()

   Call SubclassMouseWheel

End Sub

Friend Sub WheelUsed(ByVal WheelUp As Boolean)

   RaiseEvent Wheel(-1 + (-2 * WheelUp))

End Sub
