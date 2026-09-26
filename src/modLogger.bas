Attribute VB_Name = "modLogger"
Option Explicit

' Stub del log (issue 26 lo completa). ClearLog lo llama DoReset.

Private mFilas As Long

Public Sub ClearLog()
    mFilas = 0
End Sub

Public Property Get FilasLog() As Long
    FilasLog = mFilas
End Property

' Reservado para issue 26: una fila por micro-operacion.
Public Sub AppendLog(ByVal texto As String)
    mFilas = mFilas + 1
End Sub
