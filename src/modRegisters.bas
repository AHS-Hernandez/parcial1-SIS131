Attribute VB_Name = "modRegisters"
Option Explicit

' Registros de la CPU, cada uno un Byte (0-255).
' En la ventana Inmediato: PruebaModRegisters
' Casos: ResetRegisters, asignacion valida, asignacion fuera de 8 bits.

Private mPC As Byte
Private mIR As Byte
Private mMAR As Byte
Private mMDR As Byte
Private mAX As Byte
Private mBX As Byte

Private Const ERR_VALOR As Long = vbObjectError + 5201

Public Property Get PC() As Byte
    PC = mPC
End Property

Public Property Get IR() As Byte
    IR = mIR
End Property

Public Property Get MAR() As Byte
    MAR = mMAR
End Property

Public Property Get MDR() As Byte
    MDR = mMDR
End Property

Public Property Get AX() As Byte
    AX = mAX
End Property

Public Property Get BX() As Byte
    BX = mBX
End Property

Public Sub SetPC(ByVal value As Long)
    mPC = Validar(value, "PC")
End Sub

Public Sub SetIR(ByVal value As Long)
    mIR = Validar(value, "IR")
End Sub

Public Sub SetMAR(ByVal value As Long)
    mMAR = Validar(value, "MAR")
End Sub

Public Sub SetMDR(ByVal value As Long)
    mMDR = Validar(value, "MDR")
End Sub

Public Sub SetAX(ByVal value As Long)
    mAX = Validar(value, "AX")
End Sub

Public Sub SetBX(ByVal value As Long)
    mBX = Validar(value, "BX")
End Sub

Public Sub IncrementarPC()
    If mPC = 255 Then
        mPC = 0
    Else
        mPC = mPC + 1
    End If
End Sub

Public Sub ResetRegisters()
    mPC = 0
    mIR = 0
    mMAR = 0
    mMDR = 0
    mAX = 0
    mBX = 0
End Sub

Public Sub PruebaModRegisters()
    On Error GoTo Fallo
    ResetRegisters
    Debug.Print "reset", "esperado 0", "PC=" & PC, "IR=" & IR, "MAR=" & MAR, "MDR=" & MDR, "AX=" & AX, "BX=" & BX

    SetPC 10
    SetAX 255
    Debug.Print "valido", "esperado PC 10 AX 255", "PC=" & PC, "AX=" & AX

    On Error Resume Next
    Err.Clear
    SetPC 256
    Debug.Print "fuera de rango 256", "esperado error y PC=10", Err.Description, "PC=" & PC

    Err.Clear
    SetAX -1
    Debug.Print "negativo", "esperado error y AX=255", Err.Description, "AX=" & AX
    On Error GoTo Fallo

    If PC <> 10 Or AX <> 255 Then
        Debug.Print "FALLO: un registro cambio con un valor imposible"
    Else
        Debug.Print "los registros no cambiaron ante valores imposibles"
    End If

    SetPC 255
    IncrementarPC
    Debug.Print "PC despues de FFh", "esperado 0", "PC=" & PC
    Exit Sub
Fallo:
    Debug.Print "FALLO inesperado", Err.Number, Err.Description
End Sub

Private Function Validar(ByVal value As Long, ByVal registro As String) As Byte
    If value < 0 Or value > 255 Then
        Err.Raise ERR_VALOR, "modRegisters", registro & " fuera de 8 bits: " & value
    End If
    Validar = CByte(value)
End Function
