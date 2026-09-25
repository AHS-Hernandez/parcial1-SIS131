Attribute VB_Name = "modFlags"
Option Explicit

' Banderas ZF, CF y SF. El byte se lee en complemento a 2: SF es el bit 7.
' En la ventana Inmediato: PruebaModFlags
' Casos: cero, acarreo, negativo (FFh = -1), 80h = -128, valor imposible.

Private mZF As Byte
Private mCF As Byte
Private mSF As Byte

Private Const ERR_VALOR As Long = vbObjectError + 5301

Public Property Get ZF() As Byte
    ZF = mZF
End Property

Public Property Get CF() As Byte
    CF = mCF
End Property

Public Property Get SF() As Byte
    SF = mSF
End Property

' result: byte 0-255. carry: 0 o 1, lo calcula la ALU (acarreo o prestamo).
Public Sub UpdateFlags(ByVal result As Long, ByVal carry As Long)
    Dim nuevoZF As Byte
    Dim nuevoCF As Byte
    Dim nuevoSF As Byte

    If result < 0 Or result > 255 Then
        Err.Raise ERR_VALOR, "modFlags", "Resultado fuera de 8 bits: " & result
    End If
    If carry <> 0 And carry <> 1 Then
        Err.Raise ERR_VALOR, "modFlags", "Acarreo invalido: " & carry
    End If

    If result = 0 Then
        nuevoZF = 1
    Else
        nuevoZF = 0
    End If

    nuevoCF = CByte(carry)

    If (result And &H80) <> 0 Then
        nuevoSF = 1
    Else
        nuevoSF = 0
    End If

    mZF = nuevoZF
    mCF = nuevoCF
    mSF = nuevoSF
End Sub

Public Sub ResetFlags()
    mZF = 0
    mCF = 0
    mSF = 0
End Sub

' El mismo byte, leido con signo. Sirve para explicar SF en la defensa.
Public Function ValorConSigno(ByVal value As Long) As Integer
    If value < 0 Or value > 255 Then
        Err.Raise ERR_VALOR, "modFlags", "Valor fuera de 8 bits: " & value
    End If
    If value >= 128 Then
        ValorConSigno = value - 256
    Else
        ValorConSigno = value
    End If
End Function

Public Sub PruebaModFlags()
    On Error GoTo Fallo
    ResetFlags
    Debug.Print "reset", "esperado 0", "ZF=" & ZF, "CF=" & CF, "SF=" & SF

    UpdateFlags 0, 0
    Debug.Print "cero", "esperado ZF=1 CF=0 SF=0", "ZF=" & ZF, "CF=" & CF, "SF=" & SF

    UpdateFlags 0, 1
    Debug.Print "cero con acarreo", "esperado ZF=1 CF=1 SF=0", "ZF=" & ZF, "CF=" & CF, "SF=" & SF

    UpdateFlags 255, 1
    Debug.Print "FFh (-1)", "esperado ZF=0 CF=1 SF=1", "ZF=" & ZF, "CF=" & CF, "SF=" & SF, "signo=" & ValorConSigno(255)

    UpdateFlags 128, 0
    Debug.Print "80h (-128)", "esperado SF=1 ZF=0 CF=0", "ZF=" & ZF, "CF=" & CF, "SF=" & SF, "signo=" & ValorConSigno(128)

    UpdateFlags 127, 0
    Debug.Print "7Fh", "esperado SF=0", "SF=" & SF, "signo=" & ValorConSigno(127)

    On Error Resume Next
    Err.Clear
    UpdateFlags 256, 0
    Debug.Print "resultado imposible", "esperado error y SF=0", Err.Description, "SF=" & SF
    On Error GoTo Fallo

    If SF <> 0 Or ZF <> 0 Or CF <> 0 Then
        Debug.Print "FALLO: las banderas cambiaron con un resultado imposible"
    Else
        Debug.Print "las banderas no cambiaron ante un resultado imposible"
    End If
    Exit Sub
Fallo:
    Debug.Print "FALLO inesperado", Err.Number, Err.Description
End Sub
