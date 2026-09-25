Attribute VB_Name = "modALU"
Option Explicit

' Operaciones de la ALU sobre bytes. No escribe AX, BX ni la RAM.
' ADD, SUB, INC y DEC devuelven el resultado y actualizan ZF, CF y SF.
' CMP hace la resta solo para las banderas: el que llama no guarda el retorno.
' En la ventana Inmediato: PruebaModALU

Private Const ERR_VALOR As Long = vbObjectError + 5401

Public Function AluAdd(ByVal a As Long, ByVal b As Long) As Byte
    Dim suma As Long
    AsegurarByte a
    AsegurarByte b
    suma = a + b
    If suma > 255 Then
        AluAdd = CByte(suma - 256)
        UpdateFlags suma - 256, 1
    Else
        AluAdd = CByte(suma)
        UpdateFlags suma, 0
    End If
End Function

Public Function AluSub(ByVal a As Long, ByVal b As Long) As Byte
    AsegurarByte a
    AsegurarByte b
    If a >= b Then
        AluSub = CByte(a - b)
        UpdateFlags a - b, 0
    Else
        AluSub = CByte(a - b + 256)
        UpdateFlags a - b + 256, 1
    End If
End Function

Public Function AluInc(ByVal a As Long) As Byte
    AluInc = AluAdd(a, 1)
End Function

Public Function AluDec(ByVal a As Long) As Byte
    AluDec = AluSub(a, 1)
End Function

' Misma resta que AluSub. El caller descarta el valor de vuelta.
Public Function AluCmp(ByVal a As Long, ByVal b As Long) As Byte
    AluCmp = AluSub(a, b)
End Function

Public Sub PruebaModALU()
    Dim r As Byte
    On Error GoTo Fallo

    r = AluAdd(2, 3)
    Debug.Print "suma normal", "esperado 5 ZF=0 CF=0 SF=0", "r=" & r, "ZF=" & modFlags.ZF, "CF=" & modFlags.CF, "SF=" & modFlags.SF

    r = AluSub(5, 3)
    Debug.Print "resta normal", "esperado 2 ZF=0 CF=0 SF=0", "r=" & r, "ZF=" & modFlags.ZF, "CF=" & modFlags.CF, "SF=" & modFlags.SF

    r = AluSub(5, 5)
    Debug.Print "cero", "esperado 0 ZF=1 CF=0 SF=0", "r=" & r, "ZF=" & modFlags.ZF, "CF=" & modFlags.CF, "SF=" & modFlags.SF

    r = AluAdd(&HFF, 1)
    Debug.Print "acarreo FFh+1", "esperado 0 ZF=1 CF=1 SF=0", "r=" & r, "ZF=" & modFlags.ZF, "CF=" & modFlags.CF, "SF=" & modFlags.SF

    r = AluSub(0, 1)
    Debug.Print "negativo 00h-1", "esperado 255 ZF=0 CF=1 SF=1", "r=" & r, "ZF=" & modFlags.ZF, "CF=" & modFlags.CF, "SF=" & modFlags.SF

    r = AluAdd(0, 0)
    Debug.Print "00h+00h", "esperado 0 ZF=1 CF=0 SF=0", "r=" & r, "ZF=" & modFlags.ZF, "CF=" & modFlags.CF, "SF=" & modFlags.SF

    r = AluAdd(&HFF, 0)
    Debug.Print "FFh+00h", "esperado 255 ZF=0 CF=0 SF=1", "r=" & r, "ZF=" & modFlags.ZF, "CF=" & modFlags.CF, "SF=" & modFlags.SF

    r = AluSub(0, &HFF)
    Debug.Print "00h-FFh", "esperado 1 ZF=0 CF=1 SF=0", "r=" & r, "ZF=" & modFlags.ZF, "CF=" & modFlags.CF, "SF=" & modFlags.SF

    r = AluInc(&HFF)
    Debug.Print "INC FFh", "esperado 0 ZF=1 CF=1 SF=0", "r=" & r, "ZF=" & modFlags.ZF, "CF=" & modFlags.CF, "SF=" & modFlags.SF

    r = AluDec(0)
    Debug.Print "DEC 00h", "esperado 255 ZF=0 CF=1 SF=1", "r=" & r, "ZF=" & modFlags.ZF, "CF=" & modFlags.CF, "SF=" & modFlags.SF

    r = AluCmp(4, 4)
    Debug.Print "CMP iguales", "esperado retorno 0 ZF=1; no se guarda en un registro", "r=" & r, "ZF=" & modFlags.ZF, "CF=" & modFlags.CF, "SF=" & modFlags.SF

    r = AluCmp(1, 4)
    Debug.Print "CMP distintos", "esperado ZF=0 CF=1", "r=" & r, "ZF=" & modFlags.ZF, "CF=" & modFlags.CF, "SF=" & modFlags.SF

    On Error Resume Next
    Err.Clear
    r = AluAdd(1, 256)
    Debug.Print "operando imposible", "esperado error", Err.Description
    Err.Clear
    On Error GoTo Fallo
    Exit Sub
Fallo:
    Debug.Print "FALLO inesperado", Err.Number, Err.Description
End Sub

Private Sub AsegurarByte(ByVal value As Long)
    If value < 0 Or value > 255 Then
        Err.Raise ERR_VALOR, "modALU", "Operando fuera de 8 bits: " & value
    End If
End Sub
