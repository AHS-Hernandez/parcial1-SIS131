Attribute VB_Name = "modExecution"
Option Explicit

' Execute y Store de la instruccion actual. MOV y LOAD en este parcial (issue 15).
' LOAD pasa siempre por MAR y MDR. MOV no toca las banderas.
' En la ventana Inmediato: PruebaExecuteMovLoad

Private mInfo As tInstruccion
Private mOperando As Byte
Private mTemporal As Byte
Private mPreparada As Boolean

Public Property Get Temporal() As Byte
    Temporal = mTemporal
End Property

Public Property Get InstruccionLista() As Boolean
    InstruccionLista = mPreparada
End Property

Public Sub Preparar(ByRef info As tInstruccion, ByVal operando As Byte)
    mInfo = info
    mOperando = operando And &HFF
    mTemporal = 0
    mPreparada = info.Encontrada
End Sub

' Micro-operacion Execute: temporal o MAR/MDR segun la familia.
Public Sub PasoExecute()
    If Not mPreparada Then Exit Sub

    Select Case mInfo.Opcode
        Case OP_MOV_AX_IMM, OP_MOV_BX_IMM
            mTemporal = mOperando
        Case OP_MOV_AX_BX
            mTemporal = BX
        Case OP_MOV_BX_AX
            mTemporal = AX
        Case OP_LOAD_AX, OP_LOAD_BX
            SetMAR mOperando
            SetMDR ReadMem(MAR)
        Case Else
            ' Otras familias: issues siguientes.
    End Select
End Sub

' Micro-operacion Store: escribe el registro destino.
Public Sub PasoStore()
    If Not mPreparada Then Exit Sub

    Select Case mInfo.Opcode
        Case OP_MOV_AX_IMM, OP_MOV_AX_BX
            SetAX mTemporal
        Case OP_MOV_BX_IMM, OP_MOV_BX_AX
            SetBX mTemporal
        Case OP_LOAD_AX
            SetAX MDR
        Case OP_LOAD_BX
            SetBX MDR
        Case Else
            ' Otras familias: issues siguientes.
    End Select
End Sub

Public Sub PruebaExecuteMovLoad()
    Dim fallos As Long
    Dim z0 As Byte
    Dim c0 As Byte
    Dim s0 As Byte

    On Error GoTo Fallo
    fallos = 0

    ' MOV AX, imm
    ClearMem
    ResetRegisters
    UpdateFlags 0, 1
    z0 = modFlags.ZF: c0 = modFlags.CF: s0 = modFlags.SF
    Preparar DecodeByte(OP_MOV_AX_IMM), &H2A
    PasoExecute
    If Temporal <> &H2A Then
        Debug.Print "FALLO MOV AX,imm Execute", "temporal=" & Temporal
        fallos = fallos + 1
    End If
    PasoStore
    If AX <> &H2A Or Not BanderasIguales(z0, c0, s0) Then
        Debug.Print "FALLO MOV AX,imm Store", "AX=" & AX, "flags"
        fallos = fallos + 1
    Else
        Debug.Print "MOV AX, 2Ah", "ok", "AX=" & AX
    End If

    ' MOV BX, imm
    ClearMem
    ResetRegisters
    UpdateFlags &HFF, 0
    z0 = modFlags.ZF: c0 = modFlags.CF: s0 = modFlags.SF
    Preparar DecodeByte(OP_MOV_BX_IMM), &H7
    PasoExecute
    PasoStore
    If BX <> 7 Or Not BanderasIguales(z0, c0, s0) Then
        Debug.Print "FALLO MOV BX,imm", "BX=" & BX
        fallos = fallos + 1
    Else
        Debug.Print "MOV BX, 07h", "ok", "BX=" & BX
    End If

    ' MOV AX, BX
    ClearMem
    ResetRegisters
    SetBX &H55
    SetAX 0
    UpdateFlags 1, 1
    z0 = modFlags.ZF: c0 = modFlags.CF: s0 = modFlags.SF
    Preparar DecodeByte(OP_MOV_AX_BX), 0
    PasoExecute
    PasoStore
    If AX <> &H55 Or BX <> &H55 Or Not BanderasIguales(z0, c0, s0) Then
        Debug.Print "FALLO MOV AX,BX", "AX=" & AX, "BX=" & BX
        fallos = fallos + 1
    Else
        Debug.Print "MOV AX, BX", "ok", "AX=" & AX
    End If

    ' MOV BX, AX
    ClearMem
    ResetRegisters
    SetAX &H11
    SetBX 0
    UpdateFlags 0, 0
    z0 = modFlags.ZF: c0 = modFlags.CF: s0 = modFlags.SF
    Preparar DecodeByte(OP_MOV_BX_AX), 0
    PasoExecute
    PasoStore
    If BX <> &H11 Or AX <> &H11 Or Not BanderasIguales(z0, c0, s0) Then
        Debug.Print "FALLO MOV BX,AX", "AX=" & AX, "BX=" & BX
        fallos = fallos + 1
    Else
        Debug.Print "MOV BX, AX", "ok", "BX=" & BX
    End If

    ' LOAD AX, [dir] ? debe pasar por MAR y MDR
    ClearMem
    ResetRegisters
    WriteMem &H80, &H3C
    UpdateFlags &H80, 1
    z0 = modFlags.ZF: c0 = modFlags.CF: s0 = modFlags.SF
    Preparar DecodeByte(OP_LOAD_AX), &H80
    PasoExecute
    If MAR <> &H80 Or MDR <> &H3C Then
        Debug.Print "FALLO LOAD AX Execute", "MAR=" & MAR, "MDR=" & MDR
        fallos = fallos + 1
    End If
    PasoStore
    If AX <> &H3C Or Not BanderasIguales(z0, c0, s0) Then
        Debug.Print "FALLO LOAD AX Store", "AX=" & AX
        fallos = fallos + 1
    Else
        Debug.Print "LOAD AX, [80h]", "ok", "AX=" & AX, "MAR=" & MAR, "MDR=" & MDR
    End If

    ' LOAD BX, [dir]
    ClearMem
    ResetRegisters
    WriteMem &HFF, &H9A
    UpdateFlags 5, 0
    z0 = modFlags.ZF: c0 = modFlags.CF: s0 = modFlags.SF
    Preparar DecodeByte(OP_LOAD_BX), &HFF
    PasoExecute
    If MAR <> &HFF Or MDR <> &H9A Then
        Debug.Print "FALLO LOAD BX Execute", "MAR=" & MAR, "MDR=" & MDR
        fallos = fallos + 1
    End If
    PasoStore
    If BX <> &H9A Or Not BanderasIguales(z0, c0, s0) Then
        Debug.Print "FALLO LOAD BX Store", "BX=" & BX
        fallos = fallos + 1
    Else
        Debug.Print "LOAD BX, [FFh]", "ok", "BX=" & BX, "MAR=" & MAR, "MDR=" & MDR
    End If

    ' Traza corta: Decode + Execute/Store de MOV AX,imm via ControlUnit
    ClearMem
    ResetRegisters
    WriteMem 0, OP_MOV_AX_IMM
    WriteMem 1, &H5
    SetPC 0
    IniciarFetch
    PasoFetch
    PasoFetch
    PasoFetch
    PasoFetch
    PasoDecode
    TraerOperando
    Preparar InstruccionActual, OperandoActual
    PasoExecute
    PasoStore
    If AX <> 5 Or InstruccionActual.Sintaxis <> "MOV AX, imm" Then
        Debug.Print "FALLO traza Fetch+Decode+MOV", "AX=" & AX
        fallos = fallos + 1
    Else
        Debug.Print "traza MOV AX,05h", "ok", "AX=" & AX
    End If

    If fallos = 0 Then
        Debug.Print "Execute cubre MOV y LOAD sin tocar banderas"
    Else
        Debug.Print "FALLOS", fallos
    End If
    Exit Sub
Fallo:
    Debug.Print "FALLO inesperado", Err.Number, Err.Description
End Sub

Private Function BanderasIguales(ByVal z0 As Byte, ByVal c0 As Byte, ByVal s0 As Byte) As Boolean
    BanderasIguales = (modFlags.ZF = z0 And modFlags.CF = c0 And modFlags.SF = s0)
End Function
