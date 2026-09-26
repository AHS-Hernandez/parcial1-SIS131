Attribute VB_Name = "modExecution"
Option Explicit

' Execute y Store de la instruccion actual.
' MOV/LOAD/saltos/HLT no tocan banderas. ADD/SUB/INC/DEC/CMP pasan por modALU.
' CMP actualiza flags y no escribe el registro. HLT pone HALTED.
' En la ventana Inmediato: PruebaExecuteMovLoad, PruebaExecuteAritmetica, PruebaExecuteSaltos

Private mInfo As tInstruccion
Private mOperando As Byte
Private mTemporal As Byte
Private mPreparada As Boolean
Private mEscribeRegistro As Boolean

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
    mEscribeRegistro = True
End Sub

' Micro-operacion Execute: temporal, MAR/MDR o ALU segun la familia.
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

        Case OP_ADD_AX_IMM
            mTemporal = AluAdd(AX, mOperando)
        Case OP_ADD_AX_BX
            mTemporal = AluAdd(AX, BX)
        Case OP_ADD_BX_IMM
            mTemporal = AluAdd(BX, mOperando)
        Case OP_ADD_BX_AX
            mTemporal = AluAdd(BX, AX)

        Case OP_SUB_AX_IMM
            mTemporal = AluSub(AX, mOperando)
        Case OP_SUB_AX_BX
            mTemporal = AluSub(AX, BX)
        Case OP_SUB_BX_IMM
            mTemporal = AluSub(BX, mOperando)
        Case OP_SUB_BX_AX
            mTemporal = AluSub(BX, AX)

        Case OP_INC_AX
            mTemporal = AluInc(AX)
        Case OP_INC_BX
            mTemporal = AluInc(BX)
        Case OP_DEC_AX
            mTemporal = AluDec(AX)
        Case OP_DEC_BX
            mTemporal = AluDec(BX)

        Case OP_CMP_AX_IMM
            mTemporal = AluCmp(AX, mOperando)
            mEscribeRegistro = False
        Case OP_CMP_AX_BX
            mTemporal = AluCmp(AX, BX)
            mEscribeRegistro = False
        Case OP_CMP_BX_IMM
            mTemporal = AluCmp(BX, mOperando)
            mEscribeRegistro = False
        Case OP_CMP_BX_AX
            mTemporal = AluCmp(BX, AX)
            mEscribeRegistro = False

        Case OP_JMP
            SetPC mOperando
            mEscribeRegistro = False
        Case OP_JZ
            If modFlags.ZF = 1 Then SetPC mOperando
            mEscribeRegistro = False
        Case OP_JNZ
            If modFlags.ZF = 0 Then SetPC mOperando
            mEscribeRegistro = False
        Case OP_HLT
            DetenerCPU
            mEscribeRegistro = False

        Case Else
            ' STORE y logica: issues siguientes.
    End Select
End Sub

' Micro-operacion Store: escribe el registro destino (salvo CMP).
Public Sub PasoStore()
    If Not mPreparada Then Exit Sub
    If Not mEscribeRegistro Then Exit Sub

    Select Case mInfo.Opcode
        Case OP_MOV_AX_IMM, OP_MOV_AX_BX
            SetAX mTemporal
        Case OP_MOV_BX_IMM, OP_MOV_BX_AX
            SetBX mTemporal
        Case OP_LOAD_AX
            SetAX MDR
        Case OP_LOAD_BX
            SetBX MDR

        Case OP_ADD_AX_IMM, OP_ADD_AX_BX, OP_SUB_AX_IMM, OP_SUB_AX_BX, OP_INC_AX, OP_DEC_AX
            SetAX mTemporal
        Case OP_ADD_BX_IMM, OP_ADD_BX_AX, OP_SUB_BX_IMM, OP_SUB_BX_AX, OP_INC_BX, OP_DEC_BX
            SetBX mTemporal

        Case Else
            ' STORE y logica: issues siguientes.
    End Select
End Sub

Public Sub PruebaExecuteMovLoad()
    Dim fallos As Long
    Dim z0 As Byte
    Dim c0 As Byte
    Dim s0 As Byte

    On Error GoTo Fallo
    fallos = 0

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

Public Sub PruebaExecuteAritmetica()
    Dim fallos As Long
    Dim ax0 As Byte
    Dim bx0 As Byte

    On Error GoTo Fallo
    fallos = 0

    ' ADD AX, imm: 02h+03h = 05h, flags limpios
    ClearMem
    ResetRegisters
    SetAX 2
    Preparar DecodeByte(OP_ADD_AX_IMM), 3
    PasoExecute
    PasoStore
    If AX <> 5 Or modFlags.ZF <> 0 Or modFlags.CF <> 0 Or modFlags.SF <> 0 Then
        Debug.Print "FALLO ADD AX,imm", "AX=" & AX, "ZF=" & modFlags.ZF
        fallos = fallos + 1
    Else
        Debug.Print "ADD AX, 03h", "ok", "AX=" & AX
    End If

    ' ADD AX, BX con acarreo: FFh+01h
    ClearMem
    ResetRegisters
    SetAX &HFF
    SetBX 1
    Preparar DecodeByte(OP_ADD_AX_BX), 0
    PasoExecute
    PasoStore
    If AX <> 0 Or BX <> 1 Or modFlags.ZF <> 1 Or modFlags.CF <> 1 Then
        Debug.Print "FALLO ADD AX,BX acarreo", "AX=" & AX, "CF=" & modFlags.CF
        fallos = fallos + 1
    Else
        Debug.Print "ADD AX, BX (FFh+1)", "ok", "AX=0 CF=1 ZF=1"
    End If

    ' ADD BX, imm
    ClearMem
    ResetRegisters
    SetBX 10
    Preparar DecodeByte(OP_ADD_BX_IMM), 5
    PasoExecute
    PasoStore
    If BX <> 15 Then
        Debug.Print "FALLO ADD BX,imm", "BX=" & BX
        fallos = fallos + 1
    Else
        Debug.Print "ADD BX, 05h", "ok", "BX=" & BX
    End If

    ' ADD BX, AX
    ClearMem
    ResetRegisters
    SetAX 3
    SetBX 4
    Preparar DecodeByte(OP_ADD_BX_AX), 0
    PasoExecute
    PasoStore
    If BX <> 7 Or AX <> 3 Then
        Debug.Print "FALLO ADD BX,AX", "BX=" & BX
        fallos = fallos + 1
    Else
        Debug.Print "ADD BX, AX", "ok", "BX=" & BX
    End If

    ' SUB AX, imm
    ClearMem
    ResetRegisters
    SetAX 5
    Preparar DecodeByte(OP_SUB_AX_IMM), 3
    PasoExecute
    PasoStore
    If AX <> 2 Or modFlags.CF <> 0 Then
        Debug.Print "FALLO SUB AX,imm", "AX=" & AX
        fallos = fallos + 1
    Else
        Debug.Print "SUB AX, 03h", "ok", "AX=" & AX
    End If

    ' SUB BX, AX con prestamo: 00h-01h
    ClearMem
    ResetRegisters
    SetBX 0
    SetAX 1
    Preparar DecodeByte(OP_SUB_BX_AX), 0
    PasoExecute
    PasoStore
    If BX <> &HFF Or AX <> 1 Or modFlags.CF <> 1 Or modFlags.SF <> 1 Then
        Debug.Print "FALLO SUB BX,AX prestamo", "BX=" & BX, "CF=" & modFlags.CF
        fallos = fallos + 1
    Else
        Debug.Print "SUB BX, AX (0-1)", "ok", "BX=FFh CF=1 SF=1"
    End If

    ' INC AX / DEC BX
    ClearMem
    ResetRegisters
    SetAX &HFF
    Preparar DecodeByte(OP_INC_AX), 0
    PasoExecute
    PasoStore
    If AX <> 0 Or modFlags.ZF <> 1 Or modFlags.CF <> 1 Then
        Debug.Print "FALLO INC AX", "AX=" & AX
        fallos = fallos + 1
    Else
        Debug.Print "INC AX (FFh)", "ok", "AX=0 CF=1"
    End If

    ClearMem
    ResetRegisters
    SetBX 1
    Preparar DecodeByte(OP_DEC_BX), 0
    PasoExecute
    PasoStore
    If BX <> 0 Or modFlags.ZF <> 1 Or modFlags.CF <> 0 Then
        Debug.Print "FALLO DEC BX", "BX=" & BX
        fallos = fallos + 1
    Else
        Debug.Print "DEC BX (01h)", "ok", "BX=0 ZF=1"
    End If

    ' CMP AX, BX: iguales -> ZF=1, AX y BX intactos
    ClearMem
    ResetRegisters
    SetAX 5
    SetBX 5
    ax0 = AX
    bx0 = BX
    Preparar DecodeByte(OP_CMP_AX_BX), 0
    PasoExecute
    PasoStore
    If AX <> ax0 Or BX <> bx0 Or modFlags.ZF <> 1 Or modFlags.CF <> 0 Then
        Debug.Print "FALLO CMP AX,BX iguales", "AX=" & AX, "ZF=" & modFlags.ZF
        fallos = fallos + 1
    Else
        Debug.Print "CMP AX, BX (iguales)", "ok", "ZF=1 registros intactos"
    End If

    ' CMP AX, imm: AX < imm -> CF=1, AX intacto
    ClearMem
    ResetRegisters
    SetAX 3
    ax0 = AX
    Preparar DecodeByte(OP_CMP_AX_IMM), 5
    PasoExecute
    PasoStore
    If AX <> ax0 Or modFlags.CF <> 1 Or modFlags.ZF <> 0 Then
        Debug.Print "FALLO CMP AX,imm", "AX=" & AX, "CF=" & modFlags.CF
        fallos = fallos + 1
    Else
        Debug.Print "CMP AX, 05h (3<5)", "ok", "CF=1 AX intacto"
    End If

    ' CMP BX, AX
    ClearMem
    ResetRegisters
    SetBX 8
    SetAX 8
    bx0 = BX
    Preparar DecodeByte(OP_CMP_BX_AX), 0
    PasoExecute
    PasoStore
    If BX <> bx0 Or AX <> 8 Or modFlags.ZF <> 1 Then
        Debug.Print "FALLO CMP BX,AX", "BX=" & BX
        fallos = fallos + 1
    Else
        Debug.Print "CMP BX, AX", "ok", "ZF=1 BX intacto"
    End If

    If fallos = 0 Then
        Debug.Print "Execute conecta ADD/SUB/INC/DEC/CMP con la ALU"
    Else
        Debug.Print "FALLOS", fallos
    End If
    Exit Sub
Fallo:
    Debug.Print "FALLO inesperado", Err.Number, Err.Description
End Sub

Public Sub PruebaExecuteSaltos()
    Dim fallos As Long
    Dim z0 As Byte
    Dim c0 As Byte
    Dim s0 As Byte
    Dim pc0 As Byte

    On Error GoTo Fallo
    fallos = 0

    ' JMP dir: PC <- dir, banderas intactas
    ClearMem
    ResetRegisters
    SetPC &H10
    UpdateFlags 5, 1
    z0 = modFlags.ZF: c0 = modFlags.CF: s0 = modFlags.SF
    Preparar DecodeByte(OP_JMP), &H40
    PasoExecute
    PasoStore
    If PC <> &H40 Or Not BanderasIguales(z0, c0, s0) Then
        Debug.Print "FALLO JMP", "PC=" & PC
        fallos = fallos + 1
    Else
        Debug.Print "JMP 40h", "ok", "PC=" & PC
    End If

    ' JZ tomado: ZF=1 -> PC <- dir
    ClearMem
    ResetRegisters
    SetPC &H20
    UpdateFlags 0, 0
    z0 = modFlags.ZF: c0 = modFlags.CF: s0 = modFlags.SF
    Preparar DecodeByte(OP_JZ), &H80
    PasoExecute
    PasoStore
    If PC <> &H80 Or modFlags.ZF <> 1 Or Not BanderasIguales(z0, c0, s0) Then
        Debug.Print "FALLO JZ tomado", "PC=" & PC, "ZF=" & modFlags.ZF
        fallos = fallos + 1
    Else
        Debug.Print "JZ tomado (ZF=1)", "ok", "PC=" & PC
    End If

    ' JZ no tomado: ZF=0 -> PC se queda
    ClearMem
    ResetRegisters
    SetPC &H22
    pc0 = PC
    UpdateFlags 1, 0
    z0 = modFlags.ZF: c0 = modFlags.CF: s0 = modFlags.SF
    Preparar DecodeByte(OP_JZ), &H90
    PasoExecute
    PasoStore
    If PC <> pc0 Or Not BanderasIguales(z0, c0, s0) Then
        Debug.Print "FALLO JZ no tomado", "PC=" & PC
        fallos = fallos + 1
    Else
        Debug.Print "JZ no tomado (ZF=0)", "ok", "PC=" & PC
    End If

    ' JNZ tomado: ZF=0 -> PC <- dir
    ClearMem
    ResetRegisters
    SetPC &H30
    UpdateFlags 7, 0
    z0 = modFlags.ZF: c0 = modFlags.CF: s0 = modFlags.SF
    Preparar DecodeByte(OP_JNZ), &HA0
    PasoExecute
    PasoStore
    If PC <> &HA0 Or Not BanderasIguales(z0, c0, s0) Then
        Debug.Print "FALLO JNZ tomado", "PC=" & PC
        fallos = fallos + 1
    Else
        Debug.Print "JNZ tomado (ZF=0)", "ok", "PC=" & PC
    End If

    ' JNZ no tomado: ZF=1 -> PC se queda
    ClearMem
    ResetRegisters
    SetPC &H33
    pc0 = PC
    UpdateFlags 0, 1
    z0 = modFlags.ZF: c0 = modFlags.CF: s0 = modFlags.SF
    Preparar DecodeByte(OP_JNZ), &HB0
    PasoExecute
    PasoStore
    If PC <> pc0 Or Not BanderasIguales(z0, c0, s0) Then
        Debug.Print "FALLO JNZ no tomado", "PC=" & PC
        fallos = fallos + 1
    Else
        Debug.Print "JNZ no tomado (ZF=1)", "ok", "PC=" & PC
    End If

    ' HLT: estado HALTED, banderas intactas, registros intactos
    ClearMem
    ResetRegisters
    SetAX &H11
    SetBX &H22
    SetPC &H50
    IniciarFetch
    UpdateFlags 3, 1
    z0 = modFlags.ZF: c0 = modFlags.CF: s0 = modFlags.SF
    Preparar DecodeByte(OP_HLT), 0
    PasoExecute
    PasoStore
    If EstadoCPU <> HALTED Or AX <> &H11 Or BX <> &H22 Or PC <> &H50 Or Not BanderasIguales(z0, c0, s0) Then
        Debug.Print "FALLO HLT", "estado=" & EstadoCPU, "AX=" & AX
        fallos = fallos + 1
    Else
        Debug.Print "HLT", "ok", "HALTED, registros y flags intactos"
    End If

    ' Tras HLT, Fetch no avanza
    pc0 = PC
    PasoFetch
    If PC <> pc0 Or EstadoCPU <> HALTED Then
        Debug.Print "FALLO HLT frena Fetch", "PC=" & PC
        fallos = fallos + 1
    Else
        Debug.Print "HLT frena Fetch", "ok"
    End If

    If fallos = 0 Then
        Debug.Print "Execute cubre JMP, JZ, JNZ y HLT"
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
