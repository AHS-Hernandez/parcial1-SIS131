Attribute VB_Name = "modControlUnit"
Option Explicit

' Fetch y Decode. Un micro-paso por llamada.
' En la ventana Inmediato: PruebaFetch y PruebaDecode
' Traza Fetch: PC=00h y RAM(00h)=10h => MAR=00, MDR=10h, IR=10h, PC=01h

Public Enum eCPUState
    RUNNING = 0
    PAUSED = 1
    HALTED = 2
    RESET = 3
End Enum

Public Enum ePhase
    FETCH = 0
    DECODE = 1
    EXECUTE = 2
    STORE = 3
End Enum

Private mEstado As eCPUState
Private mFase As ePhase
Private mPasoFetch As Integer
Private mInstruccionActual As tInstruccion
Private mOperando As Byte
Private mTieneOperando As Boolean

Public Property Get EstadoCPU() As eCPUState
    EstadoCPU = mEstado
End Property

Public Property Get FaseActual() As ePhase
    FaseActual = mFase
End Property

Public Property Get NumeroDePaso() As Integer
    NumeroDePaso = mPasoFetch
End Property

Public Property Get InstruccionActual() As tInstruccion
    InstruccionActual = mInstruccionActual
End Property

Public Property Get OperandoActual() As Byte
    OperandoActual = mOperando
End Property

Public Property Get TieneOperando() As Boolean
    TieneOperando = mTieneOperando
End Property

Public Sub IniciarFetch()
    mEstado = RUNNING
    mFase = FETCH
    mPasoFetch = 1
    mTieneOperando = False
End Sub

' Lo llama Execute al encontrar HLT (o un opcode invalido ya lo hace Decode).
Public Sub DetenerCPU()
    mEstado = HALTED
End Sub

' Un solo micro-paso: 1 PC->MAR, 2 RAM->MDR, 3 MDR->IR, 4 PC+1.
Public Sub PasoFetch()
    If mEstado = HALTED Then Exit Sub
    mFase = FETCH
    Select Case mPasoFetch
        Case 1: Fetch1_PCaMAR
        Case 2: Fetch2_RAMaMDR
        Case 3: Fetch3_MDRaIR
        Case 4: Fetch4_IncrementarPC
        Case Else: mPasoFetch = 1
    End Select
End Sub

Public Sub Fetch1_PCaMAR()
    SetMAR PC
    mPasoFetch = 2
End Sub

Public Sub Fetch2_RAMaMDR()
    SetMDR ReadMem(MAR)
    mPasoFetch = 3
End Sub

Public Sub Fetch3_MDRaIR()
    SetIR MDR
    mPasoFetch = 4
End Sub

Public Sub Fetch4_IncrementarPC()
    IncrementarPC
    mPasoFetch = 1
End Sub

' Lee IR y traduce el opcode con la tabla ISA.
Public Sub PasoDecode()
    Dim info As tInstruccion

    If mEstado = HALTED Then Exit Sub
    mFase = DECODE
    info = DecodeByte(IR)

    If Not info.Encontrada Then
        mInstruccionActual = info
        mTieneOperando = False
        mEstado = HALTED
        Exit Sub
    End If

    mInstruccionActual = info
    If info.Bytes = 1 Then
        mTieneOperando = False
        mOperando = 0
    Else
        ' El segundo byte se trae con otro Fetch; Decode ya sabe el modo.
        mTieneOperando = False
    End If
End Sub

' Completa una instruccion de 2 bytes: Fetch del operando y lo guarda.
Public Sub TraerOperando()
    If mEstado = HALTED Then Exit Sub
    If mInstruccionActual.Bytes <> 2 Then Exit Sub

    PasoFetch
    PasoFetch
    PasoFetch
    PasoFetch
    mOperando = IR
    mTieneOperando = True
    mFase = DECODE
End Sub

Public Sub PruebaFetch()
    On Error GoTo Fallo
    ClearMem
    ResetRegisters
    WriteMem 0, &H10
    SetPC 0
    IniciarFetch

    Debug.Print "traza mano PC=00 RAM(00)=10h"
    PasoFetch
    Comparar "paso1 MAR<-PC", MAR, 0, MDR, 0, IR, 0, PC, 0

    PasoFetch
    Comparar "paso2 MDR<-RAM", MAR, 0, MDR, &H10, IR, 0, PC, 0

    PasoFetch
    Comparar "paso3 IR<-MDR", MAR, 0, MDR, &H10, IR, &H10, PC, 0

    PasoFetch
    Comparar "paso4 PC<-PC+1", MAR, 0, MDR, &H10, IR, &H10, PC, 1

    WriteMem 5, &HA0
    SetPC 5
    IniciarFetch
    PasoFetch
    PasoFetch
    PasoFetch
    PasoFetch
    Comparar "traza PC=05 RAM=A0h", MAR, 5, MDR, &HA0, IR, &HA0, PC, 6

    WriteMem 255, &HFF
    SetPC 255
    IniciarFetch
    PasoFetch
    PasoFetch
    PasoFetch
    PasoFetch
    Comparar "traza PC=FFh", MAR, 255, MDR, &HFF, IR, &HFF, PC, 0
    Exit Sub
Fallo:
    Debug.Print "FALLO inesperado", Err.Number, Err.Description
End Sub

Public Sub PruebaDecode()
    Dim fallos As Long

    On Error GoTo Fallo
    fallos = 0
    fallos = fallos + VerificarDecode(OP_MOV_AX_IMM, "MOV AX, imm", "IMM", 2)
    fallos = fallos + VerificarDecode(OP_MOV_AX_BX, "MOV AX, BX", "REG", 1)
    fallos = fallos + VerificarDecode(OP_LOAD_AX, "LOAD AX, [dir]", "DIRECT", 2)
    fallos = fallos + VerificarDecode(OP_STORE_BX, "STORE [dir], BX", "DIRECT", 2)
    fallos = fallos + VerificarDecode(OP_ADD_AX_IMM, "ADD AX, imm", "IMM", 2)
    fallos = fallos + VerificarDecode(OP_ADD_AX_BX, "ADD AX, BX", "REG", 1)
    fallos = fallos + VerificarDecode(OP_SUB_BX_AX, "SUB BX, AX", "REG", 1)
    fallos = fallos + VerificarDecode(OP_INC_AX, "INC AX", "REG", 1)
    fallos = fallos + VerificarDecode(OP_DEC_BX, "DEC BX", "REG", 1)
    fallos = fallos + VerificarDecode(OP_CMP_AX_BX, "CMP AX, BX", "REG", 1)
    fallos = fallos + VerificarDecode(OP_JMP, "JMP dir", "DIRECT", 2)
    fallos = fallos + VerificarDecode(OP_JZ, "JZ dir", "DIRECT", 2)
    fallos = fallos + VerificarDecode(OP_JNZ, "JNZ dir", "DIRECT", 2)
    fallos = fallos + VerificarDecode(OP_HLT, "HLT", "NONE", 1)
    fallos = fallos + VerificarDecode(OP_NOT_AX, "NOT AX", "REG", 1)

    ' Opcode desconocido: Decode debe dejar HALTED.
    ClearMem
    ResetRegisters
    mEstado = RUNNING
    SetIR 0
    PasoDecode
    If InstruccionActual.Encontrada Or EstadoCPU <> HALTED Then
        Debug.Print "FALLO opcode 00h", "debia ser desconocido y HALTED"
        fallos = fallos + 1
    Else
        Debug.Print "opcode 00h", "ok desconocido y HALTED"
    End If

    ' Traza corta: Fetch de MOV AX,imm (10h 05h) y Decode.
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
    If InstruccionActual.Sintaxis <> "MOV AX, imm" Or InstruccionActual.Bytes <> 2 Then
        Debug.Print "FALLO Fetch+Decode MOV", InstruccionActual.Sintaxis
        fallos = fallos + 1
    Else
        Debug.Print "Fetch+Decode MOV AX,imm", "ok", "modo " & InstruccionActual.Modo
        TraerOperando
        If Not TieneOperando Or OperandoActual <> 5 Then
            Debug.Print "FALLO operando", "esperado 5", "obtenido " & OperandoActual
            fallos = fallos + 1
        Else
            Debug.Print "operando", "ok", "imm=" & OperandoActual
        End If
    End If

    If fallos = 0 Then
        Debug.Print "Decode identifica la ISA a partir del opcode en IR"
    Else
        Debug.Print "FALLOS", fallos
    End If
    Exit Sub
Fallo:
    Debug.Print "FALLO inesperado", Err.Number, Err.Description
End Sub

Private Function VerificarDecode(ByVal opcode As Long, ByVal sintaxis As String, ByVal modo As String, ByVal bytes As Integer) As Long
    ClearMem
    ResetRegisters
    mEstado = RUNNING
    SetIR opcode
    PasoDecode
    If Not InstruccionActual.Encontrada _
       Or InstruccionActual.Sintaxis <> sintaxis _
       Or InstruccionActual.Modo <> modo _
       Or InstruccionActual.Bytes <> bytes _
       Or InstruccionActual.Opcode <> opcode Then
        Debug.Print "FALLO", Hex$(opcode) & "h", "esperado " & sintaxis, "obtenido " & InstruccionActual.Sintaxis
        VerificarDecode = 1
    Else
        Debug.Print Hex$(opcode) & "h", "ok", sintaxis, "modo " & modo, "bytes " & bytes
        VerificarDecode = 0
    End If
End Function

Private Sub Comparar(ByVal nombre As String, _
                     ByVal marReal As Long, ByVal marEsp As Long, _
                     ByVal mdrReal As Long, ByVal mdrEsp As Long, _
                     ByVal irReal As Long, ByVal irEsp As Long, _
                     ByVal pcReal As Long, ByVal pcEsp As Long)
    If marReal = marEsp And mdrReal = mdrEsp And irReal = irEsp And pcReal = pcEsp Then
        Debug.Print nombre, "ok", "MAR=" & marReal, "MDR=" & mdrReal, "IR=" & irReal, "PC=" & pcReal
    Else
        Debug.Print nombre, "FALLO", _
            "MAR esp/obt " & marEsp & "/" & marReal, _
            "MDR " & mdrEsp & "/" & mdrReal, _
            "IR " & irEsp & "/" & irReal, _
            "PC " & pcEsp & "/" & pcReal
    End If
End Sub
