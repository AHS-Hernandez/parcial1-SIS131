Attribute VB_Name = "modControlUnit"
Option Explicit

' Maquina de estados: eCPUState + ePhase. DoStep = una micro-operacion.
' DoRun agenda el siguiente DoStep con Application.OnTime (rngDelay ms).
' DoPause cancela el disparo sin perder registros. DoReset no borra la RAM.
' En la ventana Inmediato: PruebaFetch, PruebaDecode, PruebaDoStep, PruebaRunPauseReset
' Traza manual: MOV AX,01h (10h 01h) luego HLT (FFh) - ANALISIS 6.4

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
Private mEsperandoOperando As Boolean
Private mPasoGlobal As Long
Private mHaySiguiente As Boolean
Private mFaseSiguiente As ePhase
Private mRunAgendado As Boolean
Private mProximaEjecucion As Date
Private mDelayMsFallback As Long
Private Const PROC_TICK As String = "TickRun"
Private Const DELAY_DEFAULT_MS As Long = 200

Public Property Get EstadoCPU() As eCPUState
    EstadoCPU = mEstado
End Property

Public Property Get FaseActual() As ePhase
    FaseActual = mFase
End Property

Public Property Get NumeroDePaso() As Integer
    NumeroDePaso = mPasoFetch
End Property

Public Property Get PasoGlobal() As Long
    PasoGlobal = mPasoGlobal
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
    mEsperandoOperando = False
    mHaySiguiente = False
End Sub

Public Sub DetenerCPU()
    mEstado = HALTED
    CancelarRun
End Sub

' RESET: registros y flags a 0, fase FETCH, RUNNING, log vacio. No borra la RAM.
Public Sub DoReset()
    CancelarRun
    ResetRegisters
    ResetFlags
    ClearLog
    mEstado = RUNNING
    mFase = FETCH
    mPasoFetch = 1
    mTieneOperando = False
    mEsperandoOperando = False
    mPasoGlobal = 0
    mOperando = 0
    mHaySiguiente = False
    mInstruccionActual.Encontrada = False
    mInstruccionActual.Sintaxis = ""
End Sub

' PAUSE: corta RUN; registros, RAM y fase se conservan. STEP sigue pudiendo avanzar.
Public Sub DoPause()
    If mEstado = HALTED Then Exit Sub
    mEstado = PAUSED
    CancelarRun
End Sub

' RUN: estado RUNNING y agenda el siguiente micro-paso tras el delay (sin bloquear Excel).
Public Sub DoRun()
    If mEstado = HALTED Then Exit Sub
    mEstado = RUNNING
    CancelarRun
    AgendarSiguientePaso
End Sub

' Retardo en ms: celda rngDelay si existe; si no, el fallback (default 200).
Public Property Get DelayMs() As Long
    Dim v As Long
    On Error Resume Next
    Err.Clear
    v = CLng(Range("rngDelay").Value)
    If Err.Number <> 0 Or v < 0 Then
        On Error GoTo 0
        If mDelayMsFallback > 0 Then
            DelayMs = mDelayMsFallback
        Else
            DelayMs = DELAY_DEFAULT_MS
        End If
    Else
        On Error GoTo 0
        DelayMs = v
    End If
End Property

Public Sub SetDelayMs(ByVal ms As Long)
    If ms < 0 Then ms = 0
    mDelayMsFallback = ms
End Sub

' Disparo de Application.OnTime. Publico: Excel lo llama por nombre.
Public Sub TickRun()
    mRunAgendado = False
    If mEstado <> RUNNING Then Exit Sub
    DoStep
    If mEstado = RUNNING Then
        AgendarSiguientePaso
    End If
End Sub

Private Sub AgendarSiguientePaso()
    Dim segs As Double
    segs = DelayMs / 1000#
    If segs < 0.01 Then segs = 0.01
    mProximaEjecucion = Now + segs / 86400#
    Application.OnTime EarliestTime:=mProximaEjecucion, Procedure:=PROC_TICK, Schedule:=True
    mRunAgendado = True
End Sub

Private Sub CancelarRun()
    On Error Resume Next
    If mRunAgendado Then
        Application.OnTime EarliestTime:=mProximaEjecucion, Procedure:=PROC_TICK, Schedule:=False
    End If
    On Error GoTo 0
    mRunAgendado = False
End Sub

' Un clic de STEP: una sola micro-operacion. HALTED no avanza.
Public Sub DoStep()
    Dim eraF4 As Boolean

    If mEstado = HALTED Then Exit Sub
    If mEstado = RESET Then Exit Sub

    If mHaySiguiente Then
        mHaySiguiente = False
        mFase = mFaseSiguiente
        If mFase = FETCH Then
            mPasoFetch = 1
            ' Si venimos de Decode pidiendo operando, mEsperandoOperando sigue True.
            If Not mEsperandoOperando Then
                mTieneOperando = False
            End If
        ElseIf mFase = EXECUTE Then
            Preparar mInstruccionActual, mOperando
        End If
    End If

    mPasoGlobal = mPasoGlobal + 1

    Select Case mFase
        Case FETCH
            eraF4 = (mPasoFetch = 4)
            PasoFetchInterno
            If eraF4 Then
                If mEsperandoOperando Then
                    mOperando = IR
                    mTieneOperando = True
                    mEsperandoOperando = False
                    ProgramarFase EXECUTE
                Else
                    ProgramarFase DECODE
                End If
            End If

        Case DECODE
            PasoDecode
            If mEstado = HALTED Then Exit Sub
            If mInstruccionActual.Bytes = 2 And Not mTieneOperando Then
                mEsperandoOperando = True
                ProgramarFase FETCH
            Else
                ProgramarFase EXECUTE
            End If

        Case EXECUTE
            PasoExecute
            If mEstado = HALTED Then Exit Sub
            If EscribeRegistro Then
                ProgramarFase STORE
            Else
                ProgramarFase FETCH
            End If

        Case STORE
            PasoStore
            ProgramarFase FETCH
    End Select
End Sub

Private Sub ProgramarFase(ByVal f As ePhase)
    mHaySiguiente = True
    mFaseSiguiente = f
End Sub

' Un solo micro-paso de Fetch (API publica para pruebas unitarias de Fetch).
Public Sub PasoFetch()
    If mEstado = HALTED Then Exit Sub
    mFase = FETCH
    PasoFetchInterno
End Sub

Private Sub PasoFetchInterno()
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
        mTieneOperando = False
    End If
End Sub

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

' Traza manual ANALISIS 6.4: MOV AX,01h + HLT. Cada DoStep = un micro.
Public Sub PruebaDoStep()
    Dim fallos As Long

    On Error GoTo Fallo
    fallos = 0

    ClearMem
    WriteMem 0, OP_MOV_AX_IMM
    WriteMem 1, &H1
    WriteMem 2, OP_HLT
    DoReset

    Debug.Print "traza mano MOV AX,01h luego HLT"

    DoStep
    fallos = fallos + CheckPaso(1, FETCH, 0, 0, 0, 0, 0, "MAR<-PC")

    DoStep
    fallos = fallos + CheckPaso(2, FETCH, 0, 0, &H10, 0, 0, "MDR<-RAM")

    DoStep
    fallos = fallos + CheckPaso(3, FETCH, 0, 0, &H10, &H10, 0, "IR<-MDR")

    DoStep
    fallos = fallos + CheckPaso(4, FETCH, 1, 0, &H10, &H10, 0, "PC+1")

    DoStep
    If FaseActual <> DECODE Or InstruccionActual.Sintaxis <> "MOV AX, imm" Then
        Debug.Print "FALLO paso5 DECODE", "fase=" & FaseActual, InstruccionActual.Sintaxis
        fallos = fallos + 1
    Else
        Debug.Print "paso5 DECODE", "ok", "pide operando"
    End If

    DoStep
    fallos = fallos + CheckPaso(6, FETCH, 1, 1, &H10, &H10, 0, "op MAR<-PC")
    DoStep
    fallos = fallos + CheckPaso(7, FETCH, 1, 1, &H1, &H10, 0, "op MDR<-RAM")
    DoStep
    fallos = fallos + CheckPaso(8, FETCH, 1, 1, &H1, &H1, 0, "op IR<-MDR")
    DoStep
    If PC <> 2 Or FaseActual <> FETCH Or OperandoActual <> 1 Then
        Debug.Print "FALLO paso9 op PC+1", "PC=" & PC, "fase=" & FaseActual, "op=" & OperandoActual
        fallos = fallos + 1
    Else
        Debug.Print "paso9 FETCH PC+1", "ok", "operando=1"
    End If

    DoStep
    If AX <> 0 Or Temporal <> 1 Or FaseActual <> EXECUTE Then
        Debug.Print "FALLO paso10 EXECUTE", "AX=" & AX, "temp=" & Temporal, "fase=" & FaseActual
        fallos = fallos + 1
    Else
        Debug.Print "paso10 EXECUTE", "ok", "Temporal=1"
    End If

    DoStep
    If AX <> 1 Or FaseActual <> STORE Then
        Debug.Print "FALLO paso11 STORE", "AX=" & AX, "fase=" & FaseActual
        fallos = fallos + 1
    Else
        Debug.Print "paso11 STORE", "ok", "AX=1"
    End If

    ' HLT: 4x Fetch + Decode + Execute
    DoStep
    DoStep
    DoStep
    DoStep
    DoStep
    If InstruccionActual.Sintaxis <> "HLT" Or FaseActual <> DECODE Then
        Debug.Print "FALLO HLT DECODE", InstruccionActual.Sintaxis, "fase=" & FaseActual
        fallos = fallos + 1
    Else
        Debug.Print "HLT DECODE", "ok"
    End If
    DoStep
    If EstadoCPU <> HALTED Or AX <> 1 Or FaseActual <> EXECUTE Then
        Debug.Print "FALLO HLT EXECUTE", "estado=" & EstadoCPU, "AX=" & AX, "fase=" & FaseActual
        fallos = fallos + 1
    Else
        Debug.Print "HLT EXECUTE", "ok", "HALTED AX=1"
    End If

    DoStep
    If EstadoCPU <> HALTED Then
        Debug.Print "FALLO STEP en HALTED avanzo"
        fallos = fallos + 1
    Else
        Debug.Print "STEP en HALTED", "ok", "no avanza"
    End If

    ClearMem
    WriteMem 0, OP_HLT
    DoReset
    DoPause
    DoStep
    If EstadoCPU <> PAUSED Or FaseActual <> FETCH Or MAR <> 0 Then
        Debug.Print "FALLO PAUSE+STEP", "estado=" & EstadoCPU, "fase=" & FaseActual
        fallos = fallos + 1
    Else
        Debug.Print "PAUSE+STEP", "ok", "avanza y sigue PAUSED"
    End If

    If fallos = 0 Then
        Debug.Print "DoStep sigue la traza manual de MOV AX,01h y HLT"
    Else
        Debug.Print "FALLOS", fallos
    End If
    Exit Sub
Fallo:
    Debug.Print "FALLO inesperado", Err.Number, Err.Description
End Sub

Public Sub PruebaRunPauseReset()
    Dim fallos As Long
    Dim i As Long
    Dim axAntes As Byte
    Dim pcAntes As Byte
    Dim ram0 As Byte

    On Error GoTo Fallo
    fallos = 0

    ' RESET no borra la RAM; si limpia registros, flags, fase y log.
    ClearMem
    WriteMem 0, OP_MOV_AX_IMM
    WriteMem 1, &H5
    WriteMem 2, OP_HLT
    WriteMem &H80, &HAA
    SetAX &H11
    SetBX &H22
    SetPC 7
    UpdateFlags 0, 1
    AppendLog "basura"
    DoReset
    If PC <> 0 Or AX <> 0 Or BX <> 0 Or IR <> 0 Or MAR <> 0 Or MDR <> 0 Then
        Debug.Print "FALLO RESET registros", "PC=" & PC, "AX=" & AX
        fallos = fallos + 1
    ElseIf modFlags.ZF <> 0 Or modFlags.CF <> 0 Or modFlags.SF <> 0 Then
        Debug.Print "FALLO RESET flags"
        fallos = fallos + 1
    ElseIf FaseActual <> FETCH Or EstadoCPU <> RUNNING Or FilasLog <> 0 Then
        Debug.Print "FALLO RESET fase/estado/log", "fase=" & FaseActual, "filas=" & FilasLog
        fallos = fallos + 1
    ElseIf ReadMem(0) <> OP_MOV_AX_IMM Or ReadMem(&H80) <> &HAA Then
        Debug.Print "FALLO RESET borro RAM"
        fallos = fallos + 1
    Else
        Debug.Print "RESET", "ok", "CPU en cero, RAM intacta, log vacio"
    End If

    ' PAUSE conserva estado; STEP sigue avanzando
    DoReset
    DoStep
    DoStep
    axAntes = AX
    pcAntes = PC
    ram0 = ReadMem(0)
    DoPause
    If EstadoCPU <> PAUSED Then
        Debug.Print "FALLO DoPause estado", EstadoCPU
        fallos = fallos + 1
    End If
    DoStep
    If EstadoCPU <> PAUSED Or AX <> axAntes Or ReadMem(0) <> ram0 Then
        Debug.Print "FALLO PAUSE perdio estado", "AX=" & AX, "estado=" & EstadoCPU
        fallos = fallos + 1
    Else
        Debug.Print "PAUSE", "ok", "conserva CPU/RAM y STEP avanza"
    End If

    ' RUN via TickRun (sin dejar OnTime pendiente): llega a HLT
    DoReset
    SetDelayMs 50
    mEstado = RUNNING
    For i = 1 To 80
        If EstadoCPU = HALTED Then Exit For
        TickRun
        CancelarRun
        If EstadoCPU = RUNNING Then
            ' sigue el bucle de prueba
        ElseIf EstadoCPU = PAUSED Then
            Exit For
        End If
    Next i
    If EstadoCPU <> HALTED Or AX <> 5 Then
        Debug.Print "FALLO RUN hasta HLT", "estado=" & EstadoCPU, "AX=" & AX, "pasos=" & i
        fallos = fallos + 1
    Else
        Debug.Print "RUN (TickRun)", "ok", "HALTED AX=5 en " & i & " ticks"
    End If

    ' Tras HLT, DoRun no reagenda
    DoRun
    If EstadoCPU <> HALTED Or mRunAgendado Then
        Debug.Print "FALLO DoRun en HALTED", "estado=" & EstadoCPU, "agendado=" & mRunAgendado
        fallos = fallos + 1
    Else
        Debug.Print "DoRun en HALTED", "ok", "no agenda"
    End If

    ' DelayMs siempre no negativo (rngDelay o fallback)
    SetDelayMs 350
    If DelayMs < 0 Then
        Debug.Print "FALLO DelayMs", DelayMs
        fallos = fallos + 1
    Else
        Debug.Print "DelayMs", "ok", DelayMs & " ms"
    End If

    If fallos = 0 Then
        Debug.Print "RUN, PAUSE y RESET cumplen la consigna"
    Else
        Debug.Print "FALLOS", fallos
    End If
    Exit Sub
Fallo:
    Debug.Print "FALLO inesperado", Err.Number, Err.Description
End Sub

Private Function CheckPaso(ByVal n As Long, ByVal faseEsp As ePhase, _
                           ByVal pcEsp As Long, ByVal marEsp As Long, _
                           ByVal mdrEsp As Long, ByVal irEsp As Long, _
                           ByVal axEsp As Long, ByVal nombre As String) As Long
    If FaseActual <> faseEsp Or PC <> pcEsp Or MAR <> marEsp _
       Or MDR <> mdrEsp Or IR <> irEsp Or AX <> axEsp Then
        Debug.Print "FALLO paso" & n, nombre, _
            "fase " & faseEsp & "/" & FaseActual, _
            "PC " & pcEsp & "/" & PC, _
            "MAR " & marEsp & "/" & MAR, _
            "MDR " & mdrEsp & "/" & MDR, _
            "IR " & irEsp & "/" & IR, _
            "AX " & axEsp & "/" & AX
        CheckPaso = 1
    Else
        Debug.Print "paso" & n, nombre, "ok", _
            "PC=" & PC, "MAR=" & MAR, "MDR=" & MDR, "IR=" & IR, "AX=" & AX
        CheckPaso = 0
    End If
End Function

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
