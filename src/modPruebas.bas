Attribute VB_Name = "modPruebas"
Option Explicit

' Prueba aislada de las 12 familias ISA + programa corto en STEP y RUN hasta HLT.
' Programa demo (multiplicacion): PruebaProgramaDemo, PruebaDemoStepRun
' Casos limite: PruebaCasosLimite
' Flujo / HLT / validaciones: PruebaFlujoHltValidaciones
' En la ventana Inmediato: PruebaISACompleta, PruebaProgramaDemo, PruebaDemoStepRun,
'   PruebaCasosLimite, PruebaFlujoHltValidaciones
' Notas: docs/PRUEBAS.md, docs/PROGRAMA_DEMO.md

Public Sub PruebaISACompleta()
    Dim fallos As Long
    Dim axS As Byte
    Dim bxS As Byte
    Dim zS As Byte
    Dim cS As Byte
    Dim sS As Byte
    Dim ramS As Byte
    Dim axR As Byte
    Dim bxR As Byte
    Dim zR As Byte
    Dim cR As Byte
    Dim sR As Byte
    Dim ramR As Byte
    Dim lineas() As String

    On Error GoTo Fallo
    fallos = 0

    Debug.Print "=== 12 familias ISA (aislado) ==="
    fallos = fallos + ProbarMOV()
    fallos = fallos + ProbarLOAD()
    fallos = fallos + ProbarSTORE()
    fallos = fallos + ProbarADD()
    fallos = fallos + ProbarSUB()
    fallos = fallos + ProbarINC()
    fallos = fallos + ProbarDEC()
    fallos = fallos + ProbarCMP()
    fallos = fallos + ProbarJMP()
    fallos = fallos + ProbarJZ()
    fallos = fallos + ProbarJNZ()
    fallos = fallos + ProbarHLT()

    Debug.Print "=== Programa corto STEP vs RUN ==="
    ReDim lineas(0 To 3)
    lineas(0) = "MOV AX, 02h"
    lineas(1) = "ADD AX, 03h"
    lineas(2) = "STORE [80h], AX"
    lineas(3) = "HLT"

    ' STEP hasta HLT
    ClearMem
    WriteMem &H80, 0
    CargarPrograma lineas
    CorrerHastaHltStep 200
    axS = AX: bxS = BX
    zS = modFlags.ZF: cS = modFlags.CF: sS = modFlags.SF
    ramS = ReadMem(&H80)
    If EstadoCPU <> HALTED Or axS <> 5 Or ramS <> 5 Then
        Debug.Print "FALLO STEP hasta HLT", "estado=" & EstadoCPU, "AX=" & axS, "RAM80=" & ramS
        fallos = fallos + 1
    Else
        Debug.Print "STEP hasta HLT", "ok", "AX=5 RAM(80h)=5 HALTED"
    End If

    ' RUN (TickRun) hasta HLT ? mismo programa
    ClearMem
    WriteMem &H80, 0
    CargarPrograma lineas
    CorrerHastaHltRun 200
    axR = AX: bxR = BX
    zR = modFlags.ZF: cR = modFlags.CF: sR = modFlags.SF
    ramR = ReadMem(&H80)
    If EstadoCPU <> HALTED Or axR <> 5 Or ramR <> 5 Then
        Debug.Print "FALLO RUN hasta HLT", "estado=" & EstadoCPU, "AX=" & axR, "RAM80=" & ramR
        fallos = fallos + 1
    Else
        Debug.Print "RUN hasta HLT", "ok", "AX=5 RAM(80h)=5 HALTED"
    End If

    If axS <> axR Or bxS <> bxR Or zS <> zR Or cS <> cR Or sS <> sR Or ramS <> ramR Then
        Debug.Print "FALLO STEP y RUN difieren", "AX " & axS & "/" & axR, "RAM " & ramS & "/" & ramR
        fallos = fallos + 1
    Else
        Debug.Print "STEP == RUN", "ok", "mismo resultado final"
    End If

    If fallos = 0 Then
        Debug.Print "ISA completa y ciclo STEP/RUN hasta HLT OK"
    Else
        Debug.Print "FALLOS", fallos
    End If
    Exit Sub
Fallo:
    Debug.Print "FALLO inesperado", Err.Number, Err.Description
End Sub

Private Function ProbarMOV() As Long
    Dim z0 As Byte, c0 As Byte, s0 As Byte
    ClearMem
    ResetRegisters
    UpdateFlags 7, 1
    z0 = modFlags.ZF: c0 = modFlags.CF: s0 = modFlags.SF
    Preparar DecodeByte(OP_MOV_AX_IMM), &H2A
    PasoExecute
    PasoStore
    If AX <> &H2A Or modFlags.ZF <> z0 Or modFlags.CF <> c0 Or modFlags.SF <> s0 Then
        Debug.Print "FALLO MOV", "AX=" & AX
        ProbarMOV = 1
    Else
        Debug.Print "MOV", "ok", "AX=2Ah flags intactos"
        ProbarMOV = 0
    End If
End Function

Private Function ProbarLOAD() As Long
    ClearMem
    ResetRegisters
    WriteMem &H80, &H3C
    Preparar DecodeByte(OP_LOAD_AX), &H80
    PasoExecute
    PasoStore
    If AX <> &H3C Or MAR <> &H80 Or MDR <> &H3C Then
        Debug.Print "FALLO LOAD", "AX=" & AX, "MAR=" & MAR
        ProbarLOAD = 1
    Else
        Debug.Print "LOAD", "ok", "AX=3Ch via MAR/MDR"
        ProbarLOAD = 0
    End If
End Function

Private Function ProbarSTORE() As Long
    ClearMem
    ResetRegisters
    SetAX &H55
    Preparar DecodeByte(OP_STORE_AX), &H81
    PasoExecute
    PasoStore
    If ReadMem(&H81) <> &H55 Or MAR <> &H81 Or MDR <> &H55 Then
        Debug.Print "FALLO STORE", "RAM=" & ReadMem(&H81)
        ProbarSTORE = 1
    Else
        Debug.Print "STORE", "ok", "RAM(81h)=55h"
        ProbarSTORE = 0
    End If
End Function

Private Function ProbarADD() As Long
    ClearMem
    ResetRegisters
    SetAX 2
    Preparar DecodeByte(OP_ADD_AX_IMM), 3
    PasoExecute
    PasoStore
    If AX <> 5 Or modFlags.ZF <> 0 Or modFlags.CF <> 0 Then
        Debug.Print "FALLO ADD", "AX=" & AX
        ProbarADD = 1
    Else
        Debug.Print "ADD", "ok", "AX=5"
        ProbarADD = 0
    End If
End Function

Private Function ProbarSUB() As Long
    ClearMem
    ResetRegisters
    SetAX 5
    Preparar DecodeByte(OP_SUB_AX_IMM), 3
    PasoExecute
    PasoStore
    If AX <> 2 Or modFlags.CF <> 0 Then
        Debug.Print "FALLO SUB", "AX=" & AX
        ProbarSUB = 1
    Else
        Debug.Print "SUB", "ok", "AX=2"
        ProbarSUB = 0
    End If
End Function

Private Function ProbarINC() As Long
    ClearMem
    ResetRegisters
    SetAX &HFF
    Preparar DecodeByte(OP_INC_AX), 0
    PasoExecute
    PasoStore
    If AX <> 0 Or modFlags.ZF <> 1 Or modFlags.CF <> 1 Then
        Debug.Print "FALLO INC", "AX=" & AX
        ProbarINC = 1
    Else
        Debug.Print "INC", "ok", "FFh+1 -> 0 CF=1 ZF=1"
        ProbarINC = 0
    End If
End Function

Private Function ProbarDEC() As Long
    ClearMem
    ResetRegisters
    SetBX 1
    Preparar DecodeByte(OP_DEC_BX), 0
    PasoExecute
    PasoStore
    If BX <> 0 Or modFlags.ZF <> 1 Then
        Debug.Print "FALLO DEC", "BX=" & BX
        ProbarDEC = 1
    Else
        Debug.Print "DEC", "ok", "BX=0 ZF=1"
        ProbarDEC = 0
    End If
End Function

Private Function ProbarCMP() As Long
    Dim ax0 As Byte
    ClearMem
    ResetRegisters
    SetAX 5
    ax0 = AX
    Preparar DecodeByte(OP_CMP_AX_IMM), 5
    PasoExecute
    PasoStore
    If AX <> ax0 Or modFlags.ZF <> 1 Then
        Debug.Print "FALLO CMP", "AX=" & AX, "ZF=" & modFlags.ZF
        ProbarCMP = 1
    Else
        Debug.Print "CMP", "ok", "ZF=1 AX intacto"
        ProbarCMP = 0
    End If
End Function

Private Function ProbarJMP() As Long
    ClearMem
    ResetRegisters
    SetPC &H10
    Preparar DecodeByte(OP_JMP), &H40
    PasoExecute
    PasoStore
    If PC <> &H40 Then
        Debug.Print "FALLO JMP", "PC=" & PC
        ProbarJMP = 1
    Else
        Debug.Print "JMP", "ok", "PC=40h"
        ProbarJMP = 0
    End If
End Function

Private Function ProbarJZ() As Long
    ClearMem
    ResetRegisters
    SetPC &H20
    UpdateFlags 0, 0
    Preparar DecodeByte(OP_JZ), &H80
    PasoExecute
    PasoStore
    If PC <> &H80 Or modFlags.ZF <> 1 Then
        Debug.Print "FALLO JZ", "PC=" & PC
        ProbarJZ = 1
    Else
        Debug.Print "JZ", "ok", "tomado PC=80h"
        ProbarJZ = 0
    End If
End Function

Private Function ProbarJNZ() As Long
    ClearMem
    ResetRegisters
    SetPC &H30
    UpdateFlags 7, 0
    Preparar DecodeByte(OP_JNZ), &HA0
    PasoExecute
    PasoStore
    If PC <> &HA0 Or modFlags.ZF <> 0 Then
        Debug.Print "FALLO JNZ", "PC=" & PC
        ProbarJNZ = 1
    Else
        Debug.Print "JNZ", "ok", "tomado PC=A0h"
        ProbarJNZ = 0
    End If
End Function

Private Function ProbarHLT() As Long
    ClearMem
    ResetRegisters
    IniciarFetch
    Preparar DecodeByte(OP_HLT), 0
    PasoExecute
    PasoStore
    If EstadoCPU <> HALTED Then
        Debug.Print "FALLO HLT", "estado=" & EstadoCPU
        ProbarHLT = 1
    Else
        Debug.Print "HLT", "ok", "HALTED"
        ProbarHLT = 0
    End If
End Function

Private Sub CorrerHastaHltStep(ByVal maxPasos As Long)
    Dim i As Long
    For i = 1 To maxPasos
        If EstadoCPU = HALTED Then Exit Sub
        DoStep
    Next i
End Sub

Private Sub CorrerHastaHltRun(ByVal maxPasos As Long)
    Dim i As Long
    For i = 1 To maxPasos
        If EstadoCPU = HALTED Then Exit Sub
        TickRun
        DetenerRun
    Next i
End Sub

' Multiplicacion por sumas: N en 80h, M en 81h, producto en 82h (docs/PROGRAMA_DEMO.md).
Public Sub PruebaProgramaDemo()
    Dim fallos As Long
    Dim lineas() As String
    Dim i As Long
    Dim esperado As Variant
    Dim maxRun As Long

    On Error GoTo Fallo
    fallos = 0

    ReDim lineas(0 To 11)
    lineas(0) = "MOV AX, 00h"
    lineas(1) = "LOAD BX, [80h]"
    lineas(2) = "CMP BX, 00h"
    lineas(3) = "JZ 12h"
    lineas(4) = "LOAD BX, [81h]"
    lineas(5) = "ADD AX, BX"
    lineas(6) = "LOAD BX, [80h]"
    lineas(7) = "DEC BX"
    lineas(8) = "STORE [80h], BX"
    lineas(9) = "JMP 02h"
    lineas(10) = "STORE [82h], AX"
    lineas(11) = "HLT"

    ' Bytes a mano (traduccion de PROGRAMA_DEMO.md)
    esperado = Array( _
        &H10, &H0, _
        &H21, &H80, _
        &H72, &H0, _
        &H90, &H12, _
        &H21, &H81, _
        &H41, _
        &H21, &H80, _
        &H63, _
        &H31, &H80, _
        &H80, &H2, _
        &H30, &H82, _
        &HFF)

    ClearMem
    WriteMem &H80, 3
    WriteMem &H81, 4
    WriteMem &H82, 0
    CargarPrograma lineas

    For i = 0 To UBound(esperado)
        If ReadMem(i) <> CByte(esperado(i)) Then
            Debug.Print "FALLO byte", Hex$(i) & "h", "esp " & Hex$(esperado(i)), "obt " & Hex$(ReadMem(i))
            fallos = fallos + 1
        End If
    Next i
    If fallos = 0 Then
        Debug.Print "bytes demo", "ok", "00h-14h coinciden con la traduccion a mano"
    End If

    If ReadMem(&H80) <> 3 Or ReadMem(&H81) <> 4 Then
        Debug.Print "FALLO datos 80h/81h tras load"
        fallos = fallos + 1
    End If

    ' 3*4 = 12: bucle + JZ + STORE + HLT (margen de micros)
    maxRun = 500
    CorrerHastaHltStep maxRun
    If EstadoCPU <> HALTED Or AX <> 12 Or ReadMem(&H82) <> 12 Then
        Debug.Print "FALLO demo STEP", "estado=" & EstadoCPU, "AX=" & AX, "RAM82=" & ReadMem(&H82)
        fallos = fallos + 1
    Else
        Debug.Print "demo STEP", "ok", "3*4=0Ch en AX y [82h], HALTED"
    End If

    ' N=0: producto 0 por JZ inmediato
    ClearMem
    WriteMem &H80, 0
    WriteMem &H81, 5
    WriteMem &H82, &HFF
    CargarPrograma lineas
    CorrerHastaHltStep maxRun
    If EstadoCPU <> HALTED Or AX <> 0 Or ReadMem(&H82) <> 0 Then
        Debug.Print "FALLO demo N=0", "AX=" & AX, "RAM82=" & ReadMem(&H82)
        fallos = fallos + 1
    Else
        Debug.Print "demo N=0", "ok", "JZ a FIN, producto 0"
    End If

    If fallos = 0 Then
        Debug.Print "Programa demostrativo (multiplicacion) verificado"
    Else
        Debug.Print "FALLOS", fallos
    End If
    Exit Sub
Fallo:
    Debug.Print "FALLO inesperado", Err.Number, Err.Description
End Sub

' Tras DoLoad desde la hoja PROGRAM: deja N=3, M=4, producto=0 en datos.
Public Sub SembrarDatosDemo()
    Debug.Print "datos demo", "[80h]=3 [81h]=4 [82h]=0"
End Sub

' Issue 27: demo completo en STEP y RUN; log final debe coincidir.
Public Sub PruebaDemoStepRun()
    Dim fallos As Long
    Dim lineas() As String
    Dim logStep() As String
    Dim logRun() As String
    Dim filasStep As Long
    Dim filasRun As Long
    Dim axStep As Byte
    Dim axRun As Byte
    Dim ramStep As Byte
    Dim ramRun As Byte
    Dim i As Long
    Dim maxPasos As Long
    Dim faseUI As String
    Dim microUI As String

    On Error GoTo Fallo
    fallos = 0
    maxPasos = 800
    lineas = LineasDemoMultiplicacion()

    Debug.Print "=== Demo STEP (LOAD + resaltado + log) ==="
    ClearMem
    ClearMem
    CargarPrograma lineas
    DoReset
    CorrerHastaHltStep maxPasos

    filasStep = FilasLog
    axStep = AX
    ramStep = ReadMem(&H82)
    On Error Resume Next
    faseUI = CStr(Range("rngPhase").Value & "")
    microUI = CStr(Range("rngMicroOp").Value & "")
    On Error GoTo Fallo

    If EstadoCPU <> HALTED Or axStep <> 12 Or ramStep <> 12 Then
        Debug.Print "FALLO STEP resultado", "estado=" & EstadoCPU, "AX=" & axStep, "RAM82=" & ramStep
        fallos = fallos + 1
    Else
        Debug.Print "STEP resultado", "ok", "AX=0Ch [82h]=0Ch HALTED"
    End If

    If filasStep < 20 Then
        Debug.Print "FALLO STEP log corto", "filas=" & filasStep
        fallos = fallos + 1
    Else
        Debug.Print "STEP log", "ok", "filas=" & filasStep
    End If

    If Len(faseUI) = 0 Or Len(microUI) = 0 Then
        Debug.Print "FALLO UI no refleja paso", "fase=" & faseUI, "micro=" & microUI
        fallos = fallos + 1
    Else
        Debug.Print "UI tras STEP", "ok", "fase=" & faseUI, "micro=" & Left$(microUI, 40)
    End If

    logStep = CapturarLogTextos()

    Debug.Print "=== Demo RUN (mismo programa, delay via TickRun) ==="
    ClearMem
    ClearMem
    CargarPrograma lineas
    DoReset
    SetDelayMs 50
    CorrerHastaHltRun maxPasos

    filasRun = FilasLog
    axRun = AX
    ramRun = ReadMem(&H82)
    logRun = CapturarLogTextos()

    If EstadoCPU <> HALTED Or axRun <> 12 Or ramRun <> 12 Then
        Debug.Print "FALLO RUN resultado", "estado=" & EstadoCPU, "AX=" & axRun, "RAM82=" & ramRun
        fallos = fallos + 1
    Else
        Debug.Print "RUN resultado", "ok", "AX=0Ch [82h]=0Ch HALTED"
    End If

    If filasStep <> filasRun Then
        Debug.Print "FALLO filas STEP/RUN", filasStep, filasRun
        fallos = fallos + 1
    Else
        Debug.Print "filas STEP==RUN", "ok", filasRun
    End If

    If axStep <> axRun Or ramStep <> ramRun Then
        Debug.Print "FALLO resultado STEP/RUN difiere"
        fallos = fallos + 1
    End If

    If Not LogsIguales(logStep, logRun) Then
        Debug.Print "FALLO log STEP y RUN no coinciden"
        fallos = fallos + 1
        For i = LBound(logStep) To UBound(logStep)
            If i > UBound(logRun) Then Exit For
            If logStep(i) <> logRun(i) Then
                Debug.Print "  dif fila", i + 1, Left$(logStep(i), 50), Left$(logRun(i), 50)
                Exit For
            End If
        Next i
    Else
        Debug.Print "log STEP==RUN", "ok", "mismo texto por micro"
    End If

    If fallos = 0 Then
        Debug.Print "Demo STEP y RUN: resultado y log coinciden"
    Else
        Debug.Print "FALLOS", fallos
    End If
    Exit Sub
Fallo:
    Debug.Print "FALLO PruebaDemoStepRun", Err.Number, Err.Description
End Sub

Private Function LineasDemoMultiplicacion() As String()
    Dim lineas() As String
    ReDim lineas(0 To 8)
    lineas(0) = "MOV AX, 00h"
    lineas(1) = "MOV BX, 03h"
    lineas(2) = "CMP BX, 00h"
    lineas(3) = "JZ 0Dh"
    lineas(4) = "ADD AX, 04h"
    lineas(5) = "DEC BX"
    lineas(6) = "JMP 04h"
    lineas(7) = "STORE [82h], AX"
    lineas(8) = "HLT"
    LineasDemoMultiplicacion = lineas
End Function

Private Function CapturarLogTextos() As String()
    Dim ws As Worksheet
    Dim n As Long
    Dim i As Long
    Dim out() As String

    On Error GoTo Vacio
    Set ws = ThisWorkbook.Worksheets("LOG")
    n = FilasLog
    If n <= 0 Then
        ReDim out(0 To 0)
        out(0) = ""
        CapturarLogTextos = out
        Exit Function
    End If
    ReDim out(0 To n - 1)
    For i = 0 To n - 1
        out(i) = CStr(ws.Cells(2 + i, 13).Value & "")
    Next i
    CapturarLogTextos = out
    Exit Function
Vacio:
    ReDim out(0 To 0)
    out(0) = ""
    CapturarLogTextos = out
End Function

Private Function LogsIguales(ByRef a() As String, ByRef b() As String) As Boolean
    Dim i As Long
    If UBound(a) <> UBound(b) Or LBound(a) <> LBound(b) Then
        LogsIguales = False
        Exit Function
    End If
    For i = LBound(a) To UBound(a)
        If a(i) <> b(i) Then
            LogsIguales = False
            Exit Function
        End If
    Next i
    LogsIguales = True
End Function

' Issue 28: bordes 00h/FFh, valor 0/255, ZF/CF/SF, segmentacion MEMORY.
Public Sub PruebaCasosLimite()
    Dim fallos As Long
    Dim cCod As Long
    Dim cDat As Long

    On Error GoTo Fallo
    fallos = 0
    Debug.Print "=== Casos limite memoria / flags / zonas ==="

    ' --- Memoria: direccion 00h y FFh, valor 0 y 255 ---
    ClearMem
    WriteMem 0, 0
    If ReadMem(0) <> 0 Then
        Debug.Print "FALLO mem 00h=0", "obt=" & ReadMem(0)
        fallos = fallos + 1
    Else
        Debug.Print "mem 00h <- 00h", "ok", "esp=0 obt=0"
    End If

    WriteMem 0, 255
    If ReadMem(0) <> 255 Then
        Debug.Print "FALLO mem 00h=FFh", "obt=" & ReadMem(0)
        fallos = fallos + 1
    Else
        Debug.Print "mem 00h <- FFh", "ok", "esp=255 obt=255"
    End If

    WriteMem 255, 0
    If ReadMem(255) <> 0 Then
        Debug.Print "FALLO mem FFh=0", "obt=" & ReadMem(255)
        fallos = fallos + 1
    Else
        Debug.Print "mem FFh <- 00h", "ok", "esp=0 obt=0"
    End If

    WriteMem 255, 255
    If ReadMem(255) <> 255 Then
        Debug.Print "FALLO mem FFh=FFh", "obt=" & ReadMem(255)
        fallos = fallos + 1
    Else
        Debug.Print "mem FFh <- FFh", "ok", "esp=255 obt=255"
    End If

    ' --- Flags: ZF ---
    ClearMem
    ResetRegisters
    ResetFlags
    SetAX 1
    Preparar DecodeByte(OP_SUB_AX_IMM), 1
    PasoExecute
    PasoStore
    If AX <> 0 Or modFlags.ZF <> 1 Or modFlags.SF <> 0 Or modFlags.CF <> 0 Then
        Debug.Print "FALLO ZF", "AX=" & AX, "ZF=" & modFlags.ZF, "SF=" & modFlags.SF, "CF=" & modFlags.CF
        fallos = fallos + 1
    Else
        Debug.Print "ZF (1-1=0)", "ok", "esp ZF=1 SF=0 CF=0 obt iguales"
    End If

    ' --- Flags: CF acarreo FFh+1 ---
    ClearMem
    ResetRegisters
    SetAX &HFF
    Preparar DecodeByte(OP_ADD_AX_IMM), 1
    PasoExecute
    PasoStore
    If AX <> 0 Or modFlags.CF <> 1 Or modFlags.ZF <> 1 Or modFlags.SF <> 0 Then
        Debug.Print "FALLO CF suma", "AX=" & AX, "CF=" & modFlags.CF, "ZF=" & modFlags.ZF
        fallos = fallos + 1
    Else
        Debug.Print "CF (FFh+1)", "ok", "esp AX=0 CF=1 ZF=1 SF=0 obt iguales"
    End If

    ' --- Flags: SF negativo 00h-01h ---
    ClearMem
    ResetRegisters
    SetAX 0
    Preparar DecodeByte(OP_SUB_AX_IMM), 1
    PasoExecute
    PasoStore
    If AX <> &HFF Or modFlags.SF <> 1 Or modFlags.CF <> 1 Or modFlags.ZF <> 0 Then
        Debug.Print "FALLO SF", "AX=" & AX, "SF=" & modFlags.SF, "CF=" & modFlags.CF
        fallos = fallos + 1
    Else
        Debug.Print "SF (00h-01h)", "ok", "esp AX=FFh SF=1 CF=1 ZF=0 obt iguales"
    End If

    ' --- INC tope / DEC cero (bordes ALU) ---
    ClearMem
    ResetRegisters
    SetAX &HFF
    Preparar DecodeByte(OP_INC_AX), 0
    PasoExecute
    PasoStore
    If AX <> 0 Or modFlags.CF <> 1 Or modFlags.ZF <> 1 Then
        Debug.Print "FALLO INC FFh", "AX=" & AX, "CF=" & modFlags.CF
        fallos = fallos + 1
    Else
        Debug.Print "INC FFh", "ok", "esp AX=0 CF=1 ZF=1 obt iguales"
    End If

    ClearMem
    ResetRegisters
    SetAX 0
    Preparar DecodeByte(OP_DEC_AX), 0
    PasoExecute
    PasoStore
    If AX <> &HFF Or modFlags.CF <> 1 Or modFlags.SF <> 1 Then
        Debug.Print "FALLO DEC 00h", "AX=" & AX, "CF=" & modFlags.CF, "SF=" & modFlags.SF
        fallos = fallos + 1
    Else
        Debug.Print "DEC 00h", "ok", "esp AX=FFh CF=1 SF=1 obt iguales"
    End If

    ' --- Sin acarreo 01h+01h ---
    ClearMem
    ResetRegisters
    SetAX 1
    Preparar DecodeByte(OP_ADD_AX_IMM), 1
    PasoExecute
    PasoStore
    If AX <> 2 Or modFlags.CF <> 0 Or modFlags.ZF <> 0 Or modFlags.SF <> 0 Then
        Debug.Print "FALLO sin acarreo", "AX=" & AX, "CF=" & modFlags.CF
        fallos = fallos + 1
    Else
        Debug.Print "ADD 01h+01h", "ok", "esp AX=2 CF=0 ZF=0 SF=0 obt iguales"
    End If

    ' --- Segmentacion visual MEMORY (codigo vs datos) ---
    On Error GoTo SinUI
    ClearMem
    WriteMem 0, &H10
    WriteMem &H80, 3
    DoReset
    VistaMemHEX
    RefreshUI
    cCod = Range("rngRAM").Cells(1, 1).Interior.Color
    cDat = Range("rngRAM").Cells(9, 1).Interior.Color
    If cCod = cDat Then
        Debug.Print "FALLO zonas mismo color", "cod=" & cCod, "dat=" & cDat
        fallos = fallos + 1
    Else
        Debug.Print "zonas MEMORY", "ok", "codigo<>datos (colores distintos)"
    End If
    GoTo TrasUI
SinUI:
    Debug.Print "FALLO UI zonas", Err.Number, Err.Description
    fallos = fallos + 1
    On Error GoTo Fallo
TrasUI:
    On Error GoTo Fallo

    If fallos = 0 Then
        Debug.Print "Casos limite 00h/FFh, ZF/CF/SF y segmentacion OK"
    Else
        Debug.Print "FALLOS", fallos
    End If
    Exit Sub
Fallo:
    Debug.Print "FALLO PruebaCasosLimite", Err.Number, Err.Description
End Sub

' Issue 29: JZ/JNZ tomados y no tomados, HLT, direccion/opcode invalidos.
Public Sub PruebaFlujoHltValidaciones()
    Dim fallos As Long
    Dim pc0 As Byte
    Dim ax0 As Byte
    Dim bx0 As Byte
    Dim z0 As Byte
    Dim c0 As Byte
    Dim s0 As Byte
    Dim guardado As Byte
    Dim info As tInstruccion
    Dim huboError As Boolean
    Dim msgErr As String

    On Error GoTo Fallo
    fallos = 0
    Debug.Print "=== Flujo JZ/JNZ, HLT y validaciones ==="

    ' --- JZ tomado (ZF=1) ---
    ClearMem
    ResetRegisters
    SetPC &H20
    UpdateFlags 0, 0
    Preparar DecodeByte(OP_JZ), &H80
    PasoExecute
    PasoStore
    If PC <> &H80 Or modFlags.ZF <> 1 Then
        Debug.Print "FALLO JZ tomado", "esp PC=80h ZF=1", "obt PC=" & Hex$(PC) & "h ZF=" & modFlags.ZF
        fallos = fallos + 1
    Else
        Debug.Print "JZ tomado (ZF=1)", "ok", "esp PC=80h obt PC=80h"
    End If

    ' --- JZ no tomado (ZF=0) ---
    ClearMem
    ResetRegisters
    SetPC &H22
    pc0 = PC
    UpdateFlags 1, 0
    Preparar DecodeByte(OP_JZ), &H90
    PasoExecute
    PasoStore
    If PC <> pc0 Or modFlags.ZF <> 0 Then
        Debug.Print "FALLO JZ no tomado", "esp PC=" & Hex$(pc0) & "h", "obt PC=" & Hex$(PC) & "h"
        fallos = fallos + 1
    Else
        Debug.Print "JZ no tomado (ZF=0)", "ok", "esp PC=" & Hex$(pc0) & "h obt igual"
    End If

    ' --- JNZ tomado (ZF=0) ---
    ClearMem
    ResetRegisters
    SetPC &H30
    UpdateFlags 7, 0
    Preparar DecodeByte(OP_JNZ), &HA0
    PasoExecute
    PasoStore
    If PC <> &HA0 Or modFlags.ZF <> 0 Then
        Debug.Print "FALLO JNZ tomado", "esp PC=A0h", "obt PC=" & Hex$(PC) & "h"
        fallos = fallos + 1
    Else
        Debug.Print "JNZ tomado (ZF=0)", "ok", "esp PC=A0h obt PC=A0h"
    End If

    ' --- JNZ no tomado (ZF=1) ---
    ClearMem
    ResetRegisters
    SetPC &H33
    pc0 = PC
    UpdateFlags 0, 1
    Preparar DecodeByte(OP_JNZ), &HB0
    PasoExecute
    PasoStore
    If PC <> pc0 Or modFlags.ZF <> 1 Then
        Debug.Print "FALLO JNZ no tomado", "esp PC=" & Hex$(pc0) & "h", "obt PC=" & Hex$(PC) & "h"
        fallos = fallos + 1
    Else
        Debug.Print "JNZ no tomado (ZF=1)", "ok", "esp PC=" & Hex$(pc0) & "h obt igual"
    End If

    ' --- HLT: HALTED, registros/flags intactos, Fetch no avanza ---
    ClearMem
    ResetRegisters
    SetAX &H11
    SetBX &H22
    SetPC &H50
    IniciarFetch
    UpdateFlags 3, 1
    ax0 = AX: bx0 = BX: pc0 = PC
    z0 = modFlags.ZF: c0 = modFlags.CF: s0 = modFlags.SF
    Preparar DecodeByte(OP_HLT), 0
    PasoExecute
    PasoStore
    If EstadoCPU <> HALTED Or AX <> ax0 Or BX <> bx0 Or PC <> pc0 _
       Or modFlags.ZF <> z0 Or modFlags.CF <> c0 Or modFlags.SF <> s0 Then
        Debug.Print "FALLO HLT", "esp HALTED AX=11h BX=22h", "obt estado=" & EstadoCPU & " AX=" & Hex$(AX)
        fallos = fallos + 1
    Else
        Debug.Print "HLT detiene", "ok", "esp HALTED obt HALTED, regs/flags intactos"
    End If

    pc0 = PC
    PasoFetch
    If PC <> pc0 Or EstadoCPU <> HALTED Then
        Debug.Print "FALLO HLT frena Fetch", "esp PC intacto HALTED", "obt PC=" & Hex$(PC) & " estado=" & EstadoCPU
        fallos = fallos + 1
    Else
        Debug.Print "HLT frena Fetch", "ok", "esp PC intacto obt igual"
    End If

    ' --- Direccion invalida: error controlado, RAM intacta ---
    ClearMem
    WriteMem 1, &H5A
    guardado = RAM(1)
    huboError = False
    msgErr = ""
    On Error Resume Next
    Err.Clear
    WriteMem -1, 1
    If Err.Number <> 0 Then
        huboError = True
        msgErr = Err.Description
    End If
    On Error GoTo Fallo
    If Not huboError Or RAM(1) <> guardado Then
        Debug.Print "FALLO dir invalida", "esp error + RAM intacta", "obt err=" & huboError & " RAM1=" & RAM(1)
        fallos = fallos + 1
    Else
        Debug.Print "dir invalida (-1)", "ok", "esp error controlado obt " & msgErr
    End If

    huboError = False
    msgErr = ""
    On Error Resume Next
    Err.Clear
    ReadMem 256
    If Err.Number <> 0 Then
        huboError = True
        msgErr = Err.Description
    End If
    On Error GoTo Fallo
    If Not huboError Then
        Debug.Print "FALLO ReadMem 256", "esp error", "obt sin error"
        fallos = fallos + 1
    Else
        Debug.Print "ReadMem 256", "ok", "esp error controlado obt " & msgErr
    End If

    ' --- Opcode invalido: DecodeByte no encontrado; Decode -> HALTED ---
    info = DecodeByte(0)
    If info.Encontrada Or Len(info.Sintaxis) > 0 Then
        Debug.Print "FALLO opcode 00h DecodeByte", "esp Encontrada=False", "obt Encontrada=" & info.Encontrada
        fallos = fallos + 1
    Else
        Debug.Print "opcode 00h DecodeByte", "ok", "esp no encontrado obt Encontrada=False"
    End If

    ClearMem
    ResetRegisters
    IniciarFetch
    SetIR 0
    PasoDecode
    If EstadoCPU <> HALTED Or InstruccionActual.Encontrada Then
        Debug.Print "FALLO opcode 00h Decode", "esp HALTED", "obt estado=" & EstadoCPU & " enc=" & InstruccionActual.Encontrada
        fallos = fallos + 1
    Else
        Debug.Print "opcode 00h Decode", "ok", "esp HALTED obt HALTED (error controlado)"
    End If

    ' Mnemonico desconocido en ensamblado
    huboError = False
    msgErr = ""
    On Error Resume Next
    Err.Clear
    MnemonicToOpcode "NOP"
    If Err.Number <> 0 Then
        huboError = True
        msgErr = Err.Description
    End If
    On Error GoTo Fallo
    If Not huboError Then
        Debug.Print "FALLO mnemonico NOP", "esp error", "obt sin error"
        fallos = fallos + 1
    Else
        Debug.Print "mnemonico NOP", "ok", "esp error controlado obt " & msgErr
    End If

    If fallos = 0 Then
        Debug.Print "Flujo JZ/JNZ, HLT y validaciones OK"
    Else
        Debug.Print "FALLOS", fallos
    End If
    Exit Sub
Fallo:
    Debug.Print "FALLO PruebaFlujoHltValidaciones", Err.Number, Err.Description
End Sub

Public Sub DumpAXToText()
    Dim i As Long
    Dim f As Integer
    Dim lineas() As String
    
    f = FreeFile
    Open "C:to Semestre\parcial1-SIS131\scratchx_dump.txt" For Output As #f
    
    lineas = LineasDemoMultiplicacion()
    CargarPrograma lineas
    SembrarDatosDemo
    DoReset
    
    For i = 1 To 60
        DoStep
        Print #f, "Paso " & i & " - Fase " & FaseActual & " - AX: " & AX & " BX: " & BX & " Temp: " & Temporal & " PC: " & PC
    Next i
    
    Close #f
End Sub
