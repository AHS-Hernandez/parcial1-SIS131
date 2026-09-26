Attribute VB_Name = "modPruebas"
Option Explicit

' Prueba aislada de las 12 familias ISA + programa corto en STEP y RUN hasta HLT.
' En la ventana Inmediato: PruebaISACompleta
' Notas: docs/PRUEBAS.md

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

    ' RUN (TickRun) hasta HLT — mismo programa
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
