Attribute VB_Name = "modUI"
Option Explicit

' Pinta la hoja CPU/MEMORY desde el modelo. No calcula ni avanza el reloj.
' Tras cada DoStep: RefreshUI
' En la ventana Inmediato: PruebaUI

Private mUltimaCeldaMAR As Range

Public Sub RefreshUI()
    On Error GoTo Fallo
    AsegurarNombresUI
    PintarEstadoYFase
    PintarRegistrosYBanderas
    PintarInstruccionYMicro
    PintarCaminoDatos
    PintarCeldaMAR
    DoEvents
    Exit Sub
Fallo:
    Debug.Print "RefreshUI", Err.Number, Err.Description
End Sub

Public Sub PruebaUI()
    On Error GoTo Fallo
    ClearMem
    WriteMem 0, OP_MOV_AX_IMM
    WriteMem 1, &H5
    DoReset
    RefreshUI
    Debug.Print "UI tras RESET", Range("rngPhase").Value, Range("rngMicroOp").Value

    DoStep
    Debug.Print "UI tras F1", Range("rngMicroOp").Value, "MAR=" & Range("rngMAR").Value

    DoStep
    DoStep
    DoStep
    DoStep
    Debug.Print "UI tras Decode", Range("rngPhase").Value, Range("rngInstruccion").Value

    Debug.Print "RefreshUI pinta fase, flujo e instruccion"
    Exit Sub
Fallo:
    Debug.Print "FALLO PruebaUI", Err.Number, Err.Description
End Sub

Private Sub AsegurarNombresUI()
    Dim n As Name
    On Error Resume Next
    Err.Clear
    Set n = ThisWorkbook.Names("rngFlujo")
    If Err.Number <> 0 Then
        Err.Clear
        ThisWorkbook.Names.Add Name:="rngFlujo", RefersTo:="=CPU!$C$6"
    End If
    On Error GoTo 0
End Sub

Private Sub PintarEstadoYFase()
    Range("rngState").Value = NombreEstado(EstadoCPU)
    Range("rngPhase").Value = NombreFase(FaseActual)

    Select Case FaseActual
        Case FETCH: Range("rngPhase").Interior.Color = RGB(181, 214, 242)
        Case DECODE: Range("rngPhase").Interior.Color = RGB(245, 225, 181)
        Case EXECUTE: Range("rngPhase").Interior.Color = RGB(198, 242, 198)
        Case STORE: Range("rngPhase").Interior.Color = RGB(242, 198, 225)
        Case Else: Range("rngPhase").Interior.ColorIndex = xlNone
    End Select
End Sub

Private Sub PintarRegistrosYBanderas()
    Range("rngPC").Value = Hex2(PC)
    Range("rngIR").Value = Hex2(IR)
    Range("rngMAR").Value = Hex2(MAR)
    Range("rngMDR").Value = Hex2(MDR)
    Range("rngAX").Value = Hex2(AX)
    Range("rngBX").Value = Hex2(BX)
    Range("rngZF").Value = modFlags.ZF
    Range("rngCF").Value = modFlags.CF
    Range("rngSF").Value = modFlags.SF
End Sub

Private Sub PintarInstruccionYMicro()
    Dim texto As String

    If InstruccionActual.Encontrada Then
        texto = InstruccionActual.Sintaxis
        If InstruccionActual.Bytes = 2 And TieneOperando Then
            texto = texto & "  [" & Hex2(OperandoActual) & "h]"
        End If
    Else
        texto = "-"
    End If
    Range("rngInstruccion").Value = texto
    Range("rngMicroOp").Value = UltimaMicroOp
End Sub

Private Sub PintarCaminoDatos()
    Dim flu As String
    Dim op As String

    LimpiarResaltadoCamino
    op = UltimaMicroOp
    flu = ""

    If FaseActual = FETCH Or Left$(op, 1) = "F" Then
        If InStr(1, op, "F1", vbTextCompare) > 0 Then
            flu = "PC  ->  MAR"
            ResaltarRango "rngPC"
            ResaltarRango "rngMAR"
        ElseIf InStr(1, op, "F2", vbTextCompare) > 0 Then
            flu = "RAM[MAR]  ->  MDR"
            ResaltarRango "rngMAR"
            ResaltarRango "rngMDR"
        ElseIf InStr(1, op, "F3", vbTextCompare) > 0 Then
            flu = "MDR  ->  IR"
            ResaltarRango "rngMDR"
            ResaltarRango "rngIR"
        ElseIf InStr(1, op, "F4", vbTextCompare) > 0 Then
            flu = "PC  <-  PC+1"
            ResaltarRango "rngPC"
        End If
    ElseIf FaseActual = DECODE Then
        flu = "IR  ->  Decode"
        ResaltarRango "rngIR"
    ElseIf FaseActual = EXECUTE Then
        flu = "Execute"
        ResaltarRango "rngAX"
        ResaltarRango "rngBX"
    ElseIf FaseActual = STORE Then
        flu = "Store / write-back"
        ResaltarRango "rngAX"
        ResaltarRango "rngBX"
        ResaltarRango "rngMDR"
    End If

    On Error Resume Next
    Range("rngFlujo").Value = flu
    On Error GoTo 0
End Sub

Private Sub PintarCeldaMAR()
    Dim celda As Range
    Dim addr As Long

    On Error Resume Next
    If Not mUltimaCeldaMAR Is Nothing Then
        mUltimaCeldaMAR.Interior.ColorIndex = xlNone
    End If
    On Error GoTo 0

    addr = MAR
    If addr < 0 Or addr > 255 Then Exit Sub

    Set celda = Range("rngRAM").Cells((addr \ 16) + 1, (addr Mod 16) + 1)
    If InStr(1, UltimaMicroOp, "F2", vbTextCompare) > 0 _
       Or InStr(1, UltimaMicroOp, "LOAD", vbTextCompare) > 0 _
       Or InStr(1, UltimaMicroOp, "STORE", vbTextCompare) > 0 _
       Or FaseActual = STORE Then
        celda.Interior.Color = RGB(102, 255, 102)
        Set mUltimaCeldaMAR = celda
    End If

    On Error Resume Next
    celda.Value = Hex2(ReadMem(addr))
    On Error GoTo 0
End Sub

Private Sub LimpiarResaltadoCamino()
    On Error Resume Next
    Range("rngPC").Interior.ColorIndex = xlNone
    Range("rngIR").Interior.ColorIndex = xlNone
    Range("rngMAR").Interior.ColorIndex = xlNone
    Range("rngMDR").Interior.ColorIndex = xlNone
    Range("rngAX").Interior.ColorIndex = xlNone
    Range("rngBX").Interior.ColorIndex = xlNone
    On Error GoTo 0
End Sub

Private Sub ResaltarRango(ByVal nombre As String)
    On Error Resume Next
    Range(nombre).Interior.Color = RGB(255, 205, 127)
    On Error GoTo 0
End Sub

Private Function Hex2(ByVal b As Long) As String
    Hex2 = Right$("0" & Hex$(b And &HFF), 2)
End Function

Private Function NombreEstado(ByVal e As eCPUState) As String
    Select Case e
        Case RUNNING: NombreEstado = "RUNNING"
        Case PAUSED: NombreEstado = "PAUSED"
        Case HALTED: NombreEstado = "HALTED"
        Case RESET: NombreEstado = "RESET"
        Case Else: NombreEstado = "?"
    End Select
End Function

Private Function NombreFase(ByVal f As ePhase) As String
    Select Case f
        Case FETCH: NombreFase = "FETCH"
        Case DECODE: NombreFase = "DECODE"
        Case EXECUTE: NombreFase = "EXECUTE"
        Case STORE: NombreFase = "STORE"
        Case Else: NombreFase = "?"
    End Select
End Function
