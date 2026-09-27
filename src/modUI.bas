Attribute VB_Name = "modUI"
Option Explicit

' Pinta CPU y MEMORY desde el modelo. No calcula ni avanza el reloj.
' Vistas de memoria: VistaMemHEX / VistaMemBIN / VistaMemDEC / CiclarVistaMem
' Tras cada DoStep: RefreshUI
' PulirInterfaz: layout profesional (bloques, colores, tipografia, diagramas)
' En la ventana Inmediato: PruebaUI, PruebaVistaMemoria, PulirInterfaz

Private Const VISTA_HEX As Integer = 0
Private Const VISTA_BIN As Integer = 1
Private Const VISTA_DEC As Integer = 2

' Colores de marca (herramienta educativa)
Private Const COL_FETCH As Long = &HF2D6B5      ' RGB(181,214,242) BGR
Private Const COL_DECODE As Long = &HB5E1F5     ' RGB(245,225,181)
Private Const COL_EXECUTE As Long = &HC6F2C6    ' RGB(198,242,198)
Private Const COL_STORE As Long = &HE1C6F2      ' RGB(242,198,225)
Private Const COL_ACTIVO As Long = &H7FCDFF     ' RGB(255,205,127)
Private Const COL_MAR As Long = &H66FF66        ' RGB(102,255,102)
Private Const COL_PC_MEM As Long = &HFFE066     ' RGB(102,224,255)
Private Const COL_CODIGO As Long = &HFFE6D2     ' RGB(210,230,255)
Private Const COL_DATOS As Long = &HD2ECFF      ' RGB(255,236,210)
Private Const COL_TITULO As Long = &H5A3A1A     ' RGB(26,58,90)
Private Const COL_HEADER As Long = &H8B5E2E     ' RGB(46,94,139)
Private Const COL_BOTON As Long = &HD9A06B      ' RGB(107,160,217)
Private Const COL_FLAG_ON As Long = &H5AD25A    ' RGB(90,210,90)
Private Const COL_FLAG_OFF As Long = &HE0E0E0   ' RGB(224,224,224)
Private Const COL_HALTED As Long = &H6B6BFF     ' RGB(255,107,107)
Private Const COL_RUNNING As Long = &H7AD97A    ' RGB(122,217,122)
Private Const COL_PAUSED As Long = &H6BD4FF     ' RGB(255,212,107)
Private Const COL_RESET As Long = &HC0C0C0       ' RGB(192,192,192)

Private mUltimaCeldaMAR As Range
Private mUltimaCeldaPC As Range
Private mVistaMem As Integer
Private mLayoutListo As Boolean

Public Sub RefreshUI()
    On Error GoTo Fallo
    AsegurarNombresUI
    AsegurarEtiquetasPipeline
    PintarEstadoYFase
    PintarPipeline
    PintarRegistrosYBanderas
    PintarBanderasLED
    PintarInstruccionYMicro
    PintarCaminoDatos
    PintarPasoYReloj
    PintarMemoria
    PintarCeldaPC
    PintarCeldaMAR
    PintarShapesPipeline
    DoEvents
    Exit Sub
Fallo:
    Debug.Print "RefreshUI", Err.Number, Err.Description
End Sub

' Layout completo para defensa: tipografia, bloques, botones, leyendas y diagramas.
Public Sub PulirInterfaz()
    On Error GoTo Fallo
    Application.ScreenUpdating = False
    PulirHojaCPU
    PulirHojaMEMORY
    PulirHojaPROGRAM
    PulirHojaLOG
    PulirHojaISA
    PulirHojaREADME
    InsertarDiagramasSiExisten
    mLayoutListo = True
    Application.ScreenUpdating = True
    RefreshUI
    Debug.Print "PulirInterfaz", "ok", "CPU/MEMORY/PROGRAM/LOG/ISA/README"
    Exit Sub
Fallo:
    Application.ScreenUpdating = True
    Debug.Print "FALLO PulirInterfaz", Err.Number, Err.Description
End Sub

Public Sub VistaMemHEX()
    mVistaMem = VISTA_HEX
    RefreshUI
End Sub

Public Sub VistaMemBIN()
    mVistaMem = VISTA_BIN
    RefreshUI
End Sub

Public Sub VistaMemDEC()
    mVistaMem = VISTA_DEC
    RefreshUI
End Sub

' Alterna HEX -> BIN -> DEC/mnem -> HEX
Public Sub CiclarVistaMem()
    mVistaMem = (mVistaMem + 1) Mod 3
    RefreshUI
    Debug.Print "vista memoria", NombreVistaMem()
End Sub

Public Property Get VistaMemoriaActual() As String
    VistaMemoriaActual = NombreVistaMem()
End Property

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

Public Sub PruebaVistaMemoria()
    Dim fallos As Long
    Dim c00 As String
    Dim c80 As String

    On Error GoTo Fallo
    fallos = 0

    ClearMem
    WriteMem 0, OP_MOV_AX_IMM
    WriteMem 1, &H5
    WriteMem &H80, 3
    WriteMem &H81, 4
    DoReset
    mVistaMem = VISTA_HEX
    RefreshUI

    c00 = CStr(Range("rngRAM").Cells(1, 1).Value)
    c80 = CStr(Range("rngRAM").Cells(9, 1).Value)
    If UCase$(Right$("0" & c00, 2)) <> "10" Then
        Debug.Print "FALLO HEX codigo", c00
        fallos = fallos + 1
    ElseIf UCase$(Right$("0" & c80, 2)) <> "03" Then
        Debug.Print "FALLO HEX datos", c80
        fallos = fallos + 1
    Else
        Debug.Print "HEX", "ok", "00h=" & c00, "80h=" & c80
    End If

    ' Colores de zona: fila 0 (codigo) != fila 8 (datos)
    If Range("rngRAM").Cells(1, 1).Interior.Color = Range("rngRAM").Cells(9, 1).Interior.Color Then
        Debug.Print "FALLO colores codigo/datos iguales"
        fallos = fallos + 1
    Else
        Debug.Print "segmentacion", "ok", "codigo <> datos"
    End If

    VistaMemBIN
    c00 = CStr(Range("rngRAM").Cells(1, 1).Value)
    If Len(c00) <> 8 Or InStr(c00, "0") = 0 Then
        Debug.Print "FALLO BIN", c00
        fallos = fallos + 1
    Else
        Debug.Print "BIN", "ok", c00
    End If

    VistaMemDEC
    c00 = CStr(Range("rngRAM").Cells(1, 1).Value)
    c80 = CStr(Range("rngRAM").Cells(9, 1).Value)
    If InStr(1, c00, "MOV", vbTextCompare) = 0 Then
        Debug.Print "FALLO DEC/mnem codigo", c00
        fallos = fallos + 1
    Else
        Debug.Print "DEC/mnem codigo", "ok", c00
    End If
    If c80 <> "3" Then
        Debug.Print "FALLO DEC datos", c80
        fallos = fallos + 1
    Else
        Debug.Print "DEC datos", "ok", c80
    End If

    ' Operando del MOV en 01h debe verse como inmediato, no como mnemonic
    If InStr(1, CStr(Range("rngRAM").Cells(1, 2).Value), "05", vbTextCompare) = 0 _
       And CStr(Range("rngRAM").Cells(1, 2).Value) <> "5" Then
        Debug.Print "FALLO operando mnem", Range("rngRAM").Cells(1, 2).Value
        fallos = fallos + 1
    Else
        Debug.Print "operando en vista mnem", "ok", Range("rngRAM").Cells(1, 2).Value
    End If

    VistaMemHEX
    If fallos = 0 Then
        Debug.Print "MEMORY: segmentacion y vistas HEX/BIN/DEC-mnem OK"
    Else
        Debug.Print "FALLOS", fallos
    End If
    Exit Sub
Fallo:
    Debug.Print "FALLO PruebaVistaMemoria", Err.Number, Err.Description
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
    Err.Clear
    Set n = ThisWorkbook.Names("rngVistaMem")
    If Err.Number <> 0 Then
        Err.Clear
        ThisWorkbook.Names.Add Name:="rngVistaMem", RefersTo:="=MEMORY!$D$22"
    End If
    Err.Clear
    Set n = ThisWorkbook.Names("rngPaso")
    If Err.Number <> 0 Then
        Err.Clear
        ThisWorkbook.Names.Add Name:="rngPaso", RefersTo:="=CPU!$J$4"
    End If
    Err.Clear
    Set n = ThisWorkbook.Names("rngPipeFetch")
    If Err.Number <> 0 Then
        Err.Clear
        ThisWorkbook.Names.Add Name:="rngPipeFetch", RefersTo:="=CPU!$B$3"
        ThisWorkbook.Names.Add Name:="rngPipeDecode", RefersTo:="=CPU!$C$3"
        ThisWorkbook.Names.Add Name:="rngPipeExecute", RefersTo:="=CPU!$D$3"
        ThisWorkbook.Names.Add Name:="rngPipeStore", RefersTo:="=CPU!$E$3"
    End If
    On Error GoTo 0
End Sub

Private Sub AsegurarEtiquetasPipeline()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("CPU")
    If ws Is Nothing Then Exit Sub
    If Len(Trim$(CStr(ws.Range("A3").Value & ""))) = 0 Then
        ws.Range("A3").Value = "Ciclo"
    End If
    If Len(Trim$(CStr(ws.Range("B3").Value & ""))) = 0 Then ws.Range("B3").Value = "FETCH"
    If Len(Trim$(CStr(ws.Range("C3").Value & ""))) = 0 Then ws.Range("C3").Value = "DECODE"
    If Len(Trim$(CStr(ws.Range("D3").Value & ""))) = 0 Then ws.Range("D3").Value = "EXECUTE"
    If Len(Trim$(CStr(ws.Range("E3").Value & ""))) = 0 Then ws.Range("E3").Value = "STORE"
    If Len(Trim$(CStr(ws.Range("I4").Value & ""))) = 0 Then ws.Range("I4").Value = "Paso #"
    On Error GoTo 0
End Sub

Private Sub PintarEstadoYFase()
    Range("rngState").Value = NombreEstado(EstadoCPU)
    Range("rngPhase").Value = NombreFase(FaseActual)

    Select Case EstadoCPU
        Case RUNNING: Range("rngState").Interior.Color = COL_RUNNING
        Case PAUSED: Range("rngState").Interior.Color = COL_PAUSED
        Case HALTED: Range("rngState").Interior.Color = COL_HALTED
        Case RESET: Range("rngState").Interior.Color = COL_RESET
        Case Else: Range("rngState").Interior.ColorIndex = xlNone
    End Select

    Select Case FaseActual
        Case FETCH: Range("rngPhase").Interior.Color = COL_FETCH
        Case DECODE: Range("rngPhase").Interior.Color = COL_DECODE
        Case EXECUTE: Range("rngPhase").Interior.Color = COL_EXECUTE
        Case STORE: Range("rngPhase").Interior.Color = COL_STORE
        Case Else: Range("rngPhase").Interior.ColorIndex = xlNone
    End Select
End Sub

Private Sub PintarPipeline()
    On Error Resume Next
    Range("rngPipeFetch").Value = "FETCH"
    Range("rngPipeDecode").Value = "DECODE"
    Range("rngPipeExecute").Value = "EXECUTE"
    Range("rngPipeStore").Value = "STORE"

    Range("rngPipeFetch").Interior.Color = RGB(230, 230, 230)
    Range("rngPipeDecode").Interior.Color = RGB(230, 230, 230)
    Range("rngPipeExecute").Interior.Color = RGB(230, 230, 230)
    Range("rngPipeStore").Interior.Color = RGB(230, 230, 230)
    Range("rngPipeFetch").Font.Bold = False
    Range("rngPipeDecode").Font.Bold = False
    Range("rngPipeExecute").Font.Bold = False
    Range("rngPipeStore").Font.Bold = False

    Select Case FaseActual
        Case FETCH
            Range("rngPipeFetch").Interior.Color = COL_FETCH
            Range("rngPipeFetch").Font.Bold = True
        Case DECODE
            Range("rngPipeDecode").Interior.Color = COL_DECODE
            Range("rngPipeDecode").Font.Bold = True
        Case EXECUTE
            Range("rngPipeExecute").Interior.Color = COL_EXECUTE
            Range("rngPipeExecute").Font.Bold = True
        Case STORE
            Range("rngPipeStore").Interior.Color = COL_STORE
            Range("rngPipeStore").Font.Bold = True
    End Select
    On Error GoTo 0
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

Private Sub PintarBanderasLED()
    On Error Resume Next
    If modFlags.ZF <> 0 Then
        Range("rngZF").Interior.Color = COL_FLAG_ON
        Range("rngZF").Font.Bold = True
    Else
        Range("rngZF").Interior.Color = COL_FLAG_OFF
        Range("rngZF").Font.Bold = False
    End If
    If modFlags.CF <> 0 Then
        Range("rngCF").Interior.Color = COL_FLAG_ON
        Range("rngCF").Font.Bold = True
    Else
        Range("rngCF").Interior.Color = COL_FLAG_OFF
        Range("rngCF").Font.Bold = False
    End If
    If modFlags.SF <> 0 Then
        Range("rngSF").Interior.Color = COL_FLAG_ON
        Range("rngSF").Font.Bold = True
    Else
        Range("rngSF").Interior.Color = COL_FLAG_OFF
        Range("rngSF").Font.Bold = False
    End If
    On Error GoTo 0
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
    Range("rngMicroOp").Value = modControlUnit.UltimaMicroOp
End Sub

Private Sub PintarCaminoDatos()
    Dim flu As String
    Dim op As String

    LimpiarResaltadoCamino
    op = modControlUnit.UltimaMicroOp
    flu = ""

    If FaseActual = FETCH Or Left$(op, 1) = "F" Then
        If InStr(1, op, "F1", vbTextCompare) > 0 Then
            flu = "PC  ===>  MAR"
            ResaltarRango "rngPC"
            ResaltarRango "rngMAR"
        ElseIf InStr(1, op, "F2", vbTextCompare) > 0 Then
            flu = "RAM[MAR]  ===>  MDR"
            ResaltarRango "rngMAR"
            ResaltarRango "rngMDR"
        ElseIf InStr(1, op, "F3", vbTextCompare) > 0 Then
            flu = "MDR  ===>  IR"
            ResaltarRango "rngMDR"
            ResaltarRango "rngIR"
        ElseIf InStr(1, op, "F4", vbTextCompare) > 0 Then
            flu = "PC  <===  PC+1"
            ResaltarRango "rngPC"
        End If
    ElseIf FaseActual = DECODE Then
        flu = "IR  ===>  Decode (ISA)"
        ResaltarRango "rngIR"
    ElseIf FaseActual = EXECUTE Then
        flu = "Execute / ALU / saltos"
        ResaltarRango "rngAX"
        ResaltarRango "rngBX"
        If InStr(1, op, "LOAD", vbTextCompare) > 0 Or InStr(1, op, "STORE", vbTextCompare) > 0 Then
            ResaltarRango "rngMAR"
            ResaltarRango "rngMDR"
        End If
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

Private Sub PintarPasoYReloj()
    On Error Resume Next
    Range("rngPaso").Value = PasoGlobal
    Range("rngPaso").Font.Name = "Consolas"
    Range("rngPaso").Font.Bold = True
    Range("rngPaso").Font.Size = 14
    On Error GoTo 0
End Sub

' Grid 16x16: color codigo (00h-7Fh) vs datos (80h-FFh); formato segun mVistaMem.
Private Sub PintarMemoria()
    Dim addr As Long
    Dim celda As Range
    Dim b As Byte
    Dim esOperando(0 To 255) As Boolean
    Dim i As Long
    Dim info As tInstruccion
    Dim texto As String
    Dim colorCodigo As Long
    Dim colorDatos As Long

    Application.ScreenUpdating = False
    colorCodigo = COL_CODIGO
    colorDatos = COL_DATOS

    ' Marca segundos bytes de instrucciones de 2 bytes en zona codigo
    i = 0
    Do While i <= &H7F
        info = DecodeByte(ReadMem(i))
        If info.Encontrada And info.Bytes = 2 And i < &H7F Then
            esOperando(i + 1) = True
            i = i + 2
        Else
            i = i + 1
        End If
    Loop

    For addr = 0 To 255
        Set celda = Range("rngRAM").Cells((addr \ 16) + 1, (addr Mod 16) + 1)
        b = ReadMem(addr)

        If addr <= &H7F Then
            celda.Interior.Color = colorCodigo
        Else
            celda.Interior.Color = colorDatos
        End If

        Select Case mVistaMem
            Case VISTA_HEX
                texto = Hex2(b)
            Case VISTA_BIN
                texto = Bin8(b)
            Case Else
                If addr <= &H7F Then
                    If esOperando(addr) Then
                        texto = Hex2(b) & "h"
                    Else
                        info = DecodeByte(b)
                        If info.Encontrada Then
                            texto = info.Mnemonic
                        Else
                            texto = Hex2(b)
                        End If
                    End If
                Else
                    texto = CStr(b)
                End If
        End Select
        celda.Value = texto
    Next addr

    ' Leyenda en MEMORY
    On Error Resume Next
    Range("MEMORY!A22").Value = "Codigo"
    Range("MEMORY!B22").Value = "00h-7Fh"
    Range("MEMORY!A22").Interior.Color = colorCodigo
    Range("MEMORY!A23").Value = "Datos"
    Range("MEMORY!B23").Value = "80h-FFh"
    Range("MEMORY!A23").Interior.Color = colorDatos
    Range("MEMORY!C22").Value = "Vista:"
    Range("rngVistaMem").Value = NombreVistaMem()
    Range("MEMORY!E22").Value = "HEX"
    Range("MEMORY!F22").Value = "BIN"
    Range("MEMORY!G22").Value = "DEC"
    Range("MEMORY!E22").Interior.Color = IIf(mVistaMem = VISTA_HEX, COL_ACTIVO, COL_BOTON)
    Range("MEMORY!F22").Interior.Color = IIf(mVistaMem = VISTA_BIN, COL_ACTIVO, COL_BOTON)
    Range("MEMORY!G22").Interior.Color = IIf(mVistaMem = VISTA_DEC, COL_ACTIVO, COL_BOTON)
    Range("MEMORY!C23").Value = "PC=cyan  MAR=verde"
    Range("MEMORY!E23").Value = "CiclarVistaMem"
    On Error GoTo 0

    Application.ScreenUpdating = True
    Set mUltimaCeldaMAR = Nothing
    Set mUltimaCeldaPC = Nothing
End Sub

Private Sub PintarCeldaPC()
    Dim celda As Range
    Dim addr As Long

    addr = PC
    If addr < 0 Or addr > 255 Then Exit Sub
    Set celda = Range("rngRAM").Cells((addr \ 16) + 1, (addr Mod 16) + 1)
    ' No pisa el resaltado MAR si coinciden; MAR gana despues
    If celda.Interior.Color <> COL_MAR Then
        celda.Interior.Color = COL_PC_MEM
    End If
    Set mUltimaCeldaPC = celda
End Sub

Private Sub PintarCeldaMAR()
    Dim celda As Range
    Dim addr As Long

    addr = MAR
    If addr < 0 Or addr > 255 Then Exit Sub

    Set celda = Range("rngRAM").Cells((addr \ 16) + 1, (addr Mod 16) + 1)
    If InStr(1, modControlUnit.UltimaMicroOp, "F2", vbTextCompare) > 0 _
       Or InStr(1, modControlUnit.UltimaMicroOp, "LOAD", vbTextCompare) > 0 _
       Or InStr(1, modControlUnit.UltimaMicroOp, "STORE", vbTextCompare) > 0 _
       Or InStr(1, modControlUnit.UltimaMicroOp, "F1", vbTextCompare) > 0 _
       Or FaseActual = STORE Then
        celda.Interior.Color = COL_MAR
        Set mUltimaCeldaMAR = celda
    End If
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
    Range(nombre).Interior.Color = COL_ACTIVO
    On Error GoTo 0
End Sub

' --- Pulido visual de hojas (una vez / defensa) ---

Private Sub PulirHojaCPU()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("CPU")

    ws.Range("A1").Value = "Simulador de CPU 8 bits"
    ws.Range("A1").Font.Name = "Calibri"
    ws.Range("A1").Font.Size = 20
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Color = COL_TITULO

    ws.Range("A2").Value = "Von Neumann: Fetch - Decode - Execute - Store | STEP/RUN iluminan el camino"
    ws.Range("A2").Font.Size = 10
    ws.Range("A2").Font.Italic = True
    ws.Range("A2").Font.Color = RGB(89, 89, 89)

    ws.Range("A3").Value = "Ciclo"
    ws.Range("A3").Font.Bold = True
    ws.Range("B3").Value = "FETCH"
    ws.Range("C3").Value = "DECODE"
    ws.Range("D3").Value = "EXECUTE"
    ws.Range("E3").Value = "STORE"
    ws.Range("B3:E3").Font.Name = "Calibri"
    ws.Range("B3:E3").Font.Size = 11
    ws.Range("B3:E3").HorizontalAlignment = xlCenter
    ws.Range("B3:E3").Borders.LineStyle = xlContinuous

    ws.Range("A4").Value = "Estado"
    ws.Range("A4").Font.Bold = True
    ws.Range("D4").Value = "Fase"
    ws.Range("D4").Font.Bold = True
    ws.Range("I4").Value = "Paso #"
    ws.Range("I4").Font.Bold = True
    ws.Range("rngState").Font.Name = "Consolas"
    ws.Range("rngState").Font.Size = 12
    ws.Range("rngState").Font.Bold = True
    ws.Range("rngPhase").Font.Name = "Consolas"
    ws.Range("rngPhase").Font.Size = 12
    ws.Range("rngPhase").Font.Bold = True

    EstiloBloqueTitulo ws.Range("A6"), "Camino de datos"
    EstiloBloqueTitulo ws.Range("E6"), "Registros"
    EstiloBloqueTitulo ws.Range("I6"), "ALU / control"

    EstiloCeldaRegistro ws.Range("rngPC")
    EstiloCeldaRegistro ws.Range("rngIR")
    EstiloCeldaRegistro ws.Range("rngMAR")
    EstiloCeldaRegistro ws.Range("rngMDR")
    EstiloCeldaRegistro ws.Range("rngAX")
    EstiloCeldaRegistro ws.Range("rngBX")

    ws.Range("A7").Value = "PC"
    ws.Range("A8").Value = "IR"
    ws.Range("A9").Value = "MAR"
    ws.Range("A10").Value = "MDR"
    ws.Range("A7:A10").Font.Bold = True
    ws.Range("C7").Value = "Siguiente instruccion"
    ws.Range("C8").Value = "Opcode en curso"
    ws.Range("C9").Value = "Direccion hacia RAM"
    ws.Range("C10").Value = "Dato leido / a escribir"
    ws.Range("C7:C10").Font.Size = 9
    ws.Range("C7:C10").Font.Color = RGB(89, 89, 89)

    ws.Range("E7").Value = "AX"
    ws.Range("E8").Value = "BX"
    ws.Range("E7:E8").Font.Bold = True
    ws.Range("G7").Value = "Acumulador"
    ws.Range("G8").Value = "Proposito general"
    ws.Range("G7:G8").Font.Size = 9
    ws.Range("G7:G8").Font.Color = RGB(89, 89, 89)

    EstiloBloqueTitulo ws.Range("A12"), "Banderas"
    EstiloBloqueTitulo ws.Range("E12"), "Instruccion"
    ws.Range("A13").Value = "ZF"
    ws.Range("A14").Value = "CF"
    ws.Range("A15").Value = "SF"
    ws.Range("A13:A15").Font.Bold = True
    ws.Range("E13").Value = "Micro-operacion"
    ws.Range("E13").Font.Bold = True
    ws.Range("rngInstruccion").Font.Name = "Consolas"
    ws.Range("rngInstruccion").Font.Size = 12
    ws.Range("rngInstruccion").Font.Bold = True
    ws.Range("rngMicroOp").Font.Name = "Consolas"
    ws.Range("rngMicroOp").Font.Size = 10

    EstiloBloqueTitulo ws.Range("A17"), "Controles"
    ws.Range("G17").Value = "Retardo (ms)"
    ws.Range("G17").Font.Bold = True
    EstiloBotonCelda ws.Range("A18"), "STEP"
    EstiloBotonCelda ws.Range("B18"), "RUN"
    EstiloBotonCelda ws.Range("C18"), "PAUSE"
    EstiloBotonCelda ws.Range("D18"), "RESET"
    EstiloBotonCelda ws.Range("E18"), "LOAD"
    ws.Range("rngDelay").Font.Name = "Consolas"
    ws.Range("rngDelay").Font.Size = 12
    If Len(Trim$(CStr(ws.Range("rngDelay").Value & ""))) = 0 Then
        ws.Range("rngDelay").Value = 200
    End If

    ws.Columns("A").ColumnWidth = 12
    ws.Columns("B").ColumnWidth = 10
    ws.Columns("C").ColumnWidth = 28
    ws.Columns("D").ColumnWidth = 8
    ws.Columns("E").ColumnWidth = 12
    ws.Columns("F").ColumnWidth = 14
    ws.Columns("G").ColumnWidth = 18
    ws.Columns("H").ColumnWidth = 3
    ws.Columns("I").ColumnWidth = 12
    ws.Columns("J").ColumnWidth = 10
    ws.Rows("1").RowHeight = 28
    ws.Rows("3").RowHeight = 22
    ws.Rows("18").RowHeight = 24

    AsegurarShapesPipeline ws
End Sub

Private Sub PulirHojaMEMORY()
    Dim ws As Worksheet
    Dim c As Long
    Set ws = ThisWorkbook.Worksheets("MEMORY")

    ws.Range("A1").Value = "Memoria principal  00h - FFh"
    ws.Range("A1").Font.Size = 18
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Color = COL_TITULO

    ws.Range("A2").Value = "256 celdas x 8 bits. Dir = fila*16+col. Azul=codigo 00h-7Fh | Naranja=datos 80h-FFh | Cyan=PC | Verde=MAR"
    ws.Range("A2").Font.Size = 9
    ws.Range("A2").Font.Italic = True
    ws.Range("A2").Font.Color = RGB(89, 89, 89)

    ws.Range("A4").Value = ""
    For c = 0 To 15
        ws.Cells(4, c + 2).Value = Hex$(c)
        ws.Cells(4, c + 2).Font.Bold = True
        ws.Cells(4, c + 2).Font.Name = "Consolas"
        ws.Cells(4, c + 2).HorizontalAlignment = xlCenter
        ws.Cells(4, c + 2).Interior.Color = COL_HEADER
        ws.Cells(4, c + 2).Font.Color = RGB(255, 255, 255)
    Next c

    For c = 0 To 15
        ws.Cells(5 + c, 1).Value = Hex$(c)
        ws.Cells(5 + c, 1).Font.Bold = True
        ws.Cells(5 + c, 1).Font.Name = "Consolas"
        ws.Cells(5 + c, 1).HorizontalAlignment = xlCenter
        If c <= 7 Then
            ws.Cells(5 + c, 1).Interior.Color = COL_CODIGO
        Else
            ws.Cells(5 + c, 1).Interior.Color = COL_DATOS
        End If
    Next c

    ws.Range("rngRAM").Font.Name = "Consolas"
    ws.Range("rngRAM").Font.Size = 9
    ws.Range("rngRAM").HorizontalAlignment = xlCenter
    ws.Range("rngRAM").Borders.LineStyle = xlContinuous
    ws.Range("rngRAM").Borders.Color = RGB(180, 180, 180)

    EstiloBotonCelda ws.Range("E22"), "HEX"
    EstiloBotonCelda ws.Range("F22"), "BIN"
    EstiloBotonCelda ws.Range("G22"), "DEC"
    ws.Range("A22").Font.Bold = True
    ws.Range("A23").Font.Bold = True
    ws.Range("C22").Font.Bold = True
    ws.Range("rngVistaMem").Font.Name = "Consolas"
    ws.Range("rngVistaMem").Font.Bold = True

    ws.Columns("A").ColumnWidth = 6
    Dim col As Long
    For col = 2 To 17
        ws.Columns(col).ColumnWidth = 5.5
    Next col
    ws.Rows("1").RowHeight = 26
End Sub

Private Sub PulirHojaPROGRAM()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("PROGRAM")
    ws.Range("A1").Value = "PROGRAM - demo: multiplicacion por sumas (N*[80h] x M*[81h] -> [82h])"
    ws.Range("A1").Font.Size = 14
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Color = COL_TITULO
    ws.Range("A1").Interior.Color = RGB(232, 240, 254)
    On Error Resume Next
    Range("rngProgram").Font.Name = "Consolas"
    Range("rngProgram").Font.Size = 12
    Range("rngProgram").Borders.LineStyle = xlContinuous
    On Error GoTo 0
    ws.Columns("A").ColumnWidth = 72
    ws.Range("B1").Value = "LOAD escribe 00h..; datos 80h-82h se siembran aparte"
    ws.Range("B1").Font.Size = 9
    ws.Range("B1").Font.Italic = True
End Sub

Private Sub PulirHojaLOG()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("LOG")
    On Error Resume Next
    ws.Cells(1, 1).Value = "Paso"
    ws.Cells(1, 2).Value = "Fase"
    ws.Cells(1, 3).Value = "Micro-operacion"
    ws.Cells(1, 4).Value = "PC"
    ws.Cells(1, 5).Value = "MAR"
    ws.Cells(1, 6).Value = "MDR"
    ws.Cells(1, 7).Value = "IR"
    ws.Cells(1, 8).Value = "AX"
    ws.Cells(1, 9).Value = "BX"
    ws.Cells(1, 10).Value = "ZF"
    ws.Cells(1, 11).Value = "CF"
    ws.Cells(1, 12).Value = "SF"
    ws.Cells(1, 13).Value = "Texto"
    ws.Range("A1:M1").Font.Bold = True
    ws.Range("A1:M1").Font.Color = RGB(255, 255, 255)
    ws.Range("A1:M1").Interior.Color = COL_HEADER
    ws.Range("A1:M1").HorizontalAlignment = xlCenter
    ws.Columns("A").ColumnWidth = 6
    ws.Columns("B").ColumnWidth = 10
    ws.Columns("C").ColumnWidth = 36
    ws.Columns("D").ColumnWidth = 6
    ws.Columns("E").ColumnWidth = 6
    ws.Columns("F").ColumnWidth = 6
    ws.Columns("G").ColumnWidth = 6
    ws.Columns("H").ColumnWidth = 6
    ws.Columns("I").ColumnWidth = 6
    ws.Columns("J").ColumnWidth = 4
    ws.Columns("K").ColumnWidth = 4
    ws.Columns("L").ColumnWidth = 4
    ws.Columns("M").ColumnWidth = 70
    On Error GoTo 0
End Sub

Private Sub PulirHojaISA()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("ISA")
    On Error Resume Next
    ws.Range("A1").Font.Size = 16
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Color = COL_TITULO
    ws.Range("A4:G4").Font.Bold = True
    ws.Range("A4:G4").Interior.Color = COL_HEADER
    ws.Range("A4:G4").Font.Color = RGB(255, 255, 255)
    ws.Columns("A").ColumnWidth = 18
    ws.Columns("B").ColumnWidth = 8
    ws.Columns("C").ColumnWidth = 8
    ws.Columns("D").ColumnWidth = 12
    ws.Columns("E").ColumnWidth = 36
    ws.Columns("F").ColumnWidth = 14
    ws.Columns("G").ColumnWidth = 12
    On Error GoTo 0
End Sub

Private Sub PulirHojaREADME()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("README")
    On Error Resume Next
    ws.Cells.Clear
    ws.Range("A1").Value = "Simulador CPU 8 bits - guia rapida para la defensa"
    ws.Range("A1").Font.Size = 18
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Color = COL_TITULO

    ws.Range("A3").Value = "1. Habilitar macros al abrir el .xlsm"
    ws.Range("A4").Value = "2. En Inmediato (Ctrl+G): PulirInterfaz   <- tipografia, pipeline, diagramas"
    ws.Range("A5").Value = "3. Hoja PROGRAM ya tiene el demo. Boton LOAD (DoLoad) o: DoLoad"
    ws.Range("A6").Value = "4. Sembrar datos: SembrarDatosDemo   (o WriteMem &H80,3 / &H81,4 / &H82,0)"
    ws.Range("A7").Value = "5. Demo: DoReset -> STEP varias veces, o RUN con retardo 200 ms"
    ws.Range("A8").Value = "6. Mirar CPU: pipeline FETCH/DECODE/EXECUTE/STORE + camino PC->MAR->MDR->IR"
    ws.Range("A9").Value = "7. Mirar MEMORY: zona codigo (azul) vs datos (naranja); PC cyan, MAR verde"
    ws.Range("A10").Value = "8. Mirar LOG: una fila por micro-operacion (cronologico)"
    ws.Range("A11").Value = "9. Vistas memoria: VistaMemHEX / VistaMemBIN / VistaMemDEC / CiclarVistaMem"
    ws.Range("A12").Value = "10. Diagramas: docs/img/datapath-cpu.png y mapa-memoria.png (si estan junto al libro)"

    ws.Range("A14").Value = "Macros de control"
    ws.Range("A14").Font.Bold = True
    ws.Range("A15").Value = "DoStep | DoRun | DoPause | DoReset | DoLoad | RefreshUI | PulirInterfaz"
    ws.Range("A15").Font.Name = "Consolas"

    ws.Range("A17").Value = "Criterio de interfaz (rubrica): fase iluminada, flujo en tiempo real, botones claros, log detallado."
    ws.Range("A17").Font.Italic = True
    ws.Columns("A").ColumnWidth = 100
    On Error GoTo 0
End Sub

Private Sub EstiloBloqueTitulo(ByVal celda As Range, ByVal texto As String)
    celda.Value = texto
    celda.Font.Bold = True
    celda.Font.Size = 12
    celda.Font.Color = RGB(255, 255, 255)
    celda.Interior.Color = COL_HEADER
End Sub

Private Sub EstiloCeldaRegistro(ByVal celda As Range)
    celda.Font.Name = "Consolas"
    celda.Font.Size = 14
    celda.Font.Bold = True
    celda.HorizontalAlignment = xlCenter
    celda.Borders.LineStyle = xlContinuous
    celda.Borders.Color = RGB(100, 100, 100)
End Sub

Private Sub EstiloBotonCelda(ByVal celda As Range, ByVal texto As String)
    celda.Value = texto
    celda.Font.Name = "Calibri"
    celda.Font.Size = 11
    celda.Font.Bold = True
    celda.Font.Color = RGB(255, 255, 255)
    celda.Interior.Color = COL_BOTON
    celda.HorizontalAlignment = xlCenter
    celda.VerticalAlignment = xlCenter
    celda.Borders.LineStyle = xlContinuous
    celda.Borders.Color = RGB(40, 70, 110)
End Sub

Private Sub AsegurarShapesPipeline(ByVal ws As Worksheet)
    Dim left0 As Single
    Dim top0 As Single
    Dim w As Single
    Dim h As Single
    Dim gap As Single
    Dim i As Long
    Dim nombres As Variant
    Dim colores As Variant
    Dim shp As Shape

    On Error Resume Next
    nombres = Array("shpFetch", "shpDecode", "shpExecute", "shpStore")
    colores = Array(COL_FETCH, COL_DECODE, COL_EXECUTE, COL_STORE)

    ' Fila visual debajo de controles (aprox fila 20)
    left0 = ws.Range("A20").Left
    top0 = ws.Range("A20").Top
    w = 90
    h = 28
    gap = 12

    For i = 0 To 3
        Err.Clear
        Set shp = Nothing
        Set shp = ws.Shapes(CStr(nombres(i)))
        If shp Is Nothing Then
            Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, left0 + i * (w + gap), top0, w, h)
            shp.Name = CStr(nombres(i))
        End If
        shp.Fill.ForeColor.RGB = colores(i)
        shp.Line.ForeColor.RGB = RGB(60, 60, 60)
        Err.Clear
        shp.TextFrame2.TextRange.Characters.Text = UCase$(Replace(CStr(nombres(i)), "shp", ""))
        If Err.Number <> 0 Then
            Err.Clear
            shp.TextFrame.Characters.Text = UCase$(Replace(CStr(nombres(i)), "shp", ""))
            shp.TextFrame.HorizontalAlignment = xlHAlignCenter
            shp.TextFrame.VerticalAlignment = xlVAlignCenter
        Else
            shp.TextFrame2.TextRange.Font.Bold = True
            shp.TextFrame2.TextRange.Font.Size = 11
            shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(20, 20, 20)
            shp.TextFrame2.VerticalAnchor = msoAnchorMiddle
            shp.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
        End If
    Next i

    ' Flechas entre etapas
    Dim flechas As Variant
    flechas = Array("shpArr1", "shpArr2", "shpArr3")
    For i = 0 To 2
        Set shp = Nothing
        Set shp = ws.Shapes(CStr(flechas(i)))
        If shp Is Nothing Then
            Set shp = ws.Shapes.AddShape(msoShapeRightArrow, left0 + (i + 1) * w + i * gap - 8, top0 + 6, 16, 16)
            shp.Name = CStr(flechas(i))
        End If
        shp.Fill.ForeColor.RGB = RGB(80, 80, 80)
        shp.Line.Visible = msoFalse
    Next i
    On Error GoTo 0
End Sub

Private Sub PintarShapesPipeline()
    Dim ws As Worksheet
    Dim nombres As Variant
    Dim colores As Variant
    Dim i As Long
    Dim activo As Long
    Dim shp As Shape

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("CPU")
    If ws Is Nothing Then Exit Sub

    nombres = Array("shpFetch", "shpDecode", "shpExecute", "shpStore")
    colores = Array(COL_FETCH, COL_DECODE, COL_EXECUTE, COL_STORE)
    Select Case FaseActual
        Case FETCH: activo = 0
        Case DECODE: activo = 1
        Case EXECUTE: activo = 2
        Case STORE: activo = 3
        Case Else: activo = -1
    End Select

    For i = 0 To 3
        Set shp = Nothing
        Set shp = ws.Shapes(CStr(nombres(i)))
        If Not shp Is Nothing Then
            If i = activo Then
                shp.Fill.ForeColor.RGB = colores(i)
                shp.Line.Weight = 2.5
                shp.Line.ForeColor.RGB = RGB(0, 0, 0)
            Else
                shp.Fill.ForeColor.RGB = RGB(220, 220, 220)
                shp.Line.Weight = 1
                shp.Line.ForeColor.RGB = RGB(150, 150, 150)
            End If
        End If
    Next i
    On Error GoTo 0
End Sub

Private Sub InsertarDiagramasSiExisten()
    Dim ws As Worksheet
    Dim base As String
    Dim p1 As String
    Dim p2 As String
    Dim shp As Shape

    On Error Resume Next
    base = ThisWorkbook.Path
    If Len(base) = 0 Then Exit Sub
    p1 = base & "\docs\img\datapath-cpu.png"
    p2 = base & "\docs\img\mapa-memoria.png"

    Set ws = ThisWorkbook.Worksheets("CPU")
    If Dir(p1) <> "" Then
        Set shp = Nothing
        Set shp = ws.Shapes("imgDatapath")
        If Not shp Is Nothing Then shp.Delete
        Set shp = ws.Shapes.AddPicture(p1, msoFalse, msoTrue, ws.Range("A22").Left, ws.Range("A22").Top, 420, 160)
        shp.Name = "imgDatapath"
    End If

    Set ws = ThisWorkbook.Worksheets("MEMORY")
    If Dir(p2) <> "" Then
        Set shp = Nothing
        Set shp = ws.Shapes("imgMapaMem")
        If Not shp Is Nothing Then shp.Delete
        Set shp = ws.Shapes.AddPicture(p2, msoFalse, msoTrue, ws.Range("A25").Left, ws.Range("A25").Top, 380, 120)
        shp.Name = "imgMapaMem"
    End If
    On Error GoTo 0
End Sub

Private Function Hex2(ByVal b As Long) As String
    Hex2 = Right$("0" & Hex$(b And &HFF), 2)
End Function

Private Function Bin8(ByVal b As Long) As String
    Dim i As Long
    Dim s As String
    Dim bit As Long
    b = b And &HFF
    For i = 7 To 0 Step -1
        bit = 2 ^ i
        If (b And bit) <> 0 Then
            s = s & "1"
        Else
            s = s & "0"
        End If
    Next i
    Bin8 = s
End Function

Private Function NombreVistaMem() As String
    Select Case mVistaMem
        Case VISTA_HEX: NombreVistaMem = "HEX"
        Case VISTA_BIN: NombreVistaMem = "BIN"
        Case VISTA_DEC: NombreVistaMem = "DEC/mnem"
        Case Else: NombreVistaMem = "?"
    End Select
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
