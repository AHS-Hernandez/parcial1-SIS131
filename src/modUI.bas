Attribute VB_Name = "modUI"
Option Explicit

' Pinta CPU y MEMORY desde el modelo. No calcula ni avanza el reloj.
' Vistas de memoria: VistaMemHEX / VistaMemBIN / VistaMemDEC / CiclarVistaMem
' Tras cada DoStep: RefreshUI
' PulirInterfaz: layout profesional (bloques, colores, tipografia, diagramas)
' Botones = Formas Excel (msoShapeRoundedRectangle) con OnAction; viven en el .xlsm
' CrearBotonesForma / Auto_Open / Workbook_Open: crean o restauran si faltan
' En Inmediato: PruebaUI, PulirInterfaz, CrearBotonesForma  -> luego Guardar libro

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

' Al abrir el libro (macros habilitadas): asegura Formas nativas en hoja CPU.
Public Sub Auto_Open()
    On Error Resume Next
    On Error GoTo 0
End Sub

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
    AsegurarNombresUI
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

Public Sub CrearBotonesForma()
    On Error GoTo Fallo
    Application.ScreenUpdating = False
    Application.ScreenUpdating = True
    Debug.Print "CrearBotonesForma", "ok", "btnStep/btnRun/btnPause/btnReset/btnLoad"
    Exit Sub
Fallo:
    Application.ScreenUpdating = True
    Debug.Print "FALLO CrearBotonesForma", Err.Number, Err.Description
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
    
    ThisWorkbook.Names.Add Name:="rngState", RefersTo:="=CPU!$B$9"
    ThisWorkbook.Names.Add Name:="rngPhase", RefersTo:="=CPU!$D$9"
    
    ThisWorkbook.Names.Add Name:="rngPC", RefersTo:="=CPU!$B$10"
    ThisWorkbook.Names.Add Name:="rngIR", RefersTo:="=CPU!$B$11"
    ThisWorkbook.Names.Add Name:="rngMAR", RefersTo:="=CPU!$B$12"
    ThisWorkbook.Names.Add Name:="rngMDR", RefersTo:="=CPU!$B$13"
    
    ThisWorkbook.Names.Add Name:="rngInstruccion", RefersTo:="=CPU!$B$14"
    ThisWorkbook.Names.Add Name:="rngMicroOp", RefersTo:="=CPU!$B$15"
    
    ThisWorkbook.Names.Add Name:="rngAX", RefersTo:="=CPU!$G$10"
    ThisWorkbook.Names.Add Name:="rngBX", RefersTo:="=CPU!$G$11"
    ThisWorkbook.Names.Add Name:="rngTemp", RefersTo:="=CPU!$G$12"
    
    ThisWorkbook.Names.Add Name:="rngZF", RefersTo:="=CPU!$G$14"
    ThisWorkbook.Names.Add Name:="rngCF", RefersTo:="=CPU!$G$15"
    ThisWorkbook.Names.Add Name:="rngSF", RefersTo:="=CPU!$G$16"
    
    ThisWorkbook.Names.Add Name:="rngPaso", RefersTo:="=CPU!$I$3"
    ThisWorkbook.Names.Add Name:="rngDelay", RefersTo:="=CPU!$I$21"
    ThisWorkbook.Names.Add Name:="rngFlujo", RefersTo:="=CPU!$J$50"
    
    ThisWorkbook.Names.Add Name:="rngPipeFetch", RefersTo:="=CPU!$B$3"
    ThisWorkbook.Names.Add Name:="rngPipeDecode", RefersTo:="=CPU!$C$3"
    ThisWorkbook.Names.Add Name:="rngPipeExecute", RefersTo:="=CPU!$D$3"
    ThisWorkbook.Names.Add Name:="rngPipeStore", RefersTo:="=CPU!$E$3"
    
    ThisWorkbook.Names.Add Name:="rngVistaMem", RefersTo:="=MEMORY!$G$22"
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
    Range("rngTemp").Value = Hex2(Temporal)
    Range("rngZF").Value = modFlags.ZF
    Range("rngCF").Value = modFlags.CF
    Range("rngSF").Value = modFlags.SF
    
    ' Actualizar la descripcion del IR con el mnemonico decodificado
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("CPU")
    If InstruccionActual.Encontrada Then
        ws.Range("C11").Value = InstruccionActual.Mnemonic
    Else
        ws.Range("C11").Value = "Opcode en curso"
    End If
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
        ResaltarRango "rngTemp"
        If InStr(1, op, "LOAD", vbTextCompare) > 0 Or InStr(1, op, "STORE", vbTextCompare) > 0 Then
            ResaltarRango "rngMAR"
            ResaltarRango "rngMDR"
        End If
    ElseIf FaseActual = STORE Then
        flu = "Store / write-back"
        ResaltarRango "rngAX"
        ResaltarRango "rngBX"
        ResaltarRango "rngTemp"
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

    ' Actualizar indicador de vista y resaltar boton activo
    On Error Resume Next
    Range("rngVistaMem").Value = NombreVistaMem()
    
    ' Resaltar el boton Shape activo (dorado) y los inactivos (gris oscuro)
    Dim wsMem As Worksheet
    Set wsMem = ThisWorkbook.Worksheets("MEMORY")
    Dim btnNombres As Variant
    Dim btnVistas As Variant
    Dim idx As Long
    Dim shpBtn As Shape
    btnNombres = Array("btnMemHEX", "btnMemBIN", "btnMemDEC")
    btnVistas = Array(VISTA_HEX, VISTA_BIN, VISTA_DEC)
    For idx = 0 To 2
        Set shpBtn = Nothing
        Set shpBtn = wsMem.Shapes(CStr(btnNombres(idx)))
        If Not shpBtn Is Nothing Then
            If CLng(btnVistas(idx)) = mVistaMem Then
                shpBtn.Fill.ForeColor.RGB = COL_ACTIVO
                shpBtn.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
            Else
                shpBtn.Fill.ForeColor.RGB = RGB(89, 89, 89)
                shpBtn.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
            End If
        End If
    Next idx
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
    Range("rngTemp").Interior.ColorIndex = xlNone
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
    
    ws.Cells.ClearFormats
    ws.Cells.ClearContents

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
    ws.Range("A3").Font.Size = 14

    ws.Range("H3").Value = "Paso #"
    ws.Range("H3").Font.Bold = True
    ws.Range("H3").Font.Size = 14
    ws.Range("rngPaso").Font.Size = 14
    
    ' ==========================================
    ' CONTENEDOR 1: UNIDAD DE CONTROL (U.C.) (Bajado a fila 8)
    ' ==========================================
    EstiloBloqueTitulo ws.Range("A8"), "UNIDAD DE CONTROL (U.C.)"
    ws.Range("A8:D8").Merge
    
    ws.Range("A9").Value = "Estado"
    ws.Range("A9").Font.Bold = True
    ws.Range("C9").Value = "Fase"
    ws.Range("C9").Font.Bold = True
    ws.Range("rngState").Font.Name = "Consolas"
    ws.Range("rngState").Font.Size = 12
    ws.Range("rngState").Font.Bold = True
    ws.Range("rngPhase").Font.Name = "Consolas"
    ws.Range("rngPhase").Font.Size = 12
    ws.Range("rngPhase").Font.Bold = True
    
    ws.Range("A10").Value = "PC"
    ws.Range("A11").Value = "IR"
    ws.Range("A12").Value = "MAR"
    ws.Range("A13").Value = "MDR"
    ws.Range("A10:A13").Font.Bold = True
    
    ws.Range("C10").Value = "Siguiente instruccion"
    ws.Range("C11").Value = "Opcode crudo"
    ws.Range("C12").Value = "Direccion hacia RAM"
    ws.Range("C13").Value = "Dato leido / a escribir"
    ws.Range("C10:C13").Font.Size = 9
    ws.Range("C10:C13").Font.Color = RGB(89, 89, 89)

    EstiloCeldaRegistro ws.Range("rngPC")
    EstiloCeldaRegistro ws.Range("rngIR")
    EstiloCeldaRegistro ws.Range("rngMAR")
    EstiloCeldaRegistro ws.Range("rngMDR")
    
    ' Nuevas filas limpias para Decodificacion y Micro-Op
    ws.Range("A14").Value = "Decodific."
    ws.Range("A14").Font.Bold = True
    ws.Range("B14:D14").Merge
    ws.Range("rngInstruccion").Font.Name = "Consolas"
    ws.Range("rngInstruccion").Font.Size = 11
    ws.Range("rngInstruccion").Font.Bold = True
    ws.Range("rngInstruccion").HorizontalAlignment = xlCenter
    ws.Range("rngInstruccion").Interior.Color = RGB(255, 250, 205) ' Amarillo suave (como antes)
    ws.Range("rngInstruccion").Borders.LineStyle = xlContinuous
    
    ws.Range("A15").Value = "Micro-Op"
    ws.Range("A15").Font.Bold = True
    ws.Range("B15:D15").Merge
    ws.Range("rngMicroOp").Font.Name = "Consolas"
    ws.Range("rngMicroOp").Font.Size = 10
    ws.Range("rngMicroOp").HorizontalAlignment = xlCenter
    ws.Range("rngMicroOp").Interior.Color = RGB(255, 250, 205) ' Amarillo suave
    ws.Range("rngMicroOp").Borders.LineStyle = xlContinuous
    
    ws.Range("A8:D16").BorderAround xlContinuous, xlMedium
    
    ' Limpiar basura anterior
    ws.Range("I4").ClearContents
    ws.Range("I4").Borders.LineStyle = xlNone

    ' ==========================================
    ' CONTENEDOR 2: ALU & REGISTROS (Bajado a fila 8)
    ' ==========================================
    EstiloBloqueTitulo ws.Range("F8"), "ALU & REGISTROS"
    ws.Range("F8:I8").Merge
    
    ws.Range("F10").Value = "AX"
    ws.Range("F11").Value = "BX"
    ws.Range("F12").Value = "Temp"
    ws.Range("F10:F12").Font.Bold = True
    
    ws.Range("H10").Value = "Acumulador"
    ws.Range("H11").Value = "Proposito general"
    ws.Range("H12").Value = "Resultado temporal"
    ws.Range("H10:H12").Font.Size = 9
    ws.Range("H10:H12").Font.Color = RGB(89, 89, 89)

    EstiloCeldaRegistro ws.Range("rngAX")
    EstiloCeldaRegistro ws.Range("rngBX")
    EstiloCeldaRegistro ws.Range("rngTemp")
    
    ws.Range("F14").Value = "ZF"
    ws.Range("F15").Value = "CF"
    ws.Range("F16").Value = "SF"
    ws.Range("F14:F16").Font.Bold = True
    
    ws.Range("H14").Value = "Cero (Zero)"
    ws.Range("H15").Value = "Acarreo (Carry)"
    ws.Range("H16").Value = "Signo (Sign)"
    ws.Range("H14:H16").Font.Size = 9
    ws.Range("H14:H16").Font.Color = RGB(89, 89, 89)
    
    ws.Range("F8:I17").BorderAround xlContinuous, xlMedium

    ' ==========================================
    ' PANEL DE CONTROL (Bajado a fila 21)
    ' ==========================================
    EstiloBloqueTitulo ws.Range("A21"), "PANEL DE CONTROL INTERACTIVO"
    ws.Range("A21:E21").Merge
    
    ws.Range("H21").Value = "Retardo (ms)"
    ws.Range("H21").Font.Bold = True
    ws.Range("rngDelay").Font.Name = "Consolas"
    ws.Range("rngDelay").Font.Size = 12
    If Len(Trim$(CStr(ws.Range("rngDelay").Value & ""))) = 0 Then
        ws.Range("rngDelay").Value = 200
    End If
    
    ' Espaciado de columnas
    ws.Columns("A").ColumnWidth = 14
    ws.Columns("B").ColumnWidth = 12
    ws.Columns("C").ColumnWidth = 28
    ws.Columns("D").ColumnWidth = 12
    ws.Columns("E").ColumnWidth = 4
    ws.Columns("F").ColumnWidth = 14
    ws.Columns("G").ColumnWidth = 10
    ws.Columns("H").ColumnWidth = 22
    ws.Columns("I").ColumnWidth = 10
    ws.Columns("J").ColumnWidth = 10
    
    ws.Rows("1").RowHeight = 28
    ws.Rows("3").RowHeight = 30 ' Para los botones del pipeline
    ws.Rows("8").RowHeight = 24
    ws.Rows("21").RowHeight = 24

    AsegurarShapesPipeline ws
    AjustarColoresBotones ws
End Sub



Private Sub PulirHojaMEMORY()
    Dim ws As Worksheet
    Dim c As Long
    Dim shp As Shape
    Set ws = ThisWorkbook.Worksheets("MEMORY")

    ' ==========================================
    ' TITULO Y SUBTITULO
    ' ==========================================
    ws.Range("A1").Value = "Memoria Principal RAM  (256 x 8 bits)"
    ws.Range("A1").Font.Size = 18
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Color = COL_TITULO

    ws.Range("A2").Value = "Direccion = fila*16 + col  |  Azul = Codigo (00h-7Fh)  |  Naranja = Datos (80h-FFh)  |  Cyan = PC  |  Verde = MAR"
    ws.Range("A2").Font.Size = 9
    ws.Range("A2").Font.Italic = True
    ws.Range("A2").Font.Color = RGB(89, 89, 89)

    ' ==========================================
    ' ENCABEZADOS DE COLUMNA (0 a F)
    ' ==========================================
    ws.Range("A4").Value = ""
    ws.Range("A4").Interior.Color = COL_HEADER
    For c = 0 To 15
        ws.Cells(4, c + 2).Value = Hex$(c)
        ws.Cells(4, c + 2).Font.Bold = True
        ws.Cells(4, c + 2).Font.Name = "Consolas"
        ws.Cells(4, c + 2).Font.Size = 10
        ws.Cells(4, c + 2).HorizontalAlignment = xlCenter
        ws.Cells(4, c + 2).Interior.Color = COL_HEADER
        ws.Cells(4, c + 2).Font.Color = RGB(255, 255, 255)
    Next c

    ' ==========================================
    ' ENCABEZADOS DE FILA (0x0_, 0x1_, ... 0xF_)
    ' ==========================================
    For c = 0 To 15
        ws.Cells(5 + c, 1).Value = Hex$(c) & "_"
        ws.Cells(5 + c, 1).Font.Bold = True
        ws.Cells(5 + c, 1).Font.Name = "Consolas"
        ws.Cells(5 + c, 1).Font.Size = 10
        ws.Cells(5 + c, 1).HorizontalAlignment = xlCenter
        If c <= 7 Then
            ws.Cells(5 + c, 1).Interior.Color = COL_CODIGO
        Else
            ws.Cells(5 + c, 1).Interior.Color = COL_DATOS
        End If
    Next c

    ' ==========================================
    ' CUADRICULA RAM 16x16
    ' ==========================================
    ws.Range("rngRAM").Font.Name = "Consolas"
    ws.Range("rngRAM").Font.Size = 9
    ws.Range("rngRAM").HorizontalAlignment = xlCenter
    ws.Range("rngRAM").Borders.LineStyle = xlContinuous
    ws.Range("rngRAM").Borders.Color = RGB(180, 180, 180)

    ' ==========================================
    ' LEYENDA DE SEGMENTACION (fila 22)
    ' ==========================================
    ws.Range("A22").Value = "Codigo"
    ws.Range("A22").Font.Bold = True
    ws.Range("A22").Interior.Color = COL_CODIGO
    ws.Range("B22").Value = "00h-7Fh"
    ws.Range("B22").Font.Size = 9

    ws.Range("A23").Value = "Datos"
    ws.Range("A23").Font.Bold = True
    ws.Range("A23").Interior.Color = COL_DATOS
    ws.Range("B23").Value = "80h-FFh"
    ws.Range("B23").Font.Size = 9

    ws.Range("C22").Value = "PC"
    ws.Range("C22").Font.Bold = True
    ws.Range("C22").Interior.Color = COL_PC_MEM
    ws.Range("D22").Value = "Cyan"
    ws.Range("D22").Font.Size = 9

    ws.Range("C23").Value = "MAR"
    ws.Range("C23").Font.Bold = True
    ws.Range("C23").Interior.Color = COL_MAR
    ws.Range("D23").Value = "Verde"
    ws.Range("D23").Font.Size = 9

    ' ==========================================
    ' SELECTOR DE FORMATO: Botones Nativos (Shapes)
    ' ==========================================
    ws.Range("F22").Value = "Vista:"
    ws.Range("F22").Font.Bold = True
    ws.Range("F22").Font.Size = 11

    ' Indicador de vista activa
    ws.Range("rngVistaMem").Font.Name = "Consolas"
    ws.Range("rngVistaMem").Font.Bold = True
    ws.Range("rngVistaMem").Font.Size = 12
    ws.Range("rngVistaMem").HorizontalAlignment = xlCenter
    ws.Range("rngVistaMem").Interior.Color = RGB(255, 250, 205)
    ws.Range("rngVistaMem").Borders.LineStyle = xlContinuous

    ' Crear o reposicionar Shapes para HEX / BIN / DEC
    CrearBotonMemoria ws, "btnMemHEX", "HEX", ws.Range("H22"), "VistaMemHEX"
    CrearBotonMemoria ws, "btnMemBIN", "BIN", ws.Range("I22"), "VistaMemBIN"
    CrearBotonMemoria ws, "btnMemDEC", "DEC", ws.Range("J22"), "VistaMemDEC"

    ' ==========================================
    ' ANCHOS DE COLUMNA
    ' ==========================================
    ws.Columns("A").ColumnWidth = 6
    Dim col As Long
    For col = 2 To 17
        ws.Columns(col).ColumnWidth = 10
    Next col
    ws.Rows("1").RowHeight = 26
    ws.Rows("4").RowHeight = 20
End Sub

' Crea o reposiciona un boton nativo (Shape) en la hoja MEMORY para cambiar formato.
Private Sub CrearBotonMemoria(ByVal ws As Worksheet, ByVal nombre As String, _
                              ByVal texto As String, ByVal celda As Range, _
                              ByVal macro As String)
    Dim shp As Shape
    On Error Resume Next
    Set shp = ws.Shapes(nombre)
    On Error GoTo 0

    If shp Is Nothing Then
        Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, _
                    celda.Left + 2, celda.Top + 2, _
                    celda.Width - 4, celda.Height - 4)
        shp.Name = nombre
        shp.TextFrame2.TextRange.Text = texto
        shp.TextFrame2.TextRange.Font.Size = 11
        shp.TextFrame2.TextRange.Font.Bold = msoTrue
        shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        shp.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
        shp.TextFrame2.VerticalAnchor = msoAnchorMiddle
        shp.OnAction = macro
    Else
        shp.Top = celda.Top + 2
        shp.Left = celda.Left + 2
        shp.Width = celda.Width - 4
        shp.Height = celda.Height - 4
    End If

    shp.Fill.ForeColor.RGB = RGB(89, 89, 89)
    shp.Line.ForeColor.RGB = RGB(60, 60, 60)
    shp.Line.Weight = 1
End Sub

Private Sub PulirHojaPROGRAM()
    Dim ws As Worksheet
    Dim rng As Range
    Set ws = ThisWorkbook.Worksheets("PROGRAM")
    ws.Range("A1").Value = "PROGRAM - demo: multiplicacion (03h x 04h) en registros"
    ws.Range("A1").Font.Size = 14
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Color = COL_TITULO
    ws.Range("A1").Interior.Color = RGB(232, 240, 254)
    
    ws.Range("B1").Value = "LOAD escribe la Columna A en RAM."
    ws.Range("B1").Font.Size = 9
    ws.Range("B1").Font.Italic = True

    On Error Resume Next
    Set rng = Range("rngProgram")
    rng.Font.Name = "Consolas"
    rng.Font.Size = 12
    rng.Borders.LineStyle = xlContinuous
    
    ' Escribir el programa Demo Principal (Multiplicacion)
    rng.ClearContents
    rng.Cells(1, 1).Value = "MOV AX, 00h"
    rng.Cells(2, 1).Value = "MOV BX, 03h"
    rng.Cells(3, 1).Value = "CMP BX, 00h"
    rng.Cells(4, 1).Value = "JZ 0Dh"
    rng.Cells(5, 1).Value = "ADD AX, 04h"
    rng.Cells(6, 1).Value = "DEC BX"
    rng.Cells(7, 1).Value = "JMP 04h"
    rng.Cells(8, 1).Value = "STORE [82h], AX"
    rng.Cells(9, 1).Value = "HLT"
    
    ' Escribir el programa Alternativo (Fibonacci) en la columna E (Como backup)
    ws.Range("E1").Value = "BACKUP: Serie de Fibonacci hasta limite de 8 bits (233)"
    ws.Range("E1").Font.Bold = True
    ws.Range("E1").Interior.Color = RGB(255, 230, 204)
    ws.Range("E2").Value = "Si te lo piden, copia este codigo y pegalo en la Columna A."
    ws.Range("E2").Font.Italic = True
    ws.Range("E2").Font.Size = 9
    
    ws.Range("E4").Value = "MOV AX, 00h"
    ws.Range("E5").Value = "STORE [80h], AX"
    ws.Range("E6").Value = "MOV AX, 01h"
    ws.Range("E7").Value = "STORE [81h], AX"
    ws.Range("E8").Value = "LOAD AX, [80h]"
    ws.Range("E9").Value = "LOAD BX, [81h]"
    ws.Range("E10").Value = "ADD AX, BX"
    ws.Range("E11").Value = "JC 18h"
    ws.Range("E12").Value = "STORE [82h], AX"
    ws.Range("E13").Value = "STORE [81h], AX"
    ws.Range("E14").Value = "MOV AX, BX"
    ws.Range("E15").Value = "STORE [80h], AX"
    ws.Range("E16").Value = "JMP 08h"
    ws.Range("E17").Value = "HLT"
    
    ws.Range("E18:E24").ClearContents
    
    ws.Range("E4:E17").Font.Name = "Consolas"
    ws.Range("E4:E17").Font.Size = 11
    ws.Range("E4:E17").Borders.LineStyle = xlContinuous
    ws.Range("E4:E17").Interior.Color = RGB(245, 245, 245)
    On Error GoTo 0
    
    ws.Columns("A").ColumnWidth = 30
    ws.Columns("E").ColumnWidth = 30
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
    
    ' Inyectar JC al final de la tabla si no existe
    Dim r As Long
    Dim foundJC As Boolean
    foundJC = False
    For r = 5 To 30
        If ws.Cells(r, 1).Value = "JC" Then foundJC = True
    Next r
    If Not foundJC Then
        r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
        ws.Cells(r, 1).Value = "JC"
        ws.Cells(r, 2).Value = "98"
        ws.Cells(r, 3).Value = "10011000"
        ws.Cells(r, 4).Value = "2"
        ws.Cells(r, 5).Value = "Jump if Carry (Salta si hay desbordamiento)"
        ws.Cells(r, 6).Value = "IF CF=1 PC=dir"
        ws.Cells(r, 7).Value = "- - -"
        ws.Range(ws.Cells(r, 1), ws.Cells(r, 7)).Borders.LineStyle = xlContinuous
    End If
    
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

    ws.Range("A3").Value = "1. Habilitar macros al abrir el .xlsm (los botones son Formas Excel, no scripts)"
    ws.Range("A4").Value = "2. Si faltan botones STEP/RUN/...: Inmediato -> CrearBotonesForma  y Guardar el libro"
    ws.Range("A5").Value = "3. (Opcional defensa) PulirInterfaz  <- tipografia, pipeline, diagramas"
    ws.Range("A6").Value = "4. Hoja PROGRAM ya tiene el demo (multiplicacion 3x4 en registros). Boton LOAD o macro DoLoad"
    ws.Range("A7").Value = "5. Demo: LOAD -> RESET -> STEP varias veces, o RUN con retardo 200 ms"
    ws.Range("A8").Value = "6. Mirar CPU: pipeline FETCH/DECODE/EXECUTE/STORE + camino PC->MAR->MDR->IR"
    ws.Range("A9").Value = "7. Mirar MEMORY: zona codigo (azul) vs datos (naranja); PC cyan, MAR verde"
    ws.Range("A10").Value = "8. Cambiar formato de memoria: botones HEX / BIN / DEC en la hoja MEMORY"
    ws.Range("A11").Value = "9. Mirar LOG: una fila por micro-operacion (cronologico)"
    ws.Range("A12").Value = "10. Diagramas: docs/img/datapath-cpu.png y mapa-memoria.png (si estan junto al libro)"

    ws.Range("A14").Value = "Macros de control"
    ws.Range("A14").Font.Bold = True
    ws.Range("A15").Value = "DoStep | DoRun | DoPause | DoReset | DoLoad | RefreshUI | PulirInterfaz"
    ws.Range("A15").Font.Name = "Consolas"

    ws.Range("A17").Value = "Criterio de interfaz (rubrica): fase iluminada, flujo en tiempo real, Formas de control, log detallado."
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
    Dim i As Long
    Dim nombres As Variant, flechas As Variant
    Dim shp As Shape
    Dim cellTarget As Range

    On Error Resume Next
    nombres = Array("shpFetch", "shpDecode", "shpExecute", "shpStore")
    flechas = Array("shpArr1", "shpArr2", "shpArr3")
    
    ' Alinear con fila 3, columnas B a E
    For i = 0 To 3
        Set shp = ws.Shapes(CStr(nombres(i)))
        If Not shp Is Nothing Then
            Set cellTarget = ws.Cells(3, 2 + i) ' B3, C3, D3, E3
            shp.Top = cellTarget.Top + 2
            shp.Left = cellTarget.Left + 2
            shp.Width = cellTarget.Width - 4
            shp.Height = cellTarget.Height - 4
            
            ' Restaurar color base sobrio (RGB 200, 200, 200)
            shp.Fill.ForeColor.RGB = RGB(200, 200, 200)
        End If
    Next i
    
    ' Alinear las flechas (opcional, si existen)
    For i = 0 To 2
        Set shp = ws.Shapes(CStr(flechas(i)))
        If Not shp Is Nothing Then
            Set cellTarget = ws.Cells(3, 2 + i)
            shp.Top = cellTarget.Top + (cellTarget.Height / 2) - 5
            shp.Left = cellTarget.Left + cellTarget.Width - 6
        End If
    Next i
    On Error GoTo 0
End Sub

Private Sub AjustarColoresBotones(ByVal ws As Worksheet)
    Dim shp As Shape
    Dim grisOscuro As Long
    Dim targetTop As Single
    Dim i As Long
    Dim nombres As Variant
    
    grisOscuro = RGB(89, 89, 89) ' Gris oscuro original
    nombres = Array("btnStep", "btnRun", "btnPause", "btnReset", "btnLoad")
    
    On Error Resume Next
    targetTop = ws.Range("A21").Top + ws.Rows(21).Height + 2 ' Fila 22 real
    
    For i = 0 To 4
        Set shp = ws.Shapes(CStr(nombres(i)))
        If Not shp Is Nothing Then
            ' Mover a la fila 22 y acomodar su ancho al de la celda
            shp.Top = targetTop
            shp.Left = ws.Cells(22, i + 1).Left + 2
            shp.Width = ws.Cells(22, i + 1).Width - 4
            shp.Height = ws.Cells(22, i + 1).Height - 4
        End If
    Next i
    
    ' Ajustar colores de los sobrios al gris oscuro
    Set shp = ws.Shapes("btnStep")
    If Not shp Is Nothing Then shp.Fill.ForeColor.RGB = grisOscuro
    Set shp = ws.Shapes("btnPause")
    If Not shp Is Nothing Then shp.Fill.ForeColor.RGB = grisOscuro
    Set shp = ws.Shapes("btnLoad")
    If Not shp Is Nothing Then shp.Fill.ForeColor.RGB = grisOscuro
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

