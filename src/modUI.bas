Attribute VB_Name = "modUI"
Option Explicit

' Pinta CPU y MEMORY desde el modelo. No calcula ni avanza el reloj.
' Vistas de memoria: VistaMemHEX / VistaMemBIN / VistaMemDEC / CiclarVistaMem
' Tras cada DoStep: RefreshUI
' En la ventana Inmediato: PruebaUI, PruebaVistaMemoria

Private Const VISTA_HEX As Integer = 0
Private Const VISTA_BIN As Integer = 1
Private Const VISTA_DEC As Integer = 2

Private mUltimaCeldaMAR As Range
Private mVistaMem As Integer

Public Sub RefreshUI()
    On Error GoTo Fallo
    AsegurarNombresUI
    PintarEstadoYFase
    PintarRegistrosYBanderas
    PintarInstruccionYMicro
    PintarCaminoDatos
    PintarMemoria
    PintarCeldaMAR
    DoEvents
    Exit Sub
Fallo:
    Debug.Print "RefreshUI", Err.Number, Err.Description
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
    colorCodigo = RGB(210, 230, 255)
    colorDatos = RGB(255, 236, 210)

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
    On Error GoTo 0

    Application.ScreenUpdating = True
    Set mUltimaCeldaMAR = Nothing
End Sub

Private Sub PintarCeldaMAR()
    Dim celda As Range
    Dim addr As Long

    addr = MAR
    If addr < 0 Or addr > 255 Then Exit Sub

    Set celda = Range("rngRAM").Cells((addr \ 16) + 1, (addr Mod 16) + 1)
    If InStr(1, UltimaMicroOp, "F2", vbTextCompare) > 0 _
       Or InStr(1, UltimaMicroOp, "LOAD", vbTextCompare) > 0 _
       Or InStr(1, UltimaMicroOp, "STORE", vbTextCompare) > 0 _
       Or InStr(1, UltimaMicroOp, "F1", vbTextCompare) > 0 _
       Or FaseActual = STORE Then
        celda.Interior.Color = RGB(102, 255, 102)
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
    Range(nombre).Interior.Color = RGB(255, 205, 127)
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
