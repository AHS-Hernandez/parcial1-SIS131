Attribute VB_Name = "modLogger"
Option Explicit

' Una fila por micro-operacion en la hoja LOG.
' ClearLog lo llama DoReset. LogPaso lo llama DoStep.
' En la ventana Inmediato: PruebaLogger

Private mFilas As Long
Private Const FILA_HEADER As Long = 1
Private Const FILA_PRIMERA As Long = 2
Private Const COL_PASO As Long = 1
Private Const COL_FASE As Long = 2
Private Const COL_MICRO As Long = 3
Private Const COL_PC As Long = 4
Private Const COL_MAR As Long = 5
Private Const COL_MDR As Long = 6
Private Const COL_IR As Long = 7
Private Const COL_AX As Long = 8
Private Const COL_BX As Long = 9
Private Const COL_ZF As Long = 10
Private Const COL_CF As Long = 11
Private Const COL_SF As Long = 12
Private Const COL_TEXTO As Long = 13

Public Sub ClearLog()
    Dim ws As Worksheet
    Dim ultima As Long

    mFilas = 0
    On Error GoTo SoloContador
    Set ws = HojaLog()
    ultima = ws.Cells(ws.Rows.Count, COL_PASO).End(xlUp).Row
    If ultima >= FILA_PRIMERA Then
        ws.Range(ws.Cells(FILA_PRIMERA, 1), ws.Cells(ultima, COL_TEXTO)).ClearContents
    End If
    AsegurarEncabezado ws
    Exit Sub
SoloContador:
    mFilas = 0
End Sub

Public Property Get FilasLog() As Long
    FilasLog = mFilas
End Property

' Compatibilidad con pruebas anteriores.
Public Sub AppendLog(ByVal texto As String)
    mFilas = mFilas + 1
End Sub

' Escribe el estado actual de la CPU tras una micro-operacion.
Public Sub LogPaso()
    Dim ws As Worksheet
    Dim fila As Long
    Dim faseTxt As String
    Dim micro As String
    Dim texto As String

    On Error GoTo Fallo

    mFilas = mFilas + 1
    Set ws = HojaLog()
    AsegurarEncabezado ws
    fila = FILA_PRIMERA + mFilas - 1

    faseTxt = NombreFaseLog(FaseActual)
    micro = modControlUnit.UltimaMicroOp
    If Len(micro) = 0 Then micro = "-"

    ws.Cells(fila, COL_PASO).Value = mFilas
    ws.Cells(fila, COL_FASE).Value = faseTxt
    ws.Cells(fila, COL_MICRO).Value = micro
    ws.Cells(fila, COL_PC).Value = Hex2Log(PC)
    ws.Cells(fila, COL_MAR).Value = Hex2Log(MAR)
    ws.Cells(fila, COL_MDR).Value = Hex2Log(MDR)
    ws.Cells(fila, COL_IR).Value = Hex2Log(IR)
    ws.Cells(fila, COL_AX).Value = Hex2Log(AX)
    ws.Cells(fila, COL_BX).Value = Hex2Log(BX)
    ws.Cells(fila, COL_ZF).Value = modFlags.ZF
    ws.Cells(fila, COL_CF).Value = modFlags.CF
    ws.Cells(fila, COL_SF).Value = modFlags.SF

    texto = "[Paso " & Format$(mFilas, "00") & "] " & faseTxt & ": " & micro & _
            " | PC=" & Hex2Log(PC) & " MAR=" & Hex2Log(MAR) & _
            " MDR=" & Hex2Log(MDR) & " IR=" & Hex2Log(IR) & _
            " AX=" & Hex2Log(AX) & " BX=" & Hex2Log(BX)
    ws.Cells(fila, COL_TEXTO).Value = texto
    Exit Sub
Fallo:
    Debug.Print "LogPaso", Err.Number, Err.Description
End Sub

Public Sub PruebaLogger()
    Dim fallos As Long
    Dim ws As Worksheet

    On Error GoTo Fallo
    fallos = 0

    ClearMem
    WriteMem 0, OP_MOV_AX_IMM
    WriteMem 1, &H1
    WriteMem 2, OP_HLT
    DoReset

    If FilasLog <> 0 Then
        Debug.Print "FALLO ClearLog", FilasLog
        fallos = fallos + 1
    End If

    DoStep
    DoStep
    DoStep
    DoStep

    If FilasLog <> 4 Then
        Debug.Print "FALLO filas tras 4 STEP", FilasLog
        fallos = fallos + 1
    Else
        Debug.Print "4 micros", "ok", "filas=" & FilasLog
    End If

    Set ws = HojaLog()
    If CLng(ws.Cells(FILA_PRIMERA, COL_PASO).Value) <> 1 Then
        Debug.Print "FALLO paso1", ws.Cells(FILA_PRIMERA, COL_PASO).Value
        fallos = fallos + 1
    End If
    If UCase$(CStr(ws.Cells(FILA_PRIMERA, COL_FASE).Value)) <> "FETCH" Then
        Debug.Print "FALLO fase1", ws.Cells(FILA_PRIMERA, COL_FASE).Value
        fallos = fallos + 1
    Else
        Debug.Print "fase FETCH", "ok"
    End If
    If InStr(1, CStr(ws.Cells(FILA_PRIMERA, COL_MICRO).Value), "F1", vbTextCompare) = 0 Then
        Debug.Print "FALLO micro1", ws.Cells(FILA_PRIMERA, COL_MICRO).Value
        fallos = fallos + 1
    Else
        Debug.Print "micro F1", "ok", ws.Cells(FILA_PRIMERA, COL_MICRO).Value
    End If
    If InStr(1, CStr(ws.Cells(FILA_PRIMERA, COL_TEXTO).Value), "[Paso 01]", vbTextCompare) = 0 Then
        Debug.Print "FALLO texto", ws.Cells(FILA_PRIMERA, COL_TEXTO).Value
        fallos = fallos + 1
    Else
        Debug.Print "texto", "ok", Left$(CStr(ws.Cells(FILA_PRIMERA, COL_TEXTO).Value), 60)
    End If

    DoReset
    If FilasLog <> 0 Then
        Debug.Print "FALLO RESET no limpio log", FilasLog
        fallos = fallos + 1
    ElseIf Len(Trim$(CStr(ws.Cells(FILA_PRIMERA, COL_PASO).Value & ""))) <> 0 Then
        Debug.Print "FALLO hoja LOG no vacia tras RESET"
        fallos = fallos + 1
    Else
        Debug.Print "ClearLog en RESET", "ok"
    End If

    If fallos = 0 Then
        Debug.Print "modLogger escribe una fila por micro-operacion"
    Else
        Debug.Print "FALLOS", fallos
    End If
    Exit Sub
Fallo:
    Debug.Print "FALLO PruebaLogger", Err.Number, Err.Description
End Sub

Private Function HojaLog() As Worksheet
    On Error Resume Next
    Set HojaLog = ThisWorkbook.Worksheets("LOG")
    On Error GoTo 0
    If HojaLog Is Nothing Then
        Err.Raise vbObjectError + 5701, "modLogger", "No existe la hoja LOG"
    End If
End Function

Private Sub AsegurarEncabezado(ByVal ws As Worksheet)
    On Error Resume Next
    If Len(Trim$(CStr(ws.Cells(FILA_HEADER, COL_PASO).Value & ""))) = 0 _
       Or UCase$(CStr(ws.Cells(FILA_HEADER, COL_PASO).Value)) = "LOG" Then
        ws.Cells(FILA_HEADER, COL_PASO).Value = "Paso"
        ws.Cells(FILA_HEADER, COL_FASE).Value = "Fase"
        ws.Cells(FILA_HEADER, COL_MICRO).Value = "Micro-operacion"
        ws.Cells(FILA_HEADER, COL_PC).Value = "PC"
        ws.Cells(FILA_HEADER, COL_MAR).Value = "MAR"
        ws.Cells(FILA_HEADER, COL_MDR).Value = "MDR"
        ws.Cells(FILA_HEADER, COL_IR).Value = "IR"
        ws.Cells(FILA_HEADER, COL_AX).Value = "AX"
        ws.Cells(FILA_HEADER, COL_BX).Value = "BX"
        ws.Cells(FILA_HEADER, COL_ZF).Value = "ZF"
        ws.Cells(FILA_HEADER, COL_CF).Value = "CF"
        ws.Cells(FILA_HEADER, COL_SF).Value = "SF"
        ws.Cells(FILA_HEADER, COL_TEXTO).Value = "Texto"
    End If
    ' Cabecera legible para la defensa (no toca filas de datos)
    ws.Range(ws.Cells(FILA_HEADER, 1), ws.Cells(FILA_HEADER, COL_TEXTO)).Font.Bold = True
    ws.Range(ws.Cells(FILA_HEADER, 1), ws.Cells(FILA_HEADER, COL_TEXTO)).Interior.Color = RGB(46, 94, 139)
    ws.Range(ws.Cells(FILA_HEADER, 1), ws.Cells(FILA_HEADER, COL_TEXTO)).Font.Color = RGB(255, 255, 255)
    On Error GoTo 0
End Sub

Private Function Hex2Log(ByVal b As Long) As String
    Hex2Log = Right$("0" & Hex$(b And &HFF), 2) & "h"
End Function

Private Function NombreFaseLog(ByVal f As ePhase) As String
    Select Case f
        Case FETCH: NombreFaseLog = "FETCH"
        Case DECODE: NombreFaseLog = "DECODE"
        Case EXECUTE: NombreFaseLog = "EXECUTE"
        Case STORE: NombreFaseLog = "STORE"
        Case Else: NombreFaseLog = "?"
    End Select
End Function
