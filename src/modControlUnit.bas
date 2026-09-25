Attribute VB_Name = "modControlUnit"
Option Explicit

' Fase Fetch, un micro-paso por llamada. No se juntan en una sola linea.
' En la ventana Inmediato: PruebaFetch
' Traza a mano: PC=00h y RAM(00h)=10h => MAR=00, MDR=10h, IR=10h, PC=01h

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

Public Property Get EstadoCPU() As eCPUState
    EstadoCPU = mEstado
End Property

Public Property Get FaseActual() As ePhase
    FaseActual = mFase
End Property

Public Property Get NumeroDePaso() As Integer
    NumeroDePaso = mPasoFetch
End Property

Public Sub IniciarFetch()
    mEstado = RUNNING
    mFase = FETCH
    mPasoFetch = 1
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

    ' Segundo caso: PC en 05h, RAM(05h)=A0h. A mano: MAR=05, MDR=A0, IR=A0, PC=06.
    WriteMem 5, &HA0
    SetPC 5
    IniciarFetch
    PasoFetch
    PasoFetch
    PasoFetch
    PasoFetch
    Comparar "traza PC=05 RAM=A0h", MAR, 5, MDR, &HA0, IR, &HA0, PC, 6

    ' Borde: PC=FFh. Despues del +1 el PC vuelve a 00h.
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
