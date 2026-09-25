Attribute VB_Name = "modMemory"
Option Explicit

' RAM de 256 bytes. Direcciones 00h-FFh. Un byte por celda.
' En la ventana Inmediato: PruebaModMemory
' Casos: lectura/escritura normal, 00h, FFh, direccion invalida, valor fuera de 8 bits.

Public RAM(0 To 255) As Byte

Private Const ERR_DIRECCION As Long = vbObjectError + 5101
Private Const ERR_VALOR As Long = vbObjectError + 5102

Public Function ReadMem(ByVal address As Long) As Byte
    If Not DireccionValida(address) Then
        Err.Raise ERR_DIRECCION, "modMemory", "Direccion fuera de 00h-FFh: " & address
    End If
    ReadMem = RAM(address)
End Function

Public Sub WriteMem(ByVal address As Long, ByVal value As Long)
    If Not DireccionValida(address) Then
        Err.Raise ERR_DIRECCION, "modMemory", "Direccion fuera de 00h-FFh: " & address
    End If
    If Not ValorValido(value) Then
        Err.Raise ERR_VALOR, "modMemory", "Valor fuera de 8 bits: " & value
    End If
    RAM(address) = CByte(value)
End Sub

Public Sub ClearMem()
    Dim i As Long
    For i = 0 To 255
        RAM(i) = 0
    Next i
End Sub

Public Sub LoadBlock(ByVal startAddress As Long, ByRef bytes() As Byte)
    Dim i As Long
    Dim n As Long
    Dim ultimo As Long

    If (UBound(bytes) - LBound(bytes) + 1) <= 0 Then Exit Sub
    n = UBound(bytes) - LBound(bytes) + 1

    If Not DireccionValida(startAddress) Then
        Err.Raise ERR_DIRECCION, "modMemory", "Direccion fuera de 00h-FFh: " & startAddress
    End If

    ultimo = startAddress + n - 1
    If Not DireccionValida(ultimo) Then
        Err.Raise ERR_DIRECCION, "modMemory", "El bloque se sale de la RAM"
    End If

    For i = 0 To n - 1
        RAM(startAddress + i) = bytes(LBound(bytes) + i)
    Next i
End Sub

Public Sub PruebaModMemory()
    Dim guardado As Byte

    On Error GoTo Fallo
    ClearMem

    WriteMem 16, 42
    Debug.Print "normal", "esperado 42", "obtenido " & ReadMem(16)

    WriteMem 0, 7
    Debug.Print "00h", "esperado 7", "obtenido " & ReadMem(0)

    WriteMem &HFF, &HFF
    Debug.Print "FFh", "esperado 255", "obtenido " & ReadMem(255)

    guardado = RAM(1)

    On Error Resume Next
    Err.Clear
    WriteMem -1, 1
    Debug.Print "direccion invalida", "esperado error", "obtenido " & Err.Description, "RAM(1)=" & RAM(1)

    Err.Clear
    WriteMem 1, 256
    Debug.Print "valor fuera de 8 bits", "esperado error", "obtenido " & Err.Description, "RAM(1)=" & RAM(1)
    On Error GoTo Fallo

    If RAM(1) <> guardado Then
        Debug.Print "FALLO: la RAM cambio ante una entrada invalida"
    Else
        Debug.Print "la RAM no cambio ante entradas invalidas"
    End If

    Exit Sub
Fallo:
    Debug.Print "FALLO inesperado", Err.Number, Err.Description
End Sub

Private Function DireccionValida(ByVal address As Long) As Boolean
    DireccionValida = (address >= 0 And address <= 255)
End Function

Private Function ValorValido(ByVal value As Long) As Boolean
    ValorValido = (value >= 0 And value <= 255)
End Function
