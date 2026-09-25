Attribute VB_Name = "modISA"
Option Explicit

' Tabla ISA. Una constante por opcode. Decode y LOAD PROGRAM usan esta misma tabla.
' En la ventana Inmediato: PruebaModISA

Public Type tInstruccion
    Opcode As Byte
    Mnemonic As String
    Sintaxis As String
    Bytes As Integer
    Modo As String
    Encontrada As Boolean
End Type

Public Const OP_MOV_AX_IMM As Long = &H10
Public Const OP_MOV_BX_IMM As Long = &H11
Public Const OP_MOV_AX_BX As Long = &H12
Public Const OP_MOV_BX_AX As Long = &H13
Public Const OP_LOAD_AX As Long = &H20
Public Const OP_LOAD_BX As Long = &H21
Public Const OP_STORE_AX As Long = &H30
Public Const OP_STORE_BX As Long = &H31
Public Const OP_ADD_AX_IMM As Long = &H40
Public Const OP_ADD_AX_BX As Long = &H41
Public Const OP_ADD_BX_IMM As Long = &H42
Public Const OP_ADD_BX_AX As Long = &H43
Public Const OP_SUB_AX_IMM As Long = &H50
Public Const OP_SUB_AX_BX As Long = &H51
Public Const OP_SUB_BX_IMM As Long = &H52
Public Const OP_SUB_BX_AX As Long = &H53
Public Const OP_INC_AX As Long = &H60
Public Const OP_INC_BX As Long = &H61
Public Const OP_DEC_AX As Long = &H62
Public Const OP_DEC_BX As Long = &H63
Public Const OP_CMP_AX_IMM As Long = &H70
Public Const OP_CMP_AX_BX As Long = &H71
Public Const OP_CMP_BX_IMM As Long = &H72
Public Const OP_CMP_BX_AX As Long = &H73
Public Const OP_JMP As Long = &H80
Public Const OP_JZ As Long = &H90
Public Const OP_JNZ As Long = &HA0
Public Const OP_AND_AX_IMM As Long = &HB0
Public Const OP_AND_AX_BX As Long = &HB1
Public Const OP_OR_AX_IMM As Long = &HC0
Public Const OP_OR_AX_BX As Long = &HC1
Public Const OP_XOR_AX_IMM As Long = &HD0
Public Const OP_XOR_AX_BX As Long = &HD1
Public Const OP_NOT_AX As Long = &HE0
Public Const OP_NOT_BX As Long = &HE1
Public Const OP_HLT As Long = &HFF

Private Const ERR_MNEMONICO As Long = vbObjectError + 5501

Public Function DecodeByte(ByVal opcode As Long) As tInstruccion
    Dim info As tInstruccion
    info.Encontrada = True
    info.Opcode = CByte(opcode And &HFF)

    Select Case opcode
        Case OP_MOV_AX_IMM: info = Armar(opcode, "MOV", "MOV AX, imm", 2, "IMM")
        Case OP_MOV_BX_IMM: info = Armar(opcode, "MOV", "MOV BX, imm", 2, "IMM")
        Case OP_MOV_AX_BX: info = Armar(opcode, "MOV", "MOV AX, BX", 1, "REG")
        Case OP_MOV_BX_AX: info = Armar(opcode, "MOV", "MOV BX, AX", 1, "REG")
        Case OP_LOAD_AX: info = Armar(opcode, "LOAD", "LOAD AX, [dir]", 2, "DIRECT")
        Case OP_LOAD_BX: info = Armar(opcode, "LOAD", "LOAD BX, [dir]", 2, "DIRECT")
        Case OP_STORE_AX: info = Armar(opcode, "STORE", "STORE [dir], AX", 2, "DIRECT")
        Case OP_STORE_BX: info = Armar(opcode, "STORE", "STORE [dir], BX", 2, "DIRECT")
        Case OP_ADD_AX_IMM: info = Armar(opcode, "ADD", "ADD AX, imm", 2, "IMM")
        Case OP_ADD_AX_BX: info = Armar(opcode, "ADD", "ADD AX, BX", 1, "REG")
        Case OP_ADD_BX_IMM: info = Armar(opcode, "ADD", "ADD BX, imm", 2, "IMM")
        Case OP_ADD_BX_AX: info = Armar(opcode, "ADD", "ADD BX, AX", 1, "REG")
        Case OP_SUB_AX_IMM: info = Armar(opcode, "SUB", "SUB AX, imm", 2, "IMM")
        Case OP_SUB_AX_BX: info = Armar(opcode, "SUB", "SUB AX, BX", 1, "REG")
        Case OP_SUB_BX_IMM: info = Armar(opcode, "SUB", "SUB BX, imm", 2, "IMM")
        Case OP_SUB_BX_AX: info = Armar(opcode, "SUB", "SUB BX, AX", 1, "REG")
        Case OP_INC_AX: info = Armar(opcode, "INC", "INC AX", 1, "REG")
        Case OP_INC_BX: info = Armar(opcode, "INC", "INC BX", 1, "REG")
        Case OP_DEC_AX: info = Armar(opcode, "DEC", "DEC AX", 1, "REG")
        Case OP_DEC_BX: info = Armar(opcode, "DEC", "DEC BX", 1, "REG")
        Case OP_CMP_AX_IMM: info = Armar(opcode, "CMP", "CMP AX, imm", 2, "IMM")
        Case OP_CMP_AX_BX: info = Armar(opcode, "CMP", "CMP AX, BX", 1, "REG")
        Case OP_CMP_BX_IMM: info = Armar(opcode, "CMP", "CMP BX, imm", 2, "IMM")
        Case OP_CMP_BX_AX: info = Armar(opcode, "CMP", "CMP BX, AX", 1, "REG")
        Case OP_JMP: info = Armar(opcode, "JMP", "JMP dir", 2, "DIRECT")
        Case OP_JZ: info = Armar(opcode, "JZ", "JZ dir", 2, "DIRECT")
        Case OP_JNZ: info = Armar(opcode, "JNZ", "JNZ dir", 2, "DIRECT")
        Case OP_AND_AX_IMM: info = Armar(opcode, "AND", "AND AX, imm", 2, "IMM")
        Case OP_AND_AX_BX: info = Armar(opcode, "AND", "AND AX, BX", 1, "REG")
        Case OP_OR_AX_IMM: info = Armar(opcode, "OR", "OR AX, imm", 2, "IMM")
        Case OP_OR_AX_BX: info = Armar(opcode, "OR", "OR AX, BX", 1, "REG")
        Case OP_XOR_AX_IMM: info = Armar(opcode, "XOR", "XOR AX, imm", 2, "IMM")
        Case OP_XOR_AX_BX: info = Armar(opcode, "XOR", "XOR AX, BX", 1, "REG")
        Case OP_NOT_AX: info = Armar(opcode, "NOT", "NOT AX", 1, "REG")
        Case OP_NOT_BX: info = Armar(opcode, "NOT", "NOT BX", 1, "REG")
        Case OP_HLT: info = Armar(opcode, "HLT", "HLT", 1, "NONE")
        Case Else
            info.Encontrada = False
            info.Sintaxis = ""
            info.Bytes = 0
    End Select

    DecodeByte = info
End Function

Public Function MnemonicToOpcode(ByVal sintaxis As String) As Long
    Select Case Normalizar(sintaxis)
        Case "MOV AX, IMM": MnemonicToOpcode = OP_MOV_AX_IMM
        Case "MOV BX, IMM": MnemonicToOpcode = OP_MOV_BX_IMM
        Case "MOV AX, BX": MnemonicToOpcode = OP_MOV_AX_BX
        Case "MOV BX, AX": MnemonicToOpcode = OP_MOV_BX_AX
        Case "LOAD AX, [DIR]": MnemonicToOpcode = OP_LOAD_AX
        Case "LOAD BX, [DIR]": MnemonicToOpcode = OP_LOAD_BX
        Case "STORE [DIR], AX": MnemonicToOpcode = OP_STORE_AX
        Case "STORE [DIR], BX": MnemonicToOpcode = OP_STORE_BX
        Case "ADD AX, IMM": MnemonicToOpcode = OP_ADD_AX_IMM
        Case "ADD AX, BX": MnemonicToOpcode = OP_ADD_AX_BX
        Case "ADD BX, IMM": MnemonicToOpcode = OP_ADD_BX_IMM
        Case "ADD BX, AX": MnemonicToOpcode = OP_ADD_BX_AX
        Case "SUB AX, IMM": MnemonicToOpcode = OP_SUB_AX_IMM
        Case "SUB AX, BX": MnemonicToOpcode = OP_SUB_AX_BX
        Case "SUB BX, IMM": MnemonicToOpcode = OP_SUB_BX_IMM
        Case "SUB BX, AX": MnemonicToOpcode = OP_SUB_BX_AX
        Case "INC AX": MnemonicToOpcode = OP_INC_AX
        Case "INC BX": MnemonicToOpcode = OP_INC_BX
        Case "DEC AX": MnemonicToOpcode = OP_DEC_AX
        Case "DEC BX": MnemonicToOpcode = OP_DEC_BX
        Case "CMP AX, IMM": MnemonicToOpcode = OP_CMP_AX_IMM
        Case "CMP AX, BX": MnemonicToOpcode = OP_CMP_AX_BX
        Case "CMP BX, IMM": MnemonicToOpcode = OP_CMP_BX_IMM
        Case "CMP BX, AX": MnemonicToOpcode = OP_CMP_BX_AX
        Case "JMP DIR": MnemonicToOpcode = OP_JMP
        Case "JZ DIR": MnemonicToOpcode = OP_JZ
        Case "JNZ DIR": MnemonicToOpcode = OP_JNZ
        Case "AND AX, IMM": MnemonicToOpcode = OP_AND_AX_IMM
        Case "AND AX, BX": MnemonicToOpcode = OP_AND_AX_BX
        Case "OR AX, IMM": MnemonicToOpcode = OP_OR_AX_IMM
        Case "OR AX, BX": MnemonicToOpcode = OP_OR_AX_BX
        Case "XOR AX, IMM": MnemonicToOpcode = OP_XOR_AX_IMM
        Case "XOR AX, BX": MnemonicToOpcode = OP_XOR_AX_BX
        Case "NOT AX": MnemonicToOpcode = OP_NOT_AX
        Case "NOT BX": MnemonicToOpcode = OP_NOT_BX
        Case "HLT": MnemonicToOpcode = OP_HLT
        Case Else
            Err.Raise ERR_MNEMONICO, "modISA", "Mnemonico desconocido: " & sintaxis
    End Select
End Function

Public Sub PruebaModISA()
    Dim info As tInstruccion
    Dim fallos As Long

    On Error GoTo Fallo
    fallos = fallos + Verificar("MOV AX, imm", OP_MOV_AX_IMM, 2)
    fallos = fallos + Verificar("HLT", OP_HLT, 1)
    If MnemonicToOpcode("hlt") <> OP_HLT Then
        Debug.Print "FALLO hlt debia valer FFh"
        fallos = fallos + 1
    Else
        Debug.Print "hlt", "esperado FFh", "opcode " & Hex$(MnemonicToOpcode("hlt")) & "h"
    End If
    fallos = fallos + Verificar("JZ dir", OP_JZ, 2)
    fallos = fallos + Verificar("STORE [dir], BX", OP_STORE_BX, 2)
    fallos = fallos + Verificar("NOT AX", OP_NOT_AX, 1)

    info = DecodeByte(0)
    If info.Encontrada Then
        Debug.Print "FALLO opcode 00h debia ser desconocido"
        fallos = fallos + 1
    Else
        Debug.Print "opcode 00h", "esperado desconocido", "Encontrada=False"
    End If

    On Error Resume Next
    Err.Clear
    MnemonicToOpcode "NOP"
    If Err.Number = 0 Then
        Debug.Print "FALLO NOP debia dar error"
        fallos = fallos + 1
    Else
        Debug.Print "NOP", "esperado error", Err.Description
    End If
    On Error GoTo Fallo

    If fallos = 0 Then
        Debug.Print "la tabla cierra: byte y mnemonico apuntan al mismo opcode"
    Else
        Debug.Print "FALLOS", fallos
    End If
    Exit Sub
Fallo:
    Debug.Print "FALLO inesperado", Err.Number, Err.Description
End Sub

Private Function Verificar(ByVal sintaxis As String, ByVal opcode As Long, ByVal bytes As Integer) As Long
    Dim info As tInstruccion
    info = DecodeByte(opcode)
    If MnemonicToOpcode(sintaxis) <> opcode Or info.Bytes <> bytes Or info.Sintaxis <> sintaxis Or Not info.Encontrada Then
        Debug.Print "FALLO", sintaxis, "esperado " & Hex$(opcode), "obtenido " & info.Sintaxis
        Verificar = 1
    Else
        Debug.Print sintaxis, "opcode " & Hex$(opcode) & "h", "bytes " & bytes, "modo " & info.Modo
        Verificar = 0
    End If
End Function

Private Function Armar(ByVal opcode As Long, ByVal mnemonico As String, ByVal sintaxis As String, ByVal bytes As Integer, ByVal modo As String) As tInstruccion
    Dim info As tInstruccion
    info.Opcode = CByte(opcode And &HFF)
    info.Mnemonic = mnemonico
    info.Sintaxis = sintaxis
    info.Bytes = bytes
    info.Modo = modo
    info.Encontrada = True
    Armar = info
End Function

Private Function Normalizar(ByVal texto As String) As String
    Dim i As Long
    Dim salida As String
    Dim anterior As String
    texto = UCase$(Trim$(texto))
    anterior = " "
    For i = 1 To Len(texto)
        If Mid$(texto, i, 1) = " " Then
            If anterior <> " " Then salida = salida & " "
            anterior = " "
        Else
            salida = salida & Mid$(texto, i, 1)
            anterior = Mid$(texto, i, 1)
        End If
    Next i
    Normalizar = Trim$(salida)
End Function
