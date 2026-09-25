# ISA del simulador

Conjunto de instrucciones del CPU de 8 bits. Una instrucción ocupa **1 byte** si no lleva dato, o **2 bytes** si lleva un inmediato o una dirección. El segundo byte se trae con otro Fetch.

El nibble alto agrupa la familia, para poder explicarla en la defensa: `1` mueve, `2` carga, `3` guarda, `4` suma, `5` resta, `6` incrementa o decrementa, `7` compara, `8` salta, `9` salta si cero, `A` salta si no es cero, `F` parada.

`reg` es AX o BX. `imm` y `dir` son un byte, escritos en hexadecimal en la hoja PROGRAM (`05h`, `80h`).

MOV, LOAD, STORE, JMP, JZ, JNZ y HLT no cambian ZF, CF ni SF.

## Tabla mínima

Tal como se escribirá una instrucción por fila en la hoja PROGRAM.

| Sintaxis en PROGRAM | Opcode | Bytes | Operandos | Acción | Flags |
|---|---|---:|---|---|---|
| `MOV AX, imm` | 10h | 2 | inmediato | AX ← imm | no cambian |
| `MOV BX, imm` | 11h | 2 | inmediato | BX ← imm | no cambian |
| `MOV AX, BX` | 12h | 1 | registro | AX ← BX | no cambian |
| `MOV BX, AX` | 13h | 1 | registro | BX ← AX | no cambian |
| `LOAD AX, [dir]` | 20h | 2 | dirección | AX ← RAM[dir], pasando por MAR y MDR | no cambian |
| `LOAD BX, [dir]` | 21h | 2 | dirección | BX ← RAM[dir], pasando por MAR y MDR | no cambian |
| `STORE [dir], AX` | 30h | 2 | dirección | RAM[dir] ← AX, pasando por MAR y MDR | no cambian |
| `STORE [dir], BX` | 31h | 2 | dirección | RAM[dir] ← BX, pasando por MAR y MDR | no cambian |
| `ADD AX, imm` | 40h | 2 | inmediato | AX ← AX + imm | ZF, CF, SF |
| `ADD AX, BX` | 41h | 1 | registro | AX ← AX + BX | ZF, CF, SF |
| `ADD BX, imm` | 42h | 2 | inmediato | BX ← BX + imm | ZF, CF, SF |
| `ADD BX, AX` | 43h | 1 | registro | BX ← BX + AX | ZF, CF, SF |
| `SUB AX, imm` | 50h | 2 | inmediato | AX ← AX − imm | ZF, CF, SF |
| `SUB AX, BX` | 51h | 1 | registro | AX ← AX − BX | ZF, CF, SF |
| `SUB BX, imm` | 52h | 2 | inmediato | BX ← BX − imm | ZF, CF, SF |
| `SUB BX, AX` | 53h | 1 | registro | BX ← BX − AX | ZF, CF, SF |
| `INC AX` | 60h | 1 | registro | AX ← AX + 1 | ZF, CF, SF |
| `INC BX` | 61h | 1 | registro | BX ← BX + 1 | ZF, CF, SF |
| `DEC AX` | 62h | 1 | registro | AX ← AX − 1 | ZF, CF, SF |
| `DEC BX` | 63h | 1 | registro | BX ← BX − 1 | ZF, CF, SF |
| `CMP AX, imm` | 70h | 2 | inmediato | Banderas de AX − imm. AX no cambia | ZF, CF, SF |
| `CMP AX, BX` | 71h | 1 | registro | Banderas de AX − BX. AX no cambia | ZF, CF, SF |
| `CMP BX, imm` | 72h | 2 | inmediato | Banderas de BX − imm. BX no cambia | ZF, CF, SF |
| `CMP BX, AX` | 73h | 1 | registro | Banderas de BX − AX. BX no cambia | ZF, CF, SF |
| `JMP dir` | 80h | 2 | dirección | PC ← dir | no cambian |
| `JZ dir` | 90h | 2 | dirección | PC ← dir si ZF = 1; si no, sigue | no cambian |
| `JNZ dir` | A0h | 2 | dirección | PC ← dir si ZF = 0; si no, sigue | no cambian |
| `HLT` | FFh | 1 | ninguno | La CPU queda en HALTED | no cambian |

En un salto de 2 bytes el PC ya avanzó al leer el opcode y la dirección. Si la condición no se cumple, el PC queda en la instrucción siguiente. Si se cumple, se reemplaza por `dir`.

## Ejemplos de filas en PROGRAM

```
MOV AX, 05h
MOV BX, AX
LOAD AX, [80h]
ADD AX, 01h
CMP AX, 00h
JZ 10h
STORE [81h], AX
HLT
```

## Extensión de la ALU

No forman parte del mínimo de la consigna. Están aquí porque la ALU ya las define y el nibble alto queda reservado: `B` AND, `C` OR, `D` XOR, `E` NOT. CF queda en 0. ZF y SF salen del resultado.

| Sintaxis en PROGRAM | Opcode | Bytes | Acción |
|---|---|---:|---|
| `AND AX, imm` | B0h | 2 | AX ← AX AND imm |
| `AND AX, BX` | B1h | 1 | AX ← AX AND BX |
| `OR AX, imm` | C0h | 2 | AX ← AX OR imm |
| `OR AX, BX` | C1h | 1 | AX ← AX OR BX |
| `XOR AX, imm` | D0h | 2 | AX ← AX XOR imm |
| `XOR AX, BX` | D1h | 1 | AX ← AX XOR BX |
| `NOT AX` | E0h | 1 | AX ← NOT AX |
| `NOT BX` | E1h | 1 | BX ← NOT BX |
