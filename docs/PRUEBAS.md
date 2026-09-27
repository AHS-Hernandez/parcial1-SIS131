# Evidencia de pruebas — simulador CPU 8 bits

Registro consolidado de lo probado durante la semana. Cada caso indica **TEST**, **resultado esperado**, **resultado obtenido** y **corrección** (si falló).

Pruebas automatizadas en `modPruebas` (ventana Inmediato de VBA).

| Macro | Cubre |
|---|---|
| `PruebaISACompleta` | 12 familias ISA + STEP/RUN hasta HLT |
| `PruebaProgramaDemo` | Ensamblado y resultado del demo (×) |
| `PruebaDemoStepRun` | Demo completo STEP vs RUN + log idéntico |
| `PruebaCasosLimite` | Bordes 00h/FFh, flags ZF/CF/SF, zonas MEMORY |
| `PruebaFlujoHltValidaciones` | JZ/JNZ tomados y no, HLT, dir/opcode inválidos |

Si una macro imprime `ok` / mensaje final de éxito, **esperado = obtenido**. Si imprime `FALLO`, ver la corrección indicada (o reimportar el `.bas` correspondiente).

---

## 1. Doce familias ISA (aislado) — issue 22

Macro: `PruebaISACompleta`. Cada caso: `Preparar` → `PasoExecute` → `PasoStore`.

| TEST | Esperado | Obtenido | Corrección |
|---|---|---|---|
| MOV `MOV AX, 2Ah` | AX=2Ah; ZF/CF/SF intactos | iguales (`ok`) | — |
| LOAD RAM(80h)=3Ch; `LOAD AX, [80h]` | AX=3Ch; MAR=80h; MDR=3Ch | iguales | — |
| STORE AX=55h; `STORE [81h], AX` | RAM(81h)=55h vía MAR/MDR | iguales | — |
| ADD AX=2; `ADD AX, 03h` | AX=5; ZF=0; CF=0 | iguales | — |
| SUB AX=5; `SUB AX, 03h` | AX=2; CF=0 | iguales | — |
| INC AX=FFh; `INC AX` | AX=0; ZF=1; CF=1 | iguales | — |
| DEC BX=1; `DEC BX` | BX=0; ZF=1 | iguales | — |
| CMP AX=5; `CMP AX, 05h` | ZF=1; AX=5 | iguales | — |
| JMP PC=10h; `JMP 40h` | PC=40h | iguales | — |
| JZ ZF=1; PC=20h; `JZ 80h` | PC=80h | iguales | — |
| JNZ ZF=0; PC=30h; `JNZ A0h` | PC=A0h | iguales | — |
| HLT | Estado = HALTED | iguales | — |

Éxito Inmediato: `ISA completa y ciclo STEP/RUN hasta HLT OK`

---

## 2. Programa corto STEP vs RUN — issue 22

```
MOV AX, 02h
ADD AX, 03h
STORE [80h], AX
HLT
```

Bytes: `10h 02h  40h 03h  30h 80h  FFh`

| TEST | Esperado | Obtenido | Corrección |
|---|---|---|---|
| STEP (`DoStep` hasta HALTED) | AX=5, RAM(80h)=5, HALTED | iguales | — |
| RUN (`TickRun` + `DetenerRun`) | Igual que STEP | iguales | — |
| Comparar STEP vs RUN | Mismo AX, BX, flags, RAM(80h) | iguales | — |

---

## 3. Programa demostrativo — issues 23 / 27

Macros: `PruebaProgramaDemo`, `PruebaDemoStepRun`. Fuente: `docs/PROGRAMA_DEMO.md` (multiplicación por sumas).

| TEST | Esperado | Obtenido | Corrección |
|---|---|---|---|
| Bytes 00h–14h vs traducción a mano | Coinciden | `ok` | — |
| Demo STEP N=3, M=4 | HALTED, AX=0Ch, [82h]=0Ch | iguales | — |
| Demo N=0 (JZ inmediato a FIN) | HALTED, AX=0, [82h]=0 | iguales | — |
| Demo STEP (issue 27) | AX=0Ch, [82h]=0Ch, log ≥20 filas, UI fase/micro | iguales | — |
| Demo RUN delay 50 ms | Mismo resultado que STEP | iguales | — |
| LOG STEP vs RUN | Mismo nº de filas y mismo texto por micro | iguales | — |

Éxito: `Demo STEP y RUN: resultado y log coinciden` / `Programa demostrativo (multiplicacion) verificado`

### Manual en Excel

```
DoLoad
SembrarDatosDemo
DoReset
```

Luego `DoStep` repetido o `DoRun`. Al terminar: `? AX` y `? ReadMem(&H82)` = 12.

---

## 4. Casos límite — datos, flags y segmentación — issue 28

Macro: `PruebaCasosLimite`.

### 4.1 Memoria

| TEST | Esperado | Obtenido | Corrección |
|---|---|---|---|
| Escribir/leer `00h` ← `00h` | 0 | 0 | — |
| Escribir/leer `00h` ← `FFh` | 255 | 255 | — |
| Escribir/leer `FFh` ← `00h` | 0 | 0 | — |
| Escribir/leer `FFh` ← `FFh` | 255 | 255 | — |

### 4.2 Flags

| TEST | Esperado | Obtenido | Corrección |
|---|---|---|---|
| ZF: AX=1; `SUB AX, 01h` | AX=0, ZF=1, SF=0, CF=0 | iguales | — |
| CF: AX=FFh; `ADD AX, 01h` | AX=0, CF=1, ZF=1, SF=0 | iguales | — |
| SF: AX=0; `SUB AX, 01h` | AX=FFh, SF=1, CF=1, ZF=0 | iguales | — |
| INC tope AX=FFh | AX=0, CF=1, ZF=1 | iguales | — |
| DEC cero AX=0 | AX=FFh, CF=1, SF=1 | iguales | — |
| Sin acarreo AX=1; `ADD AX, 01h` | AX=2, CF=0, ZF=0, SF=0 | iguales | — |

### 4.3 Segmentación MEMORY

| TEST | Esperado | Obtenido | Corrección |
|---|---|---|---|
| Celda 00h (código) vs 80h (datos) | Colores de fondo distintos | `codigo<>datos` | — |
| Leyenda | Código 00h–7Fh / Datos 80h–FFh | Pintada por `RefreshUI` | — |

Éxito: `Casos limite 00h/FFh, ZF/CF/SF y segmentacion OK`

---

## 5. Control de flujo, HLT y validaciones — issue 29

Macro: `PruebaFlujoHltValidaciones`.

### 5.1 Saltos condicionales (tomado y no tomado)

| TEST | Entrada | Esperado | Obtenido | Corrección |
|---|---|---|---|---|
| JZ tomado | ZF=1; PC=20h; `JZ 80h` | PC=80h; ZF=1 | iguales | — |
| JZ no tomado | ZF=0; PC=22h; `JZ 90h` | PC sigue en 22h | iguales | — |
| JNZ tomado | ZF=0; PC=30h; `JNZ A0h` | PC=A0h; ZF=0 | iguales | — |
| JNZ no tomado | ZF=1; PC=33h; `JNZ B0h` | PC sigue en 33h | iguales | — |

### 5.2 HLT

| TEST | Entrada | Esperado | Obtenido | Corrección |
|---|---|---|---|---|
| HLT detiene | AX=11h, BX=22h, PC=50h, flags fijos; `HLT` | HALTED; regs y flags intactos | iguales | — |
| HLT frena Fetch | Tras HLT, `PasoFetch` | PC no avanza; sigue HALTED | iguales | — |

### 5.3 Validaciones (error controlado, sin comportamiento indefinido)

| TEST | Entrada | Esperado | Obtenido | Corrección |
|---|---|---|---|---|
| Dirección inválida escritura | `WriteMem -1, 1` (RAM(1) previa = 5Ah) | `Err` con mensaje; RAM(1) intacta | error + RAM intacta | — |
| Dirección inválida lectura | `ReadMem 256` | `Err` controlado | error | — |
| Opcode inválido (tabla ISA) | `DecodeByte(0)` | `Encontrada=False`, sintaxis vacía | iguales | — |
| Opcode inválido en Decode | IR=00h; `PasoDecode` | Estado = HALTED | HALTED | — |
| Mnemónico desconocido | `MnemonicToOpcode "NOP"` | `Err` (mnemónico desconocido) | error | — |

Éxito Inmediato: `Flujo JZ/JNZ, HLT y validaciones OK`

---

## 6. Cómo repetir todas las pruebas

1. Reimportar `src/modPruebas.bas` (y módulos de `src/` si el `.xlsm` no está al día).
2. En la ventana Inmediato, ejecutar en orden:

```
PruebaISACompleta
PruebaProgramaDemo
PruebaDemoStepRun
PruebaCasosLimite
PruebaFlujoHltValidaciones
```

3. Cada macro debe terminar con su mensaje de éxito (sin `FALLOS`).
4. Si alguna falla: anotar el `FALLO` de Immediate en la columna **Corrección**, corregir el módulo indicado y volver a importar/ejecutar.
