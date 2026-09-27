# Notas de prueba — ISA completa y ciclo STEP/RUN

Prueba automatizada: `PruebaISACompleta` en `modPruebas` (ventana Inmediato).

## 1. Doce familias ISA (aislado)

Cada caso usa `Preparar` → `PasoExecute` → `PasoStore` y comprueba registros, memoria y flags.

| # | Familia | Entrada | Esperado |
|---|---|---|---|
| 1 | MOV | `MOV AX, 2Ah` | AX=2Ah; ZF/CF/SF intactos |
| 2 | LOAD | RAM(80h)=3Ch; `LOAD AX, [80h]` | AX=3Ch; MAR=80h; MDR=3Ch |
| 3 | STORE | AX=55h; `STORE [81h], AX` | RAM(81h)=55h; vía MAR/MDR |
| 4 | ADD | AX=2; `ADD AX, 03h` | AX=5; ZF=0; CF=0 |
| 5 | SUB | AX=5; `SUB AX, 03h` | AX=2; CF=0 |
| 6 | INC | AX=FFh; `INC AX` | AX=0; ZF=1; CF=1 |
| 7 | DEC | BX=1; `DEC BX` | BX=0; ZF=1 |
| 8 | CMP | AX=5; `CMP AX, 05h` | ZF=1; AX sigue en 5 |
| 9 | JMP | PC=10h; `JMP 40h` | PC=40h |
| 10 | JZ | ZF=1; PC=20h; `JZ 80h` | PC=80h |
| 11 | JNZ | ZF=0; PC=30h; `JNZ A0h` | PC=A0h |
| 12 | HLT | `HLT` | Estado = HALTED |

## 2. Programa corto (mismo resultado en STEP y RUN)

```
MOV AX, 02h
ADD AX, 03h
STORE [80h], AX
HLT
```

Bytes en RAM: `10h 02h  40h 03h  30h 80h  FFh`

| Modo | Cómo | Resultado final |
|---|---|---|
| STEP | `DoStep` en bucle hasta HALTED | AX=5, RAM(80h)=5, HALTED |
| RUN | `TickRun` + `DetenerRun` (sin bloquear Excel) | Igual que STEP |

Si STEP y RUN no coinciden en AX, BX, flags o RAM(80h), la prueba falla.

## 3. Cómo repetir (ISA)

1. Importar `src/modPruebas.bas` (y módulos previos al día).
2. Ventana Inmediato: `PruebaISACompleta`
3. Éxito: `ISA completa y ciclo STEP/RUN hasta HLT OK`

## 4. Programa demostrativo en STEP y RUN (issue 27)

Prueba: `PruebaDemoStepRun` en `modPruebas`.

### Procedimiento

1. Carga el ensamblador de multiplicación (`docs/PROGRAMA_DEMO.md`) con `CargarPrograma`.
2. Datos: `[80h]=3`, `[81h]=4`, `[82h]=0`.
3. **STEP:** `DoReset` + `DoStep` hasta `HALTED`.
4. Comprueba UI (`rngPhase`, `rngMicroOp`) y que el log tenga muchas filas (una por micro).
5. Repite en **RUN** (`TickRun` + `DetenerRun`, delay 50 ms en la prueba).
6. Compara: mismo `AX`, mismo `RAM(82h)`, mismo número de filas y **mismo texto** de cada fila del LOG.

### Esperado (N=3, M=4)

| | Valor |
|---|---|
| Estado | HALTED |
| AX | 0Ch (12) |
| RAM(82h) | 0Ch |
| LOG STEP vs RUN | idéntico línea a línea |

Éxito en Inmediato: `Demo STEP y RUN: resultado y log coinciden`

### Manual en Excel

```
DoLoad
SembrarDatosDemo
DoReset
```

Luego `DoStep` repetido (mirar CPU + LOG), o `DoRun` con `rngDelay` visible. Al terminar, `? AX` y `? ReadMem(&H82)` deben ser 12; la hoja LOG debe tener el mismo recorrido que en STEP.
