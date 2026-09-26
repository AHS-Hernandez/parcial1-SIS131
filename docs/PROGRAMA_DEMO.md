# Programa demostrativo

## 1. Comparación de candidatos

| Candidato | ALU | Flags | Memoria | Bucle / salto | Facilidad de defensa | Decisión |
|---|---|---|---|---|---|---|
| Multiplicación por sumas sucesivas | Alta (ADD, DEC, CMP) | ZF al acabar el conteo | N, M y producto en 80h–82h | Bucle + JZ | Alta: “sumo M, N veces” | **Elegido** |
| Fibonacci hasta desborde | Alta | CF al pasar FFh | Serie en datos | Bucle | Media: dos valores vivos | Reserva |
| Factorial | Alta | Desborda pronto (6!) | Poca | Bucle | Baja: poco recorrido útil | No |
| Cuenta regresiva | Baja | ZF en cero | Un guardado | JZ | Alta, pero poca ALU | No |

**Elegido:** multiplicación por sumas sucesivas. Demuestra ALU, flags, LOAD/STORE, bucle y bifurcación condicional en un guion corto.

## 2. Especificación

| Celda | Rol | Ejemplo de defensa |
|---|---|---|
| `[80h]` | Contador N (se decrementa) | `03h` |
| `[81h]` | Multiplicando M (fijo) | `04h` |
| `[82h]` | Producto (resultado) | `0Ch` (= 3×4) |

- **AX** acumula el producto.
- **BX** es registro de trabajo (lee N o M, o el contador al decrementar).
- Si N = 0 al entrar, JZ salta al STORE final (producto 0).

## 3. Ensamblador (hoja PROGRAM)

Una instrucción por fila (copiar tal cual a PROGRAM):

```
MOV AX, 00h
LOAD BX, [80h]
CMP BX, 00h
JZ 12h
LOAD BX, [81h]
ADD AX, BX
LOAD BX, [80h]
DEC BX
STORE [80h], BX
JMP 02h
STORE [82h], AX
HLT
```

Comentarios para narrar (no van en la hoja si el ensamblador no los acepta en la misma celda; en VBA `'` al inicio de fila se ignora):

```
' LOOP en 02h; FIN (STORE resultado) en 12h
```

## 4. Traducción a mano (bytes / hex)

| Dir | Bytes | Instrucción | Notas |
|---|---|---|---|
| 00h | `10 00` | MOV AX, 00h | Producto ← 0 |
| 02h | `21 80` | LOAD BX, [80h] | **LOOP:** BX ← N |
| 04h | `72 00` | CMP BX, 00h | ¿N = 0? |
| 06h | `90 12` | JZ 12h | Sí → FIN |
| 08h | `21 81` | LOAD BX, [81h] | BX ← M |
| 0Ah | `41` | ADD AX, BX | AX ← AX + M |
| 0Bh | `21 80` | LOAD BX, [80h] | BX ← N |
| 0Dh | `63` | DEC BX | N ← N − 1 |
| 0Eh | `31 80` | STORE [80h], BX | Escribe N |
| 10h | `80 02` | JMP 02h | Vuelve al LOOP |
| 12h | `30 82` | STORE [82h], AX | **FIN:** guarda producto |
| 14h | `FF` | HLT | |

**Imagen lineal (00h–14h):**

```
10 00 21 80 72 00 90 12 21 81 41 21 80 63 31 80 80 02 30 82 FF
```

Comprobación con `LOAD PROGRAM`: esos mismos bytes deben quedar en RAM(00h..14h).

## 5. Traza corta (N=3, M=4)

Estado inicial de datos: `[80h]=03h`, `[81h]=04h`, `[82h]=00h`.

| Vuelta | N al CMP | AX tras ADD | N tras DEC |
|---|---|---|---|
| 1 | 3 | 4 | 2 |
| 2 | 2 | 8 | 1 |
| 3 | 1 | 12 (0Ch) | 0 |
| 4 | 0 → JZ FIN | — | — |

Final: **AX = 0Ch**, **RAM(82h) = 0Ch**, **HALTED**, N en 80h quedó en 0.

## 6. Cómo cargar en el simulador

1. Pegar las 12 líneas de la sección 3 en la hoja PROGRAM (`rngProgram`).
2. En MEMORY (o por Immediate): `WriteMem &H80, 3` y `WriteMem &H81, 4`.
3. Botón **LOAD PROGRAM** (`DoLoad`) o, en Immediate: `PruebaProgramaDemo`.
4. **RESET** no borra la RAM: se puede repetir STEP/RUN sobre el mismo código.
5. **RUN** hasta HLT o **STEP** narrando Fetch → Decode → Execute/Store en el bucle.

## 7. Defensa en una frase

“Multiplico sumando M (en 81h) tantas veces como diga N (en 80h); el contador baja con DEC, el bucle corta con JZ cuando ZF=1, y el producto queda en 82h.”
