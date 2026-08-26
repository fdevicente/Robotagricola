# Diseño: documentos en Google Drive

**Fecha:** 2026-08-25
**Estado:** aprobado, pendiente de plan de implementación

## Objetivo

Sacar los documentos del disco del PC y llevarlos a Google Drive, de modo que:

1. cada factura, boleta o guía que llega quede archivada en Drive y enlazada desde
   la base de datos;
2. los respaldos del Master y de la base vayan a Drive;
3. exista una carpeta donde el dueño suelte archivos y el robot los procese solo.

Esto va **antes** del salto al servidor (Fase C del plan de julio) a propósito: con
los archivos ya en la nube, el VPS solo carga aplicación y base de datos, y no hay
PDFs que migrar después ni un disco que se quede chico.

## Decisiones tomadas

| Decisión | Elegido | Por qué |
|---|---|---|
| Alcance de Drive | Documentos + respaldos + carpeta de entrada | El dueño lo quiere como lugar único |
| Dropbox | Se queda **solo con el FXP** | Juan lo edita ahí; cambiarle la rutina arriesga el archivo que alimenta el cruce con el Master |
| Cuenta | Gmail nuevo, dedicado al robot | Separa los documentos de la empresa de la cuenta personal, sin costo |
| Permiso | `drive` completo | La carpeta de entrada lo exige: `drive.file` solo alcanza archivos que crea la propia app |
| Orden de guardado | **Local primero, Drive después** | Una caída de internet no puede costar una factura |

### Gotcha de autenticación (documentado para no repetir el error)

La primera propuesta fue usar una **cuenta de servicio**. **No sirve con un Gmail
común**: los archivos que sube quedan a nombre de la cuenta de servicio, que no
tiene cuota de Drive en cuentas de consumidor, y la subida falla. Ese camino solo
funciona con Google Workspace y unidades compartidas.

Con Gmail va **OAuth**: autorización única desde el navegador, token de refresco
guardado local (fuera del repositorio, junto a `.env`). Como el permiso `drive` es
sensible y la app no está verificada por Google, la primera autorización muestra
una pantalla de advertencia. Se acepta una vez; es una herramienta interna de un
solo usuario.

## Arquitectura

```
   Telegram ──┐
              ├──► Robot (PC) ──► Google Drive  (documentos + respaldos)
   _Entrada ──┘         │              │
                        │              └── devuelve el ID del archivo
                        ▼
                  Base de datos ◄── guarda el ENLACE, nunca el archivo
```

## Estructura en Drive

```
Agrícola Santa Elisa/
├── Facturas Recibidas/2024/ 2025/ 2026/     817 archivos, 111 MB
├── Facturas Enviadas/                        19 archivos, 4 MB
├── Boletas Honorarios/                       60 archivos, 10 MB
├── Guías de Despacho/                        16 archivos, 4 MB
├── Rendiciones/                               9 archivos, 53 MB
├── Legal/                                    26 archivos, 18 MB
├── Tributario/                                5 archivos, 1 MB
├── Reportes/                                 PDF mensuales
├── Respaldos/
│   ├── Master/
│   └── Base de datos/
└── _Entrada/
    └── Sin procesar/                         lo que el robot no supo leer
```

Nombres: se mantiene `Proveedor_NroFactura.ext`, que `_renombrar_archivo` ya aplica.

**Volumen actual: 201 MB** de 15 GB (1,3%). Crecimiento de facturas ~30 MB/año.

> Medido originalmente en 2,3 GB. El 26-ago el dueño borró una carpeta que se
> había colado dentro de `Legal`, que pasó de 1,6 GB a 18 MB. La migración dejó
> de ser un problema de volumen.

## Flujos

### 1. Documento que llega por Telegram

Llega → la IA extrae los campos (ya existe) → se renombra (ya existe) → **se guarda
en el disco local** → se encola la subida → se sube a Drive → se guarda el enlace en
una columna nueva de la hoja `Facturas`.

El aviso a quien lo mandó **no espera la subida**: el documento ya está a salvo en
disco cuando se responde.

### 2. Respaldos

`infrastructure/backups.py` deja de copiar el Master a Dropbox y pasa a subirlo a
Drive, junto con la base de datos.

**Retención**: diarios de los últimos 30 días · mensuales del año en curso · anuales
para siempre. Sin esto, un Master de ~530 KB diario suma ~190 MB al año de copias
casi idénticas.

### 3. Carpeta de entrada

El robot revisa `_Entrada/` cada pocos minutos. Cada archivo nuevo se procesa como
si hubiera llegado por Telegram y **se mueve a su carpeta definitiva**. Ese
movimiento ES la marca de procesado: no hace falta llevar una lista aparte, y un
reinicio a mitad de camino no duplica nada.

## Manejo de errores

| Falla | Comportamiento |
|---|---|
| Sin internet | El archivo ya está en disco. Queda encolado y se sube al volver. **La cola se persiste en disco**: el watchdog reinicia el bot seguido y una cola en memoria perdería lo pendiente. |
| Token vencido o revocado | Aviso por Telegram con el paso a paso para reautorizar. **Nunca fallar en silencio** — es lo que pasó el 24-ago con el parte del JD 5085. |
| Drive lleno | Aviso al 80% de los 15 GB, no cuando ya no puede subir. |
| Documento repetido | Juan reenvía cosas. Se verifica nombre y contenido antes de subir; si ya está, no se duplica. |
| Archivo ilegible en `_Entrada/` | Se mueve a `_Entrada/Sin procesar/` y se avisa cuál fue. No se queda dando vueltas ni se procesa a medias. |

## Pruebas

Con un **Drive falso, sin red**:

- una subida fallida deja el archivo local y encolado;
- la cola sobrevive un reinicio del proceso;
- la carpeta de entrada mueve el archivo **solo** después de procesarlo bien;
- un archivo ilegible termina en `Sin procesar/`;
- la retención conserva exactamente los respaldos que corresponde;
- un documento repetido no genera una segunda copia.

## Fuera de alcance

- **El FXP**: sigue en Dropbox, editado por Juan. Migrarlo o eliminarlo es un
  proyecto aparte.
- **Mover el sistema al servidor**: es la Fase C del plan de julio. Este diseño la
  facilita pero no la incluye.
- **Compartir el Drive con el contador u otras personas**: no se pidió.

## Riesgos

- **La migración son 201 MB en ~950 archivos.** El volumen dejó de ser un riesgo;
  lo que sí conviene es subirla por lotes y verificar conteos por carpeta, porque
  el riesgo real es perder la pista de un archivo, no el espacio.
- **La advertencia de "app no verificada"** puede asustar al autorizar. Es esperable
  y se acepta una sola vez.
- **La pantalla de consentimiento debe quedar en "In production".** En modo
  *Testing* el token de refresco se vence a los 7 días y el robot pediría
  reautorizar todas las semanas.
