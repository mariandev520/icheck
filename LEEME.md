# eCheck · Combinador de pagos

Aplicación de escritorio para seleccionar e-cheques por importe y tandas de
fecha **sin superar el monto solicitado**. Muestra el faltante que se
debe completar por transferencia. Permite cargar **varios pagos a la vez**
y lleva un **registro (JSON) de los e-cheques ya utilizados** para no volver a
usarlos. Trabaja localmente, sin cuentas ni conexión.

## Uso en Windows

1. Abrí `eCheck.exe` de la carpeta de distribución. Si recibiste un ZIP,
   extraelo primero. El ejecutable incluye Python.
2. Presioná **Abrir Excel** y seleccioná el archivo `.xlsx`.
3. Si el libro tiene varias hojas, elegí la que contiene los pagos.
4. Ingresá el importe del pago: `500.000`, `500000` o `500.000,50`, y presioná
   **Agregar pago** (o Enter). Para **varios pagos**, repetí este paso: quedan en
   la lista **Pagos cargados**, que se corrige con **Quitar seleccionado** o **Vaciar**.
5. Antes de agregar cada pago, dejá **Todos los clientes** para combinar todo el
   listado, o elegí un cliente: cada pago tiene el suyo.
6. Elegí **Fecha F** o **Fecha G** según la columna del Excel que corresponda.
   Por defecto se usa G; el archivo no identifica cuál es emisión o vencimiento.
7. Elegí **Desde** y **Hasta** con los calendarios `…`, o escribí `dd/mm/aaaa`.
   Los extremos se incluyen. Para una fecha exacta, usá la misma en ambos campos.
   Vacío significa sin límite; **Limpiar fechas** quita ambos límites.
8. Dejá **Progresiva · mayor a menor** y, si lo deseás, activá
   **Priorizar la tanda de una semana anterior**.
9. Presioná **Buscar combinación**. Si dejaste un importe escrito sin agregar,
   se suma como último pago. Con un solo pago alcanza con escribir el importe y buscar.
10. Revisá el resultado y presioná **Registrar pagos** para guardar esos e-cheques
    como utilizados (ver más abajo). Buscar solo calcula: no registra nada.

El resultado muestra, por cada pago, el importe solicitado, el total seleccionado,
el faltante por transferencia y el detalle de cada e-cheque, en el orden de selección.
Arriba se suman los totales de todos los pagos. La tabla muestra la tanda semanal y
el saldo restante después de cada cheque. **Copiar mensaje** prepara
el texto para pegarlo donde corresponda (una línea por pago); el programa no envía
mensajes. **Exportar CSV** guarda los pagos, la selección y los totales para abrirlos en Excel.

## Varios pagos y registro de e-cheques utilizados

**Varios pagos.** Se calculan en el orden de la lista. Un pago puede usar un
e-cheque o varios, pero **un e-cheque elegido para un pago no se ofrece a los
siguientes**. El cliente, las fechas y el criterio se aplican a cada pago sobre lo
que quedó libre. El orden puede cambiar el resultado: el primer pago elige antes
que los demás, y las búsquedas de cada pago siguen siendo las de siempre (no hay
una optimización conjunta). Si detenés una búsqueda con varios pagos, los pagos
que no llegaron a calcularse figuran como **Sin calcular** y no se puede registrar el lote.

**Registro.** Al presionar **Registrar pagos** (con confirmación), los e-cheques
elegidos se guardan en `registro_echeques.json`, junto a `eCheck.exe`. El archivo
se crea al registrar el primer pago y es un JSON legible: por cada pago guarda
número, fecha y hora, archivo y hoja de origen, cliente (si filtraste), importe
solicitado, total, faltante y los datos de cada e-cheque. Los importes figuran como
texto con dos decimales (`"500000.00"`). Desde ese momento esos e-cheques **no se
vuelven a ofrecer**, tampoco al cerrar el programa ni al cargar otro Excel.

- Un e-cheque se reconoce por su **referencia (C), cliente, importe y fechas F y G**;
  no por el número de fila ni por el recibo. Por eso sigue bloqueado aunque el Excel
  nuevo traiga las filas en otro orden. Si hay varios registros idénticos en todo eso,
  se descuentan tantos como veces se usaron.
- **Registro…** muestra los pagos registrados y sus e-cheques. **Anular pago
  seleccionado** devuelve sus e-cheques a la lista de disponibles; el pago queda
  en el historial marcado como anulado. Es la forma de corregir un registro equivocado.
- Después de registrar, la lista de pagos se vacía (no se pueden recalcular por error),
  pero el resultado sigue en pantalla para copiar el mensaje o exportar.
- Si el archivo de registro está dañado o fue editado a mano con un formato inválido,
  el programa **no busca** y no lo sobrescribe: avisa cuál es el archivo para que lo
  corrijas o lo muevas. Así nunca se reutilizan e-cheques sin aviso.
- **Al actualizar el programa** (extraer un ZIP nuevo en otra carpeta), copiá
  `registro_echeques.json` junto al nuevo `eCheck.exe`. Si la carpeta no lo tiene, el
  programa muestra «Registro nuevo» y no tendrá en cuenta lo usado antes. Conviene
  hacer copia del archivo de vez en cuando.
- Cada instalación lleva su propio registro. Si varias personas comparten el archivo
  (por ejemplo desde una carpeta de red), que no registren al mismo tiempo. Al
  registrar se comprueba que los e-cheques no hayan sido tomados mientras tanto.

Ejemplo: para cubrir $ 500.000,00, si la mejor combinación es $ 450.000,00,
el programa informa $ 50.000,00 a completar por transferencia.

## Cómo se elige cada e-cheque

En **Progresiva · mayor a menor**, primero se busca un solo e-cheque del importe
exacto. Si no existe, se toma el de mayor monto que no supere lo solicitado.
Con importes iguales se prefiere la fecha más reciente, dejando anteriores
disponibles para completar. Luego se repite sobre el saldo pendiente:

1. Un importe exacto al saldo tiene prioridad, siempre respetando el orden
   descendente de montos: el siguiente no puede superar al último elegido.
2. Con la preferencia semanal activada, se busca el mayor importe que entre
   en la semana inmediatamente anterior disponible. Si esa semana no tiene
   cheques utilizables, se retrocede a la siguiente semana anterior disponible.
3. Si no hay una tanda anterior utilizable, se toma el mayor importe disponible
   que respete el saldo y el orden descendente. Puede pertenecer a la misma
   semana o a otra posterior. Sin la preferencia semanal, se busca directamente
   por importe en cada paso.

Las tandas son semanas calendario, de lunes a domingo, según F o G. Ejemplo:
para **$1.000.000**, si no hay uno exacto, **$900.000** de la semana del 21/09,
**$60.000** de la del 14/09 y **$40.000** de la del 07/09 completan el importe.
La preferencia de fecha puede elegir un cheque menor que otro de una semana
diferente. Las fechas originales no se cambian.

Esta selección prioriza los criterios anteriores: **no garantiza la menor
cantidad de cheques ni el menor faltante global**. Por ejemplo, para $1.100
con cheques de $800, $700, $600 y $500, la progresiva elige $800 y deja $300;
**Mejor suma posible** elige $600 + $500. Ese segundo modo conserva la búsqueda
exacta original, aplica los mismos filtros de cliente/fecha y no prioriza tandas.

Si se detiene una búsqueda larga, se conserva una selección válida y se indica
que se detuvo. Una progresiva que termina normalmente se identifica como tal,
sin confundirla con una búsqueda detenida. Solo el modo **Mejor suma posible**
confirma el mejor total global cuando termina; no minimiza la cantidad en empates.

## Formato del Excel

El formato esperado es el del archivo analizado. Puede tener encabezados o
comenzar directamente con los registros; las columnas deben conservarse.

| Columna | Dato |
| --- | --- |
| C | Referencia o número de e-cheque |
| F | Primera fecha |
| G | Segunda fecha |
| J | ID del cliente |
| R | Nombre del cliente |
| V | Importe positivo |
| W | Referencia del recibo |

Las fechas se llaman **Fecha F** y **Fecha G** porque el archivo original no
indica su significado. Fechas marcadas como `50/08/2025` se muestran tal cual
y no excluyen el pago cuando no hay límites de fecha. **Con Desde o Hasta,
una fecha vacía o inválida en la columna elegida queda fuera del rango** y se
informa cuántos registros se excluyeron por ese motivo. Sin rango, pueden
seleccionarse por importe, pero no se les inventa una tanda semanal.
Los ID guardados como texto conservan sus ceros iniciales
y su longitud, incluido `0419`; los ID numéricos se muestran con al menos cinco
dígitos. Los importes enteros también se incluyen.

Las filas sin importe válido o sin ID numérico se excluyen y se detallan al
cargar el archivo. Si V contiene una fórmula, se usa su resultado guardado por
Excel y se deja una observación; guardá el libro recalculado antes de importarlo.
No se admiten archivos `.xls`, archivos protegidos con contraseña ni importes
reales con más de dos decimales. Máximo: 20.000 registros por hoja.

Cada fila se usa como máximo una vez en una búsqueda, y no se repite entre los
pagos de la misma búsqueda. Registros con iguales fechas, cliente e importe se
conservan como cheques separados. La selección se hace sobre la hoja, el cliente
y el rango de fechas elegidos, entre los e-cheques que el registro no marca como
utilizados. Cambiar las fechas, el criterio o la lista de pagos invalida el
resultado anterior para evitar copiar o exportar una selección desactualizada;
escribir el importe del próximo pago no lo invalida. Al cargar otro archivo u hoja
se limpian los límites de fecha y la lista de pagos.

El Excel original no se modifica y los datos no se envían a servicios externos.
El único archivo que el programa escribe por su cuenta es `registro_echeques.json`.

## Ejecutar desde el código fuente

Con Python 3.8 o posterior y Tcl/Tk instalados, abrí `Iniciar.cmd` o ejecutá:

```powershell
py -3 app.py
```

En el ZIP de entrega, estos archivos están en la carpeta `codigo`.

La aplicación no necesita paquetes adicionales. Para correr las pruebas:

```powershell
py -3 -m unittest -v test_core test_registro test_app
```

## Volver a compilar el ejecutable y el ZIP

Después de guardar los cambios en `app.py`, `core.py` o `registro.py`, cerrá el programa
eCheck y hacé doble clic en **`Compilar.cmd`**. Mantené abierta la consola
hasta que aparezca **LISTO**.

El proceso corre las pruebas, compila el diseño actualizado, verifica que
el ejecutable funcione y genera estos archivos:

- **`dist/eCheck.exe`**: programa actualizado.
- **`entrega/eCheck-Windows-x64.zip`**: paquete completo para copiar a otras PC.

El ZIP también incluye el código actual y los archivos para volver a compilar,
dentro de `codigo`. No incluye los Excel de pagos. Si hay un error en las
pruebas o la compilación, se conserva el ZIP anterior. Ni el registro ni los Excel
van dentro del ZIP.

En esta PC el entorno de compilación ya está preparado. En otra PC, necesitás
Python de 64 bits con Tcl/Tk; la primera compilación descarga PyInstaller.
Para ejecutar el programa terminado no hace falta instalar Python.

Desde PowerShell también podés ejecutar:

```powershell
.\Compilar.cmd
```

Para distribuir una actualización, extraé el ZIP nuevo y abrí el `eCheck.exe`
de esa carpeta. Los ejecutables extraídos de un ZIP anterior no se actualizan solos.

### Compilar solamente el ejecutable de forma manual

En Windows de 64 bits, desde esta carpeta:

```powershell
py -3 -m venv .build-env
.\.build-env\Scripts\python.exe -m pip install pyinstaller==6.16.0
$env:PYINSTALLER_CONFIG_DIR = Join-Path (Get-Location) '.build-env\pyinstaller-cache'
.\.build-env\Scripts\python.exe -m PyInstaller --noconfirm --clean --onefile --windowed --name eCheck app.py
```

El resultado queda en `dist/eCheck.exe`. La generación requiere descargar
PyInstaller una vez; el uso del programa es sin conexión. El ejecutable
se debe verificar en las máquinas de destino Windows 10/11 de 64 bits.

El cálculo usa enteros en centavos. La progresiva utiliza índices de importes
y semanas y se prueba con 20.000 filas. **Mejor suma posible** usa la búsqueda
exacta con memoria acotada; para casos grandes puede tardar más.
**Detener búsqueda** recupera una solución válida sin presentarla como un
óptimo confirmado.

Documentación del empaquetador:
[PyInstaller](https://pyinstaller.org/en/stable/operating-mode.html).
