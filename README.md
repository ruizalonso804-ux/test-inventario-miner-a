# Mapa 3D de Bodega de Repuestos Mineros

**Desarrollado por Alonso Ruiz Fredes · Ingeniería Civil Industrial, UNACH**

Demo web que muestra en 3D una bodega de repuestos de una faena minera. Cada posición del rack se colorea según el riesgo de abastecimiento de su repuesto. **Todos los datos son ficticios (datos de ejemplo).**

- 10 racks (pasillos A–J) × 24 columnas × 5 niveles = 1.200 posiciones
- Corredores demarcados en el piso, zona de stage con pallets en espera, 2 puertas de recepción y 2 de despacho
- Alertas de quiebres y críticos, pedido sugerido, posiciones en riesgo, valor en riesgo y capital en sobrestock (USD)
- Modo presentación para mostrar el proyecto
- Three.js 0.170 (InstancedMesh, OrbitControls) y SheetJS 0.20.3, ambos desde CDN

## Archivos

| Archivo | Qué es |
|---|---|
| `index.html` | La aplicación completa (HTML, CSS y JS). Trae embebidos los mismos datos que `bodega_ejemplo.xlsx`. |
| `generar_datos.py` | Genera `bodega_ejemplo.xlsx` (semilla 42) y actualiza los datos embebidos en `index.html`. |
| `bodega_ejemplo.xlsx` | Excel de ejemplo: hoja `stock` con los datos y hoja `leeme` con la descripción de cada columna. |

## Cómo ejecutarlo

1. Abre la carpeta en VS Code.
2. Con la extensión **Live Server** instalada, haz clic derecho en `index.html` y elige **Open with Live Server**.
3. Se abre en el navegador con los datos de ejemplo cargados. Necesita internet para descargar Three.js y SheetJS.

Sin VS Code: `python3 -m http.server 8000` y abre `http://localhost:8000`.

Para regenerar los datos de ejemplo:

```bash
pip install pandas openpyxl
python generar_datos.py
```

## Uso

- **Encuadre:** la vista general y la planta se ajustan solas para que los 10 pasillos queden completos entre la ayuda superior y la barra inferior, con o sin el panel lateral. Se reajustan al abrir o cerrar el panel o al cambiar el tamaño de la ventana, hasta que muevas la cámara.
- **Mouse:** arrastrar para rotar, clic derecho para mover, rueda para hacer zoom. Teclas: `G` vista general, `P` planta, `F` pantalla completa, `M` mostrar u ocultar el panel lateral, `Esc` cerrar o salir de la presentación.
- **Panel lateral:** el botón del extremo izquierdo de la barra superior (o la tecla `M`) lo oculta o lo muestra, y el mapa 3D usa el espacio que queda libre. La página recuerda si lo dejaste abierto o cerrado.
- **Pantalla completa:** el botón con ícono de esquinas, a la izquierda de *Presentación*, pone toda la app en pantalla completa; se sale con el mismo botón, con `F` o con `Esc`. Combinado con el modo presentación, queda listo para proyectar.
- **Clic en una posición:** abre la ficha con SKU, descripción, criticidad, stock en esa ubicación y stock total del SKU, mínimo, consumo, lead time, cobertura, estado, **costo unitario**, valor del stock del SKU, **pedido sugerido** (en unidades y en USD) y las otras ubicaciones del mismo SKU.
- **Alertas críticas:** el panel izquierdo mezcla quiebres (de cualquier criticidad) y críticos en una sola lista. Se ordenan por criticidad (A primero) y luego por la mayor brecha entre lead time y cobertura (un quiebre tiene cobertura 0, así que su brecha es el lead time completo). Cada alerta muestra la descripción completa en dos líneas, y el texto entero en el tooltip. Un clic centra la cámara en el nicho y abre su ficha.
- **Indicadores** (barra inferior, dentro de los filtros de pasillo, nivel y criticidad):
  - *En riesgo:* posiciones en quiebre, crítico o bajo mínimo, y su % sobre las ocupadas.
  - *Valor en riesgo:* Σ pedido sugerido × `costo_unitario` de los SKU en quiebre, crítico o bajo mínimo, es decir, cuánto cuesta reponerlos.
  - *Capital en sobrestock:* Σ (stock total − consumo de 180 días) × `costo_unitario` de los SKU en sobrestock, es decir, el capital inmovilizado por sobre la cobertura máxima.
  - Pasando el mouse se ve el monto exacto. Si el archivo no trae `costo_unitario`, ambos montos muestran "—".
- **Filtros:** por pasillo, nivel y criticidad en el panel izquierdo, y por estado en la barra inferior. Un clic deja solo ese valor y Ctrl/⌘ + clic suma varios.
- **Buscador:** busca por SKU o descripción en todo el inventario (repuestos, EPP y fijación). Si no hay coincidencia exacta, tolera errores de tipeo ("barbiquejoo" encuentra "barbiquejo"). Enter o clic en un resultado vuela hasta la posición y la marca.
- **Presentación:** oculta el panel lateral, muestra durante 3 segundos una tarjeta con el nombre del proyecto, el autor y "Datos de ejemplo", y hace rotar la cámara suavemente. Si mueves la cámara, el giro se pausa y se reanuda a los 4 segundos. Se sale con el mismo botón o con `Esc`.
- **Cargar Excel…:** reemplaza los datos con tu archivo. *Usar ejemplo* restaura los originales.
- **Rótulo de datos:** la insignia, la marca de agua y la tarjeta de presentación dicen "DATOS DE EJEMPLO" con los datos embebidos. Con un Excel propio muestran su nombre tal como se llama el archivo. Si el archivo cargado tiene exactamente los datos de ejemplo (por ejemplo, uno exportado sin cambios), se rotula igual como datos de ejemplo.
- **Plantilla:** descarga un Excel con todas las ubicaciones de la bodega (pasillo, columna y nivel ya completos) para llenar.
- **Exportar:** descarga lo cargado con una fila por ubicación, más las columnas `stock_total_sku`, `cobertura_dias`, `estado`, `pedido_sugerido` y `valor_pedido_usd`. Con los datos de ejemplo, el archivo se llama `bodega_ejemplo.xlsx`. Se puede editar y volver a subir, porque las columnas extra se ignoran.

## Modelo de datos

- Cada repuesto tiene un **SKU único**. Un mismo SKU puede estar en varias ubicaciones; cada fila del Excel es una ubicación con el stock que hay **en esa ubicación**.
- El estado, la cobertura y el pedido sugerido se calculan **por SKU**, con el stock total de todas sus ubicaciones. Todas las ubicaciones de un SKU muestran el mismo color.
- Descripción, criticidad, stock mínimo, consumo y lead time son atributos del SKU. Si difieren entre filas del mismo SKU, se usan los de la primera fila y se avisa al cargar.

### Datos de ejemplo

`generar_datos.py` crea unos 420 SKU repartidos en ~1.090 posiciones (≈9 % vacías):

- **Familias:** rodamientos, sellos, revestimientos de bomba de pulpa, piezas de chancador, mangueras hidráulicas, correas y polines, filtros, componentes de camión de extracción y eléctricos (motores, variadores, contactores, sensores). EPP (guantes, lentes, respiradores, cascos, chalecos, etc.) y elementos de fijación suman cerca de 11 % de las posiciones.
- **Criticidad A** solo para repuestos de sistemas de **bombas, chancador, hidráulica, eléctrica o transmisión** en equipos críticos (chancador primario, molinos, bomba de alimentación a ciclones, camión de extracción, pala eléctrica y correa overland). EPP, fijaciones y consumibles generales (filtros, polines, fusibles, etc.) quedan siempre en B o C. El script lo verifica al final ("SKU A fuera de sistemas/equipos permitidos: 0").
- **Costo unitario** ficticio en USD por SKU, según el tipo de repuesto: desde menos de US$ 1 (golillas) hasta más de US$ 100.000 (mantos de chancador).
- **Quiebres y críticos** concentrados en ítems A y B. Las piezas pesadas (chancador y bombas de pulpa) van en los niveles 1 y 2.

## Formato del Excel

El archivo debe tener una hoja llamada **`stock`** con estas columnas (no importa el orden, las mayúsculas, los tildes ni los espacios; por ejemplo, `Lead time (días)` se reconoce como `lead_time_dias`):

| Columna | Tipo | Ejemplo |
|---|---|---|
| `pasillo` | Letra A–J | `C` |
| `columna` | Entero 1–24 | `7` |
| `nivel` | Entero 1–5 (1 = piso) | `3` |
| `sku` | Texto | `ROD-10234` |
| `descripcion` | Texto | `Rodamiento de bolas 6316-2RS · Bomba agua de proceso` |
| `criticidad` | `A`, `B` o `C` | `A` |
| `stock_actual` | Número ≥ 0 (stock en esa ubicación) | `4` |
| `stock_minimo` | Número ≥ 0 | `6` |
| `consumo_mensual` | Número ≥ 0 | `2,5` o `2.5` |
| `lead_time_dias` | Número ≥ 0 | `90` |
| `costo_unitario` | Número ≥ 0 en USD (**opcional**) | `1250` |

- **Decimales:** se aceptan con coma o con punto (`2,5`, `2.5`, `1.234,5`, `1,234.5`).
- **Encabezados:** pueden estar debajo de filas de título; se buscan en las primeras 10 filas.
- Una posición que no aparece en la hoja, o que no tiene SKU, se muestra como **vacía**.
- Si falta una columna, la página dice cuál falta y sugiere qué columna podría ser (por ejemplo, `lead_time` → `lead_time_dias`). El mapa se mantiene como estaba.
- Las filas inválidas se omiten y se informan con el número de fila de Excel y el motivo: pasillo inexistente, columna o nivel fuera de rango, criticidad distinta de A/B/C, número ilegible o número negativo.
- Si una ubicación aparece repetida, se usa la última fila válida.
- `costo_unitario` es opcional. Si falta, el mapa funciona igual y los indicadores de valor muestran "—" con un aviso. Una celda vacía significa costo desconocido para ese SKU, que no se suma.
- Las columnas adicionales se ignoran.

### Reglas de estado (por SKU)

`cobertura_dias = stock_total / (consumo_mensual / 30)`

Se evalúan en este orden y se aplica la primera que se cumple:

| Estado | Condición | Color |
|---|---|---|
| Quiebre | stock total = 0 | rojo intenso `#D7191C` |
| Crítico | criticidad A y cobertura < lead_time_dias | ámbar `#F5A300` |
| Bajo mínimo | stock total < stock_minimo | amarillo `#F0E442` |
| Sobrestock | cobertura > 180 días | azul cielo `#56B4E9` |
| Normal | ninguna de las anteriores | verde `#009E73` |
| Vacío | sin repuesto | gris `#5B6573` |

Quiebre (rojo oscuro) y crítico (ámbar claro) difieren en tono y en luminosidad, así que se distinguen aun con daltonismo. Además, los **quiebres de criticidad A** llevan un **borde rojo brillante** en el mapa, una señal de forma que no depende del color, explicada en la leyenda de la barra inferior. Cada estado también se nombra en la leyenda, el tooltip y la ficha. Un SKU con `consumo_mensual = 0` tiene cobertura infinita ("Sin consumo").

### Pedido sugerido

`pedido = consumo_mensual / 30 × (lead_time_dias + 15) − stock_total`, con mínimo 0 y redondeado hacia arriba a unidades enteras. Los 15 días de seguridad se configuran en `CONFIG.DIAS_SEGURIDAD`.

## Cambiar los parámetros de la bodega

Al inicio del `<script type="module">` en `index.html` está el objeto `CONFIG`:

```js
const CONFIG = {
  NUM_PASILLOS: 10,      // racks A, B, C...
  NUM_COLUMNAS: 24,
  NUM_NIVELES: 5,
  ANCHO_NICHO: 1.3,      // metros
  ALTO_NICHO: 1.0,
  PROF_RACK: 1.1,
  ANCHO_PASILLO: 3.2,
  ANCHO_STAGE: 6,
  DIAS_SOBRESTOCK: 180,
  DIAS_SEGURIDAD: 15,
  MAX_ALERTAS: 10,
  // ...puertas, muros y nombre del archivo de ejemplo
};
```

La geometría, las etiquetas, los filtros y las validaciones del Excel se recalculan a partir de estos valores. Si cambias la cantidad de pasillos, columnas o niveles, cambia también `NUM_PASILLOS`, `NUM_COLUMNAS` y `NUM_NIVELES` en `generar_datos.py` y vuelve a ejecutarlo. Si agregas pasillos, agrega también entradas en `FAMILIAS_POR_PASILLO`; si no, el generador reutiliza las existentes de forma cíclica.

## Rendimiento

Las 1.200 posiciones se dibujan con un solo `InstancedMesh`; los postes, las vigas, los corredores y los pallets también usan instancias, así que la escena usa pocas llamadas de dibujo. La página solo vuelve a dibujar cuando cambia algo (cámara, hover, filtros o la rotación del modo presentación), lo que ahorra batería en notebooks.

Las etiquetas HTML de los pasillos se ocultan automáticamente cuando se superponen en pantalla. La letra pintada en el piso de cada cabecera siempre queda visible.
