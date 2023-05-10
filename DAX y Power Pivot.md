# ¿Qué es DAX?
DAX (Data Analysis Expressions) es un lenguaje de fórmulas utilizado en Power Pivot, Power BI y otras herramientas de análisis de datos de Microsoft. Fue desarrollado específicamente para realizar cálculos y análisis de datos en modelos de datos tabulares.

DAX se basa en fórmulas y expresiones similares a las que se encuentran en Excel, pero con una funcionalidad más avanzada y orientada al análisis de datos. A través de DAX, puedes crear cálculos personalizados, medidas, columnas calculadas y tablas calculadas en tus modelos de datos.

Aquí hay algunas características clave de DAX:

1. **Funciones DAX:** DAX ofrece una amplia gama de funciones predefinidas que puedes utilizar para realizar cálculos y manipulaciones de datos. Hay funciones para operaciones matemáticas, manipulación de cadenas de texto, fechas y horas, agregaciones, filtrado de datos, entre otros. Algunas funciones comunes incluyen SUM, AVERAGE, MAX, MIN, COUNT, IF, RELATED, CALCULATE, entre muchas otras.

2. **Columnas calculadas:** En Power Pivot, puedes crear columnas calculadas utilizando DAX. Estas columnas se calculan en función de una fórmula o expresión definida por el usuario. Las columnas calculadas pueden realizar cálculos basados en los valores de otras columnas en la misma tabla y se calculan automáticamente cada vez que se actualiza el modelo de datos.

3. **Medidas:** Las medidas en DAX son cálculos agregados que se utilizan en tablas dinámicas y visualizaciones para resumir los datos. Puedes utilizar funciones de agregación, como SUM, AVERAGE, COUNT, para calcular valores resumidos basados en las interacciones del usuario con las tablas dinámicas o filtros aplicados. Las medidas son especialmente útiles para realizar análisis y crear indicadores clave de rendimiento (KPI).

4. **Contexto de filtro:** DAX tiene la capacidad de aplicar filtros contextuales a los cálculos. Puedes especificar filtros basados en las relaciones entre tablas, jerarquías y los valores seleccionados en una tabla dinámica o visualización. Esto permite realizar cálculos más precisos y contextuales, ajustándose dinámicamente a los filtros y selecciones realizados por el usuario.

DAX proporciona una gran flexibilidad y potencia para realizar cálculos avanzados en tus modelos de datos. A medida que te familiarices con el lenguaje y sus funciones, podrás crear análisis más sofisticados, realizar cálculos personalizados y aprovechar al máximo las capacidades de Power Pivot y Power BI.

## Sintaxis

La sintaxis en DAX tiene varias partes que te ayudarán a construir tus fórmulas. Aquí tienes una descripción de algunas de las partes más comunes en la sintaxis de DAX:

1. **Referencias a columnas y tablas:** En DAX, puedes hacer referencia a las columnas y tablas en tu modelo de datos utilizando la siguiente sintaxis:
   - Referencia a una columna: `[Nombre de la Tabla][Nombre de la Columna]`
   - Referencia a una tabla: `[Nombre de la Tabla]`

2. **Operadores aritméticos y lógicos:** Puedes utilizar operadores aritméticos y lógicos para realizar cálculos y evaluaciones en tus fórmulas DAX. Algunos ejemplos comunes son:
   - Suma: `+`
   - Resta: `-`
   - Multiplicación: `*`
   - División: `/`
   - Igualdad: `=`
   - Mayor que: `>`
   - Menor que: `<`
   - Y lógico: `&&`
   - O lógico: `||`

3. **Funciones DAX:** DAX proporciona una amplia gama de funciones predefinidas que puedes utilizar en tus fórmulas. Las funciones DAX siguen la siguiente sintaxis general:
   - `NOMBRE_DE_LA_FUNCION(argumento1, argumento2, ...)`

4. **Variables y asignaciones:** En DAX, puedes utilizar variables para almacenar valores intermedios y simplificar tus fórmulas. La sintaxis para declarar una variable en DAX es la siguiente:
   - `VAR nombre_variable = expresión`

5. **Funciones personalizadas:** Puedes crear tus propias funciones personalizadas en DAX utilizando la sintaxis `DEFINE FUNCTION`. Sin embargo, ten en cuenta que las funciones personalizadas solo están disponibles en ciertas herramientas de análisis de datos de Microsoft, como Power BI y Analysis Services Tabular.

Es importante tener en cuenta que la sintaxis exacta de DAX puede variar dependiendo de la herramienta que estés utilizando, como Power Pivot, Power BI o Analysis Services Tabular. Además, algunos conceptos avanzados, como las tablas dinámicas y los contextos de filtro, requieren un entendimiento más profundo de DAX.

Te recomendaría consultar la documentación oficial de Microsoft o recursos en línea para obtener una referencia completa de la sintaxis y las funciones disponibles en DAX, así como ejemplos prácticos de uso.

### Ejemplo 1

Supongamos que tenemos una base de datos de ventas con las siguientes tablas: "Ventas" y "Productos". La tabla "Ventas" contiene información sobre las ventas realizadas, incluyendo el producto vendido y la cantidad vendida. La tabla "Productos" contiene información detallada sobre cada producto, incluyendo su precio unitario.

1. **Calcular el total de ventas:**
Podemos utilizar la función SUM para calcular el total de ventas en la tabla "Ventas". Supongamos que queremos crear una medida llamada "Total Ventas" que muestre la suma de todas las ventas realizadas.

```
Total Ventas = SUM(Ventas[Cantidad])
```

2. **Calcular el promedio de ventas diarias:**
Podemos utilizar la función AVERAGE para calcular el promedio de ventas diarias. Supongamos que queremos crear una medida llamada "Promedio Ventas Diarias" que muestre el promedio de ventas por día.

```
Promedio Ventas Diarias = AVERAGE(Ventas[Cantidad])
```

3. **Calcular el precio total por producto:**
Podemos utilizar la función SUMX junto con la función RELATED para calcular el precio total por producto. Supongamos que queremos crear una columna calculada llamada "Precio Total" en la tabla "Productos" que muestre el precio total de ventas de cada producto.

```
Precio Total = SUMX(Ventas, Ventas[Cantidad] * RELATED(Productos[Precio]))
```

Estos son solo ejemplos básicos de cómo utilizar fórmulas DAX en Power Pivot. Puedes combinar funciones, aplicar filtros contextuales y realizar cálculos más complejos según tus necesidades y la estructura de tus datos.

### Ejemplo 2

1. **Calcular el total de ventas por categoría de producto:**
Supongamos que en nuestra base de datos tenemos una tabla adicional llamada "Categorías" que contiene información sobre las categorías de productos. Queremos calcular el total de ventas por categoría de producto.

```DAX
Total Ventas por Categoría = 
SUMX(
    FILTER(Productos, Productos[Categoría] = "Electrónicos"),
    RELATED(Ventas[Cantidad])
)
```

En este ejemplo, utilizamos la función FILTER para aplicar un filtro contextual y seleccionar solo los productos de la categoría "Electrónicos". Luego, utilizamos la función SUMX para calcular la suma de las ventas relacionadas con esos productos.

2. **Calcular el promedio de ventas mensuales para un producto específico:**
Supongamos que queremos calcular el promedio de ventas mensuales para un producto específico en particular. Utilizaremos la función CALCULATE para aplicar filtros contextuales en función del producto y la fecha.

```DAX
Promedio Ventas Mensuales = 
CALCULATE(
    AVERAGE(Ventas[Cantidad]),
    Ventas[Producto] = "Producto A",
    DATESINPERIOD(Calendar[Fecha], LASTDATE(Calendar[Fecha]), -12, MONTH)
)
```

En este ejemplo, aplicamos filtros contextuales utilizando la función CALCULATE. Filtramos por el producto "Producto A" y utilizamos la función DATESINPERIOD para especificar que queremos considerar los últimos 12 meses de ventas. Luego, calculamos el promedio de las cantidades de ventas correspondientes.

Estos ejemplos demuestran cómo puedes combinar funciones y aplicar filtros contextuales en tus fórmulas DAX. Recuerda que las posibilidades son muy amplias y puedes adaptar estas técnicas según tus necesidades y la estructura de tus datos.

---------------------------
# ¿Qué es Power Pivot?


Power Pivot es una característica de Microsoft Excel que permite el análisis de datos y la creación de modelos de datos avanzados. Te proporciona una manera eficiente de combinar, relacionar y analizar grandes volúmenes de datos provenientes de diferentes fuentes en una sola tabla dinámica o informe.

Aquí hay algunos puntos clave que debes conocer antes de abordar Power Pivot:

1. **Conocimientos básicos de Excel:** Es importante tener un buen entendimiento de las funciones y características básicas de Excel, ya que Power Pivot es una extensión de esta herramienta. Debes estar familiarizado con la creación y edición de fórmulas, manejo de hojas de cálculo y tablas, y comprensión de los conceptos de rangos y referencias.

2. **Entender los conceptos de bases de datos relacionales:** Power Pivot utiliza un enfoque de modelo de datos relacional, por lo que es útil comprender los conceptos de bases de datos, como tablas, relaciones, claves primarias y claves externas. Esto te ayudará a estructurar y relacionar tus datos correctamente en Power Pivot.

3. **Fuentes de datos externas:** Power Pivot te permite importar datos de una variedad de fuentes externas, como bases de datos SQL, archivos de texto, hojas de cálculo de Excel, archivos CSV, entre otros. Debes estar familiarizado con las fuentes de datos con las que trabajarás y saber cómo acceder a ellas desde Excel.

4. **Capacidad de análisis de datos:** Power Pivot es una herramienta poderosa para el análisis de datos. Debes estar cómodo con la manipulación y transformación de datos, la creación de cálculos personalizados, el uso de funciones y la comprensión de los principios básicos de la visualización de datos.

5. **Aprendizaje de DAX:** DAX (Data Analysis Expressions) es un lenguaje de fórmulas utilizado en Power Pivot para crear cálculos personalizados y medidas. Debes estar dispuesto a aprender los conceptos básicos de DAX, como funciones, sintaxis y operaciones, para aprovechar al máximo las capacidades de Power Pivot.

6. **Recursos de aprendizaje:** Hay una variedad de recursos disponibles para aprender Power Pivot, como tutoriales en línea, documentación de Microsoft, libros y cursos en línea. Es útil aprovechar estos recursos para mejorar tu comprensión y dominio de la herramienta.

## Más rapido

Sí, Power Pivot puede ayudarte a trabajar más rápido y eficientemente al realizar análisis de datos sin tener que depender exclusivamente de la función "BUSCARV" en Excel. Aquí hay algunas razones por las que Power Pivot puede agilizar tu trabajo:

1. **Relaciones de tablas:** Power Pivot te permite establecer relaciones entre diferentes tablas en tu modelo de datos. Esto significa que puedes vincular tablas mediante columnas comunes y, posteriormente, realizar consultas y cálculos en base a esas relaciones. En lugar de utilizar "BUSCARV" para buscar valores en tablas relacionadas, puedes simplemente utilizar las relaciones establecidas en Power Pivot para acceder a los datos relacionados.

2. **Cálculos personalizados con DAX:** Con Power Pivot, puedes utilizar el lenguaje de fórmulas DAX para crear cálculos personalizados y medidas. DAX ofrece una amplia gama de funciones y operaciones que te permiten realizar análisis complejos y realizar cálculos avanzados en tus datos. En lugar de depender únicamente de "BUSCARV", puedes crear cálculos personalizados en DAX que se ajusten a tus necesidades específicas.

3. **Capacidad de trabajar con grandes volúmenes de datos:** Power Pivot está diseñado para manejar grandes volúmenes de datos con eficiencia. Puedes importar y combinar múltiples fuentes de datos en un solo modelo de datos en Power Pivot, lo que te permite realizar análisis complejos en conjuntos de datos masivos sin comprometer el rendimiento. Esto puede ser especialmente útil cuando los datos que necesitas buscar con "BUSCARV" son demasiado grandes o complejos para manejarlos fácilmente en Excel.

4. **Tablas dinámicas avanzadas:** Power Pivot ofrece tablas dinámicas avanzadas que te permiten crear informes y análisis interactivos de tus datos. Puedes utilizar jerarquías, campos calculados y segmentación de datos para explorar y filtrar fácilmente tus datos sin tener que utilizar funciones de búsqueda como "BUSCARV". Las tablas dinámicas de Power Pivot son más flexibles y poderosas que las tablas dinámicas estándar de Excel.

## No hay limitaciones

Una de las principales ventajas de Power Pivot es su capacidad para trabajar con grandes volúmenes de datos de manera eficiente. A diferencia de las hojas de cálculo tradicionales de Excel, Power Pivot utiliza una tecnología llamada almacenamiento en columna y compresión de datos, lo que le permite manejar grandes cantidades de información de manera más rápida y efectiva. Aquí hay algunas características clave de Power Pivot relacionadas con el manejo de grandes volúmenes de datos:

1. **Importación de datos desde múltiples fuentes:** Power Pivot te permite importar datos desde múltiples fuentes, como bases de datos SQL, archivos de texto, hojas de cálculo de Excel, archivos CSV, entre otros. Esto significa que puedes combinar y consolidar datos de diferentes fuentes en un solo modelo de datos en Power Pivot.

2. **Modelo de datos en memoria:** Power Pivot utiliza un modelo de datos en memoria, lo que significa que los datos se cargan en la memoria RAM de tu computadora en lugar de trabajar directamente desde el disco duro. Esto permite un acceso más rápido a los datos y un rendimiento más eficiente al realizar consultas y cálculos.

3. **Compresión de datos:** Power Pivot aplica técnicas de compresión de datos inteligentes para reducir el tamaño de los datos almacenados en memoria. Esto ayuda a optimizar el rendimiento y permite trabajar con conjuntos de datos más grandes sin agotar los recursos del sistema.

4. **Columnas calculadas y medidas:** En Power Pivot, puedes crear columnas calculadas y medidas personalizadas utilizando el lenguaje de fórmulas DAX. Estas columnas y medidas te permiten realizar cálculos complejos y resumir los datos de manera eficiente, sin necesidad de duplicar los datos subyacentes.

5. **Optimización de consultas:** Power Pivot tiene capacidades de optimización de consultas que ayudan a acelerar el tiempo de respuesta al realizar operaciones de filtrado, agrupación y resumen de datos. Las consultas se ejecutan de manera más eficiente, lo que permite interactuar de manera rápida con grandes conjuntos de datos.

En general, Power Pivot está diseñado para manejar grandes volúmenes de datos y proporcionar un rendimiento óptimo al realizar análisis y cálculos. Su capacidad de importar datos de diferentes fuentes, la utilización de un modelo de datos en memoria y la optimización de consultas hacen de Power Pivot una herramienta valiosa para trabajar con grandes cantidades de información de manera eficiente y efectiva.

