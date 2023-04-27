# Excel y VBA

En cualquier vacante de analista de datos, científico de daros, estadístico o relacionado con datos un requisito indespensables es tener un manejo de Excel a nivel
intermedio - avanzado pero como se clasifican los niveles de excel, puede que sepas VBA a nivel de un Senior pero ignores otras funcionalidades, el más real a mi parece es la siguiente escala

- **Principiante**:
* * En este caso el usuario utiliza los libros de Excel a modo de calculadora y realiza sumas y productos principalmente, también sabe introducir datos y recopila alguna información en la hoja electrónica.

- **Básico**: 
* * En este nivel se usan funciones de Excel sencillas, sobre todo matemáticas, como divisiones, porcentajes, alguna función estadística como la Media, Contar, Autosuma.
* * Se realizan filtros sencillos con el Autofiltro. Realiza algún gráfico fácil. Y usan cambios de formatos rápidos como colores y letra. 

- **Intermedio**:
* Emplea funciones más avanzadas como las financieras, estadísticas, alguna función lógica como SI.
* Trabaja avanzadamente con gráficos. Usa  los formatos más avanzados como formatos condicionales. Es capaz de usar formularios para introducir datos.

- **Avanzado**:
* * Este tipo de usuario conoce las funciones de Excel más avanzadas como pueden ser las de búsqueda BUSCARV, BUSCARH, lógicas como SI, Y, NO, realiza funciones anidadas.
* * Conoce y realiza la herramienta Subtotales, Tablas y gráficos dinámicos, realiza grabaciones de Macros sencillas, domina la realización de gráficos.
* * Sabe y usa con fluidez el recurso importar datos desde Access y de formatos .txt.

Estos dos niveles son ya usuarios con un conocimiento alto:

- **Experto**:
* * Se trata de un usuario que domina los diferentes recursos de Excel avanzado incluso es capaz de programar con el lenguaje VBA. Trabaja con funciones matriciales como Sumaproducto, financieras avanzadas como Van, TIR, Nper.
* * Conoce las funciones de texto: Concatenar, Extraer, Carácter y similares, Importa y Consolida datos desde gran número de utilidades externas, usa la herramienta Solver y los Escenarios.
* * Es capaz de configurar y automatizar gran parte de su trabajo con la hoja electrónica. Además tiene conocimientos básicos de programación con las que realiza Macros.
* * Conoce y usa herramientas de Inteligencia de negocio como Power BI de Excel.

- **Programador Excel**:
* * Profesional de la informática con nivel experto y/o avanzado de Excel, usa el programa para obtener cálculos avanzados, estadísticas, combina Excel con otros programas. Domina la programación con VBA. Y conoce y sabe usar los complementos Power BI de Excel.

Como en mis otras entradas explique, con pandas es posible realizar todas absolutamnete todas las funciones desde la interfaz de python, pues nos permite leer los distintos archivos, cvs y xlsx y apartir de ahi podemos usar los datos para nuestros modelos estadísticos o de machine learning, quizas con ello te consideres en la clase de experto pero realmente pueda que no sea asi, si tengo una tabla con datos fiables en mi excel y me piden 3 gráficas dinamicas por ejemplo sobre ventas en un establemiento catalogada por empleado y su desempeño a lo largo de un periodo de un año, quizas estes pensando en leer la tabla como dataframe en pandas, agrupar por empleado, ventas y tienda, usar pyplot para un gráfico interactivo, mucho trabajo pues con 5 clik's se hace con excel, filtro de datos, segmentar, insertar grafico y lo puedes copiar a otra hoja o directamente en power bi, es por ello que excel nunca quedara de lado, por más python sea mejor en otras ramos para estas instancias pasa a ser complicado y ningún analista que yo conozca se quiere complicar la vida.

## Integración

Si estas en contacto con personal de otras áreas quizas no tengan los conocimentos para que les puedas pasar un código a diferencia que un documento excel que es de lenguaje universal, administrativos, contabilidad tienen conocimiento de ello.

## Poder de excel

Al ser un producto insignea de Microsft no es de sorprenderse que Excel sea constantemente actualizado y capaz de implementar nuevas funciones al nivel de un software puramente estadistico, tan sólo ser capaz de que puedas crear tus propias funciones implica que las posibilidades sean infinitas, si dominas scipy en python, quizas te preguntes si excel es capaz de competir en regresiones y así es lo hace e incluso de manera más intuitiva, sólo la herramienta de escenarios es un acercamiento al forecast, se pueden llevar a cabo métodos complejos como:
* Simulaciones de montecarlo
* Regresiones
* Machine learning


## VBA

Sería redundante que pretenda subir a este repositorio suba unas notas de excel básico pues hay miles o millones de recursos en san google que emolican de mejor manera lo que si es que presentare algunos programas en VBA los cuales automatizan tareas como reportes, limpieza, filtro de datos atravez de formularios sencillos.
```VBA
Private Sub CommandButton2_Click()
Unload Me
End Sub
```



