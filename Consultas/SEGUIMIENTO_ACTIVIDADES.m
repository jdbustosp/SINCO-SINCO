// SEGUIMIENTO_ACTIVIDADES
// Parsea "SEGUIMIENTO POR ITEMS" y devuelve solo las filas de tipo Actividad
// (Codigo, Descripcion, UM, Capitulo, Subcapitulo padre), para el Centro de
// Costos actual.
//
// Clasificacion de filas (columnas: Cod, Descripcion, Tipo, UM, ...):
//   - Capitulo:     codigo termina en ".000" y Tipo/UM vienen vacios
//   - SubCapitulo:  fila con texto "SubCapitulo :NOMBRE" en la primera columna
//   - Actividad:    codigo tiene forma N.NNN (no .000) y Tipo/UM vienen vacios
//   - Insumo:       Tipo y UM SI vienen llenos (ej "M"/"M3", "O"/"Un")
//   - vacias / encabezados: se ignoran
//
// El UM de una Actividad no viene en su propia fila (siempre esta vacio ahi),
// asi que se toma del primer Insumo que aparece inmediatamente despues de ella.
// El Subcapitulo se reinicia (vuelve a quedar vacio) cada vez que empieza un
// Capitulo nuevo, para no arrastrar el subcapitulo de un capitulo anterior.
let
    SiteUrl = "https://colsubsidio365.sharepoint.com/sites/MiGerenciaViv",
    Headers = [Accept="application/json;odata=nometadata"],
    FnEncode = F_Globales[FnEncode],
    ArchivosProyecto = Table.Buffer(SP_Archivos_Proyecto),

    Filas = Table.SelectRows(ArchivosProyecto, each Text.Contains([Name], "SEGUIMIENTO POR ITEMS", Comparer.OrdinalIgnoreCase)),
    Ordenadas = Table.Sort(Filas, {{"Centro de Costos", Order.Ascending}, {"TimeLastModified", Order.Descending}}),

    Resultado =
        if Table.RowCount(Ordenadas) = 0 then
            #table(type table[Codigo=text, Descripcion=text, UM=text, Capitulo=text, Subcapitulo=text], {})
        else
            let
                Binario = Binary.Buffer(Web.Contents(SiteUrl, [
                    RelativePath = "/_api/web/GetFileByServerRelativeUrl('" & FnEncode(Ordenadas{0}[ServerRelativeUrl]) & "')/$value",
                    Headers = Headers,
                    Timeout = #duration(0, 0, 5, 0)
                ])),
                TextoHTML = Text.FromBinary(Binario, 65001),
                // Solo se usan las primeras 4 columnas (Cod/Descripcion/Tipo/UM) para
                // clasificar filas. Pedir mas columnas de las que se usan multiplica
                // el costo de parseo del HTML sin necesidad (este reporte puede tener
                // miles de filas), y era la causa principal de que la actualizacion
                // se sintiera muy lenta.
                Columnas = List.Transform({1..4}, each {"Columna" & Text.From(_), "td:nth-child(" & Text.From(_) & "), th:nth-child(" & Text.From(_) & ")"}),
                Tabla = Html.Table(TextoHTML, Columnas, [RowSelector="tr"]),

                Limpiar = (v as any) as text =>
                    let t = if v = null then "" else Text.From(v)
                    in Text.Trim(Text.Replace(Text.Replace(Text.Replace(t, "#(lf)", " "), "#(cr)", " "), "#(00A0)", " ")),

                AddIndice = Table.AddIndexColumn(Tabla, "IndiceFila", 0, 1, Int64.Type),

                EsCodigoActividad = (c0 as text) as logical =>
                    let
                        partes = Text.Split(c0, "."),
                        dosPartes = List.Count(partes) = 2,
                        p1 = if dosPartes then partes{0} else "",
                        p2 = if dosPartes then partes{1} else "",
                        p1Num = if dosPartes then (try Number.FromText(p1) otherwise null) else null,
                        p2Num = if dosPartes then (try Number.FromText(p2) otherwise null) else null
                    in
                        dosPartes and p1Num <> null and p2Num <> null and Text.Length(p1) > 0 and Text.Length(p2) > 0,

                Clasificado = Table.AddColumn(AddIndice, "Clase", (r as record) =>
                    let
                        c0 = Limpiar(Record.Field(r, "Columna1")),
                        c2 = Limpiar(Record.Field(r, "Columna3")),
                        c3 = Limpiar(Record.Field(r, "Columna4")),
                        esActCod = c0 <> "" and EsCodigoActividad(c0),
                        esCapitulo = esActCod and Text.EndsWith(c0, ".000")
                    in
                        // El chequeo de Insumo (Tipo+UM llenos) va PRIMERO: los
                        // codigos de insumo normalmente NO tienen punto (ej "4436"),
                        // por eso antes se colaban en el "else if not esActCod then
                        // Otro" y se perdian, dejando el UM de la Actividad sin poder
                        // encontrarse (siempre devolvia null).
                        if c0 = "" or c0 = "Cod" or Text.StartsWith(c0, "SubCapitulo") then "Otro"
                        else if c2 <> "" and c3 <> "" then "Insumo"
                        else if not esActCod then "Otro"
                        else if esCapitulo then "Capitulo"
                        else "Actividad",
                    type text),

                ConCapituloRaw = Table.AddColumn(Clasificado, "CapituloRaw", each if [Clase] = "Capitulo" then Limpiar([Columna2]) else null, type text),
                ConCapituloFD = Table.FillDown(ConCapituloRaw, {"CapituloRaw"}),

                // Subcapitulo: se marca "" (no null) en cada fila de Capitulo para
                // reiniciar el relleno hacia abajo justo ahi, y se toma el nombre
                // de las filas "SubCapitulo :NOMBRE" hasta el proximo reinicio.
                ConSubRaw = Table.AddColumn(ConCapituloFD, "SubRaw", each
                    if [Clase] = "Capitulo" then ""
                    else
                        let c0 = Limpiar([Columna1]) in
                        if Text.StartsWith(c0, "SubCapitulo") then
                            let pos = Text.PositionOf(c0, ":") in if pos >= 0 then Text.Trim(Text.Range(c0, pos + 1)) else null
                        else null,
                type text),
                ConSubFD = Table.FillDown(ConSubRaw, {"SubRaw"}),
                ConSubFinal = Table.TransformColumns(ConSubFD, {"SubRaw", each if _ = "" then null else _, type text}),
                TablaCompleta = Table.Buffer(ConSubFinal),

                TotalFilas = Table.RowCount(TablaCompleta),
                BuscarUM = (indiceActividad as number) as nullable text =>
                    let
                        // Table.Range es una lectura posicional directa (rapida), a
                        // diferencia de Table.SelectRows que evalua un predicado fila
                        // por fila. Como IndiceFila es secuencial 0-based, la ventana
                        // de las siguientes 8 filas empieza justo en indiceActividad+1.
                        largoVentana = List.Min({8, TotalFilas - indiceActividad - 1}),
                        ventana = if largoVentana <= 0 then #table({"Clase","Columna4"}, {}) else Table.Range(TablaCompleta, indiceActividad + 1, largoVentana),
                        primerInsumo = List.PositionOf(ventana[Clase], "Insumo"),
                        candidatosLimite = List.Select({List.PositionOf(ventana[Clase], "Actividad"), List.PositionOf(ventana[Clase], "Capitulo")}, each _ >= 0),
                        primerLimite = if List.IsEmpty(candidatosLimite) then -1 else List.Min(candidatosLimite),
                        um = if primerInsumo < 0 or (primerLimite >= 0 and primerInsumo > primerLimite) then null else Limpiar(ventana[Columna4]{primerInsumo})
                    in um,

                SoloActividades = Table.SelectRows(TablaCompleta, each [Clase] = "Actividad"),
                ConUM = Table.AddColumn(SoloActividades, "UM", each BuscarUM([IndiceFila]), type text),
                ConCodigoLimpio = Table.AddColumn(ConUM, "Codigo", each Limpiar([Columna1]), type text),
                ConDescLimpia = Table.AddColumn(ConCodigoLimpio, "Descripcion", each Limpiar([Columna2]), type text),

                Final = Table.SelectColumns(ConDescLimpia, {"Codigo", "Descripcion", "UM", "CapituloRaw", "SubRaw"}),
                Renombrado = Table.RenameColumns(Final, {{"CapituloRaw", "Capitulo"}, {"SubRaw", "Subcapitulo"}}),
                Tipado = Table.TransformColumnTypes(Renombrado, {{"Codigo", type text}, {"Descripcion", type text}, {"UM", type text}, {"Capitulo", type text}, {"Subcapitulo", type text}})
            in
                Tipado
in
    Table.Buffer(Resultado)
