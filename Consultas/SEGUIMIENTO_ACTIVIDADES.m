// SEGUIMIENTO_ACTIVIDADES
// Parsea "SEGUIMIENTO POR ITEMS" y devuelve solo las filas de tipo Actividad
// (Codigo, Descripcion, UM, Capitulo padre), para el Centro de Costos actual.
//
// Clasificacion de filas (columnas: Cod, Descripcion, Tipo, UM, ...):
//   - Capitulo:  codigo termina en ".000" y Tipo/UM vienen vacios
//   - Actividad: codigo tiene forma N.NNN (no .000) y Tipo/UM vienen vacios
//   - Insumo:    Tipo y UM SI vienen llenos (ej "M"/"M3", "O"/"Un")
//   - SubCapitulo / vacias / encabezados: se ignoran
//
// El UM de una Actividad no viene en su propia fila (siempre esta vacio ahi),
// asi que se toma del primer Insumo que aparece inmediatamente despues de ella.
let
    SiteUrl = "https://colsubsidio365.sharepoint.com/sites/MiGerenciaViv",
    Headers = [Accept="application/json;odata=nometadata"],
    FnEncode = F_Globales[FnEncode],
    ArchivosProyecto = Table.Buffer(SP_Archivos_Proyecto),

    Filas = Table.SelectRows(ArchivosProyecto, each Text.Contains([Name], "SEGUIMIENTO POR ITEMS", Comparer.OrdinalIgnoreCase)),
    Ordenadas = Table.Sort(Filas, {{"Centro de Costos", Order.Ascending}, {"TimeLastModified", Order.Descending}}),

    Resultado =
        if Table.RowCount(Ordenadas) = 0 then
            #table(type table[Codigo=text, Descripcion=text, UM=text, Capitulo=text], {})
        else
            let
                Binario = Binary.Buffer(Web.Contents(SiteUrl, [
                    RelativePath = "/_api/web/GetFileByServerRelativeUrl('" & FnEncode(Ordenadas{0}[ServerRelativeUrl]) & "')/$value",
                    Headers = Headers,
                    Timeout = #duration(0, 0, 5, 0)
                ])),
                TextoHTML = Text.FromBinary(Binario, 65001),
                Columnas = List.Transform({1..35}, each {"Columna" & Text.From(_), "td:nth-child(" & Text.From(_) & "), th:nth-child(" & Text.From(_) & ")"}),
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
                        esCapitulo = esActCod and Text.EndsWith(c0, ".000"),
                        tipoUMVacio = c2 = "" and c3 = ""
                    in
                        if c0 = "" or c0 = "Cod" or Text.StartsWith(c0, "SubCapitulo") then "Otro"
                        else if not esActCod then "Otro"
                        else if esCapitulo and tipoUMVacio then "Capitulo"
                        else if tipoUMVacio then "Actividad"
                        else "Insumo",
                    type text),

                ConCapituloRaw = Table.AddColumn(Clasificado, "CapituloRaw", each if [Clase] = "Capitulo" then Limpiar([Columna2]) else null, type text),
                ConCapituloFD = Table.FillDown(ConCapituloRaw, {"CapituloRaw"}),
                TablaCompleta = Table.Buffer(ConCapituloFD),

                BuscarUM = (indiceActividad as number) as nullable text =>
                    let
                        ventana = Table.SelectRows(TablaCompleta, each [IndiceFila] > indiceActividad and [IndiceFila] <= indiceActividad + 8),
                        limiteFila = Table.SelectRows(ventana, each [Clase] = "Actividad" or [Clase] = "Capitulo"),
                        limiteIdx = if Table.RowCount(limiteFila) = 0 then indiceActividad + 8 else List.Min(limiteFila[IndiceFila]),
                        insumosValidos = Table.SelectRows(ventana, each [Clase] = "Insumo" and [IndiceFila] < limiteIdx),
                        um = if Table.RowCount(insumosValidos) = 0 then null else Limpiar(insumosValidos{0}[Columna4])
                    in um,

                SoloActividades = Table.SelectRows(TablaCompleta, each [Clase] = "Actividad"),
                ConUM = Table.AddColumn(SoloActividades, "UM", each BuscarUM([IndiceFila]), type text),
                ConCodigoLimpio = Table.AddColumn(ConUM, "Codigo", each Limpiar([Columna1]), type text),
                ConDescLimpia = Table.AddColumn(ConCodigoLimpio, "Descripcion", each Limpiar([Columna2]), type text),

                Final = Table.SelectColumns(ConDescLimpia, {"Codigo", "Descripcion", "UM", "CapituloRaw"}),
                Renombrado = Table.RenameColumns(Final, {{"CapituloRaw", "Capitulo"}}),
                Tipado = Table.TransformColumnTypes(Renombrado, {{"Codigo", type text}, {"Descripcion", type text}, {"UM", type text}, {"Capitulo", type text}})
            in
                Tipado
in
    Table.Buffer(Resultado)
