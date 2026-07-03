// APU_CODIGOS
// Lista ligera de codigos de actividad presentes en el catalogo "ANALISIS DE
// PRECIOS UNITARIOS" del Centro de Costos actual. Parsea SOLO la columna 1
// del HTML (el archivo pesa >1 MB y tiene miles de filas), a diferencia de
// PLANTILLA_PPTO_SINCO que ademas descarga el reporte de SEGUIMIENTO (6 MB)
// y construye la plantilla completa de 21 columnas.
//
// Existe para que ACTIVIDADES_NUEVAS (que solo necesita saber "¿que codigos
// hay en el APU?") no tenga que evaluar toda esa cadena pesada en cada
// actualizacion.
let
    // Un mismo archivo QUERY UNIFICADO sirve para cualquier proyecto: si el
    // Origen es ORACLE este dominio aun no tiene equivalente implementado,
    // asi que se devuelve la tabla vacia con el mismo esquema en vez de
    // reventar el "Actualizar todo".
    OrigenActual = try Text.Upper(Text.Trim(Origen)) otherwise "SINCO",
    VacioOracle = #table({"Codigo"}, {}),
    SiteUrl = "https://colsubsidio365.sharepoint.com/sites/MiGerenciaViv",
    Headers = [Accept="application/json;odata=nometadata"],
    FnEncode = F_Globales[FnEncode],
    ArchivosProyecto = Table.Buffer(SP_Archivos_Proyecto),

    Filas = Table.SelectRows(ArchivosProyecto, each Text.Contains([Name], "ANALISIS DE PRECIOS UNITARIOS", Comparer.OrdinalIgnoreCase)),
    Ordenadas = Table.Sort(Filas, {{"Centro de Costos", Order.Ascending}, {"TimeLastModified", Order.Descending}}),

    Resultado =
        if Table.RowCount(Ordenadas) = 0 then
            #table(type table[Codigo=text], {})
        else
            let
                Binario = Binary.Buffer(Web.Contents(SiteUrl, [
                    RelativePath = "/_api/web/GetFileByServerRelativeUrl('" & FnEncode(Ordenadas{0}[ServerRelativeUrl]) & "')/$value",
                    Headers = Headers,
                    Timeout = #duration(0, 0, 5, 0)
                ])),
                TextoHTML = Text.FromBinary(Binario, 65001),
                Columnas = {{"Columna1", "td:nth-child(1), th:nth-child(1)"}},
                Tabla = Html.Table(TextoHTML, Columnas, [RowSelector="tr"]),

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

                ConCodigo = Table.AddColumn(Tabla, "Codigo", each
                    let
                        t0 = if [Columna1] = null then "" else Text.From([Columna1]),
                        t = Text.Trim(Text.Replace(Text.Replace(Text.Replace(t0, "#(lf)", " "), "#(cr)", " "), "#(00A0)", " ")),
                        antes = if Text.Contains(t, "-") then Text.Trim(Text.BeforeDelimiter(t, "-")) else ""
                    in
                        if antes <> "" and EsCodigoActividad(antes) then antes else null,
                type text),
                SoloCodigos = Table.SelectRows(ConCodigo, each [Codigo] <> null),
                Distintos = Table.Distinct(Table.SelectColumns(SoloCodigos, {"Codigo"}))
            in
                Distintos
in
    if OrigenActual = "ORACLE" then VacioOracle else Table.Buffer(Resultado)
