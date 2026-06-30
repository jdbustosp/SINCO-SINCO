// SEGUIMIENTO_CODIGOS
// Lista de TODOS los codigos (columna 1) que aparecen en el reporte
// "SEGUIMIENTO POR ITEMS" para el Centro de Costos actual, sin importar
// si la fila es Capitulo/Actividad/Insumo.
//
// Por que no usar SP_Seguimiento_Parsed (que ya clasifica TipoFila)? Porque
// esa clasificacion usa heuristicas (longitud del codigo, si las columnas
// Tipo/UM vienen vacias, multiplos de 1000, etc.) y puede fallar para casos
// borde, dejando una Actividad sin "Codigo act" asignado y por lo tanto
// pareciendo "nueva" en ACTIVIDADES_NUEVAS aunque ya exista en SINCO.
// Aqui simplemente preguntamos "¿este codigo aparece en algun lado del
// reporte de seguimiento?" sin clasificar nada, lo cual es mucho mas
// resistente a esos casos borde.
let
    SiteUrl = "https://colsubsidio365.sharepoint.com/sites/MiGerenciaViv",
    Headers = [Accept="application/json;odata=nometadata"],
    FnEncode = F_Globales[FnEncode],
    ArchivosProyecto = Table.Buffer(SP_Archivos_Proyecto),

    Filas = Table.SelectRows(ArchivosProyecto, each Text.Contains([Name], "SEGUIMIENTO POR ITEMS", Comparer.OrdinalIgnoreCase)),
    Ordenadas = Table.Sort(Filas, {{"Centro de Costos", Order.Ascending}, {"TimeLastModified", Order.Descending}}),

    Resultado =
        if Table.RowCount(Ordenadas) = 0 then
            #table({"Codigo"}, {})
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
                Limpio = Table.AddColumn(Tabla, "Codigo", each
                    let
                        t = if [Columna1] = null then "" else Text.From([Columna1]),
                        t1 = Text.Trim(Text.Replace(Text.Replace(Text.Replace(t, "#(lf)", " "), "#(cr)", " "), "#(00A0)", " "))
                    in t1, type text),
                SoloConCodigo = Table.SelectRows(Limpio, each [Codigo] <> ""),
                Distintos = Table.Distinct(Table.SelectColumns(SoloConCodigo, {"Codigo"}))
            in
                Distintos
in
    Table.Buffer(Resultado)
