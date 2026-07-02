// PROYECCION_INSUMOS
// Ajustes de proyeccion pendientes por par (actividad, insumo), derivados del
// reporte SEGUIMIENTO POR ITEMS de la constructora:
//
//   CANTIDAD a proyectar (ajuste por "Diferencias") =
//       Proyectado(constructora) - Presupuestado(constructora), solo si > 0
//
// Racional (verificado con caso real 2.006 / S_EXCAVACION MECANICA):
// la plantilla inicial carga el Presupuestado como proyeccion de arranque
// (88.48); el ajuste por la diferencia (105.83 - 88.48 = 17.35) deja el SINCO
// propio espejado con la proyeccion oficial de la constructora, sin inflarla.
// Para insumos con Presupuestado = 0 (no venian en la plantilla), la
// diferencia es la proyeccion completa.
//
// Columnas (posiciones fijas, consumidas por el flujo RPA PROYECCION INSUMOS):
//   A ACTIVIDADES (codigo, para busqueda en 'ACItemsinput')
//   B INSUMOS (descripcion, para 'txtsearch')
//   C PROYECTADO (cantidad diferencia -> 'txtcanprN')
//   D Valor Unit (VrUnit del proyectado -> 'txtvrunitprN')
//   E Estado (el flujo escribe "Proyectado - hora"; filas marcadas no reaparecen)
let
    SiteUrl = "https://colsubsidio365.sharepoint.com/sites/MiGerenciaViv",
    Headers = [Accept="application/json;odata=nometadata"],
    FnEncode = F_Globales[FnEncode],
    ArchivosProyecto = Table.Buffer(SP_Archivos_Proyecto),

    Limpiar = (v as any) as text =>
        let t = if v = null then "" else Text.From(v)
        in Text.Trim(Text.Replace(Text.Replace(Text.Replace(t, "#(lf)", " "), "#(cr)", " "), "#(00A0)", " ")),

    FxNum = (v as any) as number =>
        let
            t = Limpiar(v),
            sinMiles = Text.Replace(t, ",", ""),
            n = try Number.FromText(sinMiles, "en-US") otherwise null
        in
            if n = null then 0 else n,

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

    EsCodigoInsumo = (c0 as text) as logical =>
        c0 <> "" and Text.Select(c0, {"0".."9"}) = c0,

    Filas = Table.SelectRows(ArchivosProyecto, each Text.Contains([Name], "SEGUIMIENTO POR ITEMS", Comparer.OrdinalIgnoreCase)),
    Ordenadas = Table.Sort(Filas, {{"Centro de Costos", Order.Ascending}, {"TimeLastModified", Order.Descending}}),

    Resultado =
        if Table.RowCount(Ordenadas) = 0 then
            #table({"ACTIVIDADES","INSUMOS","PROYECTADO","Valor Unit","Estado"}, {})
        else
            let
                Binario = Binary.Buffer(Web.Contents(SiteUrl, [
                    RelativePath = "/_api/web/GetFileByServerRelativeUrl('" & FnEncode(Ordenadas{0}[ServerRelativeUrl]) & "')/$value",
                    Headers = Headers,
                    Timeout = #duration(0, 0, 5, 0)
                ])),
                // Se necesitan las columnas 1-9: Cod, Desc, Tipo, UM,
                // Presupuestado(Cant=5, VrU=6, Tot=7), Proyectado(Cant=8, VrU=9)
                Cols = List.Transform({1..9}, each {"Columna" & Text.From(_), "td:nth-child(" & Text.From(_) & "), th:nth-child(" & Text.From(_) & ")"}),
                Tabla = Html.Table(Text.FromBinary(Binario, 65001), Cols, [RowSelector="tr"]),

                Clasificado = Table.AddColumn(Tabla, "Clase", (r as record) =>
                    let
                        c0 = Limpiar(Record.Field(r, "Columna1")),
                        c2 = Limpiar(Record.Field(r, "Columna3")),
                        c3 = Limpiar(Record.Field(r, "Columna4"))
                    in
                        if c0 = "" or c0 = "Cod" or Text.StartsWith(c0, "SubCapitulo") then "Otro"
                        else if c2 <> "" and c3 <> "" and EsCodigoInsumo(c0) then "Insumo"
                        else if EsCodigoActividad(c0) and Text.EndsWith(c0, ".000") then "Capitulo"
                        else if EsCodigoActividad(c0) then "Actividad"
                        else "Otro",
                    type text),
                ConActCod = Table.AddColumn(Clasificado, "ActCod", each if [Clase] = "Actividad" then Limpiar([Columna1]) else if [Clase] = "Capitulo" then "" else null, type text),
                FD = Table.FillDown(ConActCod, {"ActCod"}),

                Insumos = Table.SelectRows(FD, each [Clase] = "Insumo" and [ActCod] <> null and [ActCod] <> ""),
                ConCantidades = Table.AddColumn(Insumos, "__Datos", each [
                    ACTIVIDADES = [ActCod],
                    INSUMOS = Limpiar([Columna2]),
                    PresCant = FxNum([Columna5]),
                    ProyCant = FxNum([Columna8]),
                    ProyVrU = FxNum([Columna9])
                ]),
                Expandida = Table.ExpandRecordColumn(Table.SelectColumns(ConCantidades, {"__Datos"}), "__Datos", {"ACTIVIDADES","INSUMOS","PresCant","ProyCant","ProyVrU"}),

                // Diferencia a proyectar (con margen numerico para ruido de decimales)
                ConDif = Table.AddColumn(Expandida, "PROYECTADO", each Number.Round([ProyCant] - [PresCant], 6), type number),
                Pendientes = Table.SelectRows(ConDif, each [PROYECTADO] > 0.001),
                ConVrU = Table.RenameColumns(Table.SelectColumns(Pendientes, {"ACTIVIDADES","INSUMOS","PROYECTADO","ProyVrU"}), {{"ProyVrU","Valor Unit"}}),
                Dedup = Table.Distinct(ConVrU, {"ACTIVIDADES","INSUMOS"}),

                // Preservar Estado y excluir filas ya proyectadas
                RangoPrevio = try Excel.CurrentWorkbook(){[Name="RangoProyeccionInsumos"]}[Content] otherwise null,
                ColsEsperadas = {"ACTIVIDADES","INSUMOS","PROYECTADO","Valor Unit","Estado"},
                Previo =
                    if RangoPrevio = null or Table.RowCount(RangoPrevio) = 0 then #table(ColsEsperadas, {})
                    else
                        let
                            P = Table.PromoteHeaders(RangoPrevio, [PromoteAllScalars=true]),
                            Ok = if List.Contains(Table.ColumnNames(P), "INSUMOS") then P else #table(ColsEsperadas, {})
                        in
                            Table.SelectRows(Ok, each [INSUMOS] <> null and Text.Trim(Text.From([INSUMOS])) <> ""),
                PrevClave = Table.AddColumn(Previo, "__K", each Text.Upper(Text.Trim(Text.From([ACTIVIDADES]))) & "|" & Text.Upper(Text.Trim(Text.From([INSUMOS]))), type text),
                PrevSolo = Table.Distinct(Table.SelectColumns(PrevClave, {"__K","Estado"}), {"__K"}),

                ActualClave = Table.AddColumn(Dedup, "__K", each Text.Upper(Text.Trim([ACTIVIDADES])) & "|" & Text.Upper(Text.Trim([INSUMOS])), type text),
                JoinPrev = Table.NestedJoin(ActualClave, {"__K"}, PrevSolo, {"__K"}, "PrevEstado", JoinKind.LeftOuter),
                ExpPrev = Table.ExpandTableColumn(JoinPrev, "PrevEstado", {"Estado"}, {"Estado"}),
                SinHechas = Table.SelectRows(ExpPrev, each [Estado] = null or Text.Trim(Text.From([Estado])) = ""),

                Final = Table.SelectColumns(SinHechas, {"ACTIVIDADES","INSUMOS","PROYECTADO","Valor Unit","Estado"}),
                Orden = Table.Sort(Final, {{"ACTIVIDADES", Order.Ascending}, {"INSUMOS", Order.Ascending}})
            in
                Orden
in
    Resultado
