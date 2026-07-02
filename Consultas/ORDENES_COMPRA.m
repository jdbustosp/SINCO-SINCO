// ORDENES_COMPRA
// Hoja unica que alimenta los 4 flujos RPA de ordenes/almacen:
//   ORDENES DE COMPRA [B insumo, C v/u, D cantidad]
//   ASIGNACION ORDENES [A # orden]
//   ENTRADAS DE ALMACEN [A # orden, B insumo, E consumido]
//   SALIDAS ALMACEN     [A # orden, E consumido, F producto, G item]
//
// Fuente principal: InformeOrdenDeCompraDetalladoInsumos.xlsx (xlsx real, tabla
// plana con Codigo, Descripcion, Agrupacion, UM, Item, Capitulo, Orden de
// Compra, Fecha Compra, Estado Compra, Cantidad Comprada, Valor Unitario,
// IVA %, Valor Total). Complemento: "ESTADO DE ORDENES" (HTML) para la
// cantidad recibida (CONSUMIDO) por (orden, codigo insumo).
//
// Delta por fecha: si FechaVersionComparar no esta vacia, se excluyen las
// ordenes cuyo numero ya aparece en el mismo informe de la carpeta
// "Versiones previas/<fecha>". Preservacion de Estado: filas marcadas
// "OK"/"Creado" en la hoja no se vuelven a proponer.
let
    ParamProyecto = Text.Trim(Text.From(ProyectoActual)),
    ParamCC = try Text.Trim(Text.From(CentroCostoActual)) otherwise "",
    FechaVersion = try Text.Trim(Text.From(FechaVersionComparar)) otherwise "",
    SiteUrl = "https://colsubsidio365.sharepoint.com/sites/MiGerenciaViv",
    BasePath = "/sites/MiGerenciaViv/Departamento Tecnico/COORDINACION DE PRESUPUESTOS/0. Reportes EDT - Control costos interno/" & ParamProyecto,
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

    FnGetBinario = (contiene as text) as nullable binary =>
        let
            Filas = Table.SelectRows(ArchivosProyecto, each Text.Contains([Name], contiene, Comparer.OrdinalIgnoreCase)),
            Ordenadas = Table.Sort(Filas, {{"Centro de Costos", Order.Ascending}, {"TimeLastModified", Order.Descending}})
        in
            if Table.RowCount(Ordenadas) = 0 then null else
            Binary.Buffer(Web.Contents(SiteUrl, [
                RelativePath = "/_api/web/GetFileByServerRelativeUrl('" & FnEncode(Ordenadas{0}[ServerRelativeUrl]) & "')/$value",
                Headers = Headers,
                Timeout = #duration(0, 0, 10, 0)
            ])),

    // ===== 1. Informe detallado de insumos por orden (xlsx real) =====
    BinInforme = FnGetBinario("INFORMEORDEN"),
    HojaInforme =
        if BinInforme = null then #table({"Código","Descripción","Agrupación","UM","Ítem","Capitulo","Orden de Compra","Fecha Compra","Estado Compra","Cantidad Comprada","Valor Unitario","IVA %","Valor Total"}, {})
        else Table.PromoteHeaders(Excel.Workbook(BinInforme, null, true){0}[Data], [PromoteAllScalars=true]),

    BaseInforme = Table.SelectRows(HojaInforme, each Record.Field(_, "Orden de Compra") <> null and Limpiar(Record.Field(_, "Orden de Compra")) <> ""),

    // ===== 2. Cantidad recibida por (orden, cod insumo) desde ESTADO DE ORDENES =====
    BinEstado = FnGetBinario("ESTADO DE ORDENES"),
    ColsOC = List.Transform({1..7}, each {"Columna" & Text.From(_), "td:nth-child(" & Text.From(_) & "), th:nth-child(" & Text.From(_) & ")"}),
    TablaEstado = if BinEstado = null then #table({"Columna1","Columna2","Columna3","Columna4","Columna5","Columna6","Columna7"}, {}) else Html.Table(Text.FromBinary(BinEstado, 1252), ColsOC, [RowSelector="tr"]),
    ConOrden = Table.AddColumn(TablaEstado, "__Orden", each
        let t = Limpiar([Columna1])
        in if Text.Contains(Text.Upper(t), "ORDEN DE COMPRA NO") then Text.Select(t, {"0".."9"}) else null, type text),
    OrdenFD = Table.FillDown(ConOrden, {"__Orden"}),
    Recibidas = Table.SelectRows(OrdenFD, each
        [__Orden] <> null and
        Limpiar([Columna1]) <> "" and
        Text.Select(Limpiar([Columna1]), {"0".."9"}) = Limpiar([Columna1])
    ),
    RecibidasClave = Table.AddColumn(Recibidas, "__KOrdIns", each [__Orden] & "|" & Limpiar([Columna1]), type text),
    RecibidasSolo = Table.Distinct(
        Table.AddColumn(Table.SelectColumns(RecibidasClave, {"__KOrdIns","Columna7"}), "CantRecibida", each FxNum([Columna7]), type number),
        {"__KOrdIns"}
    ),

    // ===== 3. Armar filas =====
    ConCols = Table.AddColumn(BaseInforme, "__Fila", each [
        #"Cod Orden" = Limpiar(Record.Field(_, "Orden de Compra")),
        INSUMO = Limpiar(Record.Field(_, "Descripción")),
        #"V/U" = FxNum(Record.Field(_, "Valor Unitario")),
        COMPRADO = FxNum(Record.Field(_, "Cantidad Comprada")),
        __CodIns = Limpiar(Record.Field(_, "Código")),
        PRODUCTO = Limpiar(Record.Field(_, "Descripción")),
        ITEM = Limpiar(Record.Field(_, "Ítem")),
        CAPITULO = Limpiar(Record.Field(_, "Capitulo")),
        UM = Limpiar(Record.Field(_, "UM"))
    ]),
    Expandida = Table.ExpandRecordColumn(Table.SelectColumns(ConCols, {"__Fila"}), "__Fila", {"Cod Orden","INSUMO","V/U","COMPRADO","__CodIns","PRODUCTO","ITEM","CAPITULO","UM"}),
    SoloConCompra = Table.SelectRows(Expandida, each [COMPRADO] > 0),

    ConRecibida = Table.NestedJoin(
        Table.AddColumn(SoloConCompra, "__KOrdIns", each [Cod Orden] & "|" & [__CodIns], type text),
        {"__KOrdIns"}, RecibidasSolo, {"__KOrdIns"}, "Rec", JoinKind.LeftOuter),
    ExpRecibida = Table.ExpandTableColumn(ConRecibida, "Rec", {"CantRecibida"}, {"CONSUMIDO"}),
    ConsumoDefault = Table.TransformColumns(ExpRecibida, {"CONSUMIDO", each if _ = null then 0 else _, type number}),

    // ===== 4. Delta por fecha (ordenes ya presentes en la version previa) =====
    OrdenesPrevias =
        if FechaVersion = "" or ParamCC = "" then {}
        else
            let
                path = BasePath & "/" & ParamCC & "/Versiones previas/" & FechaVersion,
                raw = try Json.Document(Web.Contents(SiteUrl, [
                    RelativePath = "/_api/web/GetFolderByServerRelativeUrl('" & FnEncode(path) & "')/Files",
                    Query = [#"$select" = "Name,ServerRelativeUrl,TimeLastModified"],
                    Headers = Headers,
                    Timeout = #duration(0,0,5,0)
                ])) otherwise null,
                tbl = if raw <> null and Record.HasFields(raw, "value") then Table.FromRecords(raw[value]) else #table({"Name","ServerRelativeUrl","TimeLastModified"}, {}),
                cand = Table.Sort(Table.SelectRows(tbl, each Text.Contains([Name], "INFORMEORDEN", Comparer.OrdinalIgnoreCase) and not Text.StartsWith([Name], "~$")), {{"TimeLastModified", Order.Descending}}),
                bin = if Table.RowCount(cand) = 0 then null else Binary.Buffer(Web.Contents(SiteUrl, [
                    RelativePath = "/_api/web/GetFileByServerRelativeUrl('" & FnEncode(cand{0}[ServerRelativeUrl]) & "')/$value",
                    Headers = Headers,
                    Timeout = #duration(0,0,10,0)
                ])),
                hoja = if bin = null then #table({"Orden de Compra"}, {}) else Table.PromoteHeaders(Excel.Workbook(bin, null, true){0}[Data], [PromoteAllScalars=true]),
                ordenes = if List.Contains(Table.ColumnNames(hoja), "Orden de Compra")
                          then List.Distinct(List.Transform(Table.Column(hoja, "Orden de Compra"), each Limpiar(_)))
                          else {}
            in
                ordenes,
    OrdenesPreviasBuf = List.Buffer(OrdenesPrevias),
    SinPrevias = Table.SelectRows(ConsumoDefault, each not List.Contains(OrdenesPreviasBuf, [Cod Orden])),

    // ===== 5. Preservar Estado y excluir completadas =====
    RangoPrevio = try Excel.CurrentWorkbook(){[Name="RangoOrdenesCompra"]}[Content] otherwise null,
    ColsEsperadas = {"Cod Orden","INSUMO","V/U","COMPRADO","CONSUMIDO","PRODUCTO","ITEM","CAPITULO","UM","Estado","Fecha Hora","Error"},
    Previo =
        if RangoPrevio = null or Table.RowCount(RangoPrevio) = 0 then #table(ColsEsperadas, {})
        else
            let
                P = Table.PromoteHeaders(RangoPrevio, [PromoteAllScalars=true]),
                Ok = if List.Contains(Table.ColumnNames(P), "Cod Orden") then P else #table(ColsEsperadas, {})
            in
                Table.SelectRows(Ok, each [Cod Orden] <> null and Text.Trim(Text.From([Cod Orden])) <> ""),
    PrevClave = Table.AddColumn(Previo, "__K", each Text.Upper(Text.Trim(Text.From([Cod Orden]))) & "|" & Text.Upper(Text.Trim(Text.From([INSUMO]))) & "|" & Text.Upper(Text.Trim(Text.From([ITEM]))), type text),
    PrevSolo = Table.Distinct(Table.SelectColumns(PrevClave, {"__K","Estado","Fecha Hora","Error"}), {"__K"}),

    ActualClave = Table.AddColumn(SinPrevias, "__K", each Text.Upper(Text.Trim([Cod Orden])) & "|" & Text.Upper(Text.Trim([INSUMO])) & "|" & Text.Upper(Text.Trim([ITEM])), type text),
    JoinPrev2 = Table.NestedJoin(ActualClave, {"__K"}, PrevSolo, {"__K"}, "PrevEstado", JoinKind.LeftOuter),
    ExpPrev = Table.ExpandTableColumn(JoinPrev2, "PrevEstado", {"Estado","Fecha Hora","Error"}, {"Estado","Fecha Hora","Error"}),
    SinHechas = Table.SelectRows(ExpPrev, each
        [Estado] = null or
        not (Text.StartsWith(Text.Upper(Text.Trim(Text.From([Estado]))), "OK") or Text.StartsWith(Text.Upper(Text.Trim(Text.From([Estado]))), "CREAD"))
    ),

    Final = Table.SelectColumns(SinHechas, ColsEsperadas),
    Orden = Table.Sort(Final, {{"Cod Orden", Order.Ascending}, {"ITEM", Order.Ascending}})
in
    Orden
