// INSUMOS_NUEVOS
// Pares (actividad, insumo) que estan en SEGUIMIENTO POR ITEMS pero NO en el
// catalogo APU. La plantilla inicial solo crea/asocia los insumos del APU,
// asi que todo par presente en el seguimiento de la constructora que no este
// en el APU requiere trabajo en SINCO: asociar el insumo a la actividad, y si
// el insumo no existe en absoluto en el APU, tambien crearlo en el maestro.
//
// Columnas (posiciones fijas: las consumen los flujos RPA NUEVOS INSUMOS [B,C,D,E->G]
// y ASOCIAR INSUMOS [B,F->H]):
//   A ACTIVIDADES (codigo - descripcion)   B INSUMOS (descripcion)
//   C TIPO (letra M/O/S/V/T/X)             D UM
//   E GRUPO (agrupacion fija "OTROS")      F ITEM (solo codigo actividad, para busqueda)
//   G Estado (creacion)                    H Estado Asociacion
//   I CREAR (SI = el insumo no existe en el APU y hay que crearlo en el maestro)
//
// Los codigos de insumo son del SINCO de la constructora: se usan SOLO para
// comparar entre sus dos reportes; el insumo creado en el SINCO propio recibe
// otro codigo. La preservacion de Estado usa la clave ITEM|descripcion.
let
    // Un mismo archivo QUERY UNIFICADO sirve para cualquier proyecto: si el
    // Origen es ORACLE este dominio aun no tiene equivalente implementado,
    // asi que se devuelve la tabla vacia con el mismo esquema en vez de
    // reventar el "Actualizar todo".
    OrigenActual = try Text.Upper(Text.Trim(Origen)) otherwise "SINCO",
    VacioOracle = #table({"ACTIVIDADES","INSUMOS","TIPO","UM","GRUPO","ITEM","Estado","Estado Asociacion","CREAR"}, {}),
    SiteUrl = "https://colsubsidio365.sharepoint.com/sites/MiGerenciaViv",
    Headers = [Accept="application/json;odata=nometadata"],
    FnEncode = F_Globales[FnEncode],
    ArchivosProyecto = Table.Buffer(SP_Archivos_Proyecto),

    Limpiar = (v as any) as text =>
        let t = if v = null then "" else Text.From(v)
        in Text.Trim(Text.Replace(Text.Replace(Text.Replace(t, "#(lf)", " "), "#(cr)", " "), "#(00A0)", " ")),

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

    FnGetBinario = (contiene as text) as nullable binary =>
        let
            Filas = Table.SelectRows(ArchivosProyecto, each Text.Contains([Name], contiene, Comparer.OrdinalIgnoreCase)),
            Ordenadas = Table.Sort(Filas, {{"Centro de Costos", Order.Ascending}, {"TimeLastModified", Order.Descending}})
        in
            if Table.RowCount(Ordenadas) = 0 then null else
            Binary.Buffer(Web.Contents(SiteUrl, [
                RelativePath = "/_api/web/GetFileByServerRelativeUrl('" & FnEncode(Ordenadas{0}[ServerRelativeUrl]) & "')/$value",
                Headers = Headers,
                Timeout = #duration(0, 0, 5, 0)
            ])),

    // ===== 1. SEGUIMIENTO: pares (actividad, insumo) =====
    BinSeg = FnGetBinario("SEGUIMIENTO POR ITEMS"),
    ColsSeg = List.Transform({1..4}, each {"Columna" & Text.From(_), "td:nth-child(" & Text.From(_) & "), th:nth-child(" & Text.From(_) & ")"}),
    TablaSeg = if BinSeg = null then #table({"Columna1","Columna2","Columna3","Columna4"}, {}) else Html.Table(Text.FromBinary(BinSeg, 65001), ColsSeg, [RowSelector="tr"]),

    ClasSeg = Table.AddColumn(TablaSeg, "Clase", (r as record) =>
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

    ConActCod = Table.AddColumn(ClasSeg, "ActCod", each if [Clase] = "Actividad" then Limpiar([Columna1]) else if [Clase] = "Capitulo" then "" else null, type text),
    ConActDesc = Table.AddColumn(ConActCod, "ActDesc", each if [Clase] = "Actividad" then Limpiar([Columna2]) else if [Clase] = "Capitulo" then "" else null, type text),
    FD = Table.FillDown(ConActDesc, {"ActCod", "ActDesc"}),

    InsumosSeg = Table.SelectRows(FD, each [Clase] = "Insumo" and [ActCod] <> null and [ActCod] <> ""),
    ParesSeg = Table.Distinct(
        #table(
            {"ActCod", "ActDesc", "InsCod", "InsDesc", "Tipo", "UM"},
            Table.ToList(InsumosSeg, (r) => {r{5}, r{6}, Limpiar(r{0}), Limpiar(r{1}), Limpiar(r{2}), Limpiar(r{3})})
        ),
        {"ActCod", "InsCod"}
    ),

    // ===== 2. APU (solo columna 1): pares y catalogo de insumos =====
    BinApu = FnGetBinario("ANALISIS DE PRECIOS UNITARIOS"),
    ColsApu = {{"Columna1", "td:nth-child(1), th:nth-child(1)"}},
    TablaApu = if BinApu = null then #table({"Columna1"}, {}) else Html.Table(Text.FromBinary(BinApu, 65001), ColsApu, [RowSelector="tr"]),

    ApuClasificado = Table.AddColumn(TablaApu, "Dato", each
        let
            t = Limpiar([Columna1]),
            antes = if Text.Contains(t, "-") then Text.Trim(Text.BeforeDelimiter(t, "-")) else ""
        in
            if antes = "" then null
            else if EsCodigoActividad(antes) then [K = "A", Cod = antes]
            else if EsCodigoInsumo(antes) then [K = "I", Cod = antes]
            else null),
    ApuConAct = Table.AddColumn(ApuClasificado, "ActApu", each if [Dato] <> null and [Dato][K] = "A" then [Dato][Cod] else null, type text),
    ApuFD = Table.FillDown(ApuConAct, {"ActApu"}),
    // Columnas de ApuInsRows por indice: {0}=Columna1, {1}=Dato (record), {2}=ActApu
    ApuInsRows = Table.SelectRows(ApuFD, each [Dato] <> null and [Dato][K] = "I" and [ActApu] <> null),
    ParesApu = List.Buffer(List.Distinct(Table.ToList(ApuInsRows, (r) => Text.From(r{1}[Cod]) & "|" & Text.From(r{2})))),
    InsumosApu = List.Buffer(List.Distinct(Table.ToList(ApuInsRows, (r) => Text.From(r{1}[Cod])))),

    // ===== 3. Delta: pares del seguimiento que no estan en el APU =====
    ConClavePar = Table.AddColumn(ParesSeg, "__Par", each [InsCod] & "|" & [ActCod], type text),
    Nuevos = Table.SelectRows(ConClavePar, each not List.Contains(ParesApu, [__Par])),
    ConCrear = Table.AddColumn(Nuevos, "CREAR", each if List.Contains(InsumosApu, [InsCod]) then "NO" else "SI", type text),

    // ===== 4. Columnas de salida =====
    Base = Table.AddColumn(ConCrear, "__Fila", each [
        ACTIVIDADES = [ActCod] & " - " & [ActDesc],
        INSUMOS = [InsDesc],
        // Tipos validos del Maestro Tipos Insumos del SINCO destino; lo que
        // no exista alli (A, C, F, V, minusculas...) se mapea a "Y" (POR CLASIFICAR)
        TIPO = let t = Text.Upper(Text.Trim([Tipo])) in if List.Contains({"E","I","O","M","N","Z","Y","P","S","T","X"}, t) then t else "Y",
        UM = [UM],
        GRUPO = "OTROS",
        ITEM = [ActCod]
    ]),
    Expandida = Table.ExpandRecordColumn(Table.SelectColumns(Base, {"__Fila", "CREAR"}), "__Fila", {"ACTIVIDADES","INSUMOS","TIPO","UM","GRUPO","ITEM"}),

    // ===== 5. Preservar Estado / Estado Asociacion de la hoja actual =====
    RangoPrevio = try Excel.CurrentWorkbook(){[Name="RangoInsumosNuevos"]}[Content] otherwise null,
    ColsEsperadas = {"ACTIVIDADES","INSUMOS","TIPO","UM","GRUPO","ITEM","Estado","Estado Asociacion","CREAR"},
    Previo =
        if RangoPrevio = null or Table.RowCount(RangoPrevio) = 0 then #table(ColsEsperadas, {})
        else
            let
                P = Table.PromoteHeaders(RangoPrevio, [PromoteAllScalars=true]),
                Ok = if List.Contains(Table.ColumnNames(P), "INSUMOS") then P else #table(ColsEsperadas, {})
            in
                Table.SelectRows(Ok, each [INSUMOS] <> null and Text.Trim(Text.From([INSUMOS])) <> ""),
    PrevClave = Table.AddColumn(Previo, "__K", each Text.Upper(Text.Trim(Text.From([ITEM]))) & "|" & Text.Upper(Text.Trim(Text.From([INSUMOS]))), type text),
    PrevSolo = Table.Distinct(Table.SelectColumns(PrevClave, {"__K","Estado","Estado Asociacion"}), {"__K"}),

    ActualClave = Table.AddColumn(Expandida, "__K", each Text.Upper(Text.Trim([ITEM])) & "|" & Text.Upper(Text.Trim([INSUMOS])), type text),
    JoinPrev = Table.NestedJoin(ActualClave, {"__K"}, PrevSolo, {"__K"}, "PrevEstado", JoinKind.LeftOuter),
    ExpPrev = Table.ExpandTableColumn(JoinPrev, "PrevEstado", {"Estado","Estado Asociacion"}, {"Estado","Estado Asociacion"}),

    // Excluir pares ya asociados (paso final del proceso de insumos)
    SinHechos = Table.SelectRows(ExpPrev, each [Estado Asociacion] = null or Text.Trim(Text.From([Estado Asociacion])) = ""),

    Final = Table.SelectColumns(SinHechos, {"ACTIVIDADES","INSUMOS","TIPO","UM","GRUPO","ITEM","Estado","Estado Asociacion","CREAR"}),
    Ordenado = Table.Sort(Final, {{"ITEM", Order.Ascending}, {"INSUMOS", Order.Ascending}})
in
    if OrigenActual = "ORACLE" then VacioOracle else Ordenado
