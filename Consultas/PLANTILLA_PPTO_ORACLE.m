// PLANTILLA_PPTO_ORACLE
// Genera la MISMA plantilla de importacion de presupuesto (21 columnas,
// identica a PLANTILLA_PPTO_SINCO) pero para proyectos de origen ORACLE
// (VERSALLES, MONGUI y futuros). Fuentes en Actual/ del proyecto:
//
//   CONTROL.xlsx  - arbol CBS/WBS completo del presupuesto (codigo + nombre).
//                   El numero de niveles VARIA por proyecto (VERSALLES usa
//                   codigos de 3 segmentos tipo 01-01-001; MONGUI de 6 tipo
//                   A-10-10-30-3160-001), por eso toda la jerarquia se
//                   resuelve de forma generica: un segmento "en ceros" marca
//                   nivel no usado, y los ancestros de un codigo se generan
//                   poniendo en ceros sus segmentos finales progresivamente.
//                   OJO: el xlsx viene con la dimension declarada mal
//                   (A1:D3); Power Query lee las celdas reales sin problema.
//   ASEGURADO.xls - detalle plano por articulo (BIFF real, NO html). Las
//                   filas con Proceso = PRESUPUESTO son el presupuesto de
//                   construccion por articulo y CBS hoja. La columna
//                   "V.r Unitrio" del reporte esta ROTA (trae una constante
//                   identica en todas las filas), asi que el unitario se
//                   calcula como V.r Total / Cantidad.
//
// Mapeo de jerarquia a la plantilla (que solo maneja Capitulo > Actividad
// (+SUBCAPITULO) > Insumo):
//   Capitulo    = ancestro de nivel 1 (solo el primer segmento no-cero)
//   Actividad   = codigo CBS hoja que aparece en el PRESUPUESTO del asegurado
//                 (CANTIDAD=1 y UM=GL: Oracle no maneja cantidad/UM de
//                 actividad, los valores van en los articulos)
//   SUBCAPITULO = ancestro intermedio MAS PROFUNDO que exista en CONTROL
//                 (los niveles intermedios extra de proyectos como MONGUI
//                 se saltan; queda el padre inmediato real)
//   Insumo      = articulo agrupado por (CBS, Articulo) sumando Cantidad y
//                 V.r Total (el reporte trae multiples registros uxer por par)
//
// Tipo Insumo por prefijo del codigo de articulo contra el Maestro Tipos
// Insumos del SINCO destino: H=alquiler equipos->E, L=laboral->N (nomina),
// M y P=materiales->M, S=subcontratos->S, Z=otros->Z, desconocido->Y.
let
    SiteUrl = "https://colsubsidio365.sharepoint.com/sites/MiGerenciaViv",
    Headers = [Accept="application/json;odata=nometadata"],
    FnEncode = F_Globales[FnEncode],
    ArchivosProyecto = Table.Buffer(SP_Archivos_Proyecto),

    FxGetBinary = (textoArchivo as text) as binary =>
        let
            Filas = Table.SelectRows(ArchivosProyecto, each Text.Contains([Name], textoArchivo, Comparer.OrdinalIgnoreCase)),
            Ordenadas = Table.Sort(Filas, {{"Centro de Costos", Order.Ascending}, {"TimeLastModified", Order.Descending}}),
            Binario =
                if Table.RowCount(Ordenadas) = 0 then
                    error Error.Record(
                        "PLANTILLA_PPTO_ORACLE",
                        "No se encontró el archivo requerido: " & textoArchivo & " para ProyectoActual=" & Text.Trim(ProyectoActual),
                        [ProyectoActual = Text.Trim(ProyectoActual), ArchivoBuscado = textoArchivo]
                    )
                else
                    Binary.Buffer(Web.Contents(SiteUrl, [
                        RelativePath = "/_api/web/GetFileByServerRelativeUrl('" & FnEncode(Ordenadas{0}[ServerRelativeUrl]) & "')/$value",
                        Headers = Headers,
                        Timeout = #duration(0, 0, 5, 0)
                    ])),

    Limpiar = (v as any) as text =>
        let t = if v = null then "" else Text.From(v)
        in Text.Trim(Text.Replace(Text.Replace(Text.Replace(t, "#(lf)", " "), "#(cr)", " "), "#(00A0)", " ")),

    FnNum = (v as any) as nullable number =>
        if v = null then null
        else if Value.Is(v, type number) then v
        else try Number.FromText(Text.Trim(Text.From(v)), "en-US") otherwise try Number.FromText(Text.Trim(Text.From(v)), "es-ES") otherwise null,

    // ===== utilidades de jerarquia por segmentos =====
    EsSegCero = (s as text) as logical => Text.Select(s, {"0"}) = s and s <> "",
    // codigo valido de CBS: al menos 2 segmentos separados por "-"
    EsCodigoCbs = (c as text) as logical => List.Count(Text.Split(c, "-")) >= 2,
    // capitulo (nivel 1): primer segmento intacto, el resto en ceros del mismo ancho
    FnCapitulo = (cod as text) as text =>
        let segs = Text.Split(cod, "-")
        in Text.Combine({segs{0}} & List.Transform(List.Skip(segs, 1), each Text.Repeat("0", Text.Length(_))), "-"),
    // candidatos a ancestro, del MAS PROFUNDO al mas superficial: para cada
    // posicion no-cero k (de atras hacia adelante, sin incluir la 0) se ponen
    // en ceros los segmentos desde k hasta el final
    FnAncestros = (cod as text) as list =>
        let
            segs = Text.Split(cod, "-"),
            n = List.Count(segs),
            idxNoCero = List.Select({1..n - 1}, (i) => not EsSegCero(segs{i})),
            EnCerosDesde = (k as number) as text =>
                Text.Combine(List.Transform({0..n - 1}, (i) => if i >= k then Text.Repeat("0", Text.Length(segs{i})) else segs{i}), "-")
        in
            List.Transform(List.Reverse(List.Sort(idxNoCero)), EnCerosDesde),

    // =========================================================
    // 1. CONTROL: diccionario codigo -> nombre del arbol CBS
    // =========================================================
    BinControl = FxGetBinary("CONTROL"),
    WbControl = Excel.Workbook(BinControl, null, true),
    HojasControl = Table.SelectRows(WbControl, each [Kind] = "Sheet"),
    DataControl = HojasControl{0}[Data],
    ColsControl = Table.ColumnNames(DataControl),
    CtrlBase = Table.RenameColumns(DataControl, {{ColsControl{0}, "__Cod"}, {ColsControl{1}, "__Nom"}}),
    CtrlFilas = Table.SelectRows(Table.Skip(CtrlBase, 1), each
        let c = Limpiar([__Cod]) in c <> "" and c <> "Total" and EsCodigoCbs(c)),
    CtrlDistinct = Table.Distinct(
        Table.FromColumns({
            List.Transform(Table.Column(CtrlFilas, "__Cod"), Limpiar),
            List.Transform(Table.Column(CtrlFilas, "__Nom"), Limpiar)
        }, {"Cod", "Nom"}),
        {"Cod"}
    ),
    // record para lookups rapidos codigo -> nombre
    DictCbs = Record.FromList(Table.Column(CtrlDistinct, "Nom"), Table.Column(CtrlDistinct, "Cod")),
    FnNombre = (cod as nullable text) as nullable text =>
        if cod = null then null else Record.FieldOrDefault(DictCbs, cod, null),

    // =========================================================
    // 2. ASEGURADO: filas de PRESUPUESTO por articulo
    // Columnas posicionales (los encabezados traen "Descripción" duplicada):
    //  1 CodProy 2 NombreProy 3 Paquete 4 CodCBS 5 DescCBS 6 Articulo
    //  7 DescArticulo 8 UM 9 Proceso 10 Registro 11 Estado 12 Cantidad
    //  13 VrUnitario(ROTO) 14 VrTotal
    // =========================================================
    BinAsegurado = FxGetBinary("ASEGURADO"),
    WbAseg = Excel.Workbook(BinAsegurado, null, true),
    HojaAseg = Table.SelectRows(WbAseg, each [Kind] = "Sheet" and [Item] <> "XDO_METADATA"){0}[Data],
    ColsAseg = Table.ColumnNames(HojaAseg),
    AsegFilas = Table.SelectRows(HojaAseg, each
        Limpiar(Record.Field(_, ColsAseg{8})) = "PRESUPUESTO" and
        EsCodigoCbs(Limpiar(Record.Field(_, ColsAseg{3}))) and
        Limpiar(Record.Field(_, ColsAseg{5})) <> ""),
    AsegNorm = Table.FromRecords(Table.TransformRows(AsegFilas, (r) => [
        CodCBS = Limpiar(Record.Field(r, ColsAseg{3})),
        DescCBS = Limpiar(Record.Field(r, ColsAseg{4})),
        Articulo = Limpiar(Record.Field(r, ColsAseg{5})),
        DescArticulo = Limpiar(Record.Field(r, ColsAseg{6})),
        UM = Limpiar(Record.Field(r, ColsAseg{7})),
        Cantidad = FnNum(Record.Field(r, ColsAseg{11})),
        VrTotal = FnNum(Record.Field(r, ColsAseg{13}))
    ])),
    // agrupar duplicados (multiples registros uxer por par CBS|Articulo)
    AsegAgrupado = Table.Group(AsegNorm, {"CodCBS", "Articulo"}, {
        {"DescCBS", each List.First(List.RemoveNulls([DescCBS]), null), type nullable text},
        {"DescArticulo", each List.First(List.RemoveNulls([DescArticulo]), null), type nullable text},
        {"UM", each List.First(List.RemoveNulls([UM]), null), type nullable text},
        {"Cantidad", each List.Sum([Cantidad]), type nullable number},
        {"VrTotal", each List.Sum([VrTotal]), type nullable number}
    }),
    Insumos = Table.Buffer(AsegAgrupado),

    // =========================================================
    // 3. ACTIVIDADES: CBS hoja distintos del presupuesto, con su
    //    capitulo y subcapitulo resueltos contra el arbol de CONTROL
    // =========================================================
    ActividadesBase = Table.Group(Insumos, {"CodCBS"}, {
        {"DescCBS", each List.First(List.RemoveNulls([DescCBS]), null), type nullable text},
        {"VrTotalAct", each List.Sum([VrTotal]), type nullable number}
    }),
    ActividadesResueltas = Table.AddColumn(ActividadesBase, "__Jerarquia", each
        let
            cap = FnCapitulo([CodCBS]),
            ancExistentes = List.Select(FnAncestros([CodCBS]), (a) => a <> [CodCBS] and Record.FieldOrDefault(DictCbs, a, null) <> null),
            subcapCod = if List.IsEmpty(ancExistentes) or List.First(ancExistentes) = cap then null else List.First(ancExistentes)
        in
            [Cap = cap, SubcapCod = subcapCod]),
    Actividades = Table.Buffer(Table.ExpandRecordColumn(ActividadesResueltas, "__Jerarquia", {"Cap", "SubcapCod"})),

    Capitulos = Table.Distinct(Table.SelectColumns(Actividades, {"Cap"})),

    // =========================================================
    // 4. TIPO INSUMO por prefijo de articulo (contra el maestro destino)
    // =========================================================
    FnTipoInsumo = (articulo as text) as text =>
        let p = Text.Upper(Text.Start(articulo, 1))
        in
            if p = "H" then "E"        // equipos -> ALQUILER EQ. Y MAQ.
            else if p = "L" then "N"   // laboral -> NOMINA DE OBRA
            else if p = "M" then "M"   // materiales
            else if p = "P" then "M"   // productos -> MATERIALES
            else if p = "S" then "S"   // subcontratos
            else if p = "Z" then "Z"   // otros
            else "Y",                  // POR CLASIFICAR

    // =========================================================
    // 5. ARMADO DE FILAS (mismas 21 columnas de la plantilla SINCO).
    //    Claves de orden para intercalar: capitulo -> actividad -> insumos
    // =========================================================
    FilasCapitulo = Table.TransformRows(Capitulos, (r) => [
        Código = r[Cap], Descripción = FnNombre(r[Cap]) ?? r[Cap], Padre = "CD",
        UM = null, CANTIDAD = null, SUBCAPITULO = null,
        #"ID PROYECTO" = null, VERSION = null, #"ID APU" = null, #"Cant APU" = null,
        Rend = null, IVA = null, VrUnitSinIVA = null, #"Tipo Insumo" = null,
        Agrupacion = null, #"COD CLIENTE" = null, #"Precio Cliente" = null, Clase = null,
        Tipo = "Capítulo", #"Vr Unitario" = null, #"Vr Total" = null,
        __K1 = r[Cap], __K2 = "", __K3 = 0, __K4 = ""
    ]),
    FilasActividad = Table.TransformRows(Actividades, (r) => [
        Código = r[CodCBS], Descripción = FnNombre(r[CodCBS]) ?? r[DescCBS], Padre = r[Cap],
        UM = "GL", CANTIDAD = 1, SUBCAPITULO = FnNombre(r[SubcapCod]),
        #"ID PROYECTO" = null, VERSION = null, #"ID APU" = null, #"Cant APU" = null,
        Rend = null, IVA = null, VrUnitSinIVA = null, #"Tipo Insumo" = null,
        Agrupacion = null, #"COD CLIENTE" = null, #"Precio Cliente" = null, Clase = null,
        Tipo = "Actividad", #"Vr Unitario" = r[VrTotalAct], #"Vr Total" = r[VrTotalAct],
        __K1 = r[Cap], __K2 = r[CodCBS], __K3 = 1, __K4 = ""
    ]),
    ActPorCod = Record.FromList(Table.TransformRows(Actividades, (r) => [Cap = r[Cap]]), Table.Column(Actividades, "CodCBS")),
    FilasInsumo = Table.TransformRows(Insumos, (r) =>
        let
            cap = Record.FieldOrDefault(ActPorCod, r[CodCBS], [Cap = null])[Cap],
            vrUnit = if r[Cantidad] <> null and r[Cantidad] <> 0 and r[VrTotal] <> null then Number.Round(r[VrTotal] / r[Cantidad], 6) else r[VrTotal]
        in [
            Código = null, Descripción = r[Articulo] & " - " & (r[DescArticulo] ?? ""), Padre = r[CodCBS],
            UM = if r[UM] = null or r[UM] = "" then null else Text.Upper(r[UM]), CANTIDAD = null, SUBCAPITULO = null,
            #"ID PROYECTO" = null, VERSION = null, #"ID APU" = null, #"Cant APU" = r[Cantidad],
            Rend = null, IVA = null, VrUnitSinIVA = vrUnit, #"Tipo Insumo" = FnTipoInsumo(r[Articulo]),
            Agrupacion = "OTROS", #"COD CLIENTE" = null, #"Precio Cliente" = null, Clase = null,
            Tipo = "Insumo", #"Vr Unitario" = r[VrTotal], #"Vr Total" = null,
            __K1 = cap, __K2 = r[CodCBS], __K3 = 2, __K4 = r[Articulo]
        ]),

    TodasLasFilas = Table.FromRecords(FilasCapitulo & FilasActividad & FilasInsumo),
    Ordenada = Table.Sort(TodasLasFilas, {{"__K1", Order.Ascending}, {"__K2", Order.Ascending}, {"__K3", Order.Ascending}, {"__K4", Order.Ascending}}),
    SinClaves = Table.RemoveColumns(Ordenada, {"__K1", "__K2", "__K3", "__K4"}),

    // =========================================================
    // 6. FILA MAESTRA CD Y TIPOS (identico a la plantilla SINCO)
    // =========================================================
    FilaCD = Table.FromRecords({[Código = "CD", Descripción = "COSTOS DIRECTOS", Padre = null, UM = null, CANTIDAD = null, SUBCAPITULO = null, #"ID PROYECTO" = null, VERSION = null, #"ID APU" = null, #"Cant APU" = null, Rend = null, IVA = null, VrUnitSinIVA = null, #"Tipo Insumo" = null, Agrupacion = null, #"COD CLIENTE" = null, #"Precio Cliente" = null, Clase = null, Tipo = "CD", #"Vr Unitario" = null, #"Vr Total" = null]}),
    TablaFinal = Table.Combine({FilaCD, SinClaves}),

    TiposDefinidos = Table.TransformColumnTypes(TablaFinal,{
        {"Código", type text}, {"Descripción", type text}, {"Padre", type text}, {"UM", type text},
        {"CANTIDAD", type number}, {"SUBCAPITULO", type text}, {"ID PROYECTO", type text},
        {"VERSION", type text}, {"ID APU", type text}, {"Cant APU", type number},
        {"Rend", type number}, {"IVA", type number}, {"VrUnitSinIVA", type number},
        {"Tipo Insumo", type text}, {"Agrupacion", type text}, {"COD CLIENTE", type text},
        {"Precio Cliente", type number}, {"Clase", type text},
        {"Tipo", type text}, {"Vr Unitario", type number}, {"Vr Total", type number}
    })
in
    TiposDefinidos
