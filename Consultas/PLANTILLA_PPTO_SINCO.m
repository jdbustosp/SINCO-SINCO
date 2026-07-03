let
    // =========================================================
    // 1. CONEXIÓN Y BÚSQUEDA SOBRE EL ÍNDICE DEL PROYECTO
    // =========================================================
    ParamProyecto = Text.Trim(ProyectoActual),
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
                        "PLANTILLA_PPTO_SINCO",
                        "No se encontró el archivo requerido: " & textoArchivo & " para ProyectoActual=" & ParamProyecto,
                        [ProyectoActual = ParamProyecto, ArchivoBuscado = textoArchivo]
                    )
                else
                    Binary.Buffer(Web.Contents(SiteUrl, [
                        RelativePath = "/_api/web/GetFileByServerRelativeUrl('" & FnEncode(Ordenadas{0}[ServerRelativeUrl]) & "')/$value",
                        Headers = Headers,
                        Timeout = #duration(0, 0, 5, 0)
                    ]))
        in
            Binario,

    FxLimpiarCelda = (t as nullable text) as text =>
        if t = null then "" else Text.Trim(Text.Replace(Text.Replace(Text.Replace(Text.From(t), "#(lf)", " "), "#(cr)", " "), "#(00A0)", " ")),

    // =========================================================
    // 2. EXTRACCIÓN DEL DICCIONARIO: "SEGUIMIENTO POR ITEMS"
    // =========================================================
    BinarioSeg = FxGetBinary("SEGUIMIENTO POR ITEMS"),
    TextoHTML_Seg = Text.FromBinary(BinarioSeg, 65001),

    ColsSeg_HTML = List.Transform({1..3}, each {"Columna" & Text.From(_), "td:nth-child(" & Text.From(_) & "), th:nth-child(" & Text.From(_) & ")"}),
    TablaSegCruda = Html.Table(TextoHTML_Seg, ColsSeg_HTML, [RowSelector="tr"]),

    SegLimpiaCod = Table.AddColumn(TablaSegCruda, "CodigoSeg", each FxLimpiarCelda([Columna1])),
    SegLimpiaTipo = Table.AddColumn(SegLimpiaCod, "TipoInsumoSeg", each FxLimpiarCelda([Columna3])),

    SegFiltro = Table.SelectRows(SegLimpiaTipo, each [CodigoSeg] <> "" and [TipoInsumoSeg] <> ""),
    DiccionarioSeguimiento = Table.Buffer(Table.Distinct(Table.SelectColumns(SegFiltro, {"CodigoSeg", "TipoInsumoSeg"}), {"CodigoSeg"})),

    // =========================================================
    // 2b. MAPA ACTIVIDAD -> SUBCAPÍTULO
    // El reporte de Analisis de Precios Unitarios NO trae el subcapitulo en
    // ninguna parte, por eso antes se adivinaba por sufijo de texto contra
    // una lista fija de nombres. El reporte de Seguimiento SI trae el
    // marcador estructural de fila "SubCapitulo :X" (como ya pasa con
    // Capitulo), asi que se extrae de ahi y se cruza por codigo de actividad.
    // =========================================================
    EsCodigoCapSeg = (c as text) as logical =>
        let p = Text.Split(c, ".") in List.Count(p) = 2 and p{1} = "000" and (try Number.FromText(p{0}) otherwise null) <> null,
    EsCodigoActSeg = (c as text) as logical =>
        let p = Text.Split(c, ".") in List.Count(p) = 2 and p{1} <> "000" and Text.Length(p{1}) > 0 and (try Number.FromText(p{0}) otherwise null) <> null and (try Number.FromText(p{1}) otherwise null) <> null,

    ClasSegSubcap = Table.AddColumn(TablaSegCruda, "ClaseSeg", each
        let c0 = FxLimpiarCelda([Columna1])
        in
            if Text.StartsWith(c0, "SubCapitulo") then "SubCap"
            else if EsCodigoCapSeg(c0) then "Capitulo"
            else if EsCodigoActSeg(c0) then "Actividad"
            else "Otro",
        type text),
    AddSubCapRaw = Table.AddColumn(ClasSegSubcap, "SubCapRaw", each
        if [ClaseSeg] = "SubCap" then Text.Trim(Text.AfterDelimiter(FxLimpiarCelda([Columna1]), ":"))
        else if [ClaseSeg] = "Capitulo" then ""
        else null,
        type text),
    FillSubCap = Table.FillDown(AddSubCapRaw, {"SubCapRaw"}),
    SoloActSeg = Table.SelectRows(FillSubCap, each [ClaseSeg] = "Actividad"),
    ConCodActSeg = Table.AddColumn(SoloActSeg, "CodActSeg", each FxLimpiarCelda([Columna1]), type text),
    MapaSubcapActividad = Table.Buffer(Table.Distinct(Table.SelectColumns(ConCodActSeg, {"CodActSeg", "SubCapRaw"}), {"CodActSeg"})),

    // =========================================================
    // 3. EXTRACCIÓN DE LA BASE PRINCIPAL: "ANALISIS DE PRECIOS"
    // =========================================================
    BinarioAPU = FxGetBinary("ANALISIS DE PRECIOS UNITARIOS"),
    TextoHTML_APU = Text.FromBinary(BinarioAPU, 65001),

    Columnas_HTML = List.Transform({1..10}, each {"Columna" & Text.From(_), "td:nth-child(" & Text.From(_) & "), th:nth-child(" & Text.From(_) & ")"}),
    TablaCruda = Html.Table(TextoHTML_APU, Columnas_HTML, [RowSelector="tr"]),

    LimpiarColumnas = Table.AddColumn(TablaCruda, "Col1_Limpia", each FxLimpiarCelda([Columna1])),
    ConIndice = Table.AddIndexColumn(LimpiarColumnas, "IndiceFila", 1, 1, Int64.Type),

    FiltroBasura = Table.SelectRows(ConIndice, each
        [Col1_Limpia] <> "" and
        not Text.Contains(Text.Upper([Col1_Limpia]), "ANALISIS DE PRECIOS") and
        not Text.Contains(Text.Upper([Col1_Limpia]), "DETALLE POR ITEMS") and
        [Col1_Limpia] <> "Item" and
        [Col1_Limpia] <> "Act. Todo Costo" and
        [Col1_Limpia] <> "Total" and
        [Col1_Limpia] <> "Subcontratos" and
        [Col1_Limpia] <> "Materiales" and
        [Col1_Limpia] <> "Mano de Obra" and
        [Col1_Limpia] <> "Equipos" and
        [Col1_Limpia] <> "Transportes"
    ),

    // =========================================================
    // 4. EL CEREBRO: DEJARLO "TAL CUAL VIENE"
    // =========================================================
    Clasificador = Table.AddColumn(FiltroBasura, "Datos", each
        let
            texto = [Col1_Limpia],
            textoUpper = Text.Upper(texto),
            esCapitulo = Text.StartsWith(textoUpper, "CAPITULO") or Text.StartsWith(textoUpper, "CAPÍTULO"),
            tieneGuion = Text.Contains(texto, "-"),

            textoSinCap = Text.Trim(Text.Replace(Text.Replace(textoUpper, "CAPÍTULO", ""), "CAPITULO", "")),
            codCap = if esCapitulo then Text.Trim(Text.BeforeDelimiter(textoSinCap, " ")) else "",
            descCap = if esCapitulo then Text.Trim(Text.AfterDelimiter(textoSinCap, " ")) else "",

            codGuion = if tieneGuion then Text.Trim(Text.BeforeDelimiter(texto, "-")) else "",

            // AQUÍ ESTÁ EL CAMBIO: Extraemos todo después del código y le quitamos un posible guion extra al inicio. El resto queda INTACTO.
            descActividadRaw = if tieneGuion then Text.Trim(Text.AfterDelimiter(texto, "-")) else "",
            descActividad = if Text.StartsWith(descActividadRaw, "-") then Text.Trim(Text.Middle(descActividadRaw, 1)) else descActividadRaw,

            esActividad = tieneGuion and Text.Contains(codGuion, ".") and not esCapitulo,
            esInsumo = tieneGuion and not Text.Contains(codGuion, ".") and not esCapitulo,

            tipoFila = if esCapitulo then "Capítulo" else if esActividad then "Actividad" else if esInsumo then "Insumo" else "Otro",
            codigoFinal = if esCapitulo then codCap else if esActividad then codGuion else null,

            descProvisional = if esCapitulo then descCap else if esActividad then descActividad else if esInsumo then texto else texto,

            umRaw = if esActividad then FxLimpiarCelda([Columna3]) else if esInsumo then FxLimpiarCelda([Columna2]) else null,
            umFinal = if umRaw <> null and umRaw <> "" then Text.Upper(umRaw) else null,

            cantFinalStr = if esActividad then FxLimpiarCelda([Columna4]) else null,
            cantNum = try Number.FromText(cantFinalStr, "en-US") otherwise try Number.FromText(cantFinalStr, "es-ES") otherwise null,

            cantApuStr = if esInsumo then FxLimpiarCelda([Columna3]) else null,
            vrUnitSinIvaStr = if esInsumo then FxLimpiarCelda([Columna4]) else null,
            rendStr = if esInsumo then FxLimpiarCelda([Columna5]) else null,

            cantApuNum = try Number.FromText(cantApuStr, "en-US") otherwise try Number.FromText(cantApuStr, "es-ES") otherwise null,
            vrUnitSinIvaNum = try Number.FromText(vrUnitSinIvaStr, "en-US") otherwise try Number.FromText(vrUnitSinIvaStr, "es-ES") otherwise null,
            rendNum = try Number.FromText(rendStr, "en-US") otherwise try Number.FromText(rendStr, "es-ES") otherwise null,

            vrUnitCol_I_Str = if esInsumo then FxLimpiarCelda([Columna6]) else null,
            vrUnitNum = try Number.FromText(vrUnitCol_I_Str, "en-US") otherwise try Number.FromText(vrUnitCol_I_Str, "es-ES") otherwise null
        in
            [Tipo = tipoFila, Código = codigoFinal, Descripción = descProvisional, UM = umFinal, Cantidad = cantNum, CantAPU = cantApuNum, VrUnitSinIVA = vrUnitSinIvaNum, Rend = rendNum, VrUnitario_Temp = vrUnitNum]
    ),

    Expandido = Table.ExpandRecordColumn(Clasificador, "Datos", {"Tipo", "Código", "Descripción", "UM", "Cantidad", "CantAPU", "VrUnitSinIVA", "Rend", "VrUnitario_Temp"}),
    SoloValidos = Table.SelectRows(Expandido, each [Tipo] <> "Otro" and ([Código] <> null or [Tipo] = "Insumo")),

    // =========================================================
    // 5. IDENTIFICAR SUBCAPÍTULO (cruce estructural contra el mapa
    // Actividad->Subcapitulo construido en 2b, por codigo de actividad).
    // Respaldo: actividades presupuestadas que aun no tienen ejecucion
    // registrada en Seguimiento (por eso no estan en el mapa) pero traen
    // el subcapitulo como sufijo en su propio nombre; se compara contra
    // TODOS los subcapitulos ya vistos en el proyecto (dinamico, no una
    // lista fija de nombres).
    // =========================================================
    SubcapsConocidos = List.Buffer(List.Distinct(List.RemoveNulls(List.Transform(
        List.Select(MapaSubcapActividad[SubCapRaw], each _ <> null and Text.Trim(_) <> ""),
        each Text.Trim(_)
    )))),

    JoinSubcap = Table.NestedJoin(SoloValidos, {"Código"}, MapaSubcapActividad, {"CodActSeg"}, "SubcapGroup", JoinKind.LeftOuter),
    ExpandSubcap = Table.ExpandTableColumn(JoinSubcap, "SubcapGroup", {"SubCapRaw"}, {"SubCapRaw"}),

    ExtraccionSubcap = Table.AddColumn(ExpandSubcap, "DatosSubcap", each
        [
            DescFinal = [Descripción], // AQUÍ DEJAMOS LA DESCRIPCIÓN INTACTA
            Subcap =
                if [Tipo] <> "Actividad" then null
                else if [SubCapRaw] <> null and Text.Trim([SubCapRaw]) <> "" then Text.Trim([SubCapRaw])
                else
                    let
                        descUpper = Text.Upper(if [Descripción] = null then "" else [Descripción]),
                        coincidencias = List.Select(SubcapsConocidos, (s) => Text.EndsWith(descUpper, " " & Text.Upper(s)) or Text.EndsWith(descUpper, "-" & Text.Upper(s))),
                        masLargo = if List.IsEmpty(coincidencias) then null else List.Last(List.Sort(coincidencias, (a, b) => Text.Length(a) - Text.Length(b)))
                    in
                        masLargo
        ]
    ),

    ExpandidoSubcap = Table.RemoveColumns(Table.ExpandRecordColumn(ExtraccionSubcap, "DatosSubcap", {"DescFinal", "Subcap"}), {"SubCapRaw"}),

    // =========================================================
    // 6. LÓGICA DE PADRES Y JERARQUÍA
    // =========================================================
    AddPadreCap = Table.AddColumn(ExpandidoSubcap, "MemoriaCapitulo", each if [Tipo] = "Capítulo" then [Código] else null),
    AddPadreAct = Table.AddColumn(AddPadreCap, "MemoriaActividad", each if [Tipo] = "Actividad" then [Código] else null),
    RellenarHaciaAbajo = Table.FillDown(AddPadreAct, {"MemoriaCapitulo", "MemoriaActividad"}),

    CalcularPadre = Table.AddColumn(RellenarHaciaAbajo, "Padre", each
        if [Tipo] = "Capítulo" then "CD"
        else if [Tipo] = "Actividad" then [MemoriaCapitulo]
        else if [Tipo] = "Insumo" then [MemoriaActividad]
        else null,
        type text
    ),

    BaseEnMemoria = Table.Buffer(CalcularPadre),

    // =========================================================
    // 7. MATEMÁTICA
    // =========================================================
    TablaInsumos = Table.SelectRows(BaseEnMemoria, each [Tipo] = "Insumo"),
    SumaPorActividad = Table.Group(TablaInsumos, {"Padre"}, {{"SumaVrUnitario", each List.Sum([VrUnitario_Temp]), type nullable number}}),

    JoinSuma = Table.NestedJoin(BaseEnMemoria, {"Código"}, SumaPorActividad, {"Padre"}, "SumaGroup", JoinKind.LeftOuter),
    ExpandSuma = Table.ExpandTableColumn(JoinSuma, "SumaGroup", {"SumaVrUnitario"}, {"SumaVrUnitario"}),

    CalculoVrUnitario = Table.AddColumn(ExpandSuma, "Vr Unitario", each if [Tipo] = "Actividad" then [SumaVrUnitario] else [VrUnitario_Temp]),
    CalculoVrTotal = Table.AddColumn(CalculoVrUnitario, "Vr Total", each if [Tipo] = "Actividad" and [Cantidad] <> null and [Vr Unitario] <> null then [Cantidad] * [Vr Unitario] else null),
    Reordenado = Table.Sort(CalculoVrTotal, {{"IndiceFila", Order.Ascending}}),

    // =========================================================
    // 8. CRUCE CON SEGUIMIENTO
    // =========================================================
    ReordenadoLlave = Table.AddColumn(Reordenado, "LlaveInsumo", each if [Tipo] = "Insumo" then Text.Trim(Text.BeforeDelimiter([DescFinal], "-")) else null),
    JoinSeguimiento = Table.NestedJoin(ReordenadoLlave, {"LlaveInsumo"}, DiccionarioSeguimiento, {"CodigoSeg"}, "SegGroup", JoinKind.LeftOuter),
    ExpandSeguimiento = Table.ExpandTableColumn(JoinSeguimiento, "SegGroup", {"TipoInsumoSeg"}, {"TipoInsumoSeg"}),

    // Tipos validos en el Maestro Tipos Insumos del SINCO destino. Cualquier
    // tipo de la constructora que no exista alli (A, C, F, V, minusculas, etc.)
    // hace fallar la importacion de la plantilla, asi que se mapea a
    // "Y" (POR CLASIFICAR).
    TiposValidos = {"E","I","O","M","N","Z","Y","P","S","T","X"},
    AddTipoInsumo = Table.AddColumn(ExpandSeguimiento, "Tipo Insumo", each
        if [Tipo] = "Insumo" then
            let
                crudo = if [TipoInsumoSeg] <> null and [TipoInsumoSeg] <> "" then Text.Upper(Text.Trim([TipoInsumoSeg])) else "S"
            in
                if List.Contains(TiposValidos, crudo) then crudo else "Y"
        else null),
    AddAgrupacion = Table.AddColumn(AddTipoInsumo, "Agrupacion", each if [Tipo] = "Insumo" then "OTROS" else null),

    FinalCols = Table.AddColumn(AddAgrupacion, "IVA", each null, type number),
    AddIDP = Table.AddColumn(FinalCols, "ID PROYECTO", each null, type text),
    AddVer = Table.AddColumn(AddIDP, "VERSION", each null, type text),
    AddIDA = Table.AddColumn(AddVer, "ID APU", each null, type text),
    AddCodC = Table.AddColumn(AddIDA, "COD CLIENTE", each null, type text),
    AddPreC = Table.AddColumn(AddCodC, "Precio Cliente", each null, type number),
    AddClase = Table.AddColumn(AddPreC, "Clase", each null, type text),

    // LAS 21 COLUMNAS EXACTAS
    ColumnasOrdenadas = Table.SelectColumns(AddClase, {
        "Código", "DescFinal", "Padre", "UM", "Cantidad", "Subcap",
        "ID PROYECTO", "VERSION", "ID APU", "CantAPU", "Rend", "IVA", "VrUnitSinIVA",
        "Tipo Insumo", "Agrupacion", "COD CLIENTE", "Precio Cliente", "Clase",
        "Tipo", "Vr Unitario", "Vr Total"
    }),

    RenombrarFinal = Table.RenameColumns(ColumnasOrdenadas, {{"DescFinal", "Descripción"}, {"Cantidad", "CANTIDAD"}, {"Subcap", "SUBCAPITULO"}, {"CantAPU", "Cant APU"}}),

    // =========================================================
    // 9. FILA MAESTRA (COSTOS DIRECTOS) Y DEFINICIÓN DE TIPOS
    // =========================================================
    FilaCD = Table.FromRecords({[Código = "CD", Descripción = "COSTOS DIRECTOS", Padre = null, UM = null, CANTIDAD = null, SUBCAPITULO = null, #"ID PROYECTO" = null, VERSION = null, #"ID APU" = null, #"Cant APU" = null, Rend = null, IVA = null, VrUnitSinIVA = null, #"Tipo Insumo" = null, Agrupacion = null, #"COD CLIENTE" = null, #"Precio Cliente" = null, Clase = null, Tipo = "CD", #"Vr Unitario" = null, #"Vr Total" = null]}),

    TablaFinal = Table.Combine({FilaCD, RenombrarFinal}),

    TiposDefinidos = Table.TransformColumnTypes(TablaFinal,{
        {"Código", type text}, {"Descripción", type text}, {"Padre", type text}, {"UM", type text},
        {"CANTIDAD", type number}, {"SUBCAPITULO", type text}, {"ID PROYECTO", type text},
        {"VERSION", type text}, {"ID APU", type text}, {"Cant APU", type number},
        {"Rend", type number}, {"IVA", type number}, {"VrUnitSinIVA", type number},
        {"Tipo Insumo", type text}, {"Agrupacion", type text}, {"COD CLIENTE", type text},
        {"Precio Cliente", type number}, {"Clase", type text},
        {"Tipo", type text}, {"Vr Unitario", type number}, {"Vr Total", type number}
    }),

    TiposDefinidosBuffer = Table.Buffer(TiposDefinidos),

    // =========================================================
    // 10. LIMPIEZA DE CERO Y ELIMINACIÓN DE HUÉRFANOS
    // =========================================================
    ActividadesEnCero = Table.SelectRows(TiposDefinidosBuffer, each [Tipo] = "Actividad" and [CANTIDAD] = 0),
    CodigosListaNegra = List.Buffer(Table.Column(ActividadesEnCero, "Código")),

    TablaLimpia = Table.SelectRows(TiposDefinidosBuffer, each
        not ([Tipo] = "Actividad" and List.Contains(CodigosListaNegra, [Código] ?? "")) and
        not ([Tipo] = "Insumo" and List.Contains(CodigosListaNegra, [Padre] ?? ""))
    ),

    SinHuerfanos = Table.SelectRows(TablaLimpia, each [Padre] <> null or [Código] = "CD")
in
    SinHuerfanos
