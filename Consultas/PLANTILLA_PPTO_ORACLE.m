// PLANTILLA_PPTO_ORACLE
// Genera la MISMA plantilla de importacion de presupuesto (21 columnas,
// identica a PLANTILLA_PPTO_SINCO) pero para proyectos de origen ORACLE
// (VERSALLES, MONGUI y futuros). Fuentes en Actual/ del proyecto:
//
//   CONTROL.xlsx  - arbol CBS/WBS completo del presupuesto (codigo + nombre).
//                   El numero de niveles VARIA por proyecto (VERSALLES usa
//                   codigos de 3 segmentos tipo 01-01-001; MONGUI de 6 tipo
//                   A-10-10-30-3160-001). Se emiten TODOS los niveles del
//                   arbol como capitulos anidados (Padre = su ancestro
//                   existente mas profundo), de forma generica: un segmento
//                   "en ceros" marca nivel no usado y los ancestros de un
//                   codigo se generan poniendo en ceros sus segmentos
//                   finales progresivamente.
//                   OJO: el xlsx viene con la dimension declarada mal
//                   (A1:D3); Power Query lee las celdas reales sin problema.
//   ASEGURADO.xls - detalle plano por articulo (BIFF real, NO html). Las
//                   filas con Proceso = PRESUPUESTO son el presupuesto de
//                   construccion por articulo, CBS hoja y Paquete de
//                   Trabajo. La columna "V.r Unitrio" del reporte esta ROTA
//                   (trae una constante identica en todas las filas), asi
//                   que el unitario se calcula como V.r Total / Cantidad.
//
// Mapeo a la plantilla:
//   Capitulos   = TODOS los nodos ancestros del arbol CONTROL usados por las
//                 actividades (anidados: nivel 1 cuelga de CD, cada nivel
//                 inferior cuelga de su ancestro existente mas profundo).
//   Actividad   = par (CBS hoja, Paquete de Trabajo). El PAQUETE DE TRABAJO
//                 es el SUBCAPITULO (equivale a GENERALES/TORRES/etc. de los
//                 proyectos SINCO, donde la misma actividad se repite por
//                 subcapitulo). Como Oracle usa el MISMO codigo CBS en
//                 varios paquetes y el Codigo de la plantilla debe ser
//                 unico, el codigo de la actividad se sintetiza como
//                 "CBS - PAQUETE" cuando el CBS tiene mas de un paquete
//                 (si solo tiene uno, queda el CBS puro).
//                 CANTIDAD=1 y UM=GL (Oracle no maneja cantidad/UM de
//                 actividad; los valores van en los articulos).
//   Insumo      = articulo agrupado por (CBS, Paquete, Articulo) sumando
//                 Cantidad y V.r Total (el reporte trae multiples registros
//                 uxer por par).
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
                    ]))
        in
            Binario,

    Limpiar = (v as any) as text =>
        let t = if v = null then "" else Text.From(v)
        in Text.Trim(Text.Replace(Text.Replace(Text.Replace(t, "#(lf)", " "), "#(cr)", " "), "#(00A0)", " ")),

    // Los numeros del reporte vienen como TEXTO en formato espanol ("22,46",
    // "$ 1.935.023,33"). Probar "en-US" primero corrompe los valores (22,46
    // se vuelve 2246) y el simbolo $ hace fallar ambas culturas. Se limpia
    // el texto y se decide el separador decimal por la POSICION del ultimo
    // separador (si "," va despues de "." es decimal espanol y viceversa).
    FnNum = (v as any) as nullable number =>
        if v = null then null
        else if Value.Is(v, type number) then v
        else
            let
                t = Text.Select(Text.From(v), {"0".."9", ",", ".", "-"}),
                tieneComa = Text.Contains(t, ","),
                tienePunto = Text.Contains(t, "."),
                limpio =
                    if t = "" or t = "-" then null
                    else if tieneComa and tienePunto then
                        (if Text.PositionOf(t, ",", Occurrence.Last) > Text.PositionOf(t, ".", Occurrence.Last)
                         then Text.Replace(Text.Replace(t, ".", ""), ",", ".")   // 1.935.023,33 -> 1935023.33
                         else Text.Replace(t, ",", ""))                           // 1,935,023.33 -> 1935023.33
                    else if tieneComa then Text.Replace(t, ",", ".")              // 22,46 -> 22.46
                    else t                                                        // 22.46 o 2246
            in
                if limpio = null then null else try Number.FromText(limpio) otherwise null,

    // ===== utilidades de jerarquia por segmentos =====
    EsSegCero = (s as text) as logical => Text.Select(s, {"0"}) = s and s <> "",
    EsCodigoCbs = (c as text) as logical => List.Count(Text.Split(c, "-")) >= 2,
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
    // Seleccion de hoja SIN asumir que la tabla de navegacion trae las
    // columnas Kind/Item (en el host de Excel la navegacion de un .xls BIFF
    // puede venir sin ellas): Record.FieldOrDefault no falla si la columna
    // no existe, y el sheet de metadatos de BI Publisher se excluye por Name.
    FnHojaPrincipal = (wb as table) as table =>
        let
            SinMetadata = Table.SelectRows(wb, each
                not Text.Contains(Text.Upper(Record.FieldOrDefault(_, "Name", "")), "XDO_METADATA") and
                not Text.Contains(Text.Upper(Record.FieldOrDefault(_, "Item", "")), "XDO_METADATA")),
            SoloSheets = Table.SelectRows(SinMetadata, each Record.FieldOrDefault(_, "Kind", "Sheet") = "Sheet"),
            Candidatas = if Table.RowCount(SoloSheets) > 0 then SoloSheets else SinMetadata
        in
            Candidatas{0}[Data],

    BinControl = FxGetBinary("CONTROL"),
    // InferSheetDimensions es OBLIGATORIO: el xlsx de BI Publisher declara la
    // dimension rota (A1:D3) y sin esta opcion Power Query solo lee 3 filas
    // del arbol (por eso "se comian" los niveles intermedios).
    WbControl = Excel.Workbook(BinControl, [InferSheetDimensions = true, DelayTypes = true]),
    DataControl = FnHojaPrincipal(WbControl),
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
    // ancestros que SI existen en el arbol, del mas profundo al mas superficial
    FnAncestrosExistentes = (cod as text) as list =>
        List.Select(FnAncestros(cod), (a) => a <> cod and Record.FieldOrDefault(DictCbs, a, null) <> null),
    // padre = ancestro existente mas profundo; si no hay, CD
    FnPadre = (cod as text) as text =>
        let anc = FnAncestrosExistentes(cod)
        in if List.IsEmpty(anc) then "CD" else List.First(anc),

    // =========================================================
    // 2. ASEGURADO: filas de PRESUPUESTO por articulo
    // Columnas posicionales (los encabezados traen "Descripción" duplicada):
    //  1 CodProy 2 NombreProy 3 PaqueteTrabajo 4 CodCBS 5 DescCBS 6 Articulo
    //  7 DescArticulo 8 UM 9 Proceso 10 Registro 11 Estado 12 Cantidad
    //  13 VrUnitario(ROTO) 14 VrTotal
    // =========================================================
    BinAsegurado = FxGetBinary("ASEGURADO"),
    WbAseg = Excel.Workbook(BinAsegurado, null, true),
    HojaAseg = FnHojaPrincipal(WbAseg),
    ColsAseg = Table.ColumnNames(HojaAseg),
    AsegFilas = Table.SelectRows(HojaAseg, each
        Limpiar(Record.Field(_, ColsAseg{8})) = "PRESUPUESTO" and
        EsCodigoCbs(Limpiar(Record.Field(_, ColsAseg{3}))) and
        Limpiar(Record.Field(_, ColsAseg{5})) <> ""),
    AsegNorm = Table.FromRecords(Table.TransformRows(AsegFilas, (r) => [
        Paquete = Limpiar(Record.Field(r, ColsAseg{2})),
        CodCBS = Limpiar(Record.Field(r, ColsAseg{3})),
        DescCBS = Limpiar(Record.Field(r, ColsAseg{4})),
        Articulo = Limpiar(Record.Field(r, ColsAseg{5})),
        DescArticulo = Limpiar(Record.Field(r, ColsAseg{6})),
        UM = Limpiar(Record.Field(r, ColsAseg{7})),
        Cantidad = FnNum(Record.Field(r, ColsAseg{11})),
        VrTotal = FnNum(Record.Field(r, ColsAseg{13}))
    ])),
    // agrupar duplicados (multiples registros uxer por CBS|Paquete|Articulo)
    Insumos = Table.Buffer(Table.Group(AsegNorm, {"CodCBS", "Paquete", "Articulo"}, {
        {"DescCBS", each List.First(List.RemoveNulls([DescCBS]), null), type nullable text},
        {"DescArticulo", each List.First(List.RemoveNulls([DescArticulo]), null), type nullable text},
        {"UM", each List.First(List.RemoveNulls([UM]), null), type nullable text},
        {"Cantidad", each List.Sum([Cantidad]), type nullable number},
        {"VrTotal", each List.Sum([VrTotal]), type nullable number}
    })),

    // =========================================================
    // 3. ACTIVIDADES: par (CBS, Paquete de Trabajo). El paquete es el
    //    SUBCAPITULO. Codigo sintetizado "CBS - PAQUETE" solo cuando el CBS
    //    tiene mas de un paquete (para que el Codigo quede unico).
    // =========================================================
    PaquetesPorCbs = Table.Group(Insumos, {"CodCBS"}, {{"NumPaquetes", each List.Count(List.Distinct([Paquete])), Int64.Type}}),
    DictNumPaq = Record.FromList(Table.Column(PaquetesPorCbs, "NumPaquetes"), Table.Column(PaquetesPorCbs, "CodCBS")),
    FnCodActividad = (cbs as text, paquete as text) as text =>
        if Record.FieldOrDefault(DictNumPaq, cbs, 1) > 1 and paquete <> "" then cbs & " - " & paquete else cbs,

    // CANTIDAD de la actividad: en el reporte la Cantidad de cada insumo es
    // la cantidad de la ACTIVIDAD repetida (ej. CONCRETO CICLOPEO: todos los
    // insumos traen 22.46). Cuando todos los insumos comparten la cantidad,
    // esa es la CANTIDAD de la actividad y los insumos quedan con Cant APU=1;
    // si difieren (frecuente en VERSALLES), la actividad queda en 1 y cada
    // insumo con su cantidad real. La matematica cierra igual en ambos casos:
    // VrUnit actividad = suma(VrTotal insumos) / CANTIDAD.
    FnModa = (l as list) as nullable text =>
        let noVacios = List.Select(l, each _ <> null and _ <> "")
        in
            if List.IsEmpty(noVacios) then null
            else List.First(List.Sort(List.Distinct(noVacios), (a, b) =>
                Value.Compare(
                    List.Count(List.Select(noVacios, (x) => x = b)),
                    List.Count(List.Select(noVacios, (x) => x = a))
                ))),

    ActividadesBase = Table.Group(Insumos, {"CodCBS", "Paquete"}, {
        {"DescCBS", each List.First(List.RemoveNulls([DescCBS]), null), type nullable text},
        {"VrTotalAct", each List.Sum([VrTotal]), type nullable number},
        {"CantCompartida", each
            let d = List.Distinct(List.Transform(List.RemoveNulls([Cantidad]), (c) => Number.Round(c, 4)))
            in if List.Count(d) = 1 and d{0} <> 0 then d{0} else null,
            type nullable number},
        {"UMModa", each FnModa([UM]), type nullable text}
    }),
    Actividades = Table.Buffer(Table.AddColumn(ActividadesBase, "__Extra", each [
        CodAct = FnCodActividad([CodCBS], [Paquete]),
        Padre = FnPadre([CodCBS]),
        CantAct = if [CantCompartida] <> null then [CantCompartida] else 1,
        UMAct = if [CantCompartida] <> null and [UMModa] <> null then Text.Upper([UMModa]) else "GL"
    ])),
    // cantidad de actividad por par (CBS|Paquete), para dividir los insumos
    ListaQ = Table.TransformRows(Actividades, (r) => [K = r[CodCBS] & "|" & r[Paquete], Q = r[__Extra][CantAct]]),
    DictQ = Record.FromList(List.Transform(ListaQ, each [Q]), List.Transform(ListaQ, each [K])),
    FnCantAct = (cbs as text, paquete as text) as number =>
        Record.FieldOrDefault(DictQ, cbs & "|" & paquete, 1),

    // =========================================================
    // 4. CAPITULOS: TODOS los ancestros usados por las actividades,
    //    anidados (cada uno cuelga de su ancestro existente mas profundo)
    // =========================================================
    CbsUsados = List.Distinct(Table.Column(ActividadesBase, "CodCBS")),
    NodosArbol = List.Distinct(List.Combine(List.Transform(CbsUsados, FnAncestrosExistentes))),

    // =========================================================
    // 5. TIPO INSUMO y AGRUPACION desde el maestro INSUMOS_ORACLE.csv del
    //    repo (clasificacion real validada con las plantillas que ya se
    //    importaron: 507 insumos). Para insumos que no esten en el maestro,
    //    respaldo por prefijo segun la distribucion real del maestro:
    //    L->N (33/33), M->M (11/11), S->O (97/99), Z->S (174/207),
    //    P->M (111/145), H->S (4/9); y agrupacion OTROS.
    // =========================================================
    UrlMaestroIns = "https://raw.githubusercontent.com/jdbustosp/SINCO-SINCO/main/Datos/INSUMOS_ORACLE.csv?cb=" & Number.ToText(Number.From(DateTime.LocalNow())),
    CsvMaestro = try Table.Skip(Csv.Document(Web.Contents(UrlMaestroIns), [Delimiter = ";", Columns = 4, Encoding = 65001, QuoteStyle = QuoteStyle.None]), 1)
                 otherwise #table({"Column1", "Column2", "Column3", "Column4"}, {}),
    MaestroDist = Table.Distinct(CsvMaestro, {"Column1"}),
    DictTipoIns = Record.FromList(
        List.Transform(Table.Column(MaestroDist, "Column2"), Text.Trim),
        List.Transform(Table.Column(MaestroDist, "Column1"), Text.Trim)),
    DictGrupoIns = Record.FromList(
        List.Transform(Table.Column(MaestroDist, "Column4"), Text.Trim),
        List.Transform(Table.Column(MaestroDist, "Column1"), Text.Trim)),

    TiposValidos = {"E", "I", "O", "M", "N", "Z", "Y", "P", "S", "T", "X"},
    FnTipoInsumo = (articulo as text) as text =>
        let
            delMaestro = Record.FieldOrDefault(DictTipoIns, articulo, null),
            p = Text.Upper(Text.Start(articulo, 1)),
            heuristica =
                if p = "L" then "N"
                else if p = "M" then "M"
                else if p = "P" then "M"
                else if p = "S" then "O"
                else if p = "Z" then "S"
                else if p = "H" then "S"
                else "Y",
            t = if delMaestro <> null and delMaestro <> "" then Text.Upper(delMaestro) else heuristica
        in
            if List.Contains(TiposValidos, t) then t else "Y",
    FnGrupoInsumo = (articulo as text) as text =>
        let g = Record.FieldOrDefault(DictGrupoIns, articulo, null)
        in if g = null or g = "" then "OTROS" else g,

    // =========================================================
    // 6. ARMADO DE FILAS (mismas 21 columnas de la plantilla SINCO).
    //    El orden lexicografico de los codigos CBS ya pone cada ancestro
    //    antes que sus hijos (los segmentos en ceros ordenan primero);
    //    claves: (codigo nodo/CBS, paquete, tipoOrden, articulo)
    // =========================================================
    FilasCapitulo = List.Transform(NodosArbol, (nodo) => [
        Código = nodo, Descripción = FnNombre(nodo) ?? nodo, Padre = FnPadre(nodo),
        UM = null, CANTIDAD = null, SUBCAPITULO = null,
        #"ID PROYECTO" = null, VERSION = null, #"ID APU" = null, #"Cant APU" = null,
        Rend = null, IVA = null, VrUnitSinIVA = null, #"Tipo Insumo" = null,
        Agrupacion = null, #"COD CLIENTE" = null, #"Precio Cliente" = null, Clase = null,
        Tipo = "Capítulo", #"Vr Unitario" = null, #"Vr Total" = null,
        __K1 = nodo, __K2 = "", __K3 = 0, __K4 = ""
    ]),
    FilasActividad = Table.TransformRows(Actividades, (r) => [
        Código = r[__Extra][CodAct],
        Descripción = (FnNombre(r[CodCBS]) ?? r[DescCBS]) & (if r[__Extra][CodAct] <> r[CodCBS] then " - " & r[Paquete] else ""),
        Padre = r[__Extra][Padre],
        UM = r[__Extra][UMAct], CANTIDAD = r[__Extra][CantAct], SUBCAPITULO = if r[Paquete] = "" then null else r[Paquete],
        #"ID PROYECTO" = null, VERSION = null, #"ID APU" = null, #"Cant APU" = null,
        Rend = null, IVA = null, VrUnitSinIVA = null, #"Tipo Insumo" = null,
        Agrupacion = null, #"COD CLIENTE" = null, #"Precio Cliente" = null, Clase = null,
        Tipo = "Actividad",
        #"Vr Unitario" = if r[VrTotalAct] <> null then Number.Round(r[VrTotalAct] / r[__Extra][CantAct], 6) else null,
        #"Vr Total" = r[VrTotalAct],
        __K1 = r[CodCBS], __K2 = r[Paquete], __K3 = 1, __K4 = ""
    ]),
    FilasInsumo = Table.TransformRows(Insumos, (r) =>
        let
            // cantidad de la actividad a la que pertenece este insumo
            cantAct = FnCantAct(r[CodCBS], r[Paquete]),
            // precio unitario real del articulo (la columna V.r Unitrio del reporte esta rota)
            vrUnit = if r[Cantidad] <> null and r[Cantidad] <> 0 and r[VrTotal] <> null then Number.Round(r[VrTotal] / r[Cantidad], 6) else r[VrTotal],
            // cantidad por unidad de actividad (=1 cuando todos comparten la cantidad)
            cantApu = if r[Cantidad] <> null then Number.Round(r[Cantidad] / cantAct, 6) else null,
            // aporte del insumo por unidad de actividad (suma = VrUnit de la actividad)
            vrParcial = if r[VrTotal] <> null then Number.Round(r[VrTotal] / cantAct, 6) else null
        in [
            Código = null, Descripción = r[Articulo] & " - " & (r[DescArticulo] ?? ""), Padre = FnCodActividad(r[CodCBS], r[Paquete]),
            UM = if r[UM] = null or r[UM] = "" then null else Text.Upper(r[UM]), CANTIDAD = null, SUBCAPITULO = null,
            #"ID PROYECTO" = null, VERSION = null, #"ID APU" = null, #"Cant APU" = cantApu,
            Rend = 1, IVA = 0, VrUnitSinIVA = vrUnit, #"Tipo Insumo" = FnTipoInsumo(r[Articulo]),
            Agrupacion = FnGrupoInsumo(r[Articulo]), #"COD CLIENTE" = null, #"Precio Cliente" = null, Clase = null,
            Tipo = "Insumo", #"Vr Unitario" = vrParcial, #"Vr Total" = null,
            __K1 = r[CodCBS], __K2 = r[Paquete], __K3 = 2, __K4 = r[Articulo]
        ]),

    TodasLasFilas = Table.FromRecords(FilasCapitulo & FilasActividad & FilasInsumo),
    Ordenada = Table.Sort(TodasLasFilas, {{"__K1", Order.Ascending}, {"__K2", Order.Ascending}, {"__K3", Order.Ascending}, {"__K4", Order.Ascending}}),
    SinClaves = Table.RemoveColumns(Ordenada, {"__K1", "__K2", "__K3", "__K4"}),

    // =========================================================
    // 7. FILA MAESTRA CD Y TIPOS (identico a la plantilla SINCO)
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
