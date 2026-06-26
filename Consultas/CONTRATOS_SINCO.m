let
    // ===== 1. Parametros =====
    ParamProyecto = Text.Trim(Text.From(ProyectoActual)),
    FechaVersion = try Text.Trim(Text.From(FechaVersionComparar)) otherwise "",
    SiteUrl = "https://colsubsidio365.sharepoint.com/sites/MiGerenciaViv",
    BasePath = "/sites/MiGerenciaViv/Departamento Tecnico/COORDINACION DE PRESUPUESTOS/0. Reportes EDT - Control costos interno/" & ParamProyecto,
    Headers = [Accept = "application/json;odata=nometadata"],

    // ===== 2. Funciones auxiliares (portadas de F_Globales) =====
    FnEncode = (path as nullable text) as nullable text =>
        if path = null then null
        else Text.Combine(List.Transform(Text.Split(path, "/"), each Uri.EscapeDataString(_)), "/"),

    FnReadSPBinary = (filePath as text) as nullable binary =>
        let
            raw = try Web.Contents(SiteUrl, [
                RelativePath = "/_api/web/GetFileByServerRelativeUrl('" & FnEncode(filePath) & "')/$value",
                Headers = [Accept = "*/*"],
                Timeout = #duration(0, 0, 10, 0),
                ManualStatusHandling = {404, 429, 500, 502, 503, 504}
            ]) otherwise null,
            status = if raw = null then null else try Value.Metadata(raw)[Response.Status] otherwise 200,
            result = if raw = null or status >= 400 then null else Binary.Buffer(raw)
        in result,

    FxToNumberFlex = (value as any) as nullable number =>
        let
            v = value,
            isNum = Value.Is(v, type number),
            numeroDirecto = if isNum then Number.From(v) else null,
            t0 = if v = null then "" else Text.From(v),
            t = Text.Trim(Text.Replace(Text.Replace(t0, "#(00A0)", ""), " ", "")),
            tryUS = try Number.FromText(t, "en-US"),
            valUS = if tryUS[HasError] then null else tryUS[Value],
            tryES = try Number.FromText(t, "es-ES"),
            valES = if tryES[HasError] then null else tryES[Value],
            result = if numeroDirecto <> null then numeroDirecto
                     else if t = "" then null
                     else if valUS <> null then valUS
                     else valES
        in result,

    FnFormatCodigoAct = (raw as any) as nullable text =>
        let
            txtRaw = if raw = null then null else Text.Trim(Text.From(raw)),
            result =
                if txtRaw = null or txtRaw = "" then null
                else
                    let
                        txtNorm = Text.Replace(Text.Replace(txtRaw, ",", "."), " ", ""),
                        hasDot = Text.Contains(txtNorm, ".")
                    in
                        if hasDot then txtNorm
                        else
                            let
                                digits = Text.Select(txtNorm, {"0".."9"}),
                                len = Text.Length(digits)
                            in
                                if len <= 3 then null
                                else Text.Range(digits, 0, len - 3) & "." & Text.Range(digits, len - 3, 3)
        in result,

    FnBuildColumnas = (n as number) as list =>
        List.Transform({1..n}, each {"Columna " & Text.From(_), "td:nth-child(" & Text.From(_) & "), th:nth-child(" & Text.From(_) & ")"}),

    FnClaveLimpia = (t as nullable text) as nullable text =>
        let
            sinUnidad = if t = null then null
                        else if Text.Contains(t, "(") then Text.BeforeDelimiter(t, "(")
                        else t,
            t1 = if sinUnidad = null then null else Text.Upper(Text.Trim(sinUnidad)),
            repl = {
                {"#(00C1)","A"},{"#(00C9)","E"},{"#(00CD)","I"},
                {"#(00D3)","O"},{"#(00DA)","U"},{"#(00D1)","N"},{"#(00DC)","U"}
            },
            t2 = if t1 = null then null
                 else List.Accumulate(repl, t1, (state, current) => Text.Replace(state, current{0}, current{1})),
            t3 = if t2 = null then null else Text.Select(t2, {"A".."Z", "0".."9"}),
            result = if t3 = null or t3 = "" then null else t3
        in result,

    FnPrepareTableWithHeader = (tbl as table) as table =>
        let
            firstColName = Table.ColumnNames(tbl){0},
            firstColValues = Table.Column(tbl, firstColName),
            headerFlags = List.Transform(firstColValues, (x) =>
                let
                    txt = Text.Upper(if x = null then "" else Text.From(x)),
                    txtNorm = Text.Replace(txt, "#(00D3)", "O")
                in Text.Contains(txtNorm, "COD")),
            hasHeader = List.Contains(headerFlags, true),
            promoted = if hasHeader then
                let
                    headerIndex = List.PositionOf(headerFlags, true),
                    skipped = Table.Skip(tbl, headerIndex)
                in Table.PromoteHeaders(skipped, [PromoteAllScalars = true])
                else tbl
        in promoted,

    FnDigits = (v as any) as nullable text =>
        let t = Text.Trim(Text.From(if v = null then "" else v)), d = Text.Select(t, {"0".."9"}) in if d = "" then null else d,

    // Quita tildes/enies y caracteres de codificacion mal leida (reemplazo "?")
    FnQuitarTildes = (v as any) as nullable text =>
        let
            t0 = try (if v = null then null else Text.From(v)) otherwise null,
            repl = {
                {"á","a"},{"é","e"},{"í","i"},{"ó","o"},{"ú","u"},{"ñ","n"},{"ü","u"},
                {"Á","A"},{"É","E"},{"Í","I"},{"Ó","O"},{"Ú","U"},{"Ñ","N"},{"Ü","U"},
                {Character.FromNumber(65533), ""}
            },
            t1 = if t0 = null then null else List.Accumulate(repl, t0, (state, current) => Text.Replace(state, current{0}, current{1}))
        in t1,

    // Separa "Nombre (UM) - resto" -> [Nombre = "Nombre - resto", UM] (el parentesis puede no estar al final)
    FnSepararUM = (texto as nullable text) as record =>
        let
            t = if texto = null then "" else Text.Trim(Text.From(texto)),
            posIni = Text.PositionOf(t, "(", Occurrence.Last),
            posFin = if posIni = -1 then -1 else Text.PositionOf(t, ")", Occurrence.Last),
            valido = posIni >= 0 and posFin > posIni,
            um = if valido then Text.Trim(Text.Range(t, posIni + 1, posFin - posIni - 1)) else null,
            antes = if valido then Text.Trim(Text.Range(t, 0, posIni)) else t,
            despues = if valido then Text.Trim(Text.Range(t, posFin + 1)) else "",
            nombre = if despues = "" then antes else Text.Trim(antes & " " & despues)
        in [Nombre = nombre, UM = um],

    // ===== 3. Listado de Centros de Costo y archivos de "Actual" =====
    FolderResponse = try Json.Document(Web.Contents(SiteUrl, [
        RelativePath = "/_api/web/GetFolderByServerRelativeUrl('" & FnEncode(BasePath) & "')/Folders",
        Query = [#"$select" = "Name"],
        Headers = Headers,
        Timeout = #duration(0, 0, 5, 0)
    ])) otherwise null,

    CCFolders = if FolderResponse = null or not Record.HasFields(FolderResponse, "value")
        then #table({"Name"}, {})
        else Table.FromRecords(FolderResponse[value]),
    Centros = List.Transform(Table.Column(CCFolders, "Name"), each Text.From(_)),

    FnFilesIn = (cc as text, subfolder as text) as table =>
        let
            path = BasePath & "/" & cc & "/" & subfolder,
            raw = try Json.Document(Web.Contents(SiteUrl, [
                RelativePath = "/_api/web/GetFolderByServerRelativeUrl('" & FnEncode(path) & "')/Files",
                Query = [#"$select" = "Name,ServerRelativeUrl,TimeLastModified,Length"],
                Headers = Headers,
                Timeout = #duration(0, 0, 5, 0)
            ])) otherwise null,
            tbl = if raw <> null and Record.HasFields(raw, "value") then Table.FromRecords(raw[value]) else #table({"Name","ServerRelativeUrl","TimeLastModified","Length"}, {}),
            add = Table.AddColumn(tbl, "Centro de Costos", each cc, type text)
        in add,

    ArchivosActual = Table.Buffer(if List.Count(Centros) = 0 then #table({"Name","ServerRelativeUrl","TimeLastModified","Centro de Costos"},{}) else Table.Combine(List.Transform(Centros, each FnFilesIn(_, "Actual")))),

    FnPickLatest = (t as table, containsText as text) as nullable binary =>
        let
            candidatos = Table.Sort(Table.SelectRows(t, each Text.Contains([Name], containsText, Comparer.OrdinalIgnoreCase) and not Text.StartsWith([Name], "~$")), {{"TimeLastModified", Order.Descending}, {"Name", Order.Ascending}}),
            path = if Table.RowCount(candidatos) = 0 then null else candidatos{0}[ServerRelativeUrl]
        in if path = null then null else FnReadSPBinary(path),

    // ===== 4. Tabla maestra de Actividad(+UM)/Capitulo/Insumo oficial(+UM) =====
    Columnas_HTML25 = FnBuildColumnas(25),
    Columnas_APU = FnBuildColumnas(3),

    FxProcesarCentroCosto = (BinarioSeguimiento as binary, BinarioPresupuesto as binary) as table =>
        let
            OrigenItems = try Excel.Workbook(BinarioSeguimiento, null, true){0}[Data]
                          otherwise Html.Table(Text.FromBinary(BinarioSeguimiento, 1252), Columnas_HTML25, [RowSelector="tr"]),
            ItemsPrepared = Table.Buffer(FnPrepareTableWithHeader(OrigenItems)),

            ItemsColNames = Table.ColumnNames(ItemsPrepared),
            ItemsCodColName = ItemsColNames{0},
            ItemsDescColName = ItemsColNames{1},
            ItemsTipoColName = ItemsColNames{2},
            ItemsUMColName = ItemsColNames{3},

            ItemsWithTipoFila = Table.AddColumn(ItemsPrepared, "TipoFila", (r as record) =>
                let
                    codValue = Record.Field(r, ItemsCodColName),
                    descValue = Record.Field(r, ItemsDescColName),
                    tipoValue = Record.Field(r, ItemsTipoColName),
                    umValue = Record.Field(r, ItemsUMColName),
                    codText = if codValue = null then "" else Text.Trim(Text.From(codValue)),
                    descText = if descValue = null then "" else Text.Trim(Text.From(descValue)),
                    tipoText = if tipoValue = null then "" else Text.Trim(Text.From(tipoValue)),
                    umText = if umValue = null then "" else Text.Trim(Text.From(umValue)),
                    codUpper = Text.Upper(codText),
                    descUpper = Text.Upper(descText),
                    tryNum = try Number.FromText(codText),
                    isNumeric = not tryNum[HasError],
                    numValue = if isNumeric then tryNum[Value] else 0,
                    tipoFila =
                        if codText = "" then "Otro"
                        else if Text.StartsWith(codUpper, "SUBCAP") or Text.StartsWith(descUpper, "SUBCAP") then "SubCapitulo"
                        else if Text.Contains(codUpper, "CAPITULO") or Text.Contains(descUpper, "CAPITULO") then "Capitulo"
                        else if isNumeric and tipoText = "" and umText = "" and (Text.Length(codText) <= 2 or (numValue >= 1000 and Number.Mod(numValue, 1000) = 0)) then "Capitulo"
                        else if isNumeric and tipoText = "" and umText = "" then "Actividad"
                        else if isNumeric then "Insumo"
                        else "Otro"
                in tipoFila, type text),

            ItemsWithCodActRaw = Table.AddColumn(ItemsWithTipoFila, "CodigoActRaw", (r as record) =>
                let tipo = Record.Field(r, "TipoFila") in if tipo = "Actividad" then Text.From(Record.Field(r, ItemsCodColName)) else null, type text),
            ItemsCodActRawFillDown = Table.FillDown(ItemsWithCodActRaw, {"CodigoActRaw"}),
            ItemsWithCodigoAct = Table.AddColumn(ItemsCodActRawFillDown, "Codigo act", each FnFormatCodigoAct([CodigoActRaw]), type text),
            ItemsSoloInsumos = Table.SelectRows(ItemsWithCodigoAct, each [TipoFila] = "Insumo"),

            ItemsWithCodigoIns = Table.AddColumn(ItemsSoloInsumos, "Codigo ins", each Text.From(Record.Field(_, ItemsCodColName)), type text),
            ItemsWithIns = Table.AddColumn(ItemsWithCodigoIns, "Ins", (r as record) =>
                let
                    descIns = Record.Field(r, ItemsDescColName),
                    umIns = Record.Field(r, ItemsUMColName),
                    dTxt0 = if descIns = null then "" else Text.Trim(Text.From(descIns)),
                    umTxt = if umIns = null then "" else Text.Trim(Text.From(umIns)),
                    // la descripcion puede traer ya un "(UM)" propio; se quita antes de poner la oficial, para no duplicar
                    dTxtSinParen = FnSepararUM(dTxt0)[Nombre],
                    baseTxt = if umTxt = "" then dTxtSinParen else dTxtSinParen & " (" & umTxt & ")"
                in baseTxt, type text),

            OrigenAPU_Raw = try Excel.Workbook(BinarioPresupuesto, null, true){0}[Data]
                            otherwise Html.Table(Text.FromBinary(BinarioPresupuesto, 1252), Columnas_APU, [RowSelector="tr"]),
            OrigenAPU_Cols = Table.SelectColumns(OrigenAPU_Raw, List.FirstN(Table.ColumnNames(OrigenAPU_Raw), 3)),
            OrigenAPU = Table.RenameColumns(OrigenAPU_Cols, List.Zip({Table.ColumnNames(OrigenAPU_Cols), {"Columna 1", "Columna 2", "Columna 3"}})),

            APU_Paso1 = Table.AddColumn(OrigenAPU, "Cod_Temp", each
                let
                    c1Value = if [#"Columna 1"] = null then "" else [#"Columna 1"],
                    c1 = Text.Trim(Text.From(c1Value)),
                    hasDash = Text.Contains(c1, "-"),
                    preDash = if hasDash then Text.Trim(Text.BeforeDelimiter(c1, "-")) else "",
                    esNum = try Number.FromText(preDash) otherwise null
                in if hasDash and esNum <> null then FnFormatCodigoAct(preDash) else null),
            APU_Paso2 = Table.SelectRows(APU_Paso1, each [Cod_Temp] <> null),
            APU_Diccionario = Table.AddColumn(APU_Paso2, "NombreActAPU", each
                let
                    c1Value = if [#"Columna 1"] = null then "" else [#"Columna 1"],
                    rawName = Text.AfterDelimiter(Text.From(c1Value), "-"),
                    cleanName = Text.Trim(Text.Replace(Text.Replace(Text.Replace(rawName, "#(lf)", " "), "#(cr)", " "), "#(00A0)", " "))
                in cleanName, type text),
            APU_DiccionarioLimpio = Table.SelectColumns(APU_Diccionario, {"Cod_Temp", "NombreActAPU", "Columna 3"}, MissingField.Ignore),
            APU_DiccionarioRenombrado = Table.RenameColumns(APU_DiccionarioLimpio,
                List.Select({{"Cod_Temp", "CodigoActAPU"}, {"Columna 3", "UM_Actividad"}}, each Table.HasColumns(APU_DiccionarioLimpio, _{0}))),
            DiccionarioAPU_Unico = Table.Buffer(Table.Distinct(APU_DiccionarioRenombrado, {"CodigoActAPU"})),

            ItemsJoinAPU = Table.NestedJoin(ItemsWithIns, {"Codigo act"}, DiccionarioAPU_Unico, {"CodigoActAPU"}, "APU", JoinKind.LeftOuter),
            ItemsExpandedAPU = Table.ExpandTableColumn(ItemsJoinAPU, "APU", {"NombreActAPU", "UM_Actividad"}, {"NombreActAPU", "UM_Actividad"}),

            ItemsWithActividad = Table.AddColumn(ItemsExpandedAPU, "Actividad", each
                let
                    codTxt = if [Codigo act] = null then "" else [Codigo act],
                    nombreExtraido = Text.Trim(Text.From(if [NombreActAPU] = null then "" else [NombreActAPU])),
                    nombreReal = if nombreExtraido = "" then "Actividad " & codTxt else nombreExtraido,
                    umTxt = Text.Trim(Text.From(if [UM_Actividad] = null then "" else [UM_Actividad])),
                    nombreLimpio = Text.Combine(List.Select(Text.Split(nombreReal, " "), each _ <> ""), " "),
                    // el nombre crudo puede traer ya un "(UM)" propio; se quita antes de poner el oficial, para no duplicar
                    nombreSinParen = FnSepararUM(nombreLimpio)[Nombre],
                    actTxt = if umTxt = "" then codTxt & "-" & nombreSinParen
                             else codTxt & "-" & nombreSinParen & " (" & umTxt & ")"
                in actTxt, type text),

            Final = Table.SelectColumns(ItemsWithActividad, {"Codigo ins", "Ins", "Codigo act", "Actividad"})
        in Final,

    FnMaestroCC = (cc as text) as table =>
        let
            sub = Table.SelectRows(ArchivosActual, each [Centro de Costos] = cc),
            binSeg = FnPickLatest(sub, "SEGUIMIENTO POR ITEMS"),
            binApu = FnPickLatest(sub, "ANALISIS DE PRECIOS UNITARIOS"),
            datos = if binSeg = null or binApu = null then #table({"Codigo ins","Ins","Codigo act","Actividad"}, {}) else FxProcesarCentroCosto(binSeg, binApu),
            conCC = Table.AddColumn(datos, "Centro de Costos", each cc, type text)
        in conCC,

    Maestro = Table.Buffer(if List.Count(Centros) = 0 then #table({"Codigo ins","Ins","Codigo act","Actividad","Centro de Costos"},{}) else Table.Combine(List.Transform(Centros, each FnMaestroCC(_)))),
    MaestroJerarquia = Table.Buffer(Table.Group(Maestro, {"Centro de Costos", "Codigo act"}, {{"Ref.Act", each List.First(List.RemoveNulls([Actividad])), type text}})),
    MaestroInsumos = Table.AddColumn(Maestro, "InsClave_Cruce", each FnClaveLimpia([Ins]), type text),
    MaestroInsumosDist = Table.Buffer(Table.Group(MaestroInsumos, {"Centro de Costos", "Codigo act", "InsClave_Cruce"}, {{"Ref.InsOficial", each List.First([Ins]), type text}})),

    // ===== 5. Parsear "ESTADO DE CONTRATOS" por Centro de Costos =====
    FxParsearContratos = (binario as binary) as table =>
        let
            Source_Raw = try Excel.Workbook(binario, null, true){0}[Data]
                     otherwise Html.Table(Text.FromBinary(binario, 1252), Columnas_HTML25, [RowSelector="tr"]),
            Source_ColNames = Table.ColumnNames(Source_Raw),
            Source = Table.RenameColumns(Source_Raw, List.Zip({Source_ColNames, List.Transform({1..List.Count(Source_ColNames)}, each "Columna" & Text.From(_))})),

            AddFilaTexto = Table.AddColumn(Source, "FilaTexto", each let vals = Record.FieldValues(_), soloTexto = List.Transform(List.Select(vals, each _ <> null and _ <> ""), Text.From) in Text.Trim(Text.Combine(soloTexto, " ")), type text),
            AddOC = Table.AddColumn(AddFilaTexto, "Cod Contrato", each let txt = [FilaTexto] in if txt <> null and Text.Contains(Text.Upper(txt), "CONTRATO NO") then let after = Text.TrimStart(Text.Replace(Text.Range(txt, Text.PositionOf(Text.Upper(txt), "CONTRATO NO") + 11), "#(00A0)", " "), {".", ":", " "}), first = Text.BeforeDelimiter(after, " "), num = Text.Select(if first = "" then after else first, {"0".."9"}) in if num = "" then null else num else null, type text),
            AddDesc = Table.AddColumn(AddOC, "Descripcion contrato", each let txt = [FilaTexto] in if txt <> null and Text.Contains(Text.Upper(txt), "CONTRATO NO") then let after = Text.TrimStart(Text.Range(txt, Text.PositionOf(Text.Upper(txt), "CONTRATO NO") + 11), {".", ":", " "}), idx = Text.PositionOfAny(after, {"A".."Z","a".."z"}), desc = if idx = -1 then null else Text.Range(after, idx), lim = if desc = null then null else if Text.Contains(Text.Upper(desc), "CONTRATISTA") then Text.BeforeDelimiter(Text.Upper(desc), "CONTRATISTA") else desc in if lim = null then null else Text.Trim(lim) else null, type text),
            AddNombre = Table.AddColumn(AddDesc, "Nombre Contratista", each let txt = [FilaTexto] in if txt <> null and Text.Contains(Text.Upper(txt), "CONTRATISTA") then Text.Trim(Text.TrimStart(Text.AfterDelimiter(Text.Upper(txt), "CONTRATISTA"), {":","-"," "})) else null, type text),
            FillDown1 = Table.FillDown(AddNombre, {"Cod Contrato","Descripcion contrato","Nombre Contratista"}),

            AddCodAct = Table.AddColumn(FillDown1, "CodigoAct", each let c = [Columna1], t = if c = null then null else Text.Trim(Text.From(c)) in if t <> null and t <> "" and (try Number.From(Text.Replace(t, ".", "")) otherwise null) <> null then FnFormatCodigoAct(t) else null, type text),
            FillDown2 = Table.FillDown(AddCodAct, {"CodigoAct"}),

            AddCantC = Table.AddColumn(FillDown2, "Cantidad Contratado", each FxToNumberFlex([Columna4]), type number),
            AddVTC = Table.AddColumn(AddCantC, "VT Contratado", each FxToNumberFlex([Columna5]), type number),

            Filtered = Table.SelectRows(AddVTC, each
                [Columna2] <> null and
                [CodigoAct] <> null and
                ([Columna1] = null or Text.Trim(Text.From([Columna1])) = "") and
                not Text.Contains(Text.Upper(Text.From(if [Columna1] = null then "" else [Columna1])), "TOTAL") and
                not Text.Contains(Text.Upper(Text.From(if [Columna2] = null then "" else [Columna2])), "TOTAL")
            ),
            AddClave = Table.AddColumn(Filtered, "InsClave_Cruce", each FnClaveLimpia([Columna2]), type text),

            Seleccionado = Table.SelectColumns(AddClave, {"Cod Contrato","Descripcion contrato","Nombre Contratista","CodigoAct","Columna2","InsClave_Cruce","Cantidad Contratado","VT Contratado"})
        in Seleccionado,

    FnContratosCC = (cc as text) as table =>
        let
            sub = Table.SelectRows(ArchivosActual, each [Centro de Costos] = cc),
            bin = FnPickLatest(sub, "ESTADO DE CONTRATOS"),
            datos = if bin = null then #table({"Cod Contrato","Descripcion contrato","Nombre Contratista","CodigoAct","Columna2","InsClave_Cruce","Cantidad Contratado","VT Contratado"}, {}) else FxParsearContratos(bin),
            conCC = Table.AddColumn(datos, "Centro de Costos", each cc, type text)
        in conCC,

    Contratos = Table.Buffer(if List.Count(Centros) = 0 then #table({"Cod Contrato","Descripcion contrato","Nombre Contratista","CodigoAct","Columna2","InsClave_Cruce","Cantidad Contratado","VT Contratado","Centro de Costos"},{}) else Table.Combine(List.Transform(Centros, each FnContratosCC(_)))),
    ContratosValidos = Table.SelectRows(Contratos, each [Cod Contrato] <> null and [VT Contratado] <> null and [VT Contratado] <> 0),

    // ===== 6. Enriquecer con Actividad(+UM) e Insumo oficial(+UM) =====
    ConJerarquia = Table.NestedJoin(ContratosValidos, {"Centro de Costos", "CodigoAct"}, MaestroJerarquia, {"Centro de Costos", "Codigo act"}, "JER", JoinKind.LeftOuter),
    ExpJerarquia = Table.ExpandTableColumn(ConJerarquia, "JER", {"Ref.Act"}, {"Ref.Act"}),

    ConInsumo = Table.NestedJoin(ExpJerarquia, {"Centro de Costos", "CodigoAct", "InsClave_Cruce"}, MaestroInsumosDist, {"Centro de Costos", "Codigo act", "InsClave_Cruce"}, "INS", JoinKind.LeftOuter),
    ExpInsumo = Table.ExpandTableColumn(ConInsumo, "INS", {"Ref.InsOficial"}, {"Ref.InsOficial"}),

    AddActividadTxt = Table.AddColumn(ExpInsumo, "ActividadTxt", each if [Ref.Act] <> null then [Ref.Act] else [CodigoAct], type text),
    AddInsumoTxt = Table.AddColumn(AddActividadTxt, "InsumoTxt", each if [Ref.InsOficial] <> null then [Ref.InsOficial] else (if [Columna2] = null or Text.Trim([Columna2]) = "" then "SIN DESCRIPCION" else Text.Trim([Columna2])), type text),

    // ===== 7. Separar nombre/UM y armar columnas finales =====
    AddSepGrupo = Table.AddColumn(AddInsumoTxt, "_SepGrupo", each FnSepararUM([ActividadTxt]), type record),
    AddSepInsumo = Table.AddColumn(AddSepGrupo, "_SepInsumo", each FnSepararUM([InsumoTxt]), type record),

    AddFinal = Table.AddColumn(AddSepInsumo, "_Final", each [
        Grupo = FnQuitarTildes([_SepGrupo][Nombre]),
        UMGrupo = FnQuitarTildes([_SepGrupo][UM]),
        CantidadGrupo = [Cantidad Contratado],
        Insumo = FnQuitarTildes([_SepInsumo][Nombre]),
        UMInsumo = FnQuitarTildes([_SepInsumo][UM]),
        ValorUnitarioInsumo = if [Cantidad Contratado] <> null and [Cantidad Contratado] <> 0 then [VT Contratado] / [Cantidad Contratado] else [VT Contratado],
        Descripcion = FnQuitarTildes(Text.Combine(List.Select({Text.From([Descripcion contrato]), Text.From([Nombre Contratista]), Text.From([Cod Contrato])}, each _ <> null and Text.Trim(_) <> ""), " / "))
    ]),
    ExpFinal = Table.ExpandRecordColumn(AddFinal, "_Final", {"Grupo","UMGrupo","CantidadGrupo","Insumo","UMInsumo","ValorUnitarioInsumo","Descripcion"}),

    // ===== 8. Delta por fecha (Versiones previas) — vacio = trae todo =====
    AddKey = Table.AddColumn(ExpFinal, "__Key", each [Centro de Costos] & "|" & (let d = FnDigits([Cod Contrato]) in if d = null then "" else d), type text),

    FnKeysPrevCC = (cc as text) as table =>
        let
            path = BasePath & "/" & cc & "/Versiones previas/" & FechaVersion,
            raw = try Json.Document(Web.Contents(SiteUrl, [
                RelativePath = "/_api/web/GetFolderByServerRelativeUrl('" & FnEncode(path) & "')/Files",
                Query = [#"$select" = "Name,ServerRelativeUrl,TimeLastModified"],
                Headers = Headers,
                Timeout = #duration(0,0,5,0)
            ])) otherwise null,
            tbl = if raw <> null and Record.HasFields(raw, "value") then Table.FromRecords(raw[value]) else #table({"Name","ServerRelativeUrl","TimeLastModified"}, {}),
            bin = FnPickLatest(tbl, "ESTADO DE CONTRATOS"),
            tbl0 = if bin = null then #table({"FilaTexto"}, {}) else
                let
                    Source_Raw = try Excel.Workbook(bin, null, true){0}[Data] otherwise Html.Table(Text.FromBinary(bin, 1252), Columnas_HTML25, [RowSelector="tr"]),
                    withText = Table.AddColumn(Source_Raw, "FilaTexto", each Text.Combine(List.Transform(List.Select(Record.FieldValues(_), each _ <> null and _ <> ""), Text.From), " "), type text)
                in Table.SelectColumns(withText, {"FilaTexto"}),
            rows = Table.AddColumn(tbl0, "Contrato", each if Text.Contains(Text.Upper([FilaTexto]), "CONTRATO NO") then let after = Text.Range([FilaTexto], Text.PositionOf(Text.Upper([FilaTexto]), "CONTRATO NO") + 11) in FnDigits(Text.BeforeDelimiter(Text.TrimStart(after, {".",":"," "}), " ")) else null, type text),
            keys = Table.SelectRows(Table.AddColumn(rows, "__Key", each cc & "|" & [Contrato], type text), each [Contrato] <> null),
            out = Table.SelectColumns(keys, {"__Key"})
        in out,

    PrevKeys = Table.Distinct(if FechaVersion = "" or List.Count(Centros) = 0 then #table({"__Key"}, {}) else Table.Combine(List.Transform(Centros, each FnKeysPrevCC(_)))),
    JoinPrev = Table.NestedJoin(AddKey, {"__Key"}, PrevKeys, {"__Key"}, "Prev", JoinKind.LeftAnti),
    SinClave = Table.RemoveColumns(JoinPrev, {"__Key"}, MissingField.Ignore),

    // ===== 9. Columnas finales de la tabla CONTRATOS =====
    AddEC = Table.AddColumn(SinClave, "Estado Contrato", each "PENDIENTE", type text),
    AddEG = Table.AddColumn(AddEC, "Estado Grupo", each "PENDIENTE", type text),
    AddED = Table.AddColumn(AddEG, "Estado Detalle", each "PENDIENTE", type text),
    AddFH = Table.AddColumn(AddED, "Fecha Hora", each null, type nullable datetime),
    AddErr = Table.AddColumn(AddFH, "Error", each null, type nullable text),

    Resultado = Table.SelectColumns(AddErr, {
        "Cod Contrato","Descripcion","Grupo","UMGrupo","CantidadGrupo","Insumo","UMInsumo","ValorUnitarioInsumo",
        "Estado Contrato","Estado Grupo","Estado Detalle","Fecha Hora","Error"
    }),
    Renombrado = Table.RenameColumns(Resultado, {
        {"UMGrupo", "UM Grupo"}, {"CantidadGrupo", "Cantidad Grupo"}, {"UMInsumo", "UM Insumo"}, {"ValorUnitarioInsumo", "Valor Unitario Insumo"}
    })
in
    Renombrado
