// BD_UNIFICADO
// Tabla maestra unica (Centro de Costos / Paquete de Trabajo, Codigo act, Codigo ins) que sirve
// indistintamente para proyectos de origen SINCO u ORACLE, segun el parametro "Origen".
//
// Rama SINCO: lee la BD ya calculada en SINCO.xlsx (Cantidad/VT Presupuesto, Proyectado, Consumido,
//             Contratado, Comprado, Asegurado ya vienen separados ahi).
// Rama ORACLE: la BD publicada de ORACLE.xlsx solo trae "Asegurado" (Contratado+Comprado mezclados),
//             asi que aqui se vuelve a leer SP_Fuentes (CONTRATOS.xls / COMPRAS.xls / ASEGURADO.xls)
//             para separar Contratado de Comprado, y se reutiliza Cantidad/VT PPTO V1/V2 y
//             Cantidad corte/Valor Recepcion corte ya calculados en la BD de ORACLE.xlsx para
//             Presupuesto/Proyectado/Consumido.
//
// Parametros esperados en el libro que evalua esta consulta: Origen (texto "SINCO"|"ORACLE"),
// RutaSINCO (ruta local a SINCO.xlsx), RutaORACLE (ruta local a ORACLE.xlsx), ProyectoActual.

let
    EsOracle = Text.Upper(Text.Trim(Text.From(Origen))) = "ORACLE",

    // ---------------------------------------------------------------
    // HELPERS (portados de PowerQuery-Gestion-Costos-Oracle/Consultas/BD.m,
    // ya probados contra los reportes reales CONTRATOS.xls / COMPRAS.xls / ASEGURADO.xls)
    // ---------------------------------------------------------------
    CleanKey = (val as any) as text => if val = null then "" else Text.Upper(Text.Clean(Text.Trim(Text.From(val)))),
    ToNumberSafe = (val as any) as number => let n = try Number.From(val) otherwise 0 in if n = null then 0 else n,

    FixHeaders = (table as table) as table =>
        let
            Cols = Table.ColumnNames(table),
            NeedsPromotion = List.Contains(Cols, "Column1"),
            Rows = Table.ToRows(table),
            HeaderFlags = List.Transform(Rows, each
                let Values = List.Transform(_, (v) => Text.Upper(Text.Trim(if v = null then "" else Text.From(v))))
                in List.Contains(Values, "REGISTRO") or List.Contains(Values, "COD CBS") or (List.Contains(Values, "ORDEN") and List.Contains(Values, "CBS"))
            ),
            HeaderIndex = List.PositionOf(HeaderFlags, true),
            Skipped = if NeedsPromotion and HeaderIndex >= 0 then Table.Skip(table, HeaderIndex) else table,
            Promoted = if NeedsPromotion then Table.PromoteHeaders(Skipped, [PromoteAllScalars = true]) else Skipped,
            CleanNames = Table.TransformColumnNames(Promoted, Text.Trim)
        in
            CleanNames,

    NormalizeHeader = (name as any) as text =>
        let
            txt = Text.Upper(Text.Trim(if name = null then "" else Text.From(name))),
            repl = {{"Á","A"},{"É","E"},{"Í","I"},{"Ó","O"},{"Ú","U"},{"Ñ","N"}},
            clean = List.Accumulate(repl, txt, (state, current) => Text.Replace(state, current{0}, current{1}))
        in
            Text.Select(clean, {"A".."Z", "0".."9"}),

    GetColumn = (table as table, candidates as list, position as number) as list =>
        let
            cols = Table.ColumnNames(table),
            normalizedCandidates = List.Transform(candidates, each NormalizeHeader(_)),
            match = List.First(List.Select(cols, each List.Contains(normalizedCandidates, NormalizeHeader(_))), null),
            result =
                if match <> null then Table.Column(table, match)
                else if List.Count(cols) > position then Table.Column(table, cols{position})
                else List.Repeat({null}, Table.RowCount(table))
        in
            result,

    // tipo = "CONTRATO" -> usa columnas de CONTRATOS.xls (Cantidad/Valor Orden = lo contratado)
    // tipo = "OC"        -> usa columnas de COMPRAS.xls (Cantidad/Valor Recepcion = lo comprado/recibido)
    NormalizeMovimientos = (table as table, tipo as text) as table =>
        let
            CantidadPrincipal = if tipo = "OC"
                then GetColumn(table, {"Cantidad recepcion", "Cantidad Recepcion", "Recepciones Cantidad", "Cantidad_1", "Cantidad.1"}, 16)
                else GetColumn(table, {"Cantidad corte", "Cantidad", "Cantidad Orden", "Orden de compra Cantidad"}, 13),
            ValorPrincipal = if tipo = "OC"
                then GetColumn(table, {"Valor Recepcion", "Recepciones Valor Recepcion", "Valor Recepcion_1", "Valor Recepcion.1"}, 17)
                else GetColumn(table, {"Valor Recepcion corte", "Valor Recepcion", "Valor Orden", "Orden de compra Valor Orden"}, 15),
            NombreCantidad = if tipo = "OC" then "Cantidad Comprado" else "Cantidad Contratado",
            NombreValor = if tipo = "OC" then "VT Comprado" else "VT Contratado"
        in
            Table.FromColumns({
                GetColumn(table, {"Orden"}, 0),
                GetColumn(table, {"Razon Social", "Razón Social"}, 5),
                GetColumn(table, {"CBS"}, 7),
                GetColumn(table, {"Descripcion", "Descripción"}, 8),
                GetColumn(table, {"Paquete de trabajo", "Paquete de Trabajo"}, 9),
                GetColumn(table, {"Articulo", "Artículo"}, 10),
                GetColumn(table, {"Descripcion_1", "Descripción_1", "Descripcion.1", "Descripción.1"}, 11),
                CantidadPrincipal,
                ValorPrincipal
            }, {"# OC / Contrato", "Nombre Contratista", "Codigo act", "Actividad", "Centro de Costos", "Codigo ins", "Ins", NombreCantidad, NombreValor}),

    NormalizeAsegurado = (table as table) as table =>
        Table.FromColumns({
            GetColumn(table, {"Paquete de Trabajo"}, 2),
            GetColumn(table, {"Cod CBS", "CBS"}, 3),
            GetColumn(table, {"Articulo", "Artículo"}, 5),
            GetColumn(table, {"U Medida", "UM"}, 7),
            GetColumn(table, {"Proceso"}, 8),
            GetColumn(table, {"Registro"}, 9),
            GetColumn(table, {"Cantidad"}, 11),
            GetColumn(table, {"V.r Total", "Vr Total", "Valor Total"}, 13)
        }, {"Centro de Costos", "Codigo act", "Codigo ins", "UM", "Proceso", "# OC / Contrato", "Cantidad Asegurado", "VT Asegurado"}),

    // ---------------------------------------------------------------
    // RAMA ORACLE
    // ---------------------------------------------------------------
    BD_Oracle =
        let
            UrlOracleRepo = "https://raw.githubusercontent.com/jdbustosp/PowerQuery-Gestion-Costos-Oracle/main/Consultas/",
            FxOracle = (nombre as text) as any => Expression.Evaluate(Text.FromBinary(Web.Contents(UrlOracleRepo & nombre & ".m")), #shared),
            SP_Fuentes = FxOracle("SP_Fuentes"),

            T_Contrato = NormalizeMovimientos(FixHeaders(SP_Fuentes[CONTRATOS]), "CONTRATO"),
            T_OC = NormalizeMovimientos(FixHeaders(SP_Fuentes[COMPRAS]), "OC"),
            T_Aseg = NormalizeAsegurado(FixHeaders(SP_Fuentes[ASEGURADO])),

            G_Contrato = Table.Group(
                Table.TransformColumns(T_Contrato, {{"Centro de Costos", CleanKey}, {"Codigo act", CleanKey}, {"Codigo ins", CleanKey}}),
                {"Centro de Costos", "Codigo act", "Codigo ins"},
                {
                    {"Cantidad Contratado", each List.Sum(List.Transform([Cantidad Contratado], ToNumberSafe)), type number},
                    {"VT Contratado", each List.Sum(List.Transform([VT Contratado], ToNumberSafe)), type number},
                    {"Nombre Contratista_Ct", each List.First(List.RemoveNulls([Nombre Contratista])), type text},
                    {"# OC / Contrato_Ct", each List.First(List.RemoveNulls([# OC / Contrato])), type text},
                    {"Actividad_Ct", each List.First(List.RemoveNulls([Actividad])), type text},
                    {"Ins_Ct", each List.First(List.RemoveNulls([Ins])), type text}
                }
            ),
            G_OC = Table.Group(
                Table.TransformColumns(T_OC, {{"Centro de Costos", CleanKey}, {"Codigo act", CleanKey}, {"Codigo ins", CleanKey}}),
                {"Centro de Costos", "Codigo act", "Codigo ins"},
                {
                    {"Cantidad Comprado", each List.Sum(List.Transform([Cantidad Comprado], ToNumberSafe)), type number},
                    {"VT Comprado", each List.Sum(List.Transform([VT Comprado], ToNumberSafe)), type number},
                    {"Nombre Contratista_OC", each List.First(List.RemoveNulls([Nombre Contratista])), type text},
                    {"# OC / Contrato_OC", each List.First(List.RemoveNulls([# OC / Contrato])), type text},
                    {"Actividad_OC", each List.First(List.RemoveNulls([Actividad])), type text},
                    {"Ins_OC", each List.First(List.RemoveNulls([Ins])), type text}
                }
            ),
            G_Aseg = Table.Group(
                Table.SelectRows(
                    Table.TransformColumns(T_Aseg, {{"Centro de Costos", CleanKey}, {"Codigo act", CleanKey}, {"Codigo ins", CleanKey}}),
                    each [Proceso] <> "COSTOS DISTRIBUIBLES" and [Proceso] <> "TRANSFERENCIA"
                ),
                {"Centro de Costos", "Codigo act", "Codigo ins"},
                {
                    {"Cantidad Asegurado", each List.Sum(List.Transform([Cantidad Asegurado], ToNumberSafe)), type number},
                    {"VT Asegurado", each List.Sum(List.Transform([VT Asegurado], ToNumberSafe)), type number},
                    {"UM", each List.First(List.RemoveNulls([UM])), type text}
                }
            ),

            // Presupuesto/Proyectado/Consumido se reutilizan de la BD ya calculada en ORACLE.xlsx
            // (Cantidad/VT PPTO V1 = Presupuesto, PPTO V2 = Proyectado, Cantidad corte/Valor Recepcion corte = Consumido)
            BD_Oracle_Raw = Table.PromoteHeaders(Excel.Workbook(File.Contents(RutaORACLE), null, true){[Item = "BD", Kind = "Sheet"]}[Data], [PromoteAllScalars = true]),
            Ppto_Oracle = Table.Group(
                Table.TransformColumns(
                    Table.SelectColumns(BD_Oracle_Raw, {"Paquete de Trabajo", "Cod actividad", "Cod ins", "Cantidad PPTO V1", "VT PPTO V1", "Cantidad PPTO V2", "VT PPTO V2", "Cantidad corte", "Valor Recepcion corte", "Actividad", "Ins"}, MissingField.UseNull),
                    {{"Paquete de Trabajo", CleanKey}, {"Cod actividad", CleanKey}, {"Cod ins", CleanKey}}
                ),
                {"Paquete de Trabajo", "Cod actividad", "Cod ins"},
                {
                    {"Cantidad Presupuesto", each List.Sum(List.Transform([Cantidad PPTO V1], ToNumberSafe)), type number},
                    {"VT Presupuesto", each List.Sum(List.Transform([VT PPTO V1], ToNumberSafe)), type number},
                    {"Cantidad Proyectado", each List.Sum(List.Transform([Cantidad PPTO V2], ToNumberSafe)), type number},
                    {"VT Proyectado", each List.Sum(List.Transform([VT PPTO V2], ToNumberSafe)), type number},
                    {"Cantidad Consumido", each List.Sum(List.Transform([Cantidad corte], ToNumberSafe)), type number},
                    {"VT Consumido", each List.Sum(List.Transform([Valor Recepcion corte], ToNumberSafe)), type number},
                    {"Actividad_P", each List.First(List.RemoveNulls([Actividad])), type text},
                    {"Ins_P", each List.First(List.RemoveNulls([Ins])), type text}
                }
            ),
            Ppto_Oracle_Renamed = Table.RenameColumns(Ppto_Oracle, {{"Paquete de Trabajo", "Centro de Costos"}, {"Cod actividad", "Codigo act"}, {"Cod ins", "Codigo ins"}}),

            Join1 = Table.NestedJoin(Ppto_Oracle_Renamed, {"Centro de Costos", "Codigo act", "Codigo ins"}, G_Contrato, {"Centro de Costos", "Codigo act", "Codigo ins"}, "J1", JoinKind.FullOuter),
            Exp1 = Table.ExpandTableColumn(Join1, "J1", {"Cantidad Contratado", "VT Contratado", "Nombre Contratista_Ct", "# OC / Contrato_Ct", "Actividad_Ct", "Ins_Ct"}),
            Join2 = Table.NestedJoin(Exp1, {"Centro de Costos", "Codigo act", "Codigo ins"}, G_OC, {"Centro de Costos", "Codigo act", "Codigo ins"}, "J2", JoinKind.FullOuter),
            Exp2 = Table.ExpandTableColumn(Join2, "J2", {"Cantidad Comprado", "VT Comprado", "Nombre Contratista_OC", "# OC / Contrato_OC", "Actividad_OC", "Ins_OC"}),
            Join3 = Table.NestedJoin(Exp2, {"Centro de Costos", "Codigo act", "Codigo ins"}, G_Aseg, {"Centro de Costos", "Codigo act", "Codigo ins"}, "J3", JoinKind.FullOuter),
            Exp3 = Table.ExpandTableColumn(Join3, "J3", {"Cantidad Asegurado", "VT Asegurado", "UM"}),

            ConNombres = Table.AddColumn(Exp3, "Actividad", each if [Actividad_P] <> null then [Actividad_P] else if [Actividad_Ct] <> null then [Actividad_Ct] else [Actividad_OC], type text),
            ConIns = Table.AddColumn(ConNombres, "Ins", each if [Ins_P] <> null then [Ins_P] else if [Ins_Ct] <> null then [Ins_Ct] else [Ins_OC], type text),
            ConContratista = Table.AddColumn(ConIns, "Nombre Contratista", each if [Nombre Contratista_Ct] <> null then [Nombre Contratista_Ct] else [Nombre Contratista_OC], type text),
            ConOC = Table.AddColumn(ConContratista, "# OC / Contrato", each if [# OC / Contrato_Ct] <> null then [# OC / Contrato_Ct] else [# OC / Contrato_OC], type text),
            ConTipo = Table.AddColumn(ConOC, "Tipo", each "ORACLE", type text),

            Limpio = Table.RemoveColumns(ConTipo, {"Actividad_P", "Ins_P", "Actividad_Ct", "Ins_Ct", "Actividad_OC", "Ins_OC", "Nombre Contratista_Ct", "Nombre Contratista_OC", "# OC / Contrato_Ct", "# OC / Contrato_OC"}),

            Final = Table.SelectColumns(Limpio, {
                "Centro de Costos", "Codigo act", "Actividad", "Codigo ins", "Ins", "UM", "# OC / Contrato", "Nombre Contratista", "Tipo",
                "Cantidad Presupuesto", "VT Presupuesto", "Cantidad Proyectado", "VT Proyectado", "Cantidad Consumido", "VT Consumido",
                "Cantidad Contratado", "VT Contratado", "Cantidad Comprado", "VT Comprado", "Cantidad Asegurado", "VT Asegurado"
            }, MissingField.UseNull)
        in
            Final,

    // ---------------------------------------------------------------
    // RAMA SINCO (la BD de SINCO.xlsx ya trae todo separado)
    // ---------------------------------------------------------------
    BD_SINCO =
        let
            Raw = Table.PromoteHeaders(Excel.Workbook(File.Contents(RutaSINCO), null, true){[Item = "BD", Kind = "Sheet"]}[Data], [PromoteAllScalars = true]),
            ConUM = Table.AddColumn(Raw, "UM", each null, type text),
            Seleccionado = Table.SelectColumns(ConUM, {
                "Centro de Costos", "Codigo act", "Actividad", "Codigo ins", "Ins", "UM", "# OC / Contrato", "Nombre Contratista", "Tipo",
                "Cantidad Presupuesto", "VT Presupuesto", "Cantidad Proyectado", "VT Proyectado", "Cantidad Consumido", "VT Consumido",
                "Cantidad Contratado", "VT Contratado", "Cantidad Comprado", "VT Comprado", "Cantidad asegurada", "VT Asegurada"
            }, MissingField.UseNull),
            Renombrado = Table.RenameColumns(Seleccionado, {{"Cantidad asegurada", "Cantidad Asegurado"}, {"VT Asegurada", "VT Asegurado"}}, MissingField.Ignore)
        in
            Renombrado,

    Resultado = if EsOracle then BD_Oracle else BD_SINCO
in
    Resultado
