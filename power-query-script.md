# Script de Power Query (M) — Horas extras (Admin + Proyectos)

```powerquery
let
    // ==========================================================
    // Q_HORAS_EXTRAS_TIPIFICADO (ADMIN + PROYECTOS)
    // - Empleado unificado: EMPLEADO_ALL
    // - Nocturna: 19:00-05:59 / Diurna: 06:00-18:59
    // - Horas enteras (sin fracciones): RoundAwayFromZero
    // ==========================================================

    BaseRaw = Q_HORAS_EXTRAS,
    Tip = #"TIPIFICACION SIIGO",

    // -------- Helpers de columnas --------
    fxNormCol = (t as text) as text =>
        let
            t0 = Text.Replace(t, Character.FromNumber(160), " "),
            parts = List.Select(Text.SplitAny(t0, " #(tab)#(cr)#(lf)"), each _ <> ""),
            t1 = Text.Combine(parts, " "),
            t2 = Text.Trim(t1),
            t3 = Text.Combine(List.Select(Text.Split(t2, " "), each _ <> ""), " "),
            t4 = Text.Upper(t3)
        in t4,

    Cols = Table.ColumnNames(BaseRaw),

    FindColExact = (want as text) as nullable text =>
        let target = fxNormCol(want),
            hits = List.Select(Cols, each fxNormCol(_) = target)
        in if List.Count(hits)>0 then hits{0} else null,

    FindFirstContainsAll = (must as list, minPos as number, maxPos as number) as nullable text =>
        let
            hits =
                List.Select(
                    Cols,
                    (c) =>
                        let
                            pos = List.PositionOf(Cols, c),
                            okPos = pos >= minPos and pos <= maxPos,
                            okTxt = List.AllTrue(List.Transform(must, (m)=> Text.Contains(fxNormCol(c), fxNormCol(m))))
                        in okPos and okTxt
                )
        in if List.Count(hits)>0 then hits{0} else null,

    fxToDate = (x as any) as nullable date =>
        try Date.From(x) otherwise try Date.FromText(Text.From(x), "es-CO") otherwise null,

    // -------- Columnas del Forms --------
    cMarca = FindColExact("Marca temporal"),
    cTipoN = FindColExact("SELECCIONE TIPO DE NOVEDAD"),
    cArea  = FindColExact("SELECCIONE AREA"),

    cEmpHE  = FindColExact("SELECCIONE EMPLEADO HE"),
    cEmpHEP = FindColExact("SELECCIONE EMPLEADO H.E.P"),

    idxHEP = if cEmpHEP=null then -1 else List.PositionOf(Cols, cEmpHEP),
    maxIdx = List.Count(Cols) - 1,
    minHE  = 0,
    maxHE  = if idxHEP<0 then maxIdx else idxHEP,
    minHEP = if idxHEP<0 then 0 else idxHEP,
    maxHEP = maxIdx,

    // Fechas
    cF1  = FindColExact("FECHA 1"),
    cF2  = FindColExact("FECHA 2"),
    cF3  = FindColExact("FECHA 3"),

    cF1P = FindColExact("FECHA 1P"),
    cF2P = FindColExact("FECHA 2P"),
    cF3P = FindColExact("FECHA 3P"),

    // ADMIN (antes)
    cD1  = FindFirstContainsAll({"DESDE","INICIO","1","HORA","MILITAR"}, minHE, maxHE),
    cH1  = FindFirstContainsAll({"HASTA","FIN","1","HORA","MILITAR"},  minHE, maxHE),

    cD2  = FindFirstContainsAll({"DESDE","INICIO","2","HORA","MILITAR"}, minHE, maxHE),
    cH2  = FindFirstContainsAll({"HASTA","FIN","2","HORA","MILITAR"},  minHE, maxHE),

    cD3  = FindFirstContainsAll({"DESDE","INICIO","3","HORA","MILITAR"}, minHE, maxHE),
    cH3  = FindFirstContainsAll({"HASTA","FIN","3","HORA","MILITAR"},  minHE, maxHE),

    // PROYECTOS (después) normalmente sin "P" en headers de Desde/Hasta
    cD1P = FindFirstContainsAll({"DESDE","INICIO","1","HORA","MILITAR"}, minHEP, maxHEP),
    cH1P = FindFirstContainsAll({"HASTA","FIN","1","HORA","MILITAR"},  minHEP, maxHEP),

    cD2P = FindFirstContainsAll({"DESDE","INICIO","2","HORA","MILITAR"}, minHEP, maxHEP),
    cH2P = FindFirstContainsAll({"HASTA","FIN","2","HORA","MILITAR"},  minHEP, maxHEP),

    cD3P = FindFirstContainsAll({"DESDE","INICIO","3","HORA","MILITAR"}, minHEP, maxHEP),
    cH3P = FindFirstContainsAll({"HASTA","FIN","3","HORA","MILITAR"},  minHEP, maxHEP),

    cD3P_F = if cD3P <> null then cD3P else cD2P,

    // -------- Renombrar esenciales --------
    Base0 =
        Table.RenameColumns(
            BaseRaw,
            List.RemoveNulls({
                if cMarca<>null then {cMarca,"Marca temporal"} else null,
                if cTipoN<>null then {cTipoN,"Tipo Novedad"} else null,
                if cArea<>null then {cArea,"Área"} else null,
                if cEmpHE<>null then {cEmpHE,"EMPLEADO_HE"} else null,
                if cEmpHEP<>null then {cEmpHEP,"EMPLEADO_HEP"} else null
            }),
            MissingField.Ignore
        ),

    Base1 = Table.RemoveColumns(Base0, {"Dirección de correo electrónico","Direccion de correo electronico","email","Email"}, MissingField.Ignore),

    // ✅ Empleado unificado (esto evita que en PROYECTOS se vea "vacío")
    Base2 = Table.AddColumn(
        Base1,
        "EMPLEADO_ALL",
        each
            let
                a = try Text.Trim(Text.From([EMPLEADO_HE])) otherwise null,
                p = try Text.Trim(Text.From([EMPLEADO_HEP])) otherwise null
            in
                if a<>null and a<>"" then a
                else if p<>null and p<>"" then p
                else null,
        type text
    ),

    // -------- Festivos --------
    FestivosTbl =
        let a = try FESTIVOS_CO_2026 otherwise null,
            b = try FESTIVOS_CO otherwise null
        in
            if a <> null and Value.Is(a, type table) and Table.HasColumns(a, "Fecha") then a
            else if b <> null and Value.Is(b, type table) and Table.HasColumns(b, "Fecha") then b
            else null,

    FestivosList =
        if FestivosTbl <> null
        then List.RemoveNulls(List.Transform(Table.Column(FestivosTbl, "Fecha"), each try Date.From(_) otherwise null))
        else {},

    // -------- Time utils (parser robusto) --------
    fxParseTime = (x as any) as nullable time =>
        let
            t =
                if x = null then null
                else if Value.Is(x, type time) then Time.From(x)
                else if Value.Is(x, type datetime) then Time.From(x)
                else if Value.Is(x, type datetimezone) then Time.From(x)
                else
                    let
                        s0 = Text.Trim(Text.From(x)),
                        sDigits = Text.Select(s0, {"0".."9"}),
                        isPureDigits = (Text.Length(sDigits) = Text.Length(Text.Select(s0, {"0".."9"}))) and Text.Length(sDigits) >= 3 and not Text.Contains(s0, ":"),
                        sFix =
                            if isPureDigits then
                                let
                                    len = Text.Length(sDigits),
                                    hh = if len=3 then Text.Start(sDigits,1) else Text.Start(sDigits,2),
                                    mm = if len=3 then Text.End(sDigits,2) else Text.Range(sDigits,2,2)
                                in hh & ":" & mm & ":00"
                            else
                                let
                                    s1 = Text.Lower(s0),
                                    s2 = Text.Replace(Text.Replace(s1, ".", ""), " ", ""),
                                    hasAm = Text.EndsWith(s2, "am"),
                                    hasPm = Text.EndsWith(s2, "pm"),
                                    core0 = if hasAm or hasPm then Text.Start(s2, Text.Length(s2)-2) else s2,
                                    parts = Text.Split(core0, ":"),
                                    core1 = if List.Count(parts)=1 then core0 & ":00:00"
                                            else if List.Count(parts)=2 then core0 & ":00"
                                            else core0,
                                    sFinal = if hasAm then core1 & " AM" else if hasPm then core1 & " PM" else core1
                                in sFinal,
                        parsed =
                            if Text.Contains(Text.Upper(sFix), "AM") or Text.Contains(Text.Upper(sFix), "PM")
                            then try Time.FromText(sFix, "en-US") otherwise null
                            else try Time.FromText(sFix, "es-CO") otherwise try Time.FromText(sFix) otherwise null
                    in parsed
        in t,

    fxTimeToMin = (x as any) as nullable number =>
        let t = fxParseTime(x)
        in if t=null then null else Time.Hour(t)*60 + Time.Minute(t) + Number.From(Time.Second(t))/60,

    fxSegments = (s as number, e as number) as list =>
        if e >= s then {{s, e}} else {{s, 1440}, {0, e}},

    fxOverlap = (a as list, b as list) as number =>
        let
            combos = List.Combine(
                List.Transform(a, (sa) =>
                    List.Transform(b, (sb) =>
                        let
                            a1 = sa{0}, a2 = sa{1},
                            b1 = sb{0}, b2 = sb{1},
                            x1 = if a1 > b1 then a1 else b1,
                            x2 = if a2 < b2 then a2 else b2,
                            ov = if x2 > x1 then (x2 - x1) else 0
                        in ov
                    )
                )
            )
        in List.Sum(combos),

    NoctSegs = {{19*60, 1440}, {0, 6*60}},

    // ==========================================================
    // BLOQUES: construimos por sección, pero SIEMPRE con EMPLEADO_ALL
    // ==========================================================
    AddBlocks = Table.AddColumn(
        Base2,
        "Blocks",
        each
            let
                mk = (bk as text, f as any, d as any, h as any) as nullable record =>
                    let
                        e  = try [EMPLEADO_ALL] otherwise null,
                        fd = fxToDate(f),
                        dd = if d=null then null else Text.Trim(Text.From(d)),
                        hh = if h=null then null else Text.Trim(Text.From(h))
                    in
                        if e=null or e="" or fd=null or dd=null or dd="" or hh=null or hh="" then null
                        else [Empleado=e, BloqueKey=bk, Fecha=fd, Desde=dd, Hasta=hh],

                // ADMIN usa FECHA 1/2/3
                b1  = mk("1",  try Record.Field(_, cF1)  otherwise null, try Record.Field(_, cD1)  otherwise null, try Record.Field(_, cH1)  otherwise null),
                b2  = mk("2",  try Record.Field(_, cF2)  otherwise null, try Record.Field(_, cD2)  otherwise null, try Record.Field(_, cH2)  otherwise null),
                b3  = mk("3",  try Record.Field(_, cF3)  otherwise null, try Record.Field(_, cD3)  otherwise null, try Record.Field(_, cH3)  otherwise null),

                // PROYECTOS usa FECHA 1P/2P/3P
                bp1 = mk("1P", try Record.Field(_, cF1P) otherwise null, try Record.Field(_, cD1P) otherwise null, try Record.Field(_, cH1P) otherwise null),
                bp2 = mk("2P", try Record.Field(_, cF2P) otherwise null, try Record.Field(_, cD2P) otherwise null, try Record.Field(_, cH2P) otherwise null),
                bp3 = mk("3P", try Record.Field(_, cF3P) otherwise null, try Record.Field(_, cD3P_F) otherwise null, try Record.Field(_, cH3P) otherwise null)
            in
                List.RemoveNulls({b1,b2,b3,bp1,bp2,bp3}),
        type list
    ),

    ExpandBlocks = Table.ExpandListColumn(AddBlocks, "Blocks"),
    ExpandRec = Table.ExpandRecordColumn(ExpandBlocks, "Blocks", {"Empleado","BloqueKey","Fecha","Desde","Hasta"}, {"Empleado","BloqueKey","Fecha","Desde","Hasta"}),

    // CC / Nombre
    AddCC2 = Table.AddColumn(
        ExpandRec,
        "CC",
        each let txt = try Text.From([Empleado]) otherwise null,
                 cc = if txt=null then null else Text.Select(txt, {"0".."9"})
             in if cc="" then null else cc,
        type text
    ),

    AddNombre2 = Table.AddColumn(
        AddCC2,
        "Nombre",
        each
            let txt = try Text.From([Empleado]) otherwise null,
                sinDig = if txt=null then null else Text.Remove(txt, {"0".."9"}),
                limpio = if sinDig=null then null else Text.Trim(Text.Replace(Text.Replace(Text.Replace(Text.Replace(sinDig,"-"," "), "(", " "), ")", " "), ".", " "))
            in if limpio="" then null else limpio,
        type text
    ),

    CleanEmp = Table.RemoveColumns(AddNombre2, {"Empleado"}, MissingField.Ignore),

    // CALC en minutos -> horas enteras
    AddCalc = Table.AddColumn(
        CleanEmp,
        "Calc",
        each
            let
                f    = [Fecha],
                sMin = fxTimeToMin([Desde]),
                eMin = fxTimeToMin([Hasta]),
                segs = if sMin=null or eMin=null then {} else fxSegments(sMin, eMin),

                totalMin =
                    if sMin=null or eMin=null then null
                    else if eMin >= sMin then (eMin - sMin) else (1440 - sMin + eMin),

                noctMin = if totalMin=null then null else fxOverlap(segs, NoctSegs),

                totalH = if totalMin=null then null else Number.RoundAwayFromZero(totalMin/60, 0),
                noctH0 = if noctMin=null then null else Number.RoundAwayFromZero(noctMin/60, 0),
                noctH  = if totalH=null or noctH0=null then noctH0 else if noctH0 > totalH then totalH else noctH0,
                diurH  = if totalH=null or noctH=null then null else totalH - noctH,

                esDom   = Date.DayOfWeek(f, Day.Monday) = 6,
                esFest  = List.Contains(FestivosList, f),
                esDomFest = (esDom or esFest)
            in
                [TotalHoras=totalH, HorasDiurnas=diurH, HorasNocturnas=noctH, EsDomingoFestivo=esDomFest],
        type record
    ),

    ExpandCalc = Table.ExpandRecordColumn(AddCalc, "Calc", {"TotalHoras","HorasDiurnas","HorasNocturnas","EsDomingoFestivo"}, {"TotalHoras","HorasDiurnas","HorasNocturnas","EsDomingoFestivo"}),

    // ALLOC SIIGO
    AddAlloc = Table.AddColumn(
        ExpandCalc,
        "Alloc",
        each
            let hd=[HorasDiurnas], hn=[HorasNocturnas], df=[EsDomingoFestivo]
            in List.RemoveNulls({
                if hd<>null and hd>0 then [CODIGO_SIIGO=(if df then 7 else 10), HORAS=hd, TRAMO=(if df then "DIURNA DOM/FEST" else "DIURNA")] else null,
                if hn<>null and hn>0 then [CODIGO_SIIGO=(if df then 12 else 11), HORAS=hn, TRAMO=(if df then "NOCTURNA DOM/FEST" else "NOCTURNA")] else null
            }),
        type list
    ),

    ExpandAllocList = Table.ExpandListColumn(AddAlloc, "Alloc"),
    ExpandAllocRec = Table.ExpandRecordColumn(ExpandAllocList, "Alloc", {"CODIGO_SIIGO","HORAS","TRAMO"}, {"CODIGO_SIIGO","HORAS","TRAMO"}),
    KeepNoNulls = Table.SelectRows(ExpandAllocRec, each [HORAS] <> null and [HORAS] > 0),

    TipK = Table.TransformColumns(Tip, {{"Código", each try Int64.From(_) otherwise null, Int64.Type}}),
    Merge = Table.NestedJoin(KeepNoNulls, {"CODIGO_SIIGO"}, TipK, {"Código"}, "TIP", JoinKind.LeftOuter),
    ExpandTip = Table.ExpandTableColumn(Merge, "TIP", {"concepto","tipo in o d","Column4"}, {"CONCEPTO_SIIGO","TIPO_SIIGO","UNIDAD_SIIGO"}),

    Final = ExpandTip,
    #"Filas filtradas" = Table.SelectRows(Final, each true)
in
    #"Filas filtradas"
```
