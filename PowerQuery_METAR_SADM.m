// =====================================================================
//  METAR/SPECI -> IMC/VMC para SADM (Moron)  -- Power Query (Excel)
// ---------------------------------------------------------------------
//  Como usarlo:
//    1. Excel -> Datos -> Obtener datos -> Desde otras fuentes -> Consulta en blanco
//    2. Ver -> Editor avanzado -> borrar todo y pegar ESTE texto
//    3. Cambiar Desde / Hasta / VisMin / TechoMin arriba de todo
//    4. Cerrar y cargar -> se crea una tabla en una hoja nueva
//    5. Para actualizar: Datos -> Actualizar todo
//  La primera vez Excel pregunta por el acceso al origen web: elegir "Anonimo".
// =====================================================================
let
    // ---------- PARAMETROS ----------
    Estacion  = "SADM",
    Desde     = #date(2026, 1, 1),
    Hasta     = Date.From(DateTime.LocalNow()),   // hasta hoy
    VisMin    = 5000,    // metros
    TechoMin  = 1500,    // pies AGL

    // ---------- DESCARGA ----------
    Origen = Csv.Document(
        Web.Contents(
            "https://mesonet.agron.iastate.edu/cgi-bin/request/asos.py",
            [Query = [
                station = Estacion,      data   = "metar",
                year1 = Text.From(Date.Year(Desde)),  month1 = Text.From(Date.Month(Desde)), day1 = Text.From(Date.Day(Desde)),
                year2 = Text.From(Date.Year(Hasta)),  month2 = Text.From(Date.Month(Hasta)), day2 = Text.From(Date.Day(Hasta)),
                tz = "Etc/UTC", format = "onlycomma", missing = "M", trace = "T", direct = "no"
            ]]
        ),
        [Delimiter = ",", Columns = 3, Encoding = 65001, QuoteStyle = QuoteStyle.None]
    ),
    Encabezados = Table.PromoteHeaders(Origen, [PromoteAllScalars = true]),
    Limpio = Table.SelectRows(Encabezados, each [metar] <> null and [metar] <> "" and [metar] <> "M"),
    Tipos  = Table.TransformColumnTypes(Limpio, {{"valid", type datetime}}),

    // ---------- PARSER ----------
    Dig = {"0".."9"},
    SoloDig = (t as text) as logical => t <> "" and Text.Select(t, Dig) = t,

    Cuerpo = (metar as text) as list =>
        let
            Tk   = Text.Split(Text.Replace(Text.Trim(metar), "=", ""), " "),
            Hora = List.PositionOf(Tk, List.First(List.Select(Tk,
                       each Text.Length(_) = 7 and Text.End(_, 1) = "Z" and SoloDig(Text.Start(_, 6))) & {""})),
            Post = if Hora = -1 then Tk else List.Skip(Tk, Hora + 1),
            Cort = List.Select(List.Transform({"TEMPO","BECMG","NOSIG","RMK"}, each List.PositionOf(Post, _)), each _ >= 0),
            Fin  = if List.IsEmpty(Cort) then List.Count(Post) else List.Min(Cort)
        in  List.FirstN(Post, Fin),

    VisDe = (metar as text) as nullable number =>
        let c = Cuerpo(metar) in
        if List.Contains(c, "CAVOK") then 10000
        else let v = List.Select(c, each Text.Length(_) = 4 and SoloDig(_)) in
             if List.IsEmpty(v) then (if List.Count(List.Select(c, each List.Contains({"NSC","NCD","SKC","CLR"}, _))) > 0 then 10000 else null)
             else (let n = Number.From(List.First(v)) in if n = 9999 then 10000 else n),

    TechoDe = (metar as text) as number =>
        let
            c   = Cuerpo(metar),
            Cap = List.Transform(List.Select(c, each Text.Length(_) >= 6
                        and List.Contains({"BKN","OVC"}, Text.Start(_, 3)) and SoloDig(Text.Middle(_, 3, 3))),
                        each Number.From(Text.Middle(_, 3, 3)) * 100),
            VV  = List.Transform(List.Select(c, each Text.StartsWith(_, "VV") and SoloDig(Text.Middle(_, 2, 3))),
                        each Number.From(Text.Middle(_, 2, 3)) * 100),
            T   = Cap & VV
        in  if List.Contains(c, "CAVOK") or List.IsEmpty(T) then 99999 else List.Min(T),

    // ---------- COLUMNAS POR REPORTE ----------
    C1 = Table.AddColumn(Tipos, "TIPO",     each if Time.Minute(DateTime.Time([valid])) = 0 then "METAR" else "SPECI", type text),
    C2 = Table.AddColumn(C1,    "VIS_M",    each VisDe([metar]),   type number),
    C3 = Table.AddColumn(C2,    "TECHO_FT", each TechoDe([metar]), type number),
    C4 = Table.AddColumn(C3,    "CONDICION",
            each if [VIS_M] = null then "SIN DATO"
                 else if [VIS_M] < VisMin or [TECHO_FT] < TechoMin then "IMC" else "VMC", type text),
    C5 = Table.AddColumn(C4, "FECHA",  each DateTime.Date([valid]), type date),
    C6 = Table.AddColumn(C5, "HORA_Z", each Time.Hour(DateTime.Time([valid])), Int64.Type),
    C7 = Table.AddColumn(C6, "CLAVE",  each Date.ToText(DateTime.Date([valid]), "yyyy-MM-dd") & " " &
                                            Text.PadStart(Text.From(Time.Hour(DateTime.Time([valid]))), 2, "0"), type text),

    // ---------- AGRUPADO POR HORA (peor condicion de la hora, incluye SPECI) ----------
    Grupo = Table.Group(C7, {"CLAVE", "FECHA", "HORA_Z"}, {
        {"CONDICION",  each if List.Contains([CONDICION], "IMC") then "IMC"
                            else if List.Contains([CONDICION], "VMC") then "VMC" else "SIN DATO", type text},
        {"VIS_M",      each List.Min(List.RemoveNulls([VIS_M])),    type number},
        {"TECHO_FT",   each List.Min(List.RemoveNulls([TECHO_FT])), type number},
        {"HUBO_SPECI", each if List.Contains([TIPO], "SPECI") then "SI" else "NO", type text},
        {"METAR",      each List.First(List.Select([metar], each true)), type text},
        {"REPORTES",   each Text.Combine([metar], "  ||  "), type text}
    }),
    Orden = Table.Sort(Grupo, {{"CLAVE", Order.Ascending}}),
    Final = Table.SelectColumns(Orden,
        {"CLAVE", "FECHA", "HORA_Z", "CONDICION", "VIS_M", "TECHO_FT", "HUBO_SPECI", "METAR", "REPORTES"})
in
    Final

// =====================================================================
//  En tu planilla de movimientos, con FECHA en A y HORA_Z en B:
//    Condicion :  =BUSCARV(TEXTO(A2,"aaaa-mm-dd")&" "&TEXTO(B2,"00"); METAR_SADM[#Todo]; 4; FALSO)
//    Visibilidad: =BUSCARV(TEXTO(A2,"aaaa-mm-dd")&" "&TEXTO(B2,"00"); METAR_SADM[#Todo]; 5; FALSO)
//    METAR      : =BUSCARV(TEXTO(A2,"aaaa-mm-dd")&" "&TEXTO(B2,"00"); METAR_SADM[#Todo]; 8; FALSO)
//  (reemplazar METAR_SADM por el nombre real de la tabla que crea Power Query)
// =====================================================================
