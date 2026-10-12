// ---------------------------------------------------------------------------
// Consulta "tbOrdens": lê a exportação de ordens (CSV ";" UTF-8) e entrega a base limpa.
//
// O caminho do CSV fica na 1ª linha (CaminhoCsv). O PowerShell reescreve essa linha antes
// de atualizar, em vez de ler o caminho de uma célula: combinar File.Contents com
// Excel.CurrentWorkbook na mesma consulta dispara o erro Formula.Firewall.
//
// Tudo que vem do CSV é tratado como TEXTO e convertido com formato explícito, para não
// depender da configuração regional do Windows de quem atualiza.
// ---------------------------------------------------------------------------
let
    CaminhoCsv = "C:\caminho\para\orders-sample.csv",

    Fonte = Csv.Document(
        File.Contents(CaminhoCsv),
        [Delimiter = ";", Encoding = 65001, QuoteStyle = QuoteStyle.Csv]
    ),
    Cabecalhos = Table.PromoteHeaders(Fonte, [PromoteAllScalars = true]),

    // "SÃ£o" -> "São" (UTF-8 lido como 1252), espaços duplos e nas pontas.
    // Cuidado: "NÃO" e "SÃO" em maiúsculas têm um "Ã" legítimo. Text.FromBinary não lança
    // erro com UTF-8 inválido, só troca por "�" (U+FFFD); nesse caso o texto original é mantido.
    LimparTexto = (t as nullable text) as nullable text =>
        if t = null then null else
        let
            tentativa =
                if Text.Contains(t, "Ã") or Text.Contains(t, "Â")
                then try Text.FromBinary(Text.ToBinary(t, 1252), 65001) otherwise t
                else t,
            reparado = if Text.Contains(tentativa, Character.FromNumber(65533)) then t else tentativa,
            palavras = List.Select(Text.Split(Text.Trim(reparado), " "), each _ <> "")
        in
            Text.Combine(palavras, " "),

    // Nomes próprios em "Título", com preposições minúsculas.
    NomeProprio = (t as nullable text) as nullable text =>
        if t = null then null else
        let
            palavras = Text.Split(Text.Proper(LimparTexto(t)), " "),
            ajustadas = List.Transform(
                List.Positions(palavras),
                (i) => if i > 0 and List.Contains({"Da", "De", "Do", "Das", "Dos", "E"}, palavras{i})
                       then Text.Lower(palavras{i}) else palavras{i})
        in
            Text.Combine(ajustadas, " "),

    // Remove acentos para comparar ("Instalação" = "instalacao").
    SemAcento = (t as text) as text =>
        Text.Lower(Text.Combine(List.Transform(Text.ToList(t), each
            Record.FieldOrDefault([
                á = "a", à = "a", â = "a", ã = "a", é = "e", ê = "e", í = "i",
                ó = "o", ô = "o", õ = "o", ú = "u", ç = "c",
                Á = "a", À = "a", Â = "a", Ã = "a", É = "e", Ê = "e", Í = "i",
                Ó = "o", Ô = "o", Õ = "o", Ú = "u", Ç = "c"
            ], _, _)))),

    // Datas em formatos mistos, inclusive o número serial do Excel ("46307").
    FormatosData = {"dd/MM/yyyy", "d/M/yyyy", "dd/MM/yy", "yyyy-MM-dd", "yyyy/MM/dd", "dd-MM-yyyy"},
    LerData = (t as nullable text) as nullable date =>
        let s = if t = null then "" else Text.Trim(t) in
        if s = "" then null
        else if Text.Length(Text.Select(s, {"0".."9"})) = Text.Length(s)
            then Date.From(Number.From(s))
        else List.First(List.RemoveNulls(List.Transform(FormatosData,
            (f) => try Date.FromText(s, [Format = f, Culture = "pt-BR"]) otherwise null))),

    // "8:05", "08:05" e "08:05:00" -> time.
    LerHora = (t as nullable text) as nullable time =>
        let s = if t = null then "" else Text.Trim(t) in
        if s = "" then null else try Time.FromText(s, [Culture = "pt-BR"]) otherwise null,

    PadronizarTipo = (t as nullable text) as nullable text =>
        let k = SemAcento(LimparTexto(t ?? "")) in
        if Text.StartsWith(k, "instal") then "Instalação"
        else if Text.StartsWith(k, "repar") then "Reparo"
        else null,

    PadronizarStatus = (t as nullable text) as nullable text =>
        let k = SemAcento(LimparTexto(t ?? "")) in
        if k = "concluida" then "Concluída"
        else if k = "em execucao" then "Em execução"
        else if k = "em deslocamento" then "Em deslocamento"
        else if k = "agendada" then "Agendada"
        else if k = "nao realizada" then "Não realizada"
        else if k = "cancelada" then "Cancelada"
        else null,

    Tratada = Table.TransformColumns(Cabecalhos, {
        {"id_os",       each Text.Upper(LimparTexto(_)), type text},
        {"data",        LerData, type date},
        {"regional",    each Text.Upper(LimparTexto(_)), type text},
        {"cidade",      NomeProprio, type text},
        {"gerente",     NomeProprio, type text},
        {"coordenacao", NomeProprio, type text},
        {"supervisao",  NomeProprio, type text},
        {"tecnico",     NomeProprio, type text},
        {"tipo",        PadronizarTipo, type text},
        {"slot",        LerHora, type time},
        {"chegada",     LerHora, type time},
        {"inicio",      LerHora, type time},
        {"termino",     LerHora, type time},
        {"status",      PadronizarStatus, type text}
    }),

    // Exportações repetidas trazem a mesma ordem duas vezes: fica a primeira.
    SemDuplicadas = Table.Distinct(Tratada, {"id_os"}),
    Validas = Table.SelectRows(SemDuplicadas, each [data] <> null and [tipo] <> null and [status] <> null and [slot] <> null)
in
    Validas
