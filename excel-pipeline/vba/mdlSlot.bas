Attribute VB_Name = "mdlSlot"
'==============================================================================
' mdlSlot - cálculo do indicador "Chegada ao Slot" dentro da matriz Excel.
'
' Chamada pelo PowerShell via Application.Run("'slot-matriz.xlsm'!mdlSlot.CalcularIndicadores").
' É uma Function (não Sub) para devolver um resumo que o script registra no log.
'
' Princípios (aprendidos na prática):
'   - Nada de Select/Copy/Paste: a macro lê a tabela inteira para um array (Value2) e grava
'     o resultado de uma vez. O usuário pode continuar usando o PC (e a área de transferência)
'     enquanto o relatório roda.
'   - Round do VBA é "arredondamento bancário" (Round(0.125, 2) = 0.12). Para bater com o
'     Google Sheets, usar WorksheetFunction.Round (meio para longe do zero).
'   - Comparar horários em MINUTOS inteiros, não em frações de dia (double), para evitar
'     que 10:30 "não seja igual" a 10:30 por erro de ponto flutuante.
'   - SpecialCells(xlCellTypeVisible) falha quando a macro é chamada via Application.Run;
'     para contar linhas visíveis, SUBTOTAL(103, ...) funciona sempre (ver ContarVisiveis).
'
' Arquivo salvo no repositório em UTF-8 para leitura no GitHub. O script de montagem
' converte para Windows-1252 antes de importar, porque o editor VBA só entende ANSI.
'==============================================================================
Option Explicit

Private Const ABA_PARAMETROS As String = "Parametros"
Private Const ABA_INDICADORES As String = "Indicadores"
Private Const TABELA_BASE As String = "tbOrdens"
Private Const MAX_GRUPOS As Long = 100

' Colunas do array de contagem: hoje e D-1.
Private Const C_NOPRAZO As Long = 1
Private Const C_FORA As Long = 2
Private Const C_PEND As Long = 3
Private Const C_NOPRAZO_D1 As Long = 4
Private Const C_FORA_D1 As Long = 5

Private grupos As Object            ' Scripting.Dictionary: chave do grupo -> linha em contagens
Private contagens() As Long

'------------------------------------------------------------------------------
' Ponto de entrada. Devolve "OK|linhas lidas|grupos|segundos" ou "ERRO|mensagem".
'------------------------------------------------------------------------------
Public Function CalcularIndicadores() As String
    Dim inicio As Double: inicio = Timer
    On Error GoTo Falha

    Dim agora As Double, janela As Long, tolInstalacao As Long, tolReparo As Long
    Dim meta As Double, margem As Double
    LerParametros agora, janela, tolInstalacao, tolReparo, meta, margem

    Dim lo As ListObject
    Set lo = LocalizarTabela(TABELA_BASE)
    If lo.DataBodyRange Is Nothing Then Err.Raise vbObjectError + 1, , "Tabela " & TABELA_BASE & " vazia."

    Dim dados As Variant: dados = lo.DataBodyRange.Value2
    Dim cData As Long, cRegional As Long, cTipo As Long, cSlot As Long, cChegada As Long, cStatus As Long
    cData = lo.ListColumns("data").Index
    cRegional = lo.ListColumns("regional").Index
    cTipo = lo.ListColumns("tipo").Index
    cSlot = lo.ListColumns("slot").Index
    cChegada = lo.ListColumns("chegada").Index
    cStatus = lo.ListColumns("status").Index

    Dim hoje As Long: hoje = Int(agora)
    Dim agoraMin As Long: agoraMin = MinutosDoDia(agora)

    Set grupos = CreateObject("Scripting.Dictionary")
    ReDim contagens(1 To MAX_GRUPOS, 1 To 5)

    Dim r As Long, dia As Long, d1 As Boolean, situacao As String
    Dim limiteMin As Long, tolerancia As Long, tipo As String, slotTxt As String
    For r = 1 To UBound(dados, 1)
        If IsNumeric(dados(r, cData)) And Not IsEmpty(dados(r, cData)) Then
            dia = CLng(Int(dados(r, cData)))
            If dia = hoje Or dia = hoje - 1 Then
                d1 = (dia = hoje - 1)
                tipo = CStr(dados(r, cTipo))
                tolerancia = IIf(tipo = "Reparo", tolReparo, tolInstalacao)
                limiteMin = MinutosDoDia(dados(r, cSlot)) + janela + tolerancia
                slotTxt = Format$(MinutosDoDia(dados(r, cSlot)) \ 60, "00") & ":" & Format$(MinutosDoDia(dados(r, cSlot)) Mod 60, "00")

                situacao = Classificar(CStr(dados(r, cStatus)), dados(r, cChegada), limiteMin, d1, agoraMin)
                If situacao <> "Excluida" Then
                    Somar "Total|Total", situacao, d1
                    Somar "Regional|" & CStr(dados(r, cRegional)), situacao, d1
                    Somar "Slot|" & slotTxt, situacao, d1
                    Somar "Tipo|" & tipo, situacao, d1
                End If
            End If
        End If
    Next r

    Dim nGrupos As Long
    nGrupos = GravarResultado(meta, margem, agora)

    CalcularIndicadores = "OK|" & UBound(dados, 1) & "|" & nGrupos & "|" & Format$(Timer - inicio, "0.00")
    Exit Function

Falha:
    CalcularIndicadores = "ERRO|" & Err.Description
End Function

'------------------------------------------------------------------------------
' NoPrazo | ForaDoPrazo | Pendente | Excluida (mesma regra do PowerShell e do painel).
'------------------------------------------------------------------------------
Private Function Classificar(ByVal status As String, ByVal chegada As Variant, ByVal limiteMin As Long, _
                             ByVal ehD1 As Boolean, ByVal agoraMin As Long) As String
    If status = "Cancelada" Then
        Classificar = "Excluida"
    ElseIf Not IsEmpty(chegada) And IsNumeric(chegada) Then
        Classificar = IIf(MinutosDoDia(chegada) <= limiteMin, "NoPrazo", "ForaDoPrazo")
    ElseIf ehD1 Or agoraMin > limiteMin Then
        Classificar = "ForaDoPrazo"     ' sem chegada e o limite já passou: estourou
    Else
        Classificar = "Pendente"
    End If
End Function

Private Sub Somar(ByVal chave As String, ByVal situacao As String, ByVal ehD1 As Boolean)
    Dim i As Long
    If Not grupos.Exists(chave) Then
        If grupos.Count >= MAX_GRUPOS Then Err.Raise vbObjectError + 2, , "Grupos demais (limite " & MAX_GRUPOS & ")."
        grupos.Add chave, grupos.Count + 1
    End If
    i = grupos(chave)
    Select Case situacao
        Case "NoPrazo":     contagens(i, IIf(ehD1, C_NOPRAZO_D1, C_NOPRAZO)) = contagens(i, IIf(ehD1, C_NOPRAZO_D1, C_NOPRAZO)) + 1
        Case "ForaDoPrazo": contagens(i, IIf(ehD1, C_FORA_D1, C_FORA)) = contagens(i, IIf(ehD1, C_FORA_D1, C_FORA)) + 1
        Case "Pendente":    If Not ehD1 Then contagens(i, C_PEND) = contagens(i, C_PEND) + 1
    End Select
End Sub

'------------------------------------------------------------------------------
' Monta a aba Indicadores em um único array e grava de uma vez (sem área de transferência).
'------------------------------------------------------------------------------
Private Function GravarResultado(ByVal meta As Double, ByVal margem As Double, ByVal agora As Double) As Long
    Dim chaves As Variant: chaves = OrdenarChaves(grupos.Keys)
    Dim n As Long: n = UBound(chaves) - LBound(chaves) + 1
    Dim saida() As Variant: ReDim saida(1 To n + 1, 1 To 10)

    Dim cab As Variant
    cab = Array("grupo", "chave", "apuradas", "no_prazo", "fora", "pendentes", "percentual", "selo", "percentual_d1", "variacao_pp")
    Dim c As Long
    For c = 0 To 9: saida(1, c + 1) = cab(c): Next c

    Dim k As Long, i As Long, partes() As String, pct As Variant, pctD1 As Variant
    For k = 1 To n
        i = grupos(chaves(LBound(chaves) + k - 1))
        partes = Split(chaves(LBound(chaves) + k - 1), "|")
        pct = Percentual(contagens(i, C_NOPRAZO), contagens(i, C_FORA))
        pctD1 = Percentual(contagens(i, C_NOPRAZO_D1), contagens(i, C_FORA_D1))

        saida(k + 1, 1) = partes(0)
        saida(k + 1, 2) = partes(1)
        saida(k + 1, 3) = contagens(i, C_NOPRAZO) + contagens(i, C_FORA)
        saida(k + 1, 4) = contagens(i, C_NOPRAZO)
        saida(k + 1, 5) = contagens(i, C_FORA)
        saida(k + 1, 6) = contagens(i, C_PEND)
        saida(k + 1, 7) = pct
        saida(k + 1, 8) = Selo(pct, meta, margem)
        saida(k + 1, 9) = pctD1
        If IsEmpty(pct) Or IsEmpty(pctD1) Then
            saida(k + 1, 10) = Empty
        Else
            saida(k + 1, 10) = Application.WorksheetFunction.Round(pct - pctD1, 2)
        End If
    Next k

    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(ABA_INDICADORES)
    ws.Cells.ClearContents
    ' Coluna "chave" como texto: senão o Excel converte "08:30" em hora (0,354...) ao gravar.
    ws.Range("B:B").NumberFormat = "@"
    ws.Range("A1").Resize(n + 1, 10).Value2 = saida
    ws.Range("G2").Resize(n, 1).NumberFormat = "0.00"
    ws.Range("I2").Resize(n, 2).NumberFormat = "0.00"
    ws.Range("L1").Value2 = "atualizado_em"
    ws.Range("L2").Value2 = agora
    ws.Range("L2").NumberFormat = "dd/mm/yyyy hh:mm"

    GravarResultado = n
End Function

' % no prazo com 2 casas, "meio para cima" como no Google Sheets. Vazio se nada foi apurado.
Private Function Percentual(ByVal noPrazo As Long, ByVal fora As Long) As Variant
    If noPrazo + fora = 0 Then
        Percentual = Empty
    Else
        Percentual = Application.WorksheetFunction.Round(noPrazo / (noPrazo + fora) * 100, 2)
    End If
End Function

Private Function Selo(ByVal pct As Variant, ByVal meta As Double, ByVal margem As Double) As String
    If IsEmpty(pct) Then
        Selo = "Sem dados"
    ElseIf pct >= meta Then
        Selo = "Na meta"
    ElseIf pct >= meta - margem Then
        Selo = "Perto"
    Else
        Selo = "Abaixo"
    End If
End Function

'------------------------------------------------------------------------------
' Auxiliares
'------------------------------------------------------------------------------

' Fração de dia (ou data+hora) do Excel -> minutos desde 00:00, como inteiro.
Private Function MinutosDoDia(ByVal valor As Variant) As Long
    Dim fracao As Double: fracao = CDbl(valor) - Int(CDbl(valor))
    MinutosDoDia = CLng(Application.WorksheetFunction.Round(fracao * 1440, 0)) Mod 1440
End Function

Private Sub LerParametros(ByRef agora As Double, ByRef janela As Long, ByRef tolInstalacao As Long, _
                          ByRef tolReparo As Long, ByRef meta As Double, ByRef margem As Double)
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(ABA_PARAMETROS)
    agora = CDbl(ws.Range("prmAgora").Value2)
    janela = CLng(ws.Range("prmJanela").Value2)
    tolInstalacao = CLng(ws.Range("prmTolInstalacao").Value2)
    tolReparo = CLng(ws.Range("prmTolReparo").Value2)
    meta = CDbl(ws.Range("prmMeta").Value2)
    margem = CDbl(ws.Range("prmMargem").Value2)
End Sub

Private Function LocalizarTabela(ByVal nome As String) As ListObject
    Dim ws As Worksheet, lo As ListObject
    For Each ws In ThisWorkbook.Worksheets
        For Each lo In ws.ListObjects
            If lo.Name = nome Then Set LocalizarTabela = lo: Exit Function
        Next lo
    Next ws
    Err.Raise vbObjectError + 3, , "Tabela " & nome & " não encontrada."
End Function

' Ordena as chaves "Categoria|Valor" na ordem Total, Regional, Slot, Tipo.
Private Function OrdenarChaves(ByVal chaves As Variant) As Variant
    Dim i As Long, j As Long, tmp As Variant
    For i = LBound(chaves) To UBound(chaves) - 1
        For j = i + 1 To UBound(chaves)
            If ChaveOrdem(chaves(j)) < ChaveOrdem(chaves(i)) Then
                tmp = chaves(i): chaves(i) = chaves(j): chaves(j) = tmp
            End If
        Next j
    Next i
    OrdenarChaves = chaves
End Function

Private Function ChaveOrdem(ByVal chave As String) As String
    Dim categoria As String: categoria = Split(chave, "|")(0)
    Select Case categoria
        Case "Total":    ChaveOrdem = "1|" & chave
        Case "Regional": ChaveOrdem = "2|" & chave
        Case "Slot":     ChaveOrdem = "3|" & chave
        Case Else:       ChaveOrdem = "4|" & chave
    End Select
End Function

'------------------------------------------------------------------------------
' Conta linhas visíveis de uma coluna filtrada.
' SpecialCells(xlCellTypeVisible) devolve erro quando a macro roda via Application.Run
' (chamada externa); SUBTOTAL com função 103 (CONT.VALORES ignorando ocultas) funciona.
'------------------------------------------------------------------------------
Public Function ContarVisiveis(ByVal coluna As Range) As Long
    ContarVisiveis = CLng(Application.WorksheetFunction.Subtotal(103, coluna))
End Function

'------------------------------------------------------------------------------
' Se um dia for preciso colar valores via PasteSpecial, use as constantes nomeadas:
'   xlPasteValues = -4163  (só valores)
'   xlPasteAll    = -4104  (tudo, inclusive formatação condicional)
' Trocar uma pela outra já apagou a formatação condicional de dias inteiros de histórico.
' Nesta matriz, valores são copiados com .Value2 = .Value2, sem área de transferência.
'------------------------------------------------------------------------------
