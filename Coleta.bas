Attribute VB_Name = "Coleta"
Option Explicit
' Coleta de assinaturas da ficha da faca (aba definida em ABA).
' 1) IniciarColeta  -> inicia ou continua a coleta da faca em D2: exporta a ficha, sobe o servidor, abre o QR
' 2) FinalizarColeta -> cola as assinaturas no ficha, exporta o PDF final, encerra o servidor
'    (tambem disparado pelo celular: o servidor cria finalizar.flag e o Vigiar executa)

' AJUSTE: pasta de rede deste projeto (caminho UNC, nao letra de unidade, para funcionar em qualquer PC)
Private Const BASE As String = "\\servidor\compartilhamento\assinatura-ficha"
Private Const PORTA As String = "5055"
Private Const ABA As String = "FICHA*"   ' nome da aba da ficha (padrao Like, maiusculas)

#If VBA7 Then
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As Long)
#Else
Private Declare Sub Sleep Lib "kernel32" (ByVal ms As Long)
#End If

' coleta ativa neste Excel, para o Finalizar vindo do celular
Private mPasta As String, mFaca As String, mWb As Workbook, mProx As Date

' chave do setor | celulas da assinatura (linhas 66 e 68 do ficha)
Private Function Slots() As Variant
    Slots = Array("offset|A66:B66", "corte|D66:E66", "colagem|G66:H66", "eng|J66:K66", _
                  "coord|A68:B68", "pcp|D68:E68", "excelencia|G68:H68")
End Function

Private Function Ficha(Optional wb As Workbook) As Worksheet
    Dim ws As Worksheet
    If wb Is Nothing Then Set wb = ActiveWorkbook
    For Each ws In wb.Worksheets
        If UCase$(ws.Name) Like ABA Then Set Ficha = ws: Exit Function
    Next
    Err.Raise vbObjectError + 1, , "Aba da ficha (" & ABA & ") nao encontrada no arquivo."
End Function

Private Function PastaColeta(ByVal faca As String) As String
    Dim p As String
    p = BASE & "\Coletas"
    If Dir(p, vbDirectory) = "" Then MkDir p
    p = p & "\" & faca
    If Dir(p, vbDirectory) = "" Then MkDir p
    PastaColeta = p
End Function

Private Function LerUrl(ByVal pasta As String) As String
    Dim f As Integer
    If Dir(pasta & "\url.txt") = "" Then Exit Function
    f = FreeFile: Open pasta & "\url.txt" For Input As #f: Line Input #f, LerUrl: Close #f
End Function

Private Sub Grava(ByVal arq As String, ByVal texto As String)
    Dim f As Integer
    f = FreeFile: Open arq For Output As #f: Print #f, texto: Close #f
End Sub

Private Sub Apaga(ByVal arq As String)
    On Error Resume Next
    Kill arq
End Sub

' servidor desta coleta ainda responde? (PC nao reiniciou, ninguem fechou)
Private Function ServidorVivo(ByVal url As String) As Boolean
    Dim h As Object
    On Error GoTo Fim
    Set h = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    h.setTimeouts 1000, 1000, 1000, 1500
    h.Open "GET", Replace(url, "/?t=", "/status?t="), False
    h.send
    ServidorVivo = (h.Status = 200)
Fim:
End Function

' so cabe uma coleta na porta: encerra os servidores de outras facas (as assinaturas delas ficam salvas)
Private Function PararOutras(ByVal pasta As String) As Boolean
    Dim nomes As New Collection, n As String, item As Variant
    n = Dir(BASE & "\Coletas\*", vbDirectory)
    Do While n <> ""
        If n <> "." And n <> ".." Then nomes.Add BASE & "\Coletas\" & n
        n = Dir()
    Loop
    For Each item In nomes
        If item <> pasta And Dir(item & "\url.txt") <> "" And Dir(item & "\stop.flag") = "" Then
            Grava item & "\stop.flag", ""
            PararOutras = True
        End If
    Next
End Function

' Area de impressao como PNG nitido (2x) para ver no celular. Usa a area de transferencia.
Private Sub ExportaImagem(ws As Worksheet, ByVal arq As String)
    Dim pa As Range, co As ChartObject, sel As Object
    Set sel = Selection
    ws.Activate
    Set pa = ws.Range(ws.PageSetup.PrintArea)
    pa.CopyPicture Appearance:=xlScreen, Format:=xlPicture
    Set co = ws.ChartObjects.Add(0, 0, pa.Width * 2, pa.Height * 2)
    co.Chart.ChartArea.Format.Line.Visible = msoFalse
    co.Activate
    co.Chart.Paste
    With co.Chart.Shapes(co.Chart.Shapes.Count)
        .LockAspectRatio = msoFalse
        .Left = 0: .Top = 0
        .Width = co.Chart.ChartArea.Width: .Height = co.Chart.ChartArea.Height
    End With
    co.Chart.Export arq, "PNG"
    co.Delete
    sel.Select
End Sub

' slots.json: [chave, titulo, esq, topo, larg, alt] em fracao da area de impressao
Private Sub GravaSlots(ws As Worksheet, ByVal arq As String)
    Dim pa As Range, rg As Range, item As Variant, s As String
    Set pa = ws.Range(ws.PageSetup.PrintArea)
    For Each item In Slots()
        Set rg = ws.Range(Split(item, "|")(1))
        If s <> "" Then s = s & ","
        s = s & "[""" & Split(item, "|")(0) & """,""" & _
            Replace(Trim$(rg.Offset(1, 0).Cells(1, 1).Text), """", "'") & """," & _
            Num((rg.Left - pa.Left) / pa.Width) & "," & Num((rg.Top - pa.Top) / pa.Height) & "," & _
            Num(rg.Width / pa.Width) & "," & Num(rg.Height / pa.Height) & "]"
    Next
    Grava arq, "[" & s & "]"
End Sub

Private Function Num(ByVal v As Double) As String
    Num = Replace(Format(v, "0.0000"), ",", ".")   ' JSON exige ponto, Windows pt-BR usa virgula
End Function

Private Sub MostraQR(ByVal url As String, ByVal faca As String)
    On Error GoTo SemNavegador   ' pagina com link grande + QR code para o celular
    ActiveWorkbook.FollowHyperlink Replace(url, "/?t=", "/qr?t=")
    Exit Sub
SemNavegador:
    MsgBox "Abra no celular:" & vbCrLf & vbCrLf & url, vbInformation, "Coleta " & faca
End Sub

Public Sub IniciarColeta()
    Dim ws As Worksheet, faca As String, pasta As String, py As String, url As String, i As Integer
    Set ws = Ficha()
    faca = Trim$(CStr(ws.Range("D2").Value))
    If faca = "" Then MsgBox "Informe a faca em D2.": Exit Sub
    pasta = PastaColeta(faca)

    url = LerUrl(pasta)
    If url <> "" Then
        If ServidorVivo(url) Then   ' coleta ja rodando: so volta a vigiar e reabre o QR
            Vigia ws.Parent, faca, pasta
            MostraQR url, faca
            Exit Sub
        End If
    End If

    If PararOutras(pasta) Then Sleep 1500   ' servidor checa o stop.flag a cada 1 s
    Apaga pasta & "\url.txt": Apaga pasta & "\stop.flag"
    Apaga pasta & "\finalizar.flag": Apaga pasta & "\erro.txt": Apaga pasta & "\pdf.txt"

    ws.ExportAsFixedFormat Type:=xlTypePDF, Filename:=pasta & "\ficha.pdf", _
        IgnorePrintAreas:=False, OpenAfterPublish:=False
    GravaSlots ws, pasta & "\slots.json"
    On Error Resume Next   ' sem a imagem a pagina mostra so o PDF
    ExportaImagem ws, pasta & "\ficha.png"
    On Error GoTo 0

    py = BASE & "\python\python.exe"
    If Dir(py) = "" Then py = "python"   ' sem Python portatil, tenta o do sistema
    CreateObject("WScript.Shell").Run """" & py & """ """ & BASE & "\server.py"" """ & pasta & """ " & PORTA, 0, False

    For i = 1 To 50   ' espera ate 10 s o servidor gravar o link
        If Dir(pasta & "\url.txt") <> "" Then Exit For
        Sleep 200
    Next
    url = LerUrl(pasta)
    If url = "" Then MsgBox "Servidor nao iniciou. Verifique o Python ou se a porta " & PORTA & " esta ocupada.": Exit Sub

    Vigia ws.Parent, faca, pasta
    MostraQR url, faca
End Sub

Public Sub FinalizarColeta()
    Dim ws As Worksheet
    Set ws = Ficha()
    Finaliza ws.Parent, Trim$(CStr(ws.Range("D2").Value)), False
End Sub

' Aplica as assinaturas da faca no ficha de wb, gera o PDF, encerra o servidor e arquiva a coleta.
Private Sub Finaliza(wb As Workbook, ByVal faca As String, ByVal remoto As Boolean)
    Dim ws As Worksheet, pasta As String, item As Variant, k As String, rg As Range
    Dim png As String, shp As Shape, n As Integer, out As String, antes As String
    Set ws = Ficha(wb)
    pasta = PastaColeta(faca)
    antes = ws.Range("D2").Formula
    If Trim$(CStr(ws.Range("D2").Value)) <> faca Then   ' pedido do celular com outra faca na tela
        ws.Range("D2").Value = faca
        ws.Calculate
    End If

    For Each item In Slots()
        k = Split(item, "|")(0)
        png = pasta & "\assinaturas\" & k & ".png"
        If Dir(png) <> "" Then
            Set rg = ws.Range(Split(item, "|")(1))
            Set shp = ws.Shapes.AddPicture(png, False, True, rg.Left + 2, rg.Top + 2, -1, -1)
            shp.Name = "ASSIN_" & k
            shp.LockAspectRatio = msoTrue
            If shp.Height > rg.Height - 4 Then shp.Height = rg.Height - 4
            If shp.Width > rg.Width - 4 Then shp.Width = rg.Width - 4
            shp.Left = rg.Left + (rg.Width - shp.Width) / 2   ' centraliza, apoiada na linha
            shp.Top = rg.Top + rg.Height - shp.Height - 2
            shp.Placement = xlMoveAndSize
            n = n + 1
        End If
    Next

    out = BASE & "\Assinadas"
    If Dir(out, vbDirectory) = "" Then MkDir out
    out = out & "\" & faca & "_assinada_" & Format(Now, "yyyymmdd_hhnnss") & ".pdf"
    ws.ExportAsFixedFormat Type:=xlTypePDF, Filename:=out, IgnorePrintAreas:=False, OpenAfterPublish:=False

    For Each shp In ws.Shapes   ' deixa o ficha limpo para a proxima faca
        If Left$(shp.Name, 6) = "ASSIN_" Then shp.Delete
    Next
    If ws.Range("D2").Formula <> antes Then ws.Range("D2").Formula = antes   ' devolve a faca que estava na tela

    Grava pasta & "\pdf.txt", out          ' o celular mostra "PDF gerado"
    Grava pasta & "\stop.flag", ""         ' encerra o servidor
    If mPasta = pasta Then mPasta = ""     ' para de vigiar
    Sleep 1500   ' servidor checa o flag a cada 1 s
    ' arquiva a coleta (com log.csv) para a proxima da mesma faca comecar limpa
    On Error Resume Next
    Name pasta As pasta & "_" & Format(Now, "yyyymmdd_hhnnss")
    On Error GoTo 0
    If remoto Then
        Application.StatusBar = "Coleta " & faca & ": PDF gerado pelo celular (" & n & " assinatura(s)). " & out
    Else
        MsgBox n & " assinatura(s) aplicada(s)." & vbCrLf & out, vbInformation, "Coleta " & faca
    End If
End Sub

' ---- Finalizar pelo celular: confere a cada 3 s se o servidor criou finalizar.flag ----
Private Sub Vigia(wb As Workbook, ByVal faca As String, ByVal pasta As String)
    Set mWb = wb: mFaca = faca: mPasta = pasta
    Agenda
End Sub

Private Sub Agenda()
    On Error Resume Next   ' cancela o agendamento anterior, se houver
    Application.OnTime mProx, "'" & ThisWorkbook.Name & "'!Vigiar", , False
    On Error GoTo 0
    mProx = Now + TimeSerial(0, 0, 3)
    Application.OnTime mProx, "'" & ThisWorkbook.Name & "'!Vigiar"
End Sub

Public Sub Vigiar()
    Dim pasta As String
    If mPasta = "" Then Exit Sub
    If Dir(mPasta & "\url.txt") = "" Then mPasta = "": Exit Sub   ' coleta finalizada ou arquivada
    If Dir(mPasta & "\finalizar.flag") = "" Then Agenda: Exit Sub
    pasta = mPasta
    On Error GoTo Falhou
    Finaliza mWb, mFaca, True
    Exit Sub
Falhou:   ' ex.: planilha fechada; o celular mostra o erro e pode tentar de novo
    Grava pasta & "\erro.txt", "Excel: " & Err.Description
    Apaga pasta & "\finalizar.flag"
    Agenda
End Sub

' Botoes da guia "Assinaturas" (customUI do Coleta.xlam)
Public Sub BtnIniciar(control As IRibbonControl)
    IniciarColeta
End Sub

Public Sub BtnFinalizar(control As IRibbonControl)
    FinalizarColeta
End Sub
