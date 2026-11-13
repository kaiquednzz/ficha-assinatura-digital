Attribute VB_Name = "Coleta"
Option Explicit
' Coleta de assinaturas da ficha da faca (aba definida em ABA).
' 1) IniciarColeta  -> exporta ficha.pdf, sobe o servidor, mostra o link p/ o celular
' 2) FinalizarColeta -> cola as assinaturas na ficha, exporta o PDF final, encerra o servidor

' AJUSTE: pasta de rede deste projeto (caminho UNC, nao letra de unidade, para funcionar em qualquer PC)
Private Const BASE As String = "\\servidor\compartilhamento\assinatura-ficha"
Private Const PORTA As String = "5055"
Private Const ABA As String = "FICHA*"   ' nome da aba da ficha (padrao Like, maiusculas)

#If VBA7 Then
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As Long)
#Else
Private Declare Sub Sleep Lib "kernel32" (ByVal ms As Long)
#End If

' chave do setor | celulas da assinatura (linhas 66 e 68 da ficha)
Private Function Slots() As Variant
    Slots = Array("offset|A66:B66", "corte|D66:E66", "colagem|G66:H66", "eng|J66:K66", _
                  "coord|A68:B68", "pcp|D68:E68", "excelencia|G68:H68")
End Function

Private Function Ficha() As Worksheet
    Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
        If UCase$(ws.Name) Like ABA Then Set Ficha = ws: Exit Function
    Next
    Err.Raise vbObjectError + 1, , "Aba da ficha (" & ABA & ") nao encontrada no arquivo ativo."
End Function

Private Function PastaColeta(ByVal faca As String) As String
    Dim p As String
    p = BASE & "\Coletas"
    If Dir(p, vbDirectory) = "" Then MkDir p
    p = p & "\" & faca
    If Dir(p, vbDirectory) = "" Then MkDir p
    PastaColeta = p
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
    Dim pa As Range, rg As Range, item As Variant, s As String, f As Integer
    Set pa = ws.Range(ws.PageSetup.PrintArea)
    For Each item In Slots()
        Set rg = ws.Range(Split(item, "|")(1))
        If s <> "" Then s = s & ","
        s = s & "[""" & Split(item, "|")(0) & """,""" & _
            Replace(Trim$(rg.Offset(1, 0).Cells(1, 1).Text), """", "'") & """," & _
            Num((rg.Left - pa.Left) / pa.Width) & "," & Num((rg.Top - pa.Top) / pa.Height) & "," & _
            Num(rg.Width / pa.Width) & "," & Num(rg.Height / pa.Height) & "]"
    Next
    f = FreeFile: Open arq For Output As #f: Print #f, "[" & s & "]": Close #f
End Sub

Private Function Num(ByVal v As Double) As String
    Num = Replace(Format(v, "0.0000"), ",", ".")   ' JSON exige ponto, Windows pt-BR usa virgula
End Function

Public Sub IniciarColeta()
    Dim ws As Worksheet, faca As String, pasta As String, py As String, f As Integer, url As String, i As Integer
    Set ws = Ficha()
    faca = Trim$(CStr(ws.Range("D2").Value))
    If faca = "" Then MsgBox "Informe a faca em D2.": Exit Sub
    pasta = PastaColeta(faca)

    On Error Resume Next
    Kill pasta & "\url.txt": Kill pasta & "\stop.flag"
    On Error GoTo 0

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
    If Dir(pasta & "\url.txt") = "" Then MsgBox "Servidor nao iniciou. Verifique o Python ou finalize a coleta anterior (porta " & PORTA & " ocupada).": Exit Sub

    f = FreeFile: Open pasta & "\url.txt" For Input As #f: Line Input #f, url: Close #f
    On Error GoTo SemNavegador   ' pagina com link grande + QR code para o celular
    ActiveWorkbook.FollowHyperlink Replace(url, "/?t=", "/qr?t=")
    Exit Sub
SemNavegador:
    MsgBox "Abra no celular:" & vbCrLf & vbCrLf & url, vbInformation, "Coleta " & faca
End Sub

Public Sub FinalizarColeta()
    Dim ws As Worksheet, faca As String, pasta As String, item As Variant, k As String, rg As Range
    Dim png As String, shp As Shape, n As Integer, out As String, f As Integer
    Set ws = Ficha()
    faca = Trim$(CStr(ws.Range("D2").Value))
    pasta = PastaColeta(faca)

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

    For Each shp In ws.Shapes   ' deixa a ficha limpa para a proxima faca
        If Left$(shp.Name, 6) = "ASSIN_" Then shp.Delete
    Next

    f = FreeFile: Open pasta & "\stop.flag" For Output As #f: Close #f   ' encerra o servidor
    Sleep 1500   ' servidor checa o flag a cada 1 s
    ' arquiva a coleta (com log.csv) para a proxima da mesma faca comecar limpa
    On Error Resume Next
    Name pasta As pasta & "_" & Format(Now, "yyyymmdd_hhnnss")
    On Error GoTo 0
    MsgBox n & " assinatura(s) aplicada(s)." & vbCrLf & out, vbInformation, "Coleta " & faca
End Sub

' Botoes da guia "Assinaturas" (customUI do Coleta.xlam)
Public Sub BtnIniciar(control As IRibbonControl)
    IniciarColeta
End Sub

Public Sub BtnFinalizar(control As IRibbonControl)
    FinalizarColeta
End Sub
