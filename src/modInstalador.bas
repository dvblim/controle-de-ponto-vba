Attribute VB_Name = "modInstalador"
Option Explicit

' =====================================================================
'  Instalador do painel de ponto
'  1. Salve a planilha como .xlsm
'  2. Alt+F11 > Arquivo > Importar arquivo > modInstalador.bas
'  3. Alt+F8 > InstalarPainel > Executar e escolha a pasta com os arquivos
'  Requer: Opções > Central de Confiabilidade > Configurações de Macro >
'          "Confiar no acesso ao modelo de objeto do projeto do VBA"
' =====================================================================

Private Const ARQUIVOS As String = "modPainel.bas|clsRegra.cls|clsDia.cls|clsRegistros.cls|clsFeriados.cls|clsBotao.cls|clsMesCard.cls|clsCampo.cls|clsLista.cls"
Private Const ARQ_FORM As String = "frmPainel_codigo.txt"

Public Sub InstalarPainel()
    Dim proj As Object, pasta As String, nomes As Variant, i As Long, comp As Object

    On Error Resume Next
    Set proj = ThisWorkbook.VBProject
    i = proj.VBComponents.Count
    If Err.Number <> 0 Then
        On Error GoTo 0
        MsgBox "O Excel não permitiu acesso ao projeto VBA." & vbCrLf & vbCrLf & _
               "Arquivo > Opções > Central de Confiabilidade > Configurações da Central de Confiabilidade > " & _
               "Configurações de Macro > marque 'Confiar no acesso ao modelo de objeto do projeto do VBA'. " & _
               "Depois rode InstalarPainel de novo.", vbExclamation, "Instalar painel"
        Exit Sub
    End If
    On Error GoTo 0

    pasta = EscolherPasta()
    If pasta = "" Then Exit Sub

    nomes = Split(ARQUIVOS, "|")
    For i = 0 To UBound(nomes)
        If Dir(pasta & nomes(i)) = "" Then
            MsgBox "Arquivo não encontrado na pasta escolhida: " & nomes(i), vbExclamation, "Instalar painel"
            Exit Sub
        End If
    Next i
    If Dir(pasta & ARQ_FORM) = "" Then
        MsgBox "Arquivo não encontrado na pasta escolhida: " & ARQ_FORM, vbExclamation, "Instalar painel"
        Exit Sub
    End If

    On Error GoTo Falha

    ' remove versões anteriores (reinstalação)
    RemoverComp proj, "frmPainel"
    For i = 0 To UBound(nomes)
        RemoverComp proj, Left$(nomes(i), InStrRev(nomes(i), ".") - 1)
    Next i

    ' formulário primeiro: ele adiciona a referência ao Microsoft Forms 2.0
    Set comp = proj.VBComponents.Add(3)          ' vbext_ct_MSForm
    comp.Name = "frmPainel"
    comp.Properties("Caption") = "Controle de ponto"
    With comp.CodeModule
        If .CountOfLines > 0 Then .DeleteLines 1, .CountOfLines
        .AddFromFile pasta & ARQ_FORM
    End With

    For i = 0 To UBound(nomes)
        proj.VBComponents.Import pasta & nomes(i)
    Next i

    CriarBotao

    If ThisWorkbook.FileFormat <> 52 Then
        MsgBox "Painel instalado." & vbCrLf & vbCrLf & _
               "ATENÇÃO: salve agora como 'Pasta de Trabalho Habilitada para Macro (.xlsm)', " & _
               "senão o painel some ao fechar o arquivo.", vbInformation, "Instalar painel"
    Else
        MsgBox "Painel instalado. Use o botão 'Abrir painel' na aba Registros.", vbInformation, "Instalar painel"
    End If
    Exit Sub

Falha:
    MsgBox "Erro na instalação: " & Err.Description, vbCritical, "Instalar painel"
End Sub

Private Function EscolherPasta() As String
    With Application.FileDialog(4)               ' msoFileDialogFolderPicker
        .Title = "Escolha a pasta com os arquivos do painel"
        If ThisWorkbook.Path <> "" Then .InitialFileName = ThisWorkbook.Path & Application.PathSeparator
        If .Show = -1 Then EscolherPasta = .SelectedItems(1) & Application.PathSeparator
    End With
End Function

' Renomeia antes de remover: a remoção só termina quando a macro acaba,
' e sem renomear o Import criaria "clsRegra1" etc.
Private Sub RemoverComp(ByVal proj As Object, ByVal nome As String)
    Dim c As Object
    On Error Resume Next
    Set c = proj.VBComponents(nome)
    On Error GoTo 0
    If c Is Nothing Then Exit Sub
    c.Name = nome & "_antigo" & Format$(Now, "hhmmss")
    proj.VBComponents.Remove c
End Sub

Private Sub CriarBotao()
    Dim ws As Worksheet, r As Range, b As Object
    Set ws = ThisWorkbook.Worksheets("Registros")
    On Error Resume Next
    ws.Buttons("btnAbrirPainel").Delete
    Set r = ThisWorkbook.Names("AreaBotao").RefersToRange
    On Error GoTo 0
    If r Is Nothing Then Set r = ws.Range("A3:C3")
    If r.Cells(1, 1).Value = "Espaço do botão do painel" Then r.Cells(1, 1).ClearContents
    If r.RowHeight < 24 Then r.RowHeight = 24
    Set b = ws.Buttons.Add(r.Left + 2, r.Top + 2, 110, r.Height - 4)
    b.Name = "btnAbrirPainel"
    b.Caption = "Abrir painel"
    b.OnAction = "AbrirPainel"
End Sub
