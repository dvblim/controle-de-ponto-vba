Attribute VB_Name = "modPainel"
Option Explicit

' =====================================================================
'  Painel de controle de ponto - funções compartilhadas
'  Cores, formatação de horas/datas e criação de controles em runtime.
' =====================================================================

' Cores (formato BGR do VBA: &H00BBGGRR)
Public Const COR_FUNDO As Long = &H161616
Public Const COR_TOPO As Long = &H202020
Public Const COR_CARD As Long = &H1E1E1E
Public Const COR_INPUT As Long = &H1A1A1A
Public Const COR_BORDA As Long = &H3C3C3C
Public Const COR_BOTAO As Long = &H262626
Public Const COR_TEXTO As Long = &HEBEBEB
Public Const COR_MUTED As Long = &H9A9A9A
Public Const COR_APAGADO As Long = &H5F5F5F
Public Const COR_VERDE As Long = &H50B93F        ' RGB(63,185,80)
Public Const COR_VERMELHO As Long = &H6E6EF0     ' RGB(240,110,110)
Public Const COR_AMARELO As Long = &H4CC0E3      ' RGB(227,192,76)
Public Const COR_AZUL As Long = &HFF9A4C         ' RGB(76,154,255)
Public Const COR_AZUL_BORDA As Long = &HB0602A   ' RGB(42,96,176)
Public Const COR_AZUL_BTN As Long = &HC06A1F     ' RGB(31,106,192)
Public Const COR_BARRA_POS As Long = &H78A01E    ' RGB(30,160,120)
Public Const COR_BARRA_NEG As Long = &H4C4CE0    ' RGB(224,76,76)
Public Const COR_BRANCO As Long = &HFFFFFF

Public Const SEP As String = " · "
Public Const TRACO As String = "—"
Public Const FONTE As String = "Segoe UI"

' ---------------------------------------------------------------------
'  Abre o painel (ligado ao botão "Abrir painel" da aba Registros)
' ---------------------------------------------------------------------
Public Sub AbrirPainel()
    Dim lo As ListObject
    On Error Resume Next
    Set lo = ThisWorkbook.Worksheets("Registros").ListObjects("tbRegistros")
    On Error GoTo 0
    If lo Is Nothing Then
        MsgBox "Não encontrei a tabela tbRegistros na aba Registros.", vbExclamation, "Controle de ponto"
        Exit Sub
    End If
    frmPainel.Show vbModeless
End Sub

' ---------------------------------------------------------------------
'  Formatação
' ---------------------------------------------------------------------

' 161 -> "2:41" ; com sinal: "+2:41" / "-2:41"
Public Function FmtHM(ByVal minutos As Long, Optional ByVal comSinal As Boolean = False) As String
    Dim a As Long, s As String
    a = Abs(minutos)
    s = CStr(a \ 60) & ":" & Format$(a Mod 60, "00")
    If minutos < 0 Then
        s = "-" & s
    ElseIf minutos > 0 And comSinal Then
        s = "+" & s
    End If
    FmtHM = s
End Function

' minutos do dia -> "08:14"
Public Function HoraTxt(ByVal minutos As Long) As String
    minutos = ((minutos Mod 1440) + 1440) Mod 1440
    HoraTxt = Format$(minutos \ 60, "00") & ":" & Format$(minutos Mod 60, "00")
End Function

Public Function HoraOuVazio(ByVal minutos As Long) As String
    If minutos < 0 Then
        HoraOuVazio = ""
    Else
        HoraOuVazio = HoraTxt(minutos)
    End If
End Function

Public Function FmtData(ByVal d As Date) As String
    FmtData = Format$(d, "dd\/mm\/yyyy")
End Function

Public Function ChaveMes(ByVal d As Date) As String
    ChaveMes = Format$(Year(d), "0000") & "-" & Format$(Month(d), "00")
End Function

Public Function MesAbrev(ByVal m As Long) As String
    MesAbrev = Choose(m, "jan", "fev", "mar", "abr", "mai", "jun", "jul", "ago", "set", "out", "nov", "dez")
End Function

' "2026-08" -> "ago/26"
Public Function NomeMes(ByVal chave As String) As String
    NomeMes = MesAbrev(CLng(Mid$(chave, 6, 2))) & "/" & Mid$(chave, 3, 2)
End Function

Public Function DiaSemana(ByVal d As Date) As String
    DiaSemana = Choose(Weekday(d), "Dom", "Seg", "Ter", "Qua", "Qui", "Sex", "Sáb")
End Function

Public Function CorSaldo(ByVal minutos As Long) As Long
    If minutos > 0 Then
        CorSaldo = COR_VERDE
    ElseIf minutos < 0 Then
        CorSaldo = COR_VERMELHO
    Else
        CorSaldo = COR_MUTED
    End If
End Function

' ---------------------------------------------------------------------
'  Leitura de texto digitado
' ---------------------------------------------------------------------

Public Function SoDigitos(ByVal s As String) As Boolean
    SoDigitos = (Len(s) > 0) And Not (s Like "*[!0-9]*")
End Function

' "08:14", "8:14", "814", "0814", "8h14" -> minutos do dia
' Retorna -1 se vazio e -2 se inválido.
Public Function LerHora(ByVal s As String) As Long
    Dim p As Variant, hh As String, mm As String, h As Long, m As Long
    s = Trim$(s)
    If s = "" Then
        LerHora = -1
        Exit Function
    End If
    s = Replace(Replace(LCase$(s), "h", ":"), ".", ":")
    LerHora = -2
    If InStr(s, ":") > 0 Then
        p = Split(s, ":")
        If UBound(p) <> 1 Then Exit Function
        hh = CStr(p(0))
        mm = CStr(p(1))
        If mm = "" Then mm = "0"
    Else
        Select Case Len(s)
            Case 1, 2
                hh = s: mm = "0"
            Case 3
                hh = Left$(s, 1): mm = Right$(s, 2)
            Case 4
                hh = Left$(s, 2): mm = Right$(s, 2)
            Case Else
                Exit Function
        End Select
    End If
    If Not (SoDigitos(hh) And SoDigitos(mm)) Then Exit Function
    h = CLng(hh): m = CLng(mm)
    If h > 23 Or m > 59 Then Exit Function
    LerHora = h * 60 + m
End Function

' "08/10/2026", "8/10/26", "08/10" -> Date. Retorna 0 se inválida.
Public Function LerData(ByVal s As String) As Date
    Dim p As Variant, d As Long, m As Long, a As Long
    s = Trim$(s)
    If s = "" Then Exit Function
    p = Split(s, "/")
    If UBound(p) < 1 Or UBound(p) > 2 Then Exit Function
    If Not (SoDigitos(CStr(p(0))) And SoDigitos(CStr(p(1)))) Then Exit Function
    d = CLng(p(0)): m = CLng(p(1))
    If UBound(p) = 2 Then
        If Not SoDigitos(CStr(p(2))) Then Exit Function
        a = CLng(p(2))
        If Len(CStr(p(2))) <= 2 Then a = 2000 + a
    Else
        a = Year(Date)
    End If
    If m < 1 Or m > 12 Or d < 1 Or d > 31 Or a < 1900 Or a > 9999 Then Exit Function
    If Day(DateSerial(a, m, d)) <> d Then Exit Function
    LerData = DateSerial(a, m, d)
End Function

' Chaves de um Dictionary em ordem crescente (Empty se vazio)
Public Function ChavesOrdenadas(ByVal dic As Object) As Variant
    Dim a As Variant, i As Long, j As Long, t As Variant
    If dic.Count = 0 Then
        ChavesOrdenadas = Empty
        Exit Function
    End If
    a = dic.Keys
    For i = 0 To UBound(a) - 1
        For j = i + 1 To UBound(a)
            If a(j) < a(i) Then
                t = a(i): a(i) = a(j): a(j) = t
            End If
        Next j
    Next i
    ChavesOrdenadas = a
End Function

' ---------------------------------------------------------------------
'  Criação de controles em runtime
' ---------------------------------------------------------------------

' Altura aproximada de uma linha de texto para o tamanho de fonte
Public Function AlturaTexto(ByVal tam As Single) As Single
    AlturaTexto = Round(tam * 1.5, 0)
End Function

Public Function NovoLabel(ByVal ctls As Object, ByVal texto As String, _
                          ByVal x As Single, ByVal y As Single, ByVal w As Single, ByVal h As Single, _
                          Optional ByVal tam As Single = 9, Optional ByVal negrito As Boolean = False, _
                          Optional ByVal corTexto As Long = COR_TEXTO, Optional ByVal corFundo As Long = -1, _
                          Optional ByVal alinhar As Long = 1) As MSForms.Label
    Dim lb As MSForms.Label
    Set lb = ctls.Add("Forms.Label.1")
    With lb
        .Left = x: .Top = y: .Width = w: .Height = h
        .WordWrap = False
        .Caption = texto
        .Font.Name = FONTE
        .Font.Size = tam
        .Font.Bold = negrito
        .ForeColor = corTexto
        If corFundo = -1 Then
            .BackStyle = fmBackStyleTransparent
        Else
            .BackStyle = fmBackStyleOpaque
            .BackColor = corFundo
        End If
        .TextAlign = alinhar
    End With
    Set NovoLabel = lb
End Function

' Retângulo com borda (fundo de cartões e botões)
Public Function NovoCartao(ByVal ctls As Object, ByVal x As Single, ByVal y As Single, _
                           ByVal w As Single, ByVal h As Single, _
                           Optional ByVal corBorda As Long = COR_BORDA, _
                           Optional ByVal corFundo As Long = COR_CARD) As MSForms.Label
    Dim lb As MSForms.Label
    Set lb = NovoLabel(ctls, "", x, y, w, h, 9, False, COR_TEXTO, corFundo)
    lb.BorderStyle = fmBorderStyleSingle
    lb.BorderColor = corBorda
    Set NovoCartao = lb
End Function

Public Function NovoTextBox(ByVal ctls As Object, ByVal x As Single, ByVal y As Single, _
                            ByVal w As Single, ByVal h As Single, Optional ByVal tam As Single = 10) As MSForms.TextBox
    Dim tb As MSForms.TextBox
    Set tb = ctls.Add("Forms.TextBox.1")
    With tb
        .Left = x: .Top = y: .Width = w: .Height = h
        .SpecialEffect = fmSpecialEffectFlat
        .BorderStyle = fmBorderStyleSingle
        .BorderColor = COR_BORDA
        .BackColor = COR_INPUT
        .ForeColor = COR_TEXTO
        .Font.Name = FONTE
        .Font.Size = tam
    End With
    Set NovoTextBox = tb
End Function

Public Function NovoComboBox(ByVal ctls As Object, ByVal x As Single, ByVal y As Single, _
                             ByVal w As Single, ByVal h As Single, Optional ByVal tam As Single = 10) As MSForms.ComboBox
    Dim cb As MSForms.ComboBox
    Set cb = ctls.Add("Forms.ComboBox.1")
    With cb
        .Left = x: .Top = y: .Width = w: .Height = h
        .Style = fmStyleDropDownList
        .SpecialEffect = fmSpecialEffectFlat
        .BorderStyle = fmBorderStyleSingle
        .BorderColor = COR_BORDA
        .BackColor = COR_INPUT
        .ForeColor = COR_TEXTO
        .Font.Name = FONTE
        .Font.Size = tam
    End With
    Set NovoComboBox = cb
End Function

Public Function NovoListBox(ByVal ctls As Object, ByVal x As Single, ByVal y As Single, _
                            ByVal w As Single, ByVal h As Single) As MSForms.ListBox
    Dim lst As MSForms.ListBox
    Set lst = ctls.Add("Forms.ListBox.1")
    With lst
        .IntegralHeight = False
        .Left = x: .Top = y: .Width = w: .Height = h
        .SpecialEffect = fmSpecialEffectFlat
        .BorderStyle = fmBorderStyleSingle
        .BorderColor = COR_BORDA
        .BackColor = COR_CARD
        .ForeColor = COR_TEXTO
        .Font.Name = FONTE
        .Font.Size = 9
    End With
    Set NovoListBox = lst
End Function

Public Function NovoFrame(ByVal ctls As Object, ByVal x As Single, ByVal y As Single, _
                          ByVal w As Single, ByVal h As Single) As MSForms.Frame
    Dim f As MSForms.Frame
    Set f = ctls.Add("Forms.Frame.1")
    With f
        .Left = x: .Top = y: .Width = w: .Height = h
        .Caption = ""
        .SpecialEffect = fmSpecialEffectFlat
        .BorderStyle = fmBorderStyleNone
        .BackColor = COR_FUNDO
    End With
    Set NovoFrame = f
End Function
