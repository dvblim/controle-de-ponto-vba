VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsRegistros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' =====================================================================
'  clsRegistros - leitura e gravação da tabela tbRegistros
'  Grava só as colunas de entrada (Data, Tipo, 4 horários, Observação);
'  as colunas de fórmula da planilha continuam valendo.
' =====================================================================

Private mLo As ListObject
Private mDias As Collection
Private mIdx As Object          ' CLng(data) -> clsDia
Private cData As Long, cTipo As Long, cEnt As Long, cSai As Long
Private cRet As Long, cFim As Long, cObs As Long

Public Sub Carregar(ByVal regra As clsRegra)
    Dim v As Variant, i As Long, j As Long, n As Long, acum As Long
    Dim dia As clsDia, x As clsDia, lista() As clsDia, d As Date

    Set mLo = ThisWorkbook.Worksheets("Registros").ListObjects("tbRegistros")
    cData = mLo.ListColumns("Data").Index
    cTipo = mLo.ListColumns("Tipo").Index
    cEnt = mLo.ListColumns("Entrada").Index
    cSai = mLo.ListColumns("Saída almoço").Index
    cRet = mLo.ListColumns("Retorno almoço").Index
    cFim = mLo.ListColumns("Saída").Index
    cObs = mLo.ListColumns("Observação").Index

    Set mDias = New Collection
    Set mIdx = CreateObject("Scripting.Dictionary")
    If mLo.DataBodyRange Is Nothing Then Exit Sub

    v = mLo.DataBodyRange.Value
    ReDim lista(1 To UBound(v, 1))
    For i = 1 To UBound(v, 1)
        d = DataDe(v(i, cData))
        If d > 0 Then
            Set dia = New clsDia
            dia.Data = d
            dia.Tipo = TextoDe(v(i, cTipo))
            If dia.Tipo = "" Then dia.Tipo = "Normal"
            dia.Entrada = HoraDe(v(i, cEnt))
            dia.SaidaAlm = HoraDe(v(i, cSai))
            dia.Retorno = HoraDe(v(i, cRet))
            dia.Saida = HoraDe(v(i, cFim))
            dia.Obs = TextoDe(v(i, cObs))
            dia.Calcular regra
            n = n + 1
            Set lista(n) = dia
        End If
    Next i
    If n = 0 Then Exit Sub

    ' ordena por data (inserção)
    For i = 2 To n
        Set x = lista(i)
        j = i - 1
        Do While j >= 1
            If lista(j).Data <= x.Data Then Exit Do
            Set lista(j + 1) = lista(j)
            j = j - 1
        Loop
        Set lista(j + 1) = x
    Next i

    acum = regra.SaldoInicialMin
    For i = 1 To n
        Set dia = lista(i)
        If Not mIdx.Exists(CLng(dia.Data)) Then
            acum = acum + dia.BancoMin
            dia.AcumuladoMin = acum
            mIdx.Add CLng(dia.Data), dia
            mDias.Add dia
        End If
    Next i
End Sub

Public Property Get Dias() As Collection
    Set Dias = mDias
End Property

Public Property Get Count() As Long
    Count = mDias.Count
End Property

Public Property Get PrimeiraData() As Date
    If mDias.Count > 0 Then PrimeiraData = mDias(1).Data
End Property

Public Function DiaPorData(ByVal d As Date) As clsDia
    Dim k As Long
    k = CLng(Int(CDbl(d)))
    If mIdx.Exists(k) Then Set DiaPorData = mIdx(k)
End Function

' Inclui ou atualiza o dia (procura pela data)
Public Sub Salvar(ByVal dia As clsDia)
    Dim r As Long, lr As ListRow
    Application.ScreenUpdating = False
    r = LinhaDe(dia.Data)
    If r = 0 Then r = LinhaSemData()      ' reaproveita linha vazia (planilha modelo)
    If r = 0 Then
        Set lr = mLo.ListRows.Add(AlwaysInsert:=True)
        r = lr.Index
        CompletarFormulas r
    End If
    With mLo.DataBodyRange
        .Cells(r, cData).Value = dia.Data
        .Cells(r, cTipo).Value = dia.Tipo
        EscreverHora .Cells(r, cEnt), dia.Entrada
        EscreverHora .Cells(r, cSai), dia.SaidaAlm
        EscreverHora .Cells(r, cRet), dia.Retorno
        EscreverHora .Cells(r, cFim), dia.Saida
        If Len(dia.Obs) = 0 Then
            .Cells(r, cObs).ClearContents
        Else
            .Cells(r, cObs).Value = dia.Obs
        End If
    End With
    Ordenar
    Application.ScreenUpdating = True
End Sub

Public Function Excluir(ByVal d As Date) As Boolean
    Dim r As Long
    r = LinhaDe(d)
    If r > 0 Then
        mLo.ListRows(r).Delete
        Excluir = True
    End If
End Function

' ---------------------------------------------------------------------

Private Function LinhaDe(ByVal d As Date) As Long
    Dim i As Long, x As Variant, k As Long
    If mLo.DataBodyRange Is Nothing Then Exit Function
    k = CLng(Int(CDbl(d)))
    For i = 1 To mLo.ListRows.Count
        x = mLo.DataBodyRange.Cells(i, cData).Value
        If DataDe(x) > 0 Then
            If CLng(DataDe(x)) = k Then
                LinhaDe = i
                Exit Function
            End If
        End If
    Next i
End Function

Private Function LinhaSemData() As Long
    Dim i As Long, x As Variant
    If mLo.DataBodyRange Is Nothing Then Exit Function
    For i = 1 To mLo.ListRows.Count
        x = mLo.DataBodyRange.Cells(i, cData).Value
        If IsEmpty(x) Then
            LinhaSemData = i
            Exit Function
        End If
    Next i
End Function

' Garante que a linha nova tenha as fórmulas e formatos da linha vizinha
Private Sub CompletarFormulas(ByVal r As Long)
    Dim ref As Long, c As Long
    If r > 1 Then
        ref = r - 1
    ElseIf mLo.ListRows.Count > 1 Then
        ref = 2
    Else
        Exit Sub
    End If
    With mLo.DataBodyRange
        For c = 1 To mLo.ListColumns.Count
            .Cells(r, c).NumberFormat = .Cells(ref, c).NumberFormat
            If .Cells(ref, c).HasFormula Then
                If Not .Cells(r, c).HasFormula Then
                    .Cells(r, c).FormulaR1C1 = .Cells(ref, c).FormulaR1C1
                End If
            End If
        Next c
    End With
End Sub

Private Sub Ordenar()
    With mLo.Sort
        .SortFields.Clear
        .SortFields.Add Key:=mLo.ListColumns(cData).DataBodyRange, _
                        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .Header = xlYes
        .Apply
    End With
End Sub

Private Sub EscreverHora(ByVal cel As Range, ByVal minutos As Long)
    If minutos < 0 Then
        cel.ClearContents
    Else
        cel.Value = TimeSerial(minutos \ 60, minutos Mod 60, 0)
    End If
End Sub

Private Function DataDe(ByVal x As Variant) As Date
    If IsError(x) Or IsEmpty(x) Then Exit Function
    If VarType(x) = vbDate Then
        DataDe = CDate(Int(CDbl(x)))
    ElseIf VarType(x) = vbString Then
        DataDe = LerData(CStr(x))
    ElseIf IsNumeric(x) Then
        If CDbl(x) > 1 Then DataDe = CDate(Int(CDbl(x)))
    End If
End Function

Private Function HoraDe(ByVal x As Variant) As Long
    Dim f As Double
    HoraDe = -1
    If IsError(x) Or IsEmpty(x) Then Exit Function
    If VarType(x) = vbString Then
        HoraDe = LerHora(CStr(x))
        If HoraDe < 0 Then HoraDe = -1
    ElseIf IsNumeric(x) Or VarType(x) = vbDate Then
        f = CDbl(x)
        f = f - Int(f)
        HoraDe = CLng(Round(f * 1440, 0)) Mod 1440
    End If
End Function

Private Function TextoDe(ByVal x As Variant) As String
    If IsError(x) Or IsEmpty(x) Then Exit Function
    TextoDe = Trim$(CStr(x))
End Function
