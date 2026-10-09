VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFeriados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' =====================================================================
'  clsFeriados - feriados nacionais (calculados, inclusive os móveis)
'  + feriados locais/pontos facultativos da tabela tbFeriados (Config)
' =====================================================================

Private mDic As Object      ' CLng(data) -> nome
Private mAnos As Object     ' anos já calculados

Public Sub Inicializar(ByVal wb As Workbook)
    Dim lo As ListObject, i As Long, x As Variant, nm As String
    Set mDic = CreateObject("Scripting.Dictionary")
    Set mAnos = CreateObject("Scripting.Dictionary")

    On Error Resume Next
    Set lo = wb.Worksheets("Config").ListObjects("tbFeriados")
    On Error GoTo 0
    If lo Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub

    For i = 1 To lo.ListRows.Count
        x = lo.DataBodyRange.Cells(i, 1).Value
        If Not IsError(x) And Not IsEmpty(x) Then
            If VarType(x) = vbDate Or IsNumeric(x) Then
                nm = ""
                If Not IsError(lo.DataBodyRange.Cells(i, 2).Value) Then
                    nm = Trim$(CStr(lo.DataBodyRange.Cells(i, 2).Value))
                End If
                If nm = "" Then nm = "Feriado local"
                mDic(CLng(Int(CDbl(x)))) = nm    ' locais têm prioridade
            End If
        End If
    Next i
End Sub

' Nome do feriado na data ("" se não for feriado)
Public Function Nome(ByVal d As Date) As String
    Dim k As Long
    k = CLng(Int(CDbl(d)))
    GarantirAno Year(d)
    If mDic.Exists(k) Then Nome = mDic(k)
End Function

' Feriados entre duas datas: Collection de Array(CLng(data), nome)
Public Function Entre(ByVal d1 As Date, ByVal d2 As Date) As Collection
    Dim c As Collection, k As Long, a As Long
    Set c = New Collection
    For a = Year(d1) To Year(d2)
        GarantirAno a
    Next a
    For k = CLng(Int(CDbl(d1))) To CLng(Int(CDbl(d2)))
        If mDic.Exists(k) Then c.Add Array(k, mDic(k))
    Next k
    Set Entre = c
End Function

Private Sub GarantirAno(ByVal ano As Long)
    Dim p As Date
    If mAnos.Exists(ano) Then Exit Sub
    mAnos.Add ano, True
    p = Pascoa(ano)
    Incluir DateSerial(ano, 1, 1), "Confraternização Universal"
    Incluir p - 48, "Carnaval (ponto facultativo)"
    Incluir p - 47, "Carnaval (ponto facultativo)"
    Incluir p - 2, "Sexta-feira Santa"
    Incluir DateSerial(ano, 4, 21), "Tiradentes"
    Incluir DateSerial(ano, 5, 1), "Dia do Trabalho"
    Incluir p + 60, "Corpus Christi (ponto facultativo)"
    Incluir DateSerial(ano, 9, 7), "Independência"
    Incluir DateSerial(ano, 10, 12), "Nossa Senhora Aparecida"
    Incluir DateSerial(ano, 11, 2), "Finados"
    Incluir DateSerial(ano, 11, 15), "Proclamação da República"
    Incluir DateSerial(ano, 11, 20), "Consciência Negra"
    Incluir DateSerial(ano, 12, 25), "Natal"
End Sub

Private Sub Incluir(ByVal d As Date, ByVal nm As String)
    Dim k As Long
    k = CLng(Int(CDbl(d)))
    If Not mDic.Exists(k) Then mDic.Add k, nm
End Sub

' Domingo de Páscoa (algoritmo de Meeus/Jones/Butcher)
Private Function Pascoa(ByVal ano As Long) As Date
    Dim a As Long, b As Long, c As Long, d As Long, e As Long, f As Long
    Dim g As Long, h As Long, i As Long, k As Long, l As Long, m As Long
    a = ano Mod 19
    b = ano \ 100
    c = ano Mod 100
    d = b \ 4
    e = b Mod 4
    f = (b + 8) \ 25
    g = (b - f + 1) \ 3
    h = (19 * a + b - d - g + 15) Mod 30
    i = c \ 4
    k = c Mod 4
    l = (32 + 2 * e + 2 * i - h - k) Mod 7
    m = (a + 11 * h + 22 * l) \ 451
    Pascoa = DateSerial(ano, (h + l - 7 * m + 114) \ 31, ((h + l - 7 * m + 114) Mod 31) + 1)
End Function
