VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsRegra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' =====================================================================
'  clsRegra - regras do banco de horas (lidas da aba Config)
'  Jornada 8:48 trabalhadas, almoço não conta (horário livre),
'  tolerância de 10 min no total do dia, para mais ou para menos.
'  Passou da tolerância (11 min ou mais), entra a diferença inteira:
'  os 10 minutos + o que passou deles.
' =====================================================================

Private mJornada As Long
Private mTol As Long
Private mAlmoco As Long
Private mSaldoIni As Long

Private Sub Class_Initialize()
    mJornada = 528
    mTol = 10
    mAlmoco = 60
    mSaldoIni = 0
End Sub

Public Sub Carregar(ByVal wb As Workbook)
    On Error Resume Next
    mJornada = MinutosDe(wb.Names("Jornada").RefersToRange.Value, mJornada)
    mTol = CLng(wb.Names("TolMin").RefersToRange.Value)
    mAlmoco = MinutosDe(wb.Names("AlmocoPadrao").RefersToRange.Value, mAlmoco)
    mSaldoIni = CLng(wb.Names("SaldoInicialMin").RefersToRange.Value)
    On Error GoTo 0
    If mJornada <= 0 Then mJornada = 528
    If mTol < 0 Then mTol = 0
End Sub

' Aceita hora (fração do dia, ex.: 08:48) ou número de minutos
Private Function MinutosDe(ByVal v As Variant, ByVal padrao As Long) As Long
    MinutosDe = padrao
    If IsError(v) Or IsEmpty(v) Then Exit Function
    If Not IsNumeric(v) Then Exit Function
    If CDbl(v) < 1 Then
        MinutosDe = CLng(Round(CDbl(v) * 1440, 0))
    Else
        MinutosDe = CLng(v)
    End If
End Function

Public Property Get JornadaMin() As Long
    JornadaMin = mJornada
End Property

Public Property Get TolMin() As Long
    TolMin = mTol
End Property

Public Property Get AlmocoPadraoMin() As Long
    AlmocoPadraoMin = mAlmoco
End Property

Public Property Get SaldoInicialMin() As Long
    SaldoInicialMin = mSaldoIni
End Property

' Minutos trabalhados no dia (almoço descontado, qualquer duração)
Public Function Trabalhado(ByVal entrada As Long, ByVal saidaAlm As Long, _
                           ByVal retorno As Long, ByVal saida As Long) As Long
    Trabalhado = (saidaAlm - entrada) + (saida - retorno)
End Function

' Valor que o dia lança no banco
Public Function BancoDia(ByVal tipo As String, ByVal trabalhadoMin As Long, ByVal completo As Boolean) As Long
    Dim dif As Long
    Select Case tipo
        Case "Falta"
            BancoDia = -mJornada
        Case "Feriado", "Folga", "Atestado"
            BancoDia = 0
        Case Else
            If Not completo Then
                BancoDia = 0
            Else
                dif = trabalhadoMin - mJornada
                If Abs(dif) <= mTol Then
                    BancoDia = 0          ' dentro da tolerância
                Else
                    BancoDia = dif        ' passou: entra tudo (10 + excedente)
                End If
            End If
    End Select
End Function

' Saída que fecha exatamente a jornada
Public Function SaidaSugerida(ByVal entrada As Long, ByVal almocoMin As Long) As Long
    SaidaSugerida = entrada + mJornada + almocoMin
End Function
