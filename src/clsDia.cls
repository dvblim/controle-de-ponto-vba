VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' =====================================================================
'  clsDia - um dia de registro (horários em minutos; -1 = vazio)
' =====================================================================

Public Data As Date
Public Tipo As String
Public Entrada As Long
Public SaidaAlm As Long
Public Retorno As Long
Public Saida As Long
Public Obs As String

' Calculados por Calcular()
Public Completo As Boolean
Public AlmocoMin As Long
Public TrabalhadoMin As Long
Public DiferencaMin As Long
Public BancoMin As Long
Public AcumuladoMin As Long
Public Status As String

Private Sub Class_Initialize()
    Tipo = "Normal"
    Entrada = -1
    SaidaAlm = -1
    Retorno = -1
    Saida = -1
    AlmocoMin = -1
End Sub

Public Sub Calcular(ByVal regra As clsRegra)
    If Tipo = "" Then Tipo = "Normal"
    Completo = (Entrada >= 0 And SaidaAlm >= 0 And Retorno >= 0 And Saida >= 0)

    If SaidaAlm >= 0 And Retorno >= 0 Then
        AlmocoMin = Retorno - SaidaAlm
    Else
        AlmocoMin = -1
    End If

    If Completo Then
        TrabalhadoMin = regra.Trabalhado(Entrada, SaidaAlm, Retorno, Saida)
        DiferencaMin = TrabalhadoMin - regra.JornadaMin
    Else
        TrabalhadoMin = 0
        DiferencaMin = 0
    End If

    BancoMin = regra.BancoDia(Tipo, TrabalhadoMin, Completo)

    Select Case Tipo
        Case "Feriado", "Folga", "Atestado", "Falta"
            Status = Tipo
        Case Else
            If Not Completo Then
                Status = "Em andamento"
            ElseIf BancoMin > 0 Then
                Status = "Positivo"
            ElseIf BancoMin < 0 Then
                Status = "Negativo"
            ElseIf DiferencaMin = 0 Then
                Status = "Completo"
            Else
                Status = "Na tolerância"
            End If
    End Select
End Sub
