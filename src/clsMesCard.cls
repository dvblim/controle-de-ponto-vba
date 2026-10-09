VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMesCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' =====================================================================
'  clsMesCard - cartão de mês clicável (marca/desmarca o mês)
' =====================================================================

Private WithEvents lbFundo As MSForms.Label
Private WithEvents lbNome As MSForms.Label
Private WithEvents lbValor As MSForms.Label
Private WithEvents lbCheck As MSForms.Label
Private mPai As Object
Private mMes As Long

Public Sub Criar(ByVal pai As Object, ByVal ctls As Object, ByVal numMes As Long, _
                 ByVal x As Single, ByVal y As Single, ByVal w As Single, ByVal h As Single)
    Set mPai = pai
    mMes = numMes
    Set lbFundo = NovoCartao(ctls, x, y, w, h, COR_BORDA, COR_CARD)
    Set lbNome = NovoLabel(ctls, "", x + 6, y + 4, w - 20, 15, 9.5, True, COR_TEXTO, COR_CARD)
    Set lbCheck = NovoLabel(ctls, ChrW(&H2713), x + w - 14, y + 4, 11, 15, 9, True, COR_AZUL, COR_CARD, fmTextAlignCenter)
    Set lbValor = NovoLabel(ctls, "", x + 6, y + 19, w - 10, 13, 8, False, COR_MUTED, COR_CARD)
End Sub

Public Property Get Mes() As Long
    Mes = mMes
End Property

Public Sub Mostrar(ByVal nome As String, ByVal valor As String, ByVal corValor As Long, ByVal marcado As Boolean)
    lbNome.Caption = nome
    lbValor.Caption = valor
    lbValor.ForeColor = corValor
    lbCheck.Visible = marcado
    If marcado Then
        lbFundo.BorderColor = COR_AZUL_BORDA
    Else
        lbFundo.BorderColor = COR_BORDA
    End If
End Sub

Private Sub Clicar()
    mPai.AlternarMes mMes
End Sub

' Clique simples e duplo alternam (dois cliques rápidos = duas alternâncias)
Private Sub lbFundo_Click()
    Clicar
End Sub

Private Sub lbNome_Click()
    Clicar
End Sub

Private Sub lbValor_Click()
    Clicar
End Sub

Private Sub lbCheck_Click()
    Clicar
End Sub

Private Sub lbFundo_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Clicar
End Sub

Private Sub lbNome_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Clicar
End Sub

Private Sub lbValor_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Clicar
End Sub

Private Sub lbCheck_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Clicar
End Sub

