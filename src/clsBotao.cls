VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsBotao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' =====================================================================
'  clsBotao - botão plano (dois Labels) com evento de clique
'  O clique chama frmPainel.Acao <chave>.
' =====================================================================

Private WithEvents lbFundo As MSForms.Label
Private WithEvents lbTexto As MSForms.Label
Private mPai As Object
Private mChave As String

Public Sub Criar(ByVal pai As Object, ByVal ctls As Object, ByVal chv As String, ByVal txt As String, _
                 ByVal x As Single, ByVal y As Single, ByVal w As Single, ByVal h As Single, _
                 Optional ByVal corFundo As Long = COR_BOTAO, Optional ByVal corTexto As Long = COR_TEXTO, _
                 Optional ByVal corBorda As Long = COR_BORDA, Optional ByVal tam As Single = 9)
    Dim th As Single
    Set mPai = pai
    mChave = chv
    th = AlturaTexto(tam)
    Set lbFundo = NovoCartao(ctls, x, y, w, h, corBorda, corFundo)
    Set lbTexto = NovoLabel(ctls, txt, x + 1, y + (h - th) / 2, w - 2, th, tam, False, corTexto, corFundo, fmTextAlignCenter)
End Sub

Public Property Let Texto(ByVal s As String)
    lbTexto.Caption = s
End Property

Public Property Get Chave() As String
    Chave = mChave
End Property

Private Sub lbFundo_Click()
    mPai.Acao mChave
End Sub

Private Sub lbTexto_Click()
    mPai.Acao mChave
End Sub
