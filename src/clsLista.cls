VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsLista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' =====================================================================
'  clsLista - lista de dias; clicar num dia carrega no lançamento
' =====================================================================

Private WithEvents lst As MSForms.ListBox
Private mPai As Object

Public Sub Init(ByVal pai As Object, ByVal l As MSForms.ListBox)
    Set mPai = pai
    Set lst = l
End Sub

Private Sub lst_Click()
    mPai.DiaSelecionado lst.ListIndex
End Sub
