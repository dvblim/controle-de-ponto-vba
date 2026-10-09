VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCampo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' =====================================================================
'  clsCampo - eventos dos campos de lançamento
'  tipo "hora": só dígitos, insere ":" sozinho (0814 -> 08:14)
'  tipo "data": só dígitos, insere "/" sozinho
'  Enter em qualquer campo salva o dia.
' =====================================================================

Private WithEvents tb As MSForms.TextBox
Private WithEvents cb As MSForms.ComboBox
Private mPai As Object
Private mTipo As String

Public Sub InitTexto(ByVal pai As Object, ByVal t As MSForms.TextBox, ByVal tipo As String)
    Set mPai = pai
    Set tb = t
    mTipo = tipo
End Sub

Public Sub InitCombo(ByVal pai As Object, ByVal c As MSForms.ComboBox)
    Set mPai = pai
    Set cb = c
    mTipo = "combo"
End Sub

Private Sub tb_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Dim sepr As String, k As Long
    k = KeyAscii
    Select Case mTipo
        Case "hora": sepr = ":"
        Case "data": sepr = "/"
        Case Else: Exit Sub
    End Select

    If k = Asc(sepr) Then Exit Sub
    If k < 48 Or k > 57 Then
        KeyAscii = 0
        Exit Sub
    End If

    ' digitando no fim do campo: insere o separador automaticamente
    If tb.SelLength = 0 And tb.SelStart = Len(tb.Text) Then
        If mTipo = "hora" Then
            If Len(tb.Text) = 2 And InStr(tb.Text, ":") = 0 Then
                tb.Text = tb.Text & ":"
                tb.SelStart = Len(tb.Text)
            End If
        Else
            If Len(tb.Text) = 2 Or Len(tb.Text) = 5 Then
                tb.Text = tb.Text & "/"
                tb.SelStart = Len(tb.Text)
            End If
        End If
    End If
End Sub

Private Sub tb_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then
        KeyCode = 0
        mPai.Acao "salvar"
    End If
End Sub

Private Sub tb_Change()
    If mTipo = "data" Then
        mPai.DataMudou
    Else
        mPai.CamposMudaram
    End If
End Sub

Private Sub cb_Change()
    mPai.CamposMudaram
End Sub
