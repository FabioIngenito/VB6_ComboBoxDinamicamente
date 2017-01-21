VERSION 5.00
Begin VB.Form frmCombo 
   Caption         =   "Cria Combos Dinamicamente"
   ClientHeight    =   5790
   ClientLeft      =   1050
   ClientTop       =   2025
   ClientWidth     =   11070
   LinkTopic       =   "Form1"
   ScaleHeight     =   5790
   ScaleWidth      =   11070
   Begin VB.TextBox txtColunas 
      Height          =   315
      Left            =   2700
      TabIndex        =   5
      Text            =   "5"
      Top             =   120
      Width           =   435
   End
   Begin VB.CommandButton cmRetirarTudo 
      Caption         =   "Retirar Tudo!"
      Enabled         =   0   'False
      Height          =   315
      Left            =   4200
      TabIndex        =   4
      Top             =   120
      Width           =   1275
   End
   Begin VB.CommandButton cmdCriar 
      Caption         =   "Criar!"
      Height          =   315
      Left            =   3300
      TabIndex        =   2
      Top             =   120
      Width           =   795
   End
   Begin VB.TextBox txtNumeroCombos 
      Height          =   315
      Left            =   1380
      TabIndex        =   1
      Text            =   "32"
      Top             =   120
      Width           =   435
   End
   Begin VB.ComboBox cboCombo 
      Height          =   315
      Index           =   0
      Left            =   420
      TabIndex        =   3
      Text            =   "COMBO REAL"
      Top             =   660
      Width           =   1500
   End
   Begin VB.Label lblExplicacao 
      Caption         =   $"frmCombo.frx":0000
      Height          =   435
      Left            =   5580
      TabIndex        =   7
      Top             =   120
      Width           =   5355
   End
   Begin VB.Label lblColunas 
      Caption         =   "Colunas:"
      Height          =   195
      Left            =   1980
      TabIndex        =   6
      Top             =   180
      Width           =   615
   End
   Begin VB.Label lblCriarCombos 
      Caption         =   "Criar mais:"
      Height          =   195
      Left            =   480
      TabIndex        =   0
      Top             =   180
      Width           =   795
   End
End
Attribute VB_Name = "frmCombo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim clsC As New clsCombo

Private Sub txtColunas_KeyPress(KeyAscii As Integer)
    If InStr("0123456789", Chr(KeyAscii)) = 0 And KeyAscii <> 8 Then KeyAscii = 0
End Sub

Private Sub txtNumeroCombos_KeyPress(KeyAscii As Integer)
    If InStr("0123456789", Chr(KeyAscii)) = 0 And KeyAscii <> 8 Then KeyAscii = 0
End Sub

Private Sub cmdCriar_Click()
    cmdCriar.Enabled = False
    cmRetirarTudo.Enabled = True
    cboCombo(0).Visible = True
      
    clsC.CriaCombos frmCombo, Val(txtNumeroCombos.Text), Val(txtColunas.Text)
    'CriaCombos Me, txtNumeroCombos.Text
    PreencheCombos Me, txtNumeroCombos.Text
End Sub

Private Sub cmRetirarTudo_Click()
    cmdCriar.Enabled = True
    cmRetirarTudo.Enabled = False
    RetiraTodasCombos Me
End Sub

Public Sub CriaCombos(frm As Form, NroCombos As Long)
Dim X As Long
Dim Y As Long
Dim intAcrescenta As Integer
Dim intColuna As Byte
Dim intPosColuna As Byte

    intAcrescenta = cboCombo(0).Height
    
    If txtColunas.Text = "" Or txtColunas.Text < 1 Then txtColunas.Text = 1
    
    intColuna = txtColunas.Text
    intPosColuna = intColuna
    frm.cboCombo(0).Visible = True
        
    With frm

        For X = 1 To NroCombos
            'Parte da "física" das Combos (desenho das Combos no form)
            'Part of the "physics" of Combos (drawing of Combos in the form)
            Load .cboCombo(X)
            .cboCombo(X).Tag = 1
            .cboCombo(X).Visible = True

            PosicionaCombo frm, X, Y, intColuna

            Y = Y + intAcrescenta

            'If x >= intColuna Then intPosColuna = intColuna
        Next

    End With

End Sub

Public Sub PosicionaCombo(frm As Form, X As Long, Y As Long, cols As Byte)
Dim ContaLinha As Integer
    
    ContaLinha = 0

    Do
        ContaLinha = ContaLinha + 1
    Loop While ContaLinha < (X / cols)

    With frm
        .cboCombo(X).Top = .cboCombo(X - 1).Top
        .cboCombo(X).Left = (.cboCombo(0).Height + .cboCombo(X - 1).Left) + 1200

        If ContaLinha = X / cols Then
            .cboCombo(X).Top = (.cboCombo(0).Width + .cboCombo(X - 1).Top) - 1100
            .cboCombo(X).Left = .cboCombo(0).Left
        End If

    End With

End Sub

Public Sub PreencheCombos(frm As Form, NroCombos As Long)
Dim lngConta As Long

    For lngConta = 0 To NroCombos

        With frm
            .cboCombo(lngConta).AddItem "Item: " & lngConta
            .cboCombo(lngConta).AddItem "Ou outra opção!"
            .cboCombo(lngConta).ListIndex = 0
        End With

    Next

End Sub

Private Sub RetiraTodasCombos(frm As Form)
Dim X As Long
Dim NroCombos As Integer
Dim MyControl As Control

        X = 0
        frm.cboCombo(0).Visible = False
        frm.cboCombo(0).Clear

        For Each MyControl In frm
        
            If TypeOf MyControl Is ComboBox Then X = X + 1
        
        Next
        
        For NroCombos = 1 To X - 1
            Unload frm.cboCombo(NroCombos)
        Next

End Sub
