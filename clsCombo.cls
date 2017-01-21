VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCombo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub CriaCombos(frm As Form, NroCombos As Long, NroColunas As Long)
Dim lngX As Long
Dim lngY As Long
Dim lngAcrescenta As Long
Dim intColuna As Integer

    lngAcrescenta = frm.cboCombo(0).Height
    intColuna = NroColunas

    With frm

        For lngX = 1 To NroCombos
            'Parte da "física" das Combos (desenho das Combos no form)
            'Part of the "physics" of Combos (drawing of Combos in the form)
            Load .cboCombo(lngX)
            .cboCombo(lngX).Tag = 1
            .cboCombo(lngX).Visible = True

            PosicionaCombo lngX, lngY, intColuna, frmCombo

            lngY = lngY + lngAcrescenta

            If lngX >= intColuna Then intColuna = intColuna + NroColunas
        Next

    End With

End Sub

Public Sub PosicionaCombo(X As Long, Y As Long, cols As Integer, frm As Form)

    With frm
        .cboCombo(X).Top = .cboCombo(X - 1).Top
        .cboCombo(X).Left = (.cboCombo(0).Height + .cboCombo(X - 1).Left) + 1200

        If X >= cols Then
            .cboCombo(X).Top = (.cboCombo(0).Width + .cboCombo(X - 1).Top) - 1100
            .cboCombo(X).Left = .cboCombo(0).Left
        End If

    End With

End Sub

Public Sub PreencheCombos(frm As Form, NroCombos As Integer)
Dim lngConta As Long

    For lngConta = 0 To NroCombos

        With frm
            .cboCombo(lngConta).AddItem "Item: " & lngConta
            .cboCombo(lngConta).AddItem "Ou outra opção!"
            .cboCombo(lngConta).ListIndex = 0
        End With

    Next

End Sub
