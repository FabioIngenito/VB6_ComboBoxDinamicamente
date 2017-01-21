Attribute VB_Name = "mdlCombo"
Option Explicit

'Public Sub CriaCombos(frm As Form, NroCombos As Integer)
'Dim X As Byte
'Dim Y As Integer
'Dim intAcrescenta As Integer
'Dim intColuna As Byte
'
'    intAcrescenta = cboC(0).Height
'    intColuna = 5
'
'    With frm
'
'        For X = 1 To NroCombos
'            'Parte da "física" das Combos (desenho das Combos no form)
'            Load .cboC(X)
'            .cboC(X).Tag = 1
'            .cboC(X).Visible = True
'
'            PosicionaCombo X, Y, intColuna, Me
'
'            Y = Y + intAcrescenta
'
'            If X >= intColuna Then intColuna = intColuna + 5
'        Next
'
'    End With
'
'End Sub
'
'Public Sub PosicionaCombo(X As Byte, Y As Integer, cols As Byte, frm As Form)
'
'    With frm
'        .cboC(X).Top = .cboC(X - 1).Top
'        .cboC(X).Left = (.cboC(0).Height + .cboC(X - 1).Left) + 1200
'
'        If X >= cols Then
'            .cboC(X).Top = (.cboC(0).Width + .cboC(X - 1).Top) - 1100
'            .cboC(X).Left = .cboC(0).Left
'        End If
'
'    End With
'
'End Sub
'
'Public Sub PreencheCombos(frm As Form, NroCombos As Integer)
'Dim bytConta As Byte
'
'    For bytConta = 0 To NroCombos
'
'        With frm
'            .cboC(bytConta).AddItem "Item: " & bytConta
'            .cboC(bytConta).AddItem "Ou outra opção!"
'            .cboC(bytConta).ListIndex = 0
'        End With
'
'    Next
'
'End Sub
