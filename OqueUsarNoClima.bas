Attribute VB_Name = "M�dulo1"
Sub Competicao()
linha = 3

While Cells(linha, 2) <> ""
   If Cells(linha, 2) = "Sol" Then
      Cells(linha, 3) = "Levar chap�u e protetor"
   Else
      Cells(linha, 3) = "Levar botas e toalha"
   End If
      linha = linha + 1
      
While Cells(linha, 8) <> ""
   If Cells(linha, 8) = "Neblida" Then
      Cells(linha, 3) = "Levar �culos"
   Else
      Cells(linha, 3) = "Levar nada"
   End If
      linha = linha + 1
Wend
Wend
End Sub
