
' Procedure : RemoveDuplicates
' Source    : https://ferramentasexcelvba.wordpress.com/
' Author    : Arnaldo Gunzi
' Purpose   : Remove duplicates from a unique column array
' https://ferramentasexcelvba.wordpress.com/
Function RemoveDuplicates(ByVal varArray As Variant)
    ' \\ Declaração de variáveis
    Dim varValue As Variant
    
    ' \\ Cria o objeto dictionary
    With CreateObject("scripting.dictionary")
      .CompareMode = vbTextCompare ' \\ Compara texto
      For Each varValue In varArray '\\ Para cada valor na matriz
       If Not Strings.Len(varValue) = 0 And Not .exists(varValue) Then '\\ Desconsidera valores vazios, alterar esta linha caso queira considerar
          .Add varValue, Nothing
        End If
      Next
      remove_duplicate = .keys
    End With
End Function