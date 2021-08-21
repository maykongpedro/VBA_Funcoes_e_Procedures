Attribute VB_Name = "modTools"

Option Explicit

' Procedure : TurnOffFunctionality
' Source    : www.ExcelMacroMastery.com
' Author    : Paul Kelly
' Purpose   : Turn off automatic calculations, events and screen updating
' https://excelmacromastery.com/
Public Sub TurnOffFunctionality()
    Application.Calculation = xlCalculationManual
    Application.DisplayStatusBar = False
    Application.EnableEvents = False
    Application.ScreenUpdating = False
End Sub

' Procedure : TurnOnFunctionality
' Source    : www.ExcelMacroMastery.com
' Author    : Paul Kelly
' Purpose   : turn on automatic calculations, events and screen updating
' https://excelmacromastery.com/
Public Sub TurnOnFunctionality()
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayStatusBar = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub


' Procedure : CopiarDados
' Source    : https://ferramentasexcelvba.wordpress.com/
' Author    : Arnaldo Gunzi
' Purpose   : Copy data from a range to an array inside the vba memory
' https://ferramentasexcelvba.wordpress.com/
Sub CopiarDados(linini As Integer, colini As Integer, ncols As Long, ByRef varRef As Variant, Optional nomeSht As String = "", Optional maxLin As Long = 10 ^ 6)
'Copia dados da planilha nomeSht, a comecas da linIni e colIni, para varRef
    Dim nl As Long, nc As Long
    If nomeSht <> "" Then
            Sheets(nomeSht).Activate
    End If
    
    nl = Cells(Rows.count, colini).End(xlUp).Row
    nc = ncols
    If nl > 0 And nc > 0 Then
        varRef = Range(Cells(linini, colini), Cells(nl, colini + nc - 1))
    End If
End Sub

' Procedure : ColarDados
' Source    : https://ferramentasexcelvba.wordpress.com/
' Author    : Arnaldo Gunzi
' Purpose   : Paste data from an array to a range
' https://ferramentasexcelvba.wordpress.com/
Sub ColarDados(linini As Integer, colini As Integer, ncols As Long, ByRef varRef As Variant, Optional nomeSht As String = "", Optional maxLin As Long = 10 ^ 6)
    If nomeSht <> "" Then
            Sheets(nomeSht).Activate
    End If
    
    Range(Cells(linini, colini), Cells(linini + 500000, colini + ncols - 1)).ClearContents
    Range(Cells(linini, colini), Cells(linini, colini)).Resize(UBound(varRef, 1), ncols) = varRef
End Sub

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
      RemoveDuplicates = .keys
    End With
End Function

' Procedure : VisualizarPlanilha
' Source    : 
' Author    : Maykon G. Pedro
' Purpose   : Hide all sheets but turn visible the sheeet 'PlanName'
' Utils to organize menu and panels in a workbook
Public Sub VisualizarPlanilha(ByVal PlanName As String)
    
    On Error GoTo ErrorHandler
    Dim i As Integer, planilha As Worksheet, nome_planilha As String
    Worksheets(PlanName).Visible = True
    For i = 1 To Sheets.Count
        Set planilha = Sheets(i)
        nome_planilha = planilha.Name
        If nome_planilha <> PlanName Then
            planilha.Visible = False
        End If
    Next i
    
    Exit Sub
    
ErrorHandler:
    If Err.Number = 9 Then
        MsgBox "Planilha não encontrada!", Title:="Aviso!"
    Else
        MsgBox "Erro encontrado, tipo: " & Err.Description, Title:="Aviso!"
    End If
End Sub