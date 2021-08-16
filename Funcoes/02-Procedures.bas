
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