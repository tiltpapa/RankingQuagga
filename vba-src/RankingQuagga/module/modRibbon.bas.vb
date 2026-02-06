Private rankRibbon As IRibbonUI
Dim SE As New ShowEvent

Public Sub rankRibbonOnLoad(ribbon As IRibbonUI)
    Set rankRibbon = ribbon
    rankRibbon.Invalidate
End Sub

Sub ShowSettingsForm(control As IRibbonControl)
    UserForm1.Show
End Sub

'Sub LinkShapeToPeriodMacro(control As IRibbonControl)
'    Dim shp As Shape
'    If ActiveWindow.Selection.Type = ppSelectionShapes Then
'        Set shp = ActiveWindow.Selection.ShapeRange(1)
'        shp.ActionSettings(ppMouseClick).Action = ppActionRunMacro
'        shp.ActionSettings(ppMouseClick).Run = "UpdatePeriodRanking"
'    Else
'        MsgBox "シェイプを選択してください", vbExclamation
'    End If
'End Sub

Sub UpdateTopRanking(control As IRibbonControl)
    Call modLastAggr.UpdateLastAggrRanking(top:=True)
End Sub

Sub UpdateBottomRanking(control As IRibbonControl)
    Call modLastAggr.UpdateLastAggrRanking(top:=False)
End Sub

Sub UpdatePeriodRanking(control As IRibbonControl)
    Call modPeriodResult.UpdatePeriodRanking ' 既に定義済みのマクロを呼び出す
End Sub

Sub UpdateTotalRanking(control As IRibbonControl)
    Call modTotalResult.UpdateTotalRanking ' 既に定義済みのマクロを呼び出す
End Sub

Sub UpdatePrizeRanking(control As IRibbonControl)
    Call modTotalPrize.UpdatePrizeRanking ' 既に定義済みのマクロを呼び出す
End Sub

Sub AddTrigShape(control As IRibbonControl)
    UserForm2.Show
End Sub

Sub InitializeApp(control As IRibbonControl)
    Set SE.App = Application
End Sub

Sub FinalizeApp(control As IRibbonControl)
    Set SE.App = Nothing
End Sub
