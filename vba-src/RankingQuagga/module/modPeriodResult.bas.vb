Public Sub UpdatePeriodRanking(Optional slide As slide)
    Dim eventId As String
    eventId = ActivePresentation.Tags("event_id")
    If eventId = "" Then
        MsgBox "イベントIDを設定してください", vbCritical
        Exit Sub
    End If

    Dim data As Dictionary
    Set data = GetApiData("/programs/" & eventId & "/result/period")
    If data Is Nothing Then Exit Sub

    Dim results As Collection
    Set results = data("results")
    If slide Is Nothing Then Set slide = ActiveWindow.Selection.SlideRange(1)
    Call ResetField(slide)

    Dim shp As Shape, i As Integer
    For i = 1 To Min(results.Count, 10)
        Call SetShapeTextByZOrder(slide, i, "番号", results(i)("rank"))
        Call SetShapeTextByZOrder(slide, i, "ポイント", results(i)("point"))
        Call SetShapeTextByZOrder(slide, i, "タイム", Format(Val(results(i)("time")), "0.00"))
        Call SetShapeTextByZOrder(slide, i, "名前", results(i)("member")("user")("name"))
    Next
    
    Call SetShapeTextByZOrder(slide, 1, "番号色変", results(1)("rank"))
    Call SetShapeTextByZOrder(slide, 1, "ポイント色変", results(1)("point"))
    Call SetShapeTextByZOrder(slide, 1, "タイム色変", Format(Val(results(1)("time")), "0.00"))
    Call SetShapeTextByZOrder(slide, 1, "名前色変", results(1)("member")("user")("name"))
    
    MsgBox "反映しました", Title:="完了"
End Sub

