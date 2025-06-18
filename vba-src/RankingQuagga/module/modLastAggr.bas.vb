Public Sub UpdateLastAggrRanking(Optional slide As slide, Optional top As Boolean = True)
    Dim eventId As String
    eventId = ActivePresentation.Tags("event_id")
    If eventId = "" Then
        MsgBox "イベントIDを設定してください", vbCritical
        Exit Sub
    End If

'    If top Is Nothing Then
'        top = IIf(MsgBox("早押しトップ10を表示しますか？" & vbCrLf & "Noなら予選落ち", vbYesNo) = vbYes, True, False)
'    End If
    
    Dim data As Dictionary
    Set data = GetApiData("/programs/" & eventId & "/questions/last_aggregate")
    If data Is Nothing Then Exit Sub

    Dim results As Collection, i As Integer
    Set results = data("answers")
    For i = results.Count To 1 Step -1
        If results(i)("correct") = False Then results.Remove i
    Next
    
    If slide Is Nothing Then Set slide = ActiveWindow.Selection.SlideRange(1)
    Call ResetField(slide)

    Dim idxEnd As Integer, idxStart As Integer
    If top Then
        idxStart = 1
        idxEnd = Min(results.Count, 10)
    Else
        idxStart = results.Count
        idxEnd = Max(idxStart - 9, 1)
    End If
    
    Dim shp As Shape, pos As Integer
    pos = 10 + 1
    For i = idxStart To idxEnd Step IIf(top, 1, -1)
        pos = IIf(top, i, pos - 1)
        Call SetShapeTextByZOrder(slide, pos, "番号", i)
        Call SetShapeTextByZOrder(slide, pos, "タイム", Format(Val(results(i)("time")), "0.00"))
        Call SetShapeTextByZOrder(slide, pos, "名前", results(i)("member")("user")("name"))
    Next
    
    'pos = IIf(top, 1, 10)
    Call SetShapeTextByZOrder(slide, 1, "番号色変", idxStart)
    Call SetShapeTextByZOrder(slide, 1, "タイム色変", Format(Val(results(idxStart)("time")), "0.00"))
    Call SetShapeTextByZOrder(slide, 1, "名前色変", results(idxStart)("member")("user")("name"))
    
    MsgBox "反映しました", Title:="完了"
End Sub