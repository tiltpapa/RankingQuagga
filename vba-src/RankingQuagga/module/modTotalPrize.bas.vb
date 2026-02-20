Public Sub UpdatePrizeRanking()
    Dim eventId As String
    eventId = ActivePresentation.Tags("event_id")
    If eventId = "" Then
        MsgBox "イベントIDを設定してください", vbCritical
        Exit Sub
    End If

    Dim totalTop As Integer, perSlide As Integer
    totalTop = Val(InputBox("上位何名の成績を出力しますか？（未入力で全員）", "人数指定"))
    'perSlide = Val(InputBox("1スライドあたりの人数は？（未入力で10名）", "スライド人数"))
    'If perSlide = 0 Then perSlide = 10

    Dim data As Dictionary
    Dim apiEnd As String
    apiEnd = "/programs/" & eventId & "/result/total"
    If Not (ActivePresentation.Tags("server") Like "*quagga.studio*") Then
        apiEnd = apiEnd & IIf(totalTop <> 0, "?top=" & Str(totalTop + 5), "")
    '    perSlide = Val(InputBox("1スライドあたりの人数は？（未入力で10名）", "スライド人数"))
    End If
    If perSlide = 0 Then perSlide = 10
    Set data = GetApiData(apiEnd)
    If data Is Nothing Then Exit Sub

    Dim results As Collection
    Set results = data("results")
    
    Call SortCollectionByMoney(results)
    
    Dim i As Integer
    If totalTop = 0 Or totalTop >= results.Count Then
        totalTop = results.Count
    Else
        For i = results.Count To totalTop + 1 Step -1
            results.Remove i
        Next i
    End If

    Dim baseSlide As slide
    Set baseSlide = ActiveWindow.Selection.SlideRange(1)

    ' --- スライド分割のための人数ロジック ---
    ' 最後のスライドは1～perSlide位のため、それを除いた人数
    Dim remainingCount As Integer
    remainingCount = totalTop - perSlide

    Dim firstSlideCount As Integer
    firstSlideCount = totalTop Mod perSlide
    If firstSlideCount = 0 And totalTop > perSlide Then firstSlideCount = perSlide

    Dim slideCount As Integer
    If totalTop <= perSlide Then
        slideCount = 0
    Else
        ' 必要なスライド数（最後の1～perSlide位スライドは除外）
        slideCount = (totalTop - firstSlideCount) \ perSlide
    End If

    ' 必要枚数だけスライド複製
    'Dim i As Integer
    For i = 1 To slideCount - 1
        baseSlide.Duplicate.MoveTo baseSlide.SlideIndex + i
    Next

    Dim currentSlide As slide
    Set currentSlide = baseSlide

    Dim idxEnd As Integer, idxStart As Integer
    idxEnd = totalTop
    idxStart = totalTop - firstSlideCount + 1

    '最初のスライドへの反映
    If remainingCount > 0 Then
        Call ResetField(currentSlide)
        Call ReflectRankingToSlide(currentSlide, results, idxStart, idxEnd)
        idxEnd = idxStart - 1

        '次以降のスライドへ反映
        For i = 1 To slideCount - 1
            Set currentSlide = ActivePresentation.Slides(baseSlide.SlideIndex + i)
            idxStart = idxEnd - perSlide + 1
            Call ResetField(currentSlide)
            Call ReflectRankingToSlide(currentSlide, results, idxStart, idxEnd)
            idxEnd = idxStart - 1
        Next

        '最後のスライド（既存スライド）
        If baseSlide.SlideIndex + slideCount > ActivePresentation.Slides.Count Then
            MsgBox "1位～" & perSlide & "位用のスライドがありません。", vbCritical
            Exit Sub
        Else
            Set currentSlide = ActivePresentation.Slides(baseSlide.SlideIndex + slideCount)
            Call ResetField(currentSlide)
            Call ReflectRankingToSlide(currentSlide, results, 1, perSlide)
        End If
    Else
        ' 人数がperSlide以下の場合（複製なし）
        Call ResetField(currentSlide)
        Call ReflectRankingToSlide(currentSlide, results, 1, totalTop)
    End If
    
    Call SetShapeTextByZOrder(currentSlide, 1, "番号色変", results(1)("rank"))
    Call SetShapeTextByZOrder(currentSlide, 1, "ポイント色変", results(1)("point"), False)
    Call SetShapeTextByZOrder(currentSlide, 1, "賞金色変", FormatPrize(Val(results(1)("money"))))
    Call SetShapeTextByZOrder(currentSlide, 1, "名前色変", results(1)("member")("user")("name"))
    
    MsgBox "反映しました", Title:="完了"
End Sub

'順位範囲を指定してスライドに反映する関数
Private Sub ReflectRankingToSlide(sld As slide, results As Collection, idxStart As Integer, idxEnd As Integer)
    Dim dispPos As Integer, participantIdx As Integer
    
    dispPos = 1
    For participantIdx = idxStart To idxEnd
        SetShapeText sld, dispPos, results(participantIdx)
        dispPos = dispPos + 1
    Next
End Sub

Private Sub SetShapeText(sld As slide, pos As Integer, score As Dictionary)
    Dim shp As Shape
    Call SetShapeTextByZOrder(sld, pos, "番号", score("rank"))
    Call SetShapeTextByZOrder(sld, pos, "ポイント", score("point"), False)
    Call SetShapeTextByZOrder(sld, pos, "賞金", FormatPrize(Val(score("money"))))
    Call SetShapeTextByZOrder(sld, pos, "名前", score("member")("user")("name"))
End Sub

Private Sub SortCollectionByMoney(ByRef results As Collection)
    Dim i As Long, j As Long
    Dim vi As Dictionary, vj As Dictionary
    For i = 1 To results.Count - 1
        For j = i + 1 To results.Count
            If Val(results(i)("money")) < Val(results(j)("money")) Then
                Set vi = results(i)
                Set vj = results(j)

                results.Remove i
                results.Add vj, , i
                
                results.Remove j
                
                If j > results.Count Then
                    results.Add vi
                Else
                    results.Add vi, , j
                End If
            End If
        Next j
    Next i
    
    For i = 1 To results.Count
        results(i)("rank") = i
    Next i
End Sub

