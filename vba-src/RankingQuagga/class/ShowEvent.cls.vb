Public WithEvents App As Application
Attribute App.VB_VarHelpID = -1

Private Sub App_SlideShowNextSlide(ByVal Wn As SlideShowWindow)
    Dim sld As slide, col As New Collection
    Set sld = Wn.View.slide
    
    Call CollectMatchingShapes(sld.shapes.Range, "自動_*", col)

    If col.Count = 0 Then
        Exit Sub
    End If
    
    Dim shp As Shape
    For Each shp In col
        Select Case shp.name
            Case "自動_早押しトップ10"
                Call modLastAggr.UpdateLastAggrRanking(sld, True)
            Case "自動_予選落ち"
                Call modLastAggr.UpdateLastAggrRanking(sld, False)
            Case "自動_ピリオド成績"
                Call modPeriodResult.UpdatePeriodRanking(sld)
        End Select
    Next
End Sub