Public Function GetApiData(endpoint As String) As Dictionary
    Dim xhr As Object
    Set xhr = CreateObject("MSXML2.XMLHTTP.6.0")
    
    Dim baseUrl As String, token As String
    baseUrl = ActivePresentation.Tags("server")
    token = ActivePresentation.Tags("token")
    
    If baseUrl = "" Then
        MsgBox "サーバーを設定してください", vbCritical
        Exit Function
    End If
    
    If baseUrl Like "*quagga.studio*" And token = "" Then
        MsgBox "秘密の暗号を設定してください", vbCritical
        Exit Function
    End If
    
    Dim fullUrl As String
    fullUrl = baseUrl & endpoint
    
    On Error GoTo ConnectionErr
    xhr.Open "GET", fullUrl, False
    If baseUrl Like "*quagga.studio*" Then
        xhr.setRequestHeader "Authorization", "Bearer " & token
    End If
    xhr.setRequestHeader "Content-Type", "application/json"
    xhr.setRequestHeader "Pragma", "no-cache"
    xhr.setRequestHeader "Cache-Control", "no-cache"
    xhr.setRequestHeader "If-Modified-Since", "Thu, 01 Jun 1970 00:00:00 GMT"
    xhr.Send
    
    If xhr.Status <> 200 Then
        MsgBox "成績が取得できません（エラーコード：" & xhr.Status & "）", vbCritical
        Exit Function
    End If
    
    Debug.Print xhr.ResponseText
    Set GetApiData = JsonConverter.ParseJson(xhr.ResponseText)
    Exit Function
    
ConnectionErr:
    MsgBox "サーバーに接続できません", vbCritical
End Function

Public Function FormatTime(sec As Double) As String
    Dim mins As Long, secs As Double
    mins = Int(sec / 60)
    secs = sec - mins * 60
    FormatTime = mins & ":" & Format(secs, "00.00")
End Function

Public Sub SetShapeTextByZOrder(sld As slide, pos As Integer, name As String, txt As Variant, Optional critical As Boolean = True)
    Dim shp As Shape
    Set shp = GetShapeByZOrder(sld, pos, name, critical): If shp Is Nothing Then Exit Sub Else shp.TextFrame.TextRange.Text = txt
End Sub

Private Function GetShapeByZOrder(sld As slide, index As Integer, keyword As String, Optional critical As Boolean = True) As Shape
    'Dim sld As Slide
    'Set sld = ActiveWindow.Selection.SlideRange(1)

    Dim colShapes As New Collection
    Call CollectMatchingShapes(sld.shapes.Range, keyword, colShapes)

    If colShapes.Count = 0 Then
        If critical Then MsgBox keyword & "のオブジェクトがありません", vbCritical
        Exit Function
    End If

    ' コレクション内をTopプロパティで並べ替え
    Dim sortedShapes As Variant
    sortedShapes = SortShapesByTop(colShapes)

    If index > UBound(sortedShapes) + 1 Then
        If critical Then MsgBox keyword & "のオブジェクトが足りません", vbCritical
        Exit Function
    End If

    Set GetShapeByZOrder = sortedShapes(index - 1)
End Function

Public Sub CollectMatchingShapes(shapes As ShapeRange, keyword As String, ByRef col As Collection)
    Dim shp As Shape
    Dim i As Long
    'Dim childShp As Shape
    
    For Each shp In shapes
        If shp.Visible Then
            If shp.Type = msoGroup Then
                For i = 1 To shp.GroupItems.Count
                    Call CollectMatchingShapes(shp.GroupItems.Range(i), keyword, col)
                Next i
            ElseIf shp.name Like keyword Then
                col.Add shp
            End If
        End If
    Next
End Sub

Public Function SortShapesByTop(colShapes As Collection) As Shape()
    Dim shpArray() As Shape
    ReDim shpArray(colShapes.Count - 1)

    Dim i As Long, j As Long
    For i = 1 To colShapes.Count
        Set shpArray(i - 1) = colShapes(i)
    Next i

    Dim temp As Shape
    For i = LBound(shpArray) To UBound(shpArray) - 1
        For j = i + 1 To UBound(shpArray)
            If shpArray(i).top > shpArray(j).top Then
                Set temp = shpArray(i)
                Set shpArray(i) = shpArray(j)
                Set shpArray(j) = temp
            End If
        Next j
    Next i

    SortShapesByTop = shpArray()
End Function

Public Sub ResetField(sld As slide)
    Dim pos As Integer
    For pos = 1 To 10
        Call SetShapeTextByZOrder(sld, pos, "番号", "", False)
        Call SetShapeTextByZOrder(sld, pos, "ポイント", "", False)
        Call SetShapeTextByZOrder(sld, pos, "タイム", "", False)
        Call SetShapeTextByZOrder(sld, pos, "名前", "", False)
    Next
    
    Call SetShapeTextByZOrder(sld, 1, "番号色変", "", False)
    Call SetShapeTextByZOrder(sld, 1, "ポイント色変", "", False)
    Call SetShapeTextByZOrder(sld, 1, "タイム色変", "", False)
    Call SetShapeTextByZOrder(sld, 1, "名前色変", "", False)
End Sub

' 独自Max関数
Public Function Max(a As Long, b As Long) As Long
    If a > b Then Max = a Else Max = b
End Function

' 独自Min関数
Public Function Min(a As Long, b As Long) As Long
    If a < b Then Min = a Else Min = b
End Function

Public Sub AddNameShape(name As String)
    Dim sld As slide
    If ActiveWindow.Selection.Type = ppSelectionSlides Then
        Set sld = ActiveWindow.Selection.SlideRange(1)
        With sld.shapes.AddShape(msoShapeRectangle, 0, -55, 100, 50)
            .name = name
            .TextFrame.TextRange.Font.Size = 10
            .TextFrame.TextRange.Text = name
            .Select
        End With
    Else
        MsgBox "スライドを選択してください", vbExclamation
    End If
End Sub