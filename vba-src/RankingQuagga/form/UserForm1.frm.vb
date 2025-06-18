Private Sub CommandButton1_Click()
    ActivePresentation.Tags.Add "server", IIf(optQuagga.Value, "https://quagga.studio/api/v1", "https://ranking-quagga.tiltpapa.workers.dev")
    ActivePresentation.Tags.Add "token", txtToken.Text
    ActivePresentation.Tags.Add "event_id", txtEventID.Text
    MsgBox "設定を保存しました"
    Unload Me
End Sub

Private Sub UserForm_Initialize()
    ' サーバー設定の表示
    Dim serverUrl As String
    serverUrl = ActivePresentation.Tags("server")
    If serverUrl = "" Or serverUrl Like "*quagga.studio*" Then
        optQuagga.Value = True
    Else
        optTestServer.Value = True
    End If
    
    ' 秘密の暗号を表示
    txtToken.Text = ActivePresentation.Tags("token")
    
    ' イベントIDを表示
    txtEventID.Text = ActivePresentation.Tags("event_id")
End Sub