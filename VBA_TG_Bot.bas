Attribute VB_Name = "VBA_TG_Bot"
Sub Tes()
    KirimPesanTelegram "OK"
End Sub
Sub KirimPesanTelegram(ByRef Pesan As String)
    Dim Token As String, ChatID As String
    Dim sURL As String, oHttp As Object, sHTML As String
    
    Token = "TOKENBOT" 'Token Bot dari @BotFather
    ChatID = "IDUSER" 'ID User Telegram yang ingin di kirim pesan
    sURL = "https://api.telegram.org/bot" & Token & "/sendMessage?chat_id=" & ChatID & "&text=" & Pesan
    
    Set oHttp = CreateObject("Msxml2.XMLHTTP")
    oHttp.Open "POST", sURL, False
    oHttp.Send

    sHTML = oHttp.ResponseText

    Debug.Print sHTML
End Sub

