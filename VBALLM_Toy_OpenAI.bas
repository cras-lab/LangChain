Sub CallOpenAI()

    Dim apiKey As String
    Dim http As Object
    Dim url As String
    Dim prompt As String
    Dim body As String

    apiKey = InputBox("OpenAI API Key를 입력하세요")
    If apiKey = "" Then Exit Sub

    prompt = Range("A1").Value

    url = "https://api.openai.com/v1/responses"

    body = "{""model"":""gpt-4.1-mini"",""input"":""" & prompt & """}"

    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")

    http.Open "POST", url, False
    http.SetRequestHeader "Content-Type", "application/json"
    http.SetRequestHeader "Authorization", "Bearer " & apiKey

    http.Send body

    Dim json As Object
    
    Set json = JsonConverter.ParseJson(http.ResponseText)

    Range("B1") = json("output")(1)("content")(1)("text")

End Sub
