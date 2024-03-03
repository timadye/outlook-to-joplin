Public Sub SendToJoplin()
    Dim sToken As String, sURL As String
    Dim sURLNotes, sURLResources, sEscapedBody, sJSONString, sFolderID As String
    Dim objItem As Outlook.MailItem
    
    sToken = "REPLACE ME WITH YOUR TOKEN"
    sURL = "http://127.0.0.1:41184"
    sURLNotes = sURL & "/notes?token=" & sToken
    sURLResources = sURL & "/resources?token=" & sToken

    For Each objItem In ActiveExplorer.Selection
        sEscapedBody = EscapeBody( _
                "Date: " & objItem.ReceivedTime & "<br>" _
                & "To: " & objItem.To & "<br>" _
                & objItem.HTMLBody)
        
        sFolderID = GetFoldersFromJoplin(sToken, sURL)
                
        With CreateObject("MSXML2.XMLHTTP")
            .Open "POST", sURLNotes, False
            .Send "{ ""is_todo"": 1, ""title"": """ & objItem.ConversationTopic & attachmentName & """" _
            & ", ""parent_id"": """ & sFolderID & """" _
            & ", ""body_html"": """ & sEscapedBody & """" _
            & " }"
            Do Until .ReadyState = 4: DoEvents: Loop
                sJSONString = .ResponseText
        End With
    Next
    'Debug.Print sJSONString 'Uncomment to see joplin response
End Sub

Private Function EscapeBody(sText As String)
    EscapeBody = sText
    EscapeBody = Replace(EscapeBody, "\", "\\")                 'Backslash is replaced with \\
    EscapeBody = Replace(EscapeBody, Chr(34), "\" & Chr(34))    'Double quote is replaced with \"
    EscapeBody = Replace(EscapeBody, vbCr, "\r")                'Carriage return is replaced with \r
    EscapeBody = Replace(EscapeBody, vbLf, "\n")                'Newline is replaced with \n
    EscapeBody = Replace(EscapeBody, Chr(8), "\b")              'Backspace is replaced with \b
    EscapeBody = Replace(EscapeBody, Chr(12), "\f")             'Form feed is replaced with \f
    EscapeBody = Replace(EscapeBody, vbTab, "\t")               'Tab is replaced with \t
End Function

Private Function GetFoldersFromJoplin(sToken As String, sURL As String)    
    'Input token, url
    'Output folder id
    
    Dim sJSONString, sMessage, sTitle, sDefault, sMyChoice As String
    Dim vJSON As Variant
    Dim sState As String
    Dim aData(), aHeader()
    Dim i As Integer

    sURL = sURL & "/folders?token=" & sToken
    
    'Get folders list
    With CreateObject("MSXML2.XMLHTTP")
        .Open "GET", sURL, False
        .Send
        Do Until .ReadyState = 4: DoEvents: Loop
            sJSONString = .ResponseText
    End With
        
    'Parse JSON response
    JSON.Parse sJSONString, vJSON, sState
    JSON.ToArray vJSON, aData(), aHeader()
    
    'Dsiplay a choices of folders
    'Set prompt
    sMessage = "Enter a value between " & LBound(aData) & " To " & UBound(aData)
    For i = LBound(aData) To UBound(aData)
        sMessage = sMessage & Chr(10) & i & " " & aData(i, 2)
    Next i
    sTitle = "Choose Joplin folder..."    'Set title
    sDefault = "1"    'Set default
    sMyChoice = InputBox(sMessage, sTitle, sDefault)
    GetFoldersFromJoplin = aData(sMyChoice, 0)
End Function