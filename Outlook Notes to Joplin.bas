Public Sub SendToJoplin()
    Dim sToken As String, sURL As String
    Dim sURLNotes, sURLResources, sEscapedBody, sJSONString, sFolderID As String
    Dim objItem As Outlook.NoteItem
    
    sToken = "REPLACE ME WITH YOUR TOKEN"
    sURL = "http://127.0.0.1:41184"
    sURLNotes = sURL & "/notes?token=" & sToken
    sURLResources = sURL & "/resources?token=" & sToken

    For Each objItem In Application.ActiveExplorer.Selection
        sEscapedBody = "Categories: " & objItem.Categories & "\n\n" _
                & EscapeBody(objItem.Body)
        If objItem.Categories = "" Then
            sEscapedBody = EscapeBody(objItem.Body)
        Else
            sEscapedBody = "Categories: " & objItem.Categories & "\n\n" & EscapeBody(objItem.Body)
        End If

        ' sFolderID = GetFoldersFromJoplin(sToken, sURL)
        sFolderID = "831460035a924f688eda5c7bd83ddcbc"
        
        sPost = "{ ""is_todo"": 0, ""title"": """ & objItem.Subject & """" _
            & ", ""parent_id"": """ & sFolderID & """" _
            & ", ""user_created_time"": """ & ToUnix(objItem.CreationTime) & """" _
            & ", ""user_updated_time"": """ & ToUnix(objItem.LastModificationTime) & """" _
            & ", ""body"": """ & sEscapedBody & """" _
            & " }"
        ' Debug.Print sPost

        With CreateObject("MSXML2.XMLHTTP")
            .Open "POST", sURLNotes, False
            .Send sPost
            Do Until .ReadyState = 4: DoEvents: Loop
                sJSONString = .ResponseText
        End With
        Debug.Print sJSONString 'Uncomment to see joplin response
    Next
End Sub

Private Function EscapeBody(sText As String)
    EscapeBody = sText
    EscapeBody = Replace(EscapeBody, "\", "\\")                 'Backslash is replaced with \\
    EscapeBody = Replace(EscapeBody, Chr(34), "\" & Chr(34))    'Double quote is replaced with \"
    EscapeBody = Replace(EscapeBody, vbCr + vbLf, "\n")              'Carriage return + Newline is replaced with <br>\n
    EscapeBody = Replace(EscapeBody, vbCr, "\r")                'Carriage return is replaced with <br>\r
    EscapeBody = Replace(EscapeBody, vbLf, "\n")                'Newline is replaced with <br>\n
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
        
    Debug.Print sJSONString
    
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

Public Function ToUnix(ByVal dt As Date) As LongLong
    ToUnix = DateDiff("s", "1/1/1970 00:00:00", dt) * 1000
    '    Debug.Print ToUnix
End Function
