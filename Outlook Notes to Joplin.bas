Option Explicit

Private has_more As Boolean

Public Sub SendToJoplin()
    Dim sToken As String
    Dim sURL As String
    Dim sJSONString As String
    Dim sMailFolderID As String
    Dim sNotesFolderID As String
    Dim sMailFolderName As String
    Dim sNotesFolderName As String
    Dim sItemID As String
    Dim sMsg As String
    Dim oItem As Object  ' Outlook.MailItem or Outlook.PostItem or Outlook.NoteItem
    Dim nExport As Integer
    Dim nError As Integer
    Dim sTagID As String
    Dim aCategories() As String
    Dim i As Integer
    Dim sTaggedID As String
    Dim oNoteIDs

    sToken = "REPLACE ME WITH YOUR TOKEN"
    sURL = "http://127.0.0.1:41184"
    sMailFolderName = "Outlook Mail"
    sNotesFolderName = "Outlook Notes"

    Set oNoteIDs = CreateObject("Scripting.Dictionary")
    sMailFolderID = ""
    sNotesFolderID = ""
    nExport = 0
    nError = 0
    For Each oItem In Application.ActiveExplorer.Selection

        If TypeOf oItem Is Outlook.MailItem Or TypeOf oItem Is Outlook.PostItem Then

            If sMailFolderID = "" Then
                sMailFolderID = CreateJoplinItem("folder", sMailFolderName, sURL, sToken)
                If sMailFolderID = "" Then Return
            End If
            sJSONString = HttpRequest(sURL & "/notes?token=" & sToken, "POST", "{ " _
                            & """is_todo"": 0, ""title"": """ & EscapeBody(oItem.ConversationTopic) & """" _
                            & ", ""parent_id"": """ & sMailFolderID & """" _
                            & ", ""user_created_time"": """ & ToUnixTime(oItem.CreationTime) & """" _
                            & ", ""user_updated_time"": """ & ToUnixTime(oItem.ReceivedTime) & """" _
                            & ", """ & IIf(oItem.BodyFormat = olFormatHTML, "body_html", "body") & """: """ & EscapeBody(MakeBody(oItem)) & """" _
                            & " }")
            sItemID = ParseJsonResponse(sJSONString, "id", "AddNote")
        
        ElseIf TypeOf oItem Is Outlook.NoteItem Then
        
            If sNotesFolderID = "" Then
                sNotesFolderID = CreateJoplinItem("folder", sNotesFolderName, sURL, sToken)
                If sNotesFolderID = "" Then Return
            End If
        
            sJSONString = HttpRequest(sURL & "/notes?token=" & sToken, "POST", "{ " _
                            & """is_todo"": 0, ""title"": """ & EscapeBody(oItem.Subject) & """" _
                            & ", ""parent_id"": """ & sNotesFolderID & """" _
                            & ", ""user_created_time"": """ & ToUnixTime(oItem.CreationTime) & """" _
                            & ", ""user_updated_time"": """ & ToUnixTime(oItem.LastModificationTime) & """" _
                            & ", ""body"": """ & EscapeBody(oItem.Body) & """" _
                            & " }")
            sItemID = ParseJsonResponse(sJSONString, "id", "AddNote")

        Else
            MsgBox "Outlook " & TypeName(oItem) & " is not supported: " & oItem.Subject
            sItemID = ""
        End If

        If sItemID <> "" Then
            nExport = nExport + 1
            Debug.Print nExport & " " & EscapeBody(oItem.Subject)
        Else
            nError = nError + 1
        End If
    
    
        If oItem.Categories <> "" And sItemID <> "" Then
            aCategories = Split(oItem.Categories, ", ")
            For i = LBound(aCategories, 1) To UBound(aCategories, 1)
                If oNoteIDs.Exists(aCategories(i)) Then
                    sTagID = oNoteIDs.item(aCategories(i))
                Else
                    sTagID = CreateJoplinItem("tag", aCategories(i), sURL, sToken)
                    oNoteIDs.Add aCategories(i), sTagID
                End If
                If sTagID = "" Then
                    nError = nError + 1
                Else
                    sJSONString = HttpRequest(sURL & "/tags/" & sTagID & "/notes?token=" & sToken, "POST", "{ ""id"": """ & sItemID & """ }")
                    sTaggedID = ParseJsonResponse(sJSONString, "id", "AddNote")
                    If sTaggedID = "" Then
                        nError = nError + 1
                    End If
                End If
            Next
        End If
    Next
    sMsg = nExport & " notes exported to Joplin folder "
    If sMailFolderID = "" Then
        sMsg = sMsg & """" & sNotesFolderName & """"
    ElseIf sNotesFolderID = "" Then
        sMsg = sMsg & """" & sMailFolderName & """"
    Else
        sMsg = sMsg & """" & sNotesFolderName & """ and """ & sMailFolderName & """"
    End If
    If nError = 0 Then
        MsgBox sMsg
    Else
        MsgBox nError & " errors encountered. " & sMsg
    End If
End Sub

Private Function MakeBody(oItem As Object) As String
    Dim sFrom As String
    Dim sNl As String

    sFrom = oItem.SenderEmailAddress
    If oItem.SenderName <> "" Then
        If sFrom <> "" And oItem.SenderEmailType = "SMTP" Then
            sFrom = oItem.SenderName & " <" & sFrom & ">"
        Else
            sFrom = oItem.SenderName
        End If
    End If
    If oItem.BodyFormat = olFormatHTML Then
        MakeBody = oItem.HTMLBody
        sFrom = EscapeHtml(sFrom)
        sNl = "<br/>" & vbLf
    Else
        MakeBody = oItem.Body
        sNl = vbLf
    End If
    If TypeOf oItem Is Outlook.MailItem Then
        If oItem.To <> "" Then
            MakeBody = "From: " & sFrom & sNl & _
                       "To: " & oItem.To & sNl & sNl & _
                       MakeBody
        End If
    End If
End Function


Private Function EscapeBody(sText As String) As String
    EscapeBody = sText
    EscapeBody = Replace(EscapeBody, "\", "\\")                 'Backslash is replaced with \\
    EscapeBody = Replace(EscapeBody, Chr(34), "\" & Chr(34))    'Double quote is replaced with \"
    EscapeBody = Replace(EscapeBody, vbCr + vbLf, "\n")         'Carriage return + Newline is replaced with \n
    EscapeBody = Replace(EscapeBody, vbCr, "\r")                'Carriage return is replaced with \r
    EscapeBody = Replace(EscapeBody, vbLf, "\n")                'Newline is replaced with \n
    EscapeBody = Replace(EscapeBody, Chr(8), "\b")              'Backspace is replaced with \b
    EscapeBody = Replace(EscapeBody, Chr(12), "\f")             'Form feed is replaced with \f
    EscapeBody = Replace(EscapeBody, vbTab, "\t")               'Tab is replaced with \t
End Function

Private Function EscapeHtml(sText As String) As String
    EscapeHtml = sText
    EscapeHtml = Replace(EscapeHtml, "&", "&amp;")
    EscapeHtml = Replace(EscapeHtml, "<", "&lt;")
    EscapeHtml = Replace(EscapeHtml, ">", "&gt;")
End Function

Private Function FindJoplinItem(sType As String, sItemName As String, sURL As String, sToken As String) As String
    Dim sJSONString As String
    Dim aItems As Variant
    Dim i As Integer
    Dim page As Integer
    Dim jItem As Object

    page = 1
    Do
        sJSONString = HttpRequest(sURL & "/search?query=" & sItemName & "&type=" & sType & "&page=" & page & "&token=" & sToken)
        page = page + 1
        aItems = ParseJsonResponse(sJSONString, "items", "FindJoplinItem")
    
        If IsArray(aItems) Then
            For i = LBound(aItems) To UBound(aItems)
                If VarType(aItems(i)) = vbObject Then
                    Set jItem = aItems(i)
                    If jItem.Exists("id") And jItem.Exists("title") And jItem.Exists("parent_id") Then
'                        Debug.Print jItem.Item("id") & " " & jItem.item("title")
                        If jItem.item("parent_id") = "" And LCase(jItem.item("title")) = LCase(sItemName) Then
                            FindJoplinItem = jItem.item("id")
                            Exit Function
                        End If
                    End If
                End If
            Next
        End If
    Loop While has_more
    FindJoplinItem = ""
End Function
    
Private Function CreateJoplinItem(sType As String, sItemName As String, sURL As String, sToken As String) As String
    Dim sJSONString As String

    CreateJoplinItem = FindJoplinItem(sType, sItemName, sURL, sToken)
    If CreateJoplinItem <> "" Then Exit Function

    sJSONString = HttpRequest(sURL & "/" & sType & "s?token=" & sToken, "POST", "{ ""title"": """ & EscapeBody(sItemName) & """ }")
    CreateJoplinItem = ParseJsonResponse(sJSONString, "id", "CreateJoplinItem")
End Function

Private Function ParseJsonResponse(sJSONString As String, sItem As String, sOp As String)
    Dim vJSON As Variant
    Dim sState As String
    
    has_more = False
    JSON.Parse sJSONString, vJSON, sState
    ParseJsonResponse = ""
    If sState <> "Object" Then
        MsgBox sOp & ": invalid response from Joplin server: " & sJSONString
    ElseIf vJSON.Exists("error") Then
        MsgBox sOp & " error: " & vJSON.item("error")
    ElseIf Not vJSON.Exists(sItem) Then
        MsgBox sOp & ": no item """ & sItem & """ in response from Joplin server: " & sJSONString
    Else
        ParseJsonResponse = vJSON.item(sItem)
    End If
    If vJSON.Exists("has_more") Then
        has_more = vJSON.item("has_more")
    End If
End Function

Private Function HttpRequest(sURL As String, Optional sMethod As String = "GET", Optional sPost As String = "") As String
    Dim sResponse As String
    
    With CreateObject("Msxml2.ServerXMLHTTP")
        .Open sMethod, sURL, False
        .setRequestHeader "Cache-Control", "no-cache"
        .setRequestHeader "Pragma", "no-cache"
        .Send sPost
        Do Until .ReadyState = 4: DoEvents: Loop
            sResponse = .ResponseText
    End With
'    Debug.Print sResponse & " <- " & sMethod & " " & sURL & " " & sPost
    HttpRequest = sResponse
End Function

Private Function ToUnixTime(ByVal dt As Date) As LongLong
   ' ToUnixTime convert Date value in the local timezone to Unix timestamp in milliseconds, UTC
   ' Based on example from https://gist.github.com/seakintruth/ddcc3d5e400a5083458494ae30d55466
    Dim objDateTime
    Set objDateTime = CreateObject("WbemScripting.SWbemDateTime")
    objDateTime.SetVarDate dt
    ToUnixTime = DateDiff("s", "01/01/1970 00:00:00", CDate(objDateTime.GetVarDate(False))) * 1000 + Fix((dt - Fix(dt)) * 1000)
'    Debug.Print dt & Format(dt - Fix(dt), ".000") & " -> " & ToUnixTime
End Function
