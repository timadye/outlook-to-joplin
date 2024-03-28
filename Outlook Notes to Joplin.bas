Option Explicit
Private Declare PtrSafe Function CoCreateGuid Lib "ole32" (ByRef GUID As Byte) As Long
Private has_more As Boolean

Public Sub SendToJoplin()
    Dim sToken As String
    Dim sUrl As String
    Dim sMailFolderName As String
    Dim sNotesFolderName As String

    sToken = "REPLACE ME WITH YOUR TOKEN"
    sUrl = "http://127.0.0.1:41184"
    sMailFolderName = "Outlook Mail"
    sNotesFolderName = "Outlook Notes"

    Dim sMailFolderID As String
    Dim sNotesFolderID As String
    Dim oNoteIDs
    Set oNoteIDs = CreateObject("Scripting.Dictionary")
    sMailFolderID = ""
    sNotesFolderID = ""
    
    Dim nExport As Integer
    Dim nError As Integer
    nExport = 0
    nError = 0
    
    Dim oItem As Object  ' Outlook.MailItem or Outlook.PostItem or Outlook.DocumentItem or Outlook.NoteItem
    For Each oItem In Application.ActiveExplorer.Selection
        Dim sJSONString As String
        Dim sItemID As String

        If TypeOf oItem Is Outlook.MailItem Or TypeOf oItem Is Outlook.PostItem Or TypeOf oItem Is Outlook.DocumentItem Then

            If sMailFolderID = "" Then
                sMailFolderID = CreateJoplinItem("folder", sMailFolderName, sUrl, sToken)
                If sMailFolderID = "" Then Return
            End If
            
            Dim sAttachmentInfo As String
            sAttachmentInfo = ImportAttachments(oItem, sUrl, sToken)
            sJSONString = HttpRequest(sUrl & "/notes?token=" & sToken, "POST", "{ " _
                            & """is_todo"": 0, ""title"": """ & EscapeBody(oItem.ConversationTopic) & """" _
                            & ", ""parent_id"": """ & sMailFolderID & """" _
                            & ", ""user_created_time"": """ & ToUnixTime(oItem.CreationTime) & """" _
                            & ", ""user_updated_time"": """ & ToUnixTime(UpdateTime(oItem)) & """" _
                            & ", """ & IIf(IsHtml(oItem), "body_html", "body") & """: """ & EscapeBody(MakeBody(oItem, sAttachmentInfo)) & """" _
                            & " }")
            sItemID = ParseJsonResponse(sJSONString, "id", "AddNote")
        
        ElseIf TypeOf oItem Is Outlook.NoteItem Then
        
            If sNotesFolderID = "" Then
                sNotesFolderID = CreateJoplinItem("folder", sNotesFolderName, sUrl, sToken)
                If sNotesFolderID = "" Then Return
            End If
        
            sJSONString = HttpRequest(sUrl & "/notes?token=" & sToken, "POST", "{ " _
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
            Dim aCategories() As String
            aCategories = Split(oItem.Categories, ", ")
            Dim vCategory As Variant
            For Each vCategory In aCategories
                Dim sCategory As String
                Dim sTagID As String
                sCategory = vCategory
                If oNoteIDs.Exists(sCategory) Then
                    sTagID = oNoteIDs.item(sCategory)
                Else
                    sTagID = CreateJoplinItem("tag", sCategory, sUrl, sToken)
                    oNoteIDs.Add sCategory, sTagID
                End If
                If sTagID = "" Then
                    nError = nError + 1
                Else
                    sJSONString = HttpRequest(sUrl & "/tags/" & sTagID & "/notes?token=" & sToken, "POST", "{ ""id"": """ & sItemID & """ }")
                    Dim sTaggedID As String
                    sTaggedID = ParseJsonResponse(sJSONString, "id", "AddNote")
                    If sTaggedID = "" Then
                        nError = nError + 1
                    End If
                End If
            Next
        End If
    Next
    Dim sMsg As String
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

Private Function UpdateTime(oItem As Object) As Date
    If TypeOf oItem Is Outlook.DocumentItem Then
        UpdateTime = oItem.LastModificationTime
    Else
        UpdateTime = oItem.ReceivedTime
    End If
End Function

Private Function ImportAttachments(oItem As Object, sUrl As String, sToken As String) As String
    Dim oAttachment As Object
    Dim sTemp As String
    Dim sJSONString As String
    Dim sFileID As String

    ImportAttachments = ""
    sTemp = Environ("TEMP")
    For Each oAttachment In oItem.Attachments
        Dim sSaveFile As String
        sSaveFile = sTemp & "\" & NewGuid()
        If InStrRev(oAttachment.FileName, ".") > 0 Then sSaveFile = sSaveFile & Mid(oAttachment.FileName, InStrRev(oAttachment.FileName, "."))
        oAttachment.SaveAsFile sSaveFile
        sJSONString = HttpUpload(sUrl & "/resources?token=" & sToken, sSaveFile, "{ ""title"":""" & oAttachment.DisplayName & """ }")
        Kill sSaveFile
        sFileID = ParseJsonResponse(sJSONString, "id", "ImportAttachments")
        If ImportAttachments <> "" Then ImportAttachments = ImportAttachments & ", "
        If IsHtml(oItem) Then
            ImportAttachments = ImportAttachments & "<a href=""" & sFileID & """>" & oAttachment.FileName & "</a>"
        Else
            ImportAttachments = ImportAttachments & "[" & oAttachment.FileName & "](:/" & sFileID & ")"
        End If
    Next
End Function

Private Function IsHtml(oItem As Object) As Boolean
    If TypeOf oItem Is Outlook.DocumentItem Then
        IsHtml = False
    Else
        IsHtml = (oItem.BodyFormat = olFormatHTML)
    End If
End Function

Private Function MakeBody(oItem As Object, sAttachmentInfo As String) As String
    Dim sFrom As String
    Dim sNl As String

    If Not (TypeOf oItem Is Outlook.DocumentItem) Then
        sFrom = oItem.SenderEmailAddress
        If oItem.SenderName <> "" Then
            If sFrom <> "" And oItem.SenderEmailType = "SMTP" Then
                sFrom = oItem.SenderName & " <" & sFrom & ">"
            Else
                sFrom = oItem.SenderName
            End If
        End If
    End If
    If IsHtml(oItem) Then
        MakeBody = oItem.HTMLBody
        sFrom = EscapeHtml(sFrom)
        sNl = "<br/>" & vbLf
    Else
        MakeBody = oItem.Body
        sNl = vbLf
    End If
    If sAttachmentInfo <> "" Then MakeBody = "Attachments: " & sAttachmentInfo & sNl & sNl & MakeBody
    If TypeOf oItem Is Outlook.MailItem Then
        If oItem.To <> "" Then
            If sAttachmentInfo = "" Then MakeBody = sNl & MakeBody
            MakeBody = "To: " & oItem.To & sNl & MakeBody
            If sFrom <> "" Then MakeBody = "From: " & sFrom & sNl & MakeBody
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

Private Function FindJoplinItem(sType As String, sItemName As String, sUrl As String, sToken As String) As String
    Dim page As Integer

    page = 1
    Do
        Dim sJSONString As String
        Dim aItems As Variant
        sJSONString = HttpRequest(sUrl & "/search?query=" & sItemName & "&type=" & sType & "&page=" & page & "&token=" & sToken)
        page = page + 1
        aItems = ParseJsonResponse(sJSONString, "items", "FindJoplinItem")
    
        If IsArray(aItems) Then
            Dim jItem As Variant
            For Each jItem In aItems
                If VarType(jItem) = vbObject Then
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
    
Private Function CreateJoplinItem(sType As String, sItemName As String, sUrl As String, sToken As String) As String

    CreateJoplinItem = FindJoplinItem(sType, sItemName, sUrl, sToken)
    If CreateJoplinItem <> "" Then Exit Function

    Dim sJSONString As String
    sJSONString = HttpRequest(sUrl & "/" & sType & "s?token=" & sToken, "POST", "{ ""title"": """ & EscapeBody(sItemName) & """ }")
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

Private Function HttpRequest(sUrl As String, Optional sMethod As String = "GET", Optional sPost As String = "") As String
    Dim sResponse As String
    
    With CreateObject("Msxml2.ServerXMLHTTP")
        .Open sMethod, sUrl, False
        .setRequestHeader "Cache-Control", "no-cache"
        .setRequestHeader "Pragma", "no-cache"
        .Send sPost
        Do Until .ReadyState = 4: DoEvents: Loop
            sResponse = .ResponseText
    End With
'    Debug.Print sResponse & " <- " & sMethod & " " & sURL & " " & sPost
    HttpRequest = sResponse
End Function

Private Function HttpUpload(sUrl As String, sFileName As String, sPost As String) As String
    ' upload file based on XMLHTTP example from https://wqweto.wordpress.com/2011/07/12/vb6-using-wininet-to-post-binary-file/
    Dim STR_BOUNDARY As String
    STR_BOUNDARY = "SendToJoplin-" & NewGuid()
    Dim fileNo As Integer
    Dim baFileData() As Byte
    Dim sResponse As String

    ' read file
    fileNo = FreeFile
    Open sFileName For Binary Access Read As fileNo
    If LOF(fileNo) > 0 Then
        ReDim baFileData(0 To LOF(fileNo) - 1) As Byte
        Get fileNo, , baFileData
    End If
    Close fileNo

    ' upload file
    With CreateObject("Msxml2.ServerXMLHTTP")
        .Open "POST", sUrl, False
        .setRequestHeader "Cache-Control", "no-cache"
        .setRequestHeader "Pragma", "no-cache"
        .setRequestHeader "Content-Type", "multipart/form-data; boundary=" & STR_BOUNDARY
        .Send CombineArrays( _
            pvToByteArray( _
                "--" & STR_BOUNDARY & vbCrLf & _
                "Content-Disposition: form-data; name=""props""" & vbCrLf & vbCrLf & _
                sPost & vbCrLf & _
                "--" & STR_BOUNDARY & vbCrLf & _
                "Content-Disposition: form-data; name=""data""; filename=""" & Mid(sFileName, InStrRev(sFileName, "\") + 1) & """" & vbCrLf & _
                "Content-Type: application/octet-stream" & vbCrLf & vbCrLf _
            ), _
            baFileData, _
            pvToByteArray(vbCrLf & "--" & STR_BOUNDARY & "--") _
        )
        Do Until .ReadyState = 4: DoEvents: Loop
            sResponse = .ResponseText
    End With
'    Debug.Print sResponse
    HttpUpload = sResponse
End Function

Public Function CombineArrays(ParamArray arraysToMerge() As Variant) As Byte()
    ' Adapted from https://stackoverflow.com/a/51407942/6199960
    Dim CombinedArrayLength As Long
    Dim i As Long, j As Long
    
    CombinedArrayLength = 0
    For i = LBound(arraysToMerge) To UBound(arraysToMerge)
        CombinedArrayLength = CombinedArrayLength + (UBound(arraysToMerge(i)) - LBound(arraysToMerge(i)) + 1)
    Next i

    Dim combinedArray() As Byte
    ReDim combinedArray(0 To CombinedArrayLength - 1)

    Dim combinedArrayIndex As Long
    combinedArrayIndex = LBound(combinedArray)
    For i = LBound(arraysToMerge) To UBound(arraysToMerge)
        For j = LBound(arraysToMerge(i)) To UBound(arraysToMerge(i))
            combinedArray(combinedArrayIndex) = arraysToMerge(i)(j)
            combinedArrayIndex = combinedArrayIndex + 1
        Next j
    Next i
'    Debug.Print StrConv(combinedArray, vbUnicode)
    CombineArrays = combinedArray
End Function

Private Function pvToByteArray(sText As String) As Byte()
    pvToByteArray = StrConv(sText, vbFromUnicode)
End Function

Private Function NewGuid() As String
    ' based on https://stackoverflow.com/a/23126614/6199960
    Dim ID(0 To 15) As Byte
    Dim N As Integer
    Dim GUID As String
    Dim Res As Long

    Res = CoCreateGuid(ID(0))
    For N = 0 To 15
        GUID = GUID & Right("0" & Hex(ID(N)), 2)
        If N = 3 Or N = 5 Or N = 7 Or N = 9 Then GUID = GUID & "-"
    Next N
    NewGuid = GUID
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
