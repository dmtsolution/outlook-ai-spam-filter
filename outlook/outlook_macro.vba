' ============================================================================
' Macro a coller dans ThisOutlookSession (Outlook classique, Alt+F11)
' Appelle le serveur local spam_api.py (http://127.0.0.1:5000/predict)
' a chaque nouveau mail recu : tague "IA-Spam" puis deplace vers le dossier
' Courrier indesirable (Junk) du MEME compte que le mail recu.
' ============================================================================

Private Sub Application_NewMailEx(ByVal EntryIDCollection As String)
    Dim ns As Outlook.NameSpace
    Dim ids() As String
    Dim i As Long
    Dim mail As Object

    Set ns = Application.GetNamespace("MAPI")
    ids = Split(EntryIDCollection, ",")

    For i = LBound(ids) To UBound(ids)
        On Error Resume Next
        Set mail = ns.GetItemFromID(CStr(ids(i)))
        On Error GoTo 0
        If Not mail Is Nothing Then
            If mail.Class = olMail Then
                ClassifySpam mail
            End If
        End If
        Set mail = Nothing
    Next i
End Sub

Private Sub ClassifySpam(mail As Outlook.MailItem)
    Dim http As Object
    Dim json As String
    Dim respText As String
    Dim pos As Long
    Dim probaStr As String
    Dim proba As Double

    On Error GoTo ErrHandler

    json = "{""subject"":""" & EscapeJson(mail.Subject) & _
           """,""body"":""" & EscapeJson(Left(mail.Body, 5000)) & """}"

    Set http = CreateObject("MSXML2.XMLHTTP.6.0")
    http.Open "POST", "http://127.0.0.1:5000/predict", False
    http.setRequestHeader "Content-Type", "application/json; charset=utf-8"
    http.Send json

    respText = http.responseText

    pos = InStr(respText, """proba_spam"":")
    If pos > 0 Then
        probaStr = Mid(respText, pos + Len("""proba_spam"":"))
        probaStr = Left(probaStr, InStr(probaStr, "}") - 1)
        probaStr = Replace(probaStr, ",", "")
        proba = Val(probaStr)   ' Val() lit toujours "." comme separateur decimal
    End If

    If proba >= 0.5 Then
        If InStr(1, mail.Categories, "IA-Spam", vbTextCompare) = 0 Then
            If Len(Trim(mail.Categories)) > 0 Then
                mail.Categories = mail.Categories & ", IA-Spam"
            Else
                mail.Categories = "IA-Spam"
            End If
            mail.Save
        End If

        ' Deplacement vers le dossier Courrier indesirable du MEME compte
        ' (important en multi-comptes : ne pas utiliser le Junk par defaut global)
        Dim parentFolder As Outlook.Folder
        Dim junkFolder As Outlook.Folder
        Set parentFolder = mail.Parent
        Set junkFolder = parentFolder.Store.GetDefaultFolder(olFolderJunk)
        mail.Move junkFolder
    End If

    Exit Sub

ErrHandler:
    ' Le serveur local n'est probablement pas lance (python spam_api.py)
    Debug.Print "Erreur classification spam : " & Err.Description
End Sub

Private Function EscapeJson(ByVal s As String) As String
    If IsNull(s) Then
        EscapeJson = ""
        Exit Function
    End If
    s = Replace(s, "\", "\\")
    s = Replace(s, """", "\""")
    s = Replace(s, vbCrLf, " ")
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    s = Replace(s, vbTab, " ")
    EscapeJson = s
End Function
