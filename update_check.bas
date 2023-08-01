Private Declare PtrSafe Function ShellExecute _
  Lib "shell32.dll" Alias "ShellExecuteA" ( _
  ByVal hWnd As Long, _
  ByVal Operation As String, _
  ByVal Filename As String, _
  Optional ByVal Parameters As String, _
  Optional ByVal Directory As String, _
  Optional ByVal WindowStyle As Long = vbMinimizedFocus _
  ) As Long

Public Sub OpenUrl(url As String)

    Dim lSuccess As Long
    lSuccess = ShellExecute(0, "Open", url)

End Sub


Sub UpdateCheck()
    
    Dim current_version As String
    Dim repo As String
    
    current_version = "2023-08-01T10-55"
    repo = "impcymru/impcymru-chart-formatter"
    
    Dim hReq As Object
    Dim strUrl As String

    strUrl = "https://api.github.com/repos/" & repo & "/releases/latest"
    Set hReq = CreateObject("MSXML2.XMLHTTP")
    With hReq
        .Open "GET", strUrl, False
        .Send
    End With

    Dim response As String
    response = hReq.responseText
    
    Dim regexObject As RegExp
    Set regexObject = CreateObject("VBScript.RegExp")
    
    ' This is a bit of hack
    regexObject.Pattern = """?name""?:\s*?""?(.+?)""?,"

    Set matches = regexObject.Execute(response)
    
    If matches.Count < 1 Then
        Exit Sub
    End If
    
    
    Dim latest_version As String
    
    latest_version = matches(0).SubMatches.Item(0)

    If latest_version > current_version Then
        ret = MsgBox("A newer version (" & latest_version & ") of the add-in is available." & vbNewLine & _
        "Do you want to download this new version now?" _
        , vbYesNo + vbInformation, "Update available")
        
        If ret = vbYes Then
            OpenUrl ("https://github.com/" & repo & "/releases/latest")
        End If
        
        
    End If

End Sub
