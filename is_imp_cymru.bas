Dim CurrentUserIsImprovementCymruCachedValue As Boolean
Dim CurrentUserIsImprovementCymruCachedFlag As Boolean

Function CurrentUserIsImprovementCymru() As Boolean
    CurrentUserIsImprovementCymru = False
    
    If CurrentUserIsImprovementCymruCachedFlag Then
        CurrentUserIsImprovementCymru = CurrentUserIsImprovementCymruCachedValue
        Exit Function
    End If
    
    On Error Resume Next
    
    Dim strPath As String
    
    Set objNetwork = CreateObject("WScript.Network")
    
    If objNetwork.UserDomain <> "CYMRU" Then
        On Error GoTo 0
        Exit Function
    End If
    
    strPath = "WinNT://" & objNetwork.UserDomain & "/"
    
    Set objImpCymruGrp = GetObject(strPath & "Improvement Cymru" & ",group")
    
    If objImpCymruGrp.IsMember(strPath & objNetwork.UserName) Then
        CurrentUserIsImprovementCymruCachedFlag = True
        CurrentUserIsImprovementCymruCachedValue = True
        CurrentUserIsImprovementCymru = True
    End If
    
    CurrentUserIsImprovementCymruCachedFlag = True
    On Error GoTo 0
End Function

Sub TestCurrentUserIsImprovementCymru()
    If CurrentUserIsImprovementCymru() Then
        MsgBox "You *are* a member of Improvement Cymru!", vbInformation
    Else
        MsgBox "You are *not* a member of Improvement Cymru!", vbExclamation
    End If

End Sub
