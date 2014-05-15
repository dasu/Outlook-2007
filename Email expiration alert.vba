Private Sub Application_MAPILogonComplete()

    Dim filter As String
    Dim folder As Outlook.folder
    Dim row As Outlook.row
    Dim table As Outlook.table
    Dim subfolder As Outlook.folder
    Dim subfolderloc() As String
    Dim subname As String
    
    ReDim subfolderloc(0 To 0) As String
    
    
    
    Set folder = Application.Session.GetDefaultFolder(olFolderInbox)
    num = 0
    date5 = Format((DateAdd("d", -65, Now())), "mm/dd/yyyy")
    filter = "[CreationTime] < '" & date5 & "'"
    filter2 = "[MessageClass] = 'IPM.Note'"
    Set table = folder.GetTable(filter)
    Set table2 = table.Restrict(filter2)

    Do Until (table2.EndOfTable)
        Set row = table2.GetNextRow()
        num = num + 1
        subname = folder.Name
    Loop
    If subname <> "" Then
        subfolderloc(UBound(subfolderloc)) = subname
        ReDim Preserve subfolderloc(0 To UBound(subfolderloc) + 1) As String
        subname = ""
    End If
    
            
    For Each subfolder In folder.Folders
        Set table = subfolder.GetTable(filter)
        Set table2 = table.Restrict(filter2)
        Do Until (table2.EndOfTable)
            Set row = table2.GetNextRow()
            num = num + 1
            subname = subfolder.Name
        Loop
        If subname <> "" Then
            subfolderloc(UBound(subfolderloc)) = subname
            ReDim Preserve subfolderloc(0 To UBound(subfolderloc) + 1) As String
            subname = ""
        End If
    Next
    If Len(Join(subfolderloc)) > 0 Then
        ReDim Preserve subfolderloc(0 To UBound(subfolderloc) - 1) As String
        locs = Join(subfolderloc, vbCrLf)
    End If
    
    Set folder = Application.Session.GetDefaultFolder(olFolderSentMail)
    Set table = folder.GetTable(filter)
    Set table2 = table.Restrict(filter2)
    Do Until (table2.EndOfTable)
        Set row = table2.GetNextRow()
        num = num + 1
        innum = 1
    Loop
    
    If innum = 1 Then
        If locs = "" Then
            locs = locs & "Sent Items"
        Else
            locs = locs & vbCrLf & "Sent Items"
        End If
    End If
    
    If num > 0 Then
        Msg = MsgBox("Currently " & num & " emails will expire in 5 days from the following folders:" & vbCrLf & vbCrLf & locs & vbCrLf & vbCrLf & "Consider filing these emails as soon as possible.", vbMsgBoxSetForeground + vbExclamation)
    End If
    
    Debug.Print filter
End Sub


