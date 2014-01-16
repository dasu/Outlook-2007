Private Sub Application_MAPILogonComplete()
    Dim filter As String
    Dim folder As Outlook.folder
    Dim row As Outlook.row
    Dim table As Outlook.table
    
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
    Loop
    If num > 0 Then
    MsgBox ("Currently " & num & " emails will expire in 5 days")
    End If
    Debug.Print filter
    
End Sub
