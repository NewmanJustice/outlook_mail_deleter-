Option Explicit

Public Sub Purge_ByReceivedDate_First150()
    Dim fld As Outlook.MAPIFolder
    On Error Resume Next
    Set fld = Application.ActiveExplorer.CurrentFolder
    On Error GoTo 0
    If fld Is Nothing Then
        MsgBox "Select a folder in Outlook, then run the macro.", vbExclamation
        Exit Sub
    End If

    ' 1) Ask for cutoff date (items with ReceivedTime strictly BEFORE this date)
    Dim vCut As Variant, cut As Date
    vCut = PromptForCutoffDate(#10/1/2023#) ' default shows as 01/10/2023 on UK systems
    If IsEmpty(vCut) Then
        MsgBox "Cancelled.", vbInformation
        Exit Sub
    End If
    cut = CDate(vCut) ' ensure Date

    ' 2) Find matches: strictly before cutoff, oldest first, MailItem only, cap 150
    Dim ids() As String: ReDim ids(1 To 150)
    Dim storeId As String: storeId = fld.StoreID
    Dim kept As Long: kept = 0
    Dim totalMailMatches As Long: totalMailMatches = 0

    Dim itms As Outlook.Items, matches As Outlook.Items
    Set itms = fld.Items
    ' Sort ascending so index 1 is oldest
    itms.Sort "[ReceivedTime]", False

    ' Outlook Restrict date format is safest as ISO yyyy-mm-dd hh:nn
    Dim filter As String
    filter = "[ReceivedTime] < '" & Format$(cut, "yyyy-mm-dd 00:00") & "'"

    On Error Resume Next
    Set matches = itms.Restrict(filter)
    On Error GoTo 0
    If matches Is Nothing Then
        MsgBox "No items matched (or folder not readable).", vbInformation
        Exit Sub
    End If

    ' Iterate once: count all MailItems, keep EntryIDs for the first 150 (oldest first)
    Dim i As Long
    For i = 1 To matches.Count
        Dim obj As Object
        Set obj = matches(i)
        If TypeName(obj) = "MailItem" Then
            totalMailMatches = totalMailMatches + 1
            If kept < 150 Then
                kept = kept + 1
                ids(kept) = obj.EntryID
            End If
        End If
        DoEvents
    Next

    If totalMailMatches = 0 Then
        MsgBox "Found 0 emails with Received date before " & Format$(cut, "dd/mm/yyyy") & _
               " in '" & fld.FolderPath & "'.", vbInformation
        Exit Sub
    End If

    ' 3) Confirm action
    Dim prompt As String
    If totalMailMatches <= 150 Then
        prompt = "Found " & totalMailMatches & " email(s) received BEFORE " & _
                 Format$(cut, "dd/mm/yyyy") & " in:" & vbCrLf & _
                 fld.FolderPath & vbCrLf & vbCrLf & _
                 "Do you want to PERMANENTLY DELETE these email(s)?"
    Else
        prompt = "Found more than 150 email(s) received BEFORE " & _
                 Format$(cut, "dd/mm/yyyy") & "." & vbCrLf & _
                 "This will delete the FIRST 150 (oldest) from:" & vbCrLf & _
                 fld.FolderPath & vbCrLf & vbCrLf & _
                 "Do you want to PERMANENTLY DELETE these 150 email(s)?"
    End If

    If MsgBox(prompt, vbYesNo + vbExclamation + vbDefaultButton2, "Confirm permanent delete") <> vbYes Then
        MsgBox "No changes made.", vbInformation
        Exit Sub
    End If

    ' 4) Permanently delete selected items (by EntryID to avoid index shifts)
    Dim ns As Outlook.NameSpace
    Set ns = Application.GetNamespace("MAPI")

    Dim deleted As Long: deleted = 0
    For i = 1 To kept
        If LenB(ids(i)) > 0 Then
            On Error Resume Next
            Dim itm As Object
            Set itm = ns.GetItemFromID(ids(i), storeId)
            If Not itm Is Nothing Then
                ' Check if itm supports Delete method
                If TypeName(itm) = "MailItem" Or TypeName(itm) = "PostItem" Then
                    ' StatusBar update removed (not supported in this Outlook version)
                    itm.Delete
                    deleted = deleted + 1
                Else
                    MsgBox "Warning: Item type '" & TypeName(itm) & "' does not support Delete method.", vbExclamation
                End If
            End If
            If Err.Number <> 0 Then
                MsgBox "Error deleting item: " & Err.Description, vbCritical
                Err.Clear
            End If
            On Error GoTo 0
        End If
        DoEvents
    Next
    Application.ActiveExplorer.StatusBar = False

    MsgBox "Done. Permanently deleted " & deleted & " email(s) from:" & vbCrLf & _
           fld.FolderPath, vbInformation
End Sub

' === Helpers ===

' Prompts for a date; accepts dd/mm/yyyy or ISO yyyy-mm-dd. Returns Empty if cancelled.
Private Function PromptForCutoffDate(ByVal defaultDate As Date) As Variant
    Dim tries As Long
    For tries = 1 To 3
        Dim s As String
        s = InputBox("Enter the cutoff date (emails with ReceivedTime BEFORE this date are affected)." & vbCrLf & _
                     "Examples: 30/10/2023  or  2023-10-30", _
                     "Cutoff date", _
                     Format$(defaultDate, "dd/mm/yyyy"))
        If s = "" Then
            PromptForCutoffDate = Empty
            Exit Function
        End If

        Dim d As Date
        If TryParseDate(s, d) Then
            PromptForCutoffDate = d
            Exit Function
        Else
            MsgBox "Sorry, I couldn't understand that date. Try again.", vbExclamation
        End If
    Next
    PromptForCutoffDate = Empty
End Function

' Tries ISO first, then locale (UK dd/mm/yyyy).
Private Function TryParseDate(ByVal s As String, ByRef d As Date) As Boolean
    On Error GoTo fail
    Dim t As String: t = Trim$(s)
    If InStr(t, "-") > 0 Then
        Dim a() As String: a = Split(t, "-")
        If UBound(a) = 2 Then
            d = DateSerial(CInt(a(0)), CInt(a(1)), CInt(a(2)))
            TryParseDate = True
            Exit Function
        End If
    End If
    d = CDate(t)
    TryParseDate = True
    Exit Function
fail:
    TryParseDate = False
End Function
