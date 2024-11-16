' ************************************************************
' License Disclaimer:
' This code is provided by SmaRTy Saini Corp.
' Use, modify, and distribute this code freely, but please
' retain the credit to SmaRTy Saini Corp. in any usage.
' Unauthorized redistribution for profit is prohibited.
' ************************************************************


Private Sub Application_ItemSend(ByVal Item As Object, Cancel As Boolean)
    On Error GoTo ErrorHandler

    Dim mail As Outlook.mailItem
    Dim currentSubject As String
    Dim threadCount As Integer
    Dim subj As String
    Dim existingCount As Integer

    ' Check if the item is an email
    If TypeName(Item) = "MailItem" Then
        Set mail = Item
        currentSubject = mail.subject

        ' Initialize thread count to 1 if no count is found
        threadCount = 1

        ' Check if the subject already contains a thread count (e.g., "[2]")
        If InStr(currentSubject, "[") > 0 And InStr(currentSubject, "]") > 0 Then
            ' Extract the existing thread count if it exists in the subject
            existingCount = ExtractThreadCount(currentSubject)
            threadCount = existingCount + 1 ' Increment the existing count by 1
        End If

        ' If subject starts with "RE:", modify the subject with the new thread count
        If Left(currentSubject, 3) = "RE:" Then
            ' If there is an existing thread count, replace it; otherwise, add a new count
            If InStr(currentSubject, "[") > 0 And InStr(currentSubject, "]") > 0 Then
                subj = Replace(currentSubject, "[" & existingCount & "]", "[" & threadCount & "]")
            Else
                subj = "RE: [" & threadCount & "] " & Mid(currentSubject, 5)
            End If
        Else
            ' If the subject doesn't start with "RE:", just prepend the thread count
            subj = "[" & threadCount & "] " & currentSubject
        End If

        ' Update the subject with the new thread count
        mail.subject = subj
    End If

    Exit Sub

ErrorHandler:
    MsgBox "An error occurred: " & Err.Description, vbCritical, "Error"
End Sub

' Function to extract the thread count from the existing subject line (if present)
Private Function ExtractThreadCount(subject As String) As Integer
    Dim startPos As Integer
    Dim endPos As Integer
    Dim threadCount As Integer

    ' Look for the first "[" and "]" in the subject to extract the number inside
    startPos = InStr(subject, "[") + 1
    endPos = InStr(subject, "]")
    
    If startPos > 0 And endPos > startPos Then
        ' Extract the number inside the brackets
        threadCount = CInt(Mid(subject, startPos, endPos - startPos))
    Else
        threadCount = 1 ' Default to 1 if no count is found
    End If
    
    ExtractThreadCount = threadCount
End Function

