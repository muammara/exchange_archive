Attribute VB_Name = "Module1"
Sub SaveAttachements()
Dim objOL As Outlook.Application
Dim objMsg As Outlook.MailItem 'Object
Dim objAttachments As Outlook.Attachments
Dim objSelection As Outlook.Selection
Dim i As Long
Dim lngCount As Long
Dim strFile As String
Dim strFolderpath As String
Dim strDeletedFiles As String





' Get the path to your My Documents folder
strFolderpath = "C:\Exchange"
On Error Resume Next

' Instantiate an Outlook Application object.
Set objOL = CreateObject("Outlook.Application")

' Get the collection of selected objects.
Set objSelection = objOL.ActiveExplorer.Selection

' Set the Attachment folder.
strFolderpath = "C:\Exchange\Attachments\"

' Check each selected item for attachments. If attachments exist,
' save them to the strFolderPath folder and strip them from the item.
For Each objMsg In objSelection

    ' This code only strips attachments from mail items.
    ' If objMsg.class=olMail Then
    ' Get the Attachments collection of the item.
    Set objAttachments = objMsg.Attachments
    lngCount = objAttachments.Count
    strDeletedFiles = ""

    If lngCount > 0 Then

        ' We need to use a count down loop for removing items
        ' from a collection. Otherwise, the loop counter gets
        ' confused and only every other item is removed.

        For i = lngCount To 1 Step -1

            ' Save attachment before deleting from item.
            ' Get the file name.
            strFile = objAttachments.Item(i).FileName
            'MsgBox strFile
            
            If InStr(1, strFile, "png", vbTextCompare) Or InStr(1, strFile, "jpg", vbTextCompare) Or InStr(1, strFile, "gif", vbTextCompare) Then
                  objAttachments.Item(i).Delete
                GoTo JPG
                
                
            End If
            
            
            

            ' Combine with the path to the Temp folder.
            strFile = strFolderpath & strFile

            ' Save the attachment as a file.
            objAttachments.Item(i).SaveAsFile strFile

            ' Delete the attachment.
            objAttachments.Item(i).Delete

            'write the save as path to a string to add to the message
            'check for html and use html tags in link
            If objMsg.BodyFormat <> olFormatHTML Then
                strDeletedFiles = strDeletedFiles & vbCrLf & "<file://" & strFile & ">"
            Else
                strDeletedFiles = strDeletedFiles & "<br>" & "<a href='file://" & _
                strFile & "'>" & strFile & "</a>"
            End If

            'Use the MsgBox command to troubleshoot. Remove it from the final code.
            'MsgBox strDeletedFiles
JPG:
        
        Next i

        ' Adds the filename string to the message body and save it
        ' Check for HTML body
        If objMsg.BodyFormat <> olFormatHTML Then
            objMsg.Body = vbCrLf & "The file(s) were saved to " & strDeletedFiles & vbCrLf & objMsg.Body
        Else
            objMsg.HTMLBody = "<p>" & "The file(s) were saved to " & strDeletedFiles & "</p>" & objMsg.HTMLBody
        End If
        objMsg.Save
    End If
Next

ExitSub:

Set objAttachments = Nothing
Set objMsg = Nothing
Set objSelection = Nothing
Set objOL = Nothing



End Sub
Sub Orders()

Dim xlApp As Object
Dim xlWB As Object
Dim xlSheet As Object
Dim olItem As Outlook.MailItem
Dim vText As Variant
Dim sText As String
Dim vItem As Variant
Dim i As Long
Dim rCount As Long
Dim bXStarted As Boolean
Const strPath As String = "C:\temp\Orders.xlsx" 'the path of the workbook

If Application.ActiveExplorer.Selection.Count = 0 Then
     MsgBox "No Items selected!", vbCritical, "Error"
     Exit Sub
End If
On Error Resume Next
Set xlApp = GetObject(, "Excel.Application")
If Err <> 0 Then
     Application.StatusBar = "Please wait while Excel source is opened ... "
     Set xlApp = CreateObject("Excel.Application")
     bXStarted = True
End If
On Error GoTo 0
'Open the workbook to input the data
Set xlWB = xlApp.Workbooks.Open(strPath)
Set xlSheet = xlWB.Sheets("Sheet1")

'Process each selected record

  rCount = xlSheet.UsedRange.Rows.Count
       'Insert the title
     rCount = 4
     xlSheet.Range("A" & rCount) = "Product Quote Number"
     xlSheet.Range("B" & rCount) = "Currency"
     xlSheet.Range("C" & rCount) = "Value"
     'xlSheet.Range("C" & rCountformat
     xlSheet.Range("D" & rCount) = "Customer"
     xlSheet.Range("E" & rCount) = "Partner"
     xlSheet.Range("F" & rCount) = "City"
     xlSheet.Range("G" & rCount) = "Time"

  
  
   For Each olItem In Application.ActiveExplorer.Selection
     If InStr(1, olItem.Subject, "CRMDBP1 - ZA", vbTextCompare) Then GoTo CONT: 'Remove SA
      If InStr(1, olItem.Subject, "CRMDBP1 - AF", vbTextCompare) Then GoTo CONT: 'Remove SA
     sText = olItem.Body
     vText = Split(sText, Chr(13))

     
      rCount = rCount + 1
     'Check each line of text in the message body
     For i = UBound(vText) To 0 Step -1
         
         
         
         If InStr(1, vText(i), "Product Quote Number:") > 0 Then
             vItem = Split(vText(i), Chr(58))
             xlSheet.Range("A" & rCount) = Trim(vItem(1))
         End If

         If InStr(1, vText(i), "Quote Currency:") > 0 Then
             vItem = Split(vText(i), Chr(58))
             xlSheet.Range("B" & rCount) = Trim(vItem(1))
         End If

         If InStr(1, vText(i), "Quote Value (Including Freight):") > 0 Then
             vItem = Split(vText(i), Chr(58))
             xlSheet.Range("C" & rCount) = Trim(vItem(1))
            xlSheet.Range("C" & rCount).NumberFormat = "$#,##0.00_);($#,##0.00)"
            
         End If

         If InStr(1, vText(i), "Install At Address:") > 0 Then
             vItem = Split(vText(i), Chr(58))
             vItem = Split(vItem(1), ",")
             xlSheet.Range("D" & rCount) = Trim(vItem(0))
             
            On Error Resume Next
              xlSheet.Range("F" & rCount) = Trim(vItem(5))
            
         
         End If

         If InStr(1, vText(i), "Ship To Address:") > 0 Then
             vItem = Split(vText(i), Chr(58))
             vItem = Split(vItem(1), ",")
             xlSheet.Range("E" & rCount) = Trim(vItem(0))
         End If
         
         xlSheet.Range("G" & rCount) = olItem.ReceivedTime

     Next i
     xlWB.Save
CONT:

 Next olItem
xlWB.Close SaveChanges:=True
If bXStarted Then
     xlApp.Quit
End If
Set xlApp = Nothing
Set xlWB = Nothing
Set xlSheet = Nothing
Set olItem = Nothing


End Sub


