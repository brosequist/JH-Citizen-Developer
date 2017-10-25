Attribute VB_Name = "ArrayToDistributionList"
Sub SendToDistributionList(EmailList() As String, EmailSubject As String)

    'Send the active workbook to the distribution list defined on the template;
    'close the workbook after sendmail process is complete

    With ActiveWorkbook
        .SendMail Recipients:=EmailList, Subject:=EmailSubject
    End With

End Sub
