Public Sub SepareDrafts() 
     

    Dim lDraftItem As Long 
    Dim myOutlook As Outlook.Application 
    Dim myNameSpace As Outlook.NameSpace 
    Dim myFolders As Outlook.Folders 
    Dim myDraftsFolder As Outlook.MAPIFolder 
    Dim objMailMessage As Outlook.MailItem 
    Dim emlBody, sendTo As String 
    Dim TOs 
     
    Set myOutlook = Outlook.Application 
    Set myNameSpace = myOutlook.GetNamespace("MAPI") 
    Set myFolders = myNameSpace.Folders 
    Set myDraftsFolder = myNameSpace.PickFolder 
     
    For lDraftItem = myDraftsFolder.Items.Count To 1 Step -1 
        TOs = Split(myDraftsFolder.Items.Item(lDraftItem).To, ";") 
        For i = 0 To UBound(TOs) 
            Set objMailMessage = myOutlook.CreateItem(0) 
            With objMailMessage 
                .To = TOs(i) 
                .Body = myDraftsFolder.Items.Item(lDraftItem).Body 
                .Subject = myDraftsFolder.Items.Item(lDraftItem).Subject 
                .Display 
                .Send 
            End With 
        Next 
    Next lDraftItem 
    Set myDraftsFolder = Nothing 
    Set myNameSpace = Nothing 
    Set myOutlook = Nothing 

End Sub
