Attribute VB_Name = "Module1"
Sub SaveAttachmentsFromSelectedEmails()
    Dim objSelection As Outlook.Selection
    Dim Item As Object
    Dim Atmt As Attachment
    Dim FilePath As String
    Dim i As Integer
    
    ' Define the folder path where attachments should be saved
    FilePath = "C:\Users\gaujat\Documents\Data Analytics\MITvsPR_Utitilty_SA360\ZNew folder\" ' Replace with your desired folder path
    
    ' Create the folder if it doesn't exist
    If Dir(FilePath, vbDirectory) = "" Then
        MkDir FilePath
    End If
    
    ' Get the selected emails
    Set objSelection = Application.ActiveExplorer.Selection
    
    i = 0
    
    ' Loop through each selected email
    For Each Item In objSelection
        If Item.Attachments.Count > 0 Then
            For Each Atmt In Item.Attachments
                ' Save each attachment with its original file name in the specified folder
                Atmt.SaveAsFile FilePath & Atmt.FileName
                i = i + 1
            Next Atmt
        End If
    Next Item
    
    MsgBox i & " attachments were saved to " & FilePath, vbInformation, "Task Completed"
End Sub

