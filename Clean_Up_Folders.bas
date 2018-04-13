Attribute VB_Name = "Clean_Up_Folders"

Sub CleanFolders()

    Dim LDate As String             'Declare the current Date
    
    LDate = Date
    createLogDelete ("------------------------------------------------------------------")
    createLogDelete (LDate)
    createLogDelete ("------------------------------------------------------------------")
    
    'Module5.DeleteEmailFromFolder ("Deleted Items")  'Clean Deleted Items folder
    
    'Module5.createLogDelete ("##################################################################")
    
    DeleteEmailFromFolder ("FMS")  'Clean FMS folder
    
    createLogDelete ("##################################################################")
    
    DeleteEmailFromFolder ("OpenView")  'Clean OpenView folder
       
    createLogDelete ("##################################################################")
       
    DeleteEmailFromFolder ("RUBBISH")  'Clean RUBBISH folder
       
    createLogDelete ("##################################################################")
       
    DeleteEmailFromFolder ("SERV.REQUESTS") 'Clean SERV.REQUESTS folder
    
    createLogDelete ("##################################################################")
    
    DeleteEmailFromFolder ("test")  'Clean test folder
    
    createLogDelete ("##################################################################")

    DeleteEmailFromFolder ("test1") 'Clean test1 folder
    
    'MsgBox ("All folders have been cleaned up.")
    
    createLogDelete ("------------------------------------------------------------------")
    createLogDelete ("------------------------------------------------------------------")
    
    UserForm_Dogs.Show 'This is how the picture is displayed
    
    
End Sub

Sub DeleteEmailFromFolder(ByVal nameFile As String)

    'Declare some Variables
    Dim Msg As Outlook.MailItem
    Dim objNS As Outlook.NameSpace
    Dim objFolder As Outlook.MAPIFolder
    Dim myItems As Outlook.Items
    Dim title As String
    Dim howMany As Integer
    Dim mySpace As String
    
    Set objNS = GetNamespace("MAPI")
    Set objFolder = objNS.Folders("E.MAZARAKIS@wind.gr")    'Folders of your account
    Set objFolder = objFolder.Folders(nameFile)             'Specified the folder
    
    Set myItems = objFolder.Items                           'Returns an Items collection as a collection of Microsoft Outlook items in the specified folder.
    howMany = myItems.count                                 'Count the number of e-mails in the specified folder
    'MsgBox (nameFile & " contains: " & CStr(howMany))
    mySpace = "                        "
    createLogDelete (mySpace & nameFile & " contains: " & CStr(howMany))
    createLogDelete ("..................................................................")
    
    For i = howMany To 1 Step -1        ' For all the e-mails on the specified folder
        Set Msg = myItems.Item(i)
        title = Msg.Subject
        createLogDelete (title & " " & CStr(i))
        Msg.Delete      'Deleting the message
    Next

End Sub


Sub createLogDelete(ByVal line As String)
'Write a line  to a text file

    Dim logFile As String
    logFile = "C:\Users\e.mazarakis\Desktop\LogDeleteMails.txt"   'It contains the path of the log File
    
    Open logFile For Append As #1
        'To do the actual writing to the file you need this:
        Write #1, line
    Close #1     'You have to close the file
    
End Sub
