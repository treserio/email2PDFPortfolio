'List of required references: "Visual Basic For Applications", "Microsoft Outlook 16.0 Object Library", "OLE Automation", "Microsoft Office 16.0 Object Library", "Microsoft VBScript Regular Expressions 5.5", "Microsoft Scripting Runtime", "Microsoft Word 16.0 Object Library", "Microsoft Forms 2.0 Object Library", "Adobe Acrobat 10.0 Type Library", "AdobePDFMakerForOffice" <-- attempting to view the portfolio object.

Private Declare Function SetCurrentDirectoryA Lib "kernel32" (ByVal lpPathName As String) As Long

'Function to use the UNC path for a shared drive, so you don't run into issues where users have them mapped to seperate letters.
Function setUNCPath(dPath As String) As Long
    Dim pathVal As Long
    pathVal = SetCurrentDirectoryA(dPath)
    setUNCPath = pathVal
End Function

'Clean Name of any chars that'll keep it from saving
Function CleanName(dirtyText As String) As String
    'Chars that cause errors when saving
    Dim cleaner As String: cleaner = "/\[]:=,(" & Chr(34)
    Dim lngth As Integer: lngth = Len(cleaner)
    Dim cntr As Integer
    'Remove leading and trailing spaces from dirtyText
    dirtyText = Trim(dirtyText)
    'iterate through dirtyText replacing each char in cleaner with ""
    For cntr = 1 To lngth
        dirtyText = Replace(dirtyText, Mid(cleaner, cntr, 1), "")
    Next
    'output cleanName
    CleanName = dirtyText
End Function

'Return 2D array of folder keys and folder names, keys acquired through regex looking for patterns #.### or ##.### or ###.###
Function findFolders(ByRef dPath As String) As Variant
    'Create an instance of the FileSystemObject
    Dim objFSO As Object
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    'Get the folder object
    Dim objFolder As Object
    'Final path "\\HOLDENDATA\Common\Cases\"
    Set objFolder = objFSO.GetFolder(dPath)
    'Create 2D Array for storage along with collections for specific text and full folder name
    Dim storageArray(1) As Collection
    Dim fldrKey As New Collection
    Dim fullName As New Collection
        Set storageArray(0) = fldrKey
        Set storageArray(1) = fullName
    'Regex object and pattern to identify HM #s in folder names and email subjects
    Dim regX As New RegExp
    regX.Pattern = "\d+\.\d+"
    'Object to store the results
    Dim regXResult As Object

    'For every subfolder: identify regex pattern in text and push matches to fldrKey, then folder's full name to fullName that are both in storageArray
    For Each objSubFolder In objFolder.subfolders
        'add each folder name to the fullName collection
        fullName.Add objSubFolder.Name
        'Check if regex can run, then assign matches to fldrKey, if no matches exist add NA
        If regX.Test(objSubFolder.Name) Then
            Set regXResult = regX.Execute(objSubFolder.Name)
            fldrKey.Add regXResult(0).Value
        Else
            fldrKey.Add "noMtchFnd"
        End If
    Next objSubFolder

    findFolders = storageArray()
End Function

'Return 2D array where array(0) is the parentFolderNames and array(1) is a 2D array of [Keys][subFolderNames], accepts 2D array as input
Function findSubFolders(ByRef pFolders As Variant, dPath As String) As Variant
    
    'Hold subfolder keys and names returned from findFolders
    Dim tempArray As Variant
    'Hold pfolder collection and collection of subfolder keys & names
    Dim Container(1) As Collection
    Dim pFldr As New Collection
    Dim keysNPath As New Collection
    'Add arrays to Container
    Set Container(0) = pFldr
    Set Container(1) = keysNPath
    
    For Each fParent In pFolders(1)
        tempArray = findFolders(dPath & fParent & "\")
        pFldr.Add fParent
        keysNPath.Add tempArray
    Next fParent

    findSubFolders = Container
End Function

'correct namespace.store "FileRoom@holdenlitigation.com"
'Access shared email accounts and match subject text to keys from findFolders to determine correct path
Sub emailIdent()
    'Open Outlook Application
    Dim olApp As Outlook.Application
    Set olApp = Outlook.Application
    'Open Word Application to pass to saveAsPDF() so that you only have to use one instance of word.
    Dim wrdApp As Word.Application
    Set wrdApp = CreateObject("Word.Application")
    'I tried opening one instance of Acrobat to pass in to do everything, but it keeps opening in acrobat reader.
    'Open Acrobat Application to pass to pdf manipulation, saveAsPDF() & import2PDF(), so that you only have to use one instance of Acrobat.
    Dim acroApp As Object
    Set acroApp = CreateObject("AcroExch.App")
    'Set namespace for mail to mapi, unsure exactly what this is doing but it works
    Dim olNamespace As Outlook.NameSpace: Set olNamespace = GetNamespace("MAPI")
    'Create Shared folder variable with name of share, and set olFolder to the default shared folder
    Dim olShareName As Outlook.Recipient: Set olShareName = olNamespace.CreateRecipient("FileRoom@holdenlitigation.com")
    Dim olFolder As Outlook.MAPIFolder: Set olFolder = olNamespace.GetSharedDefaultFolder(olShareName, olFolderInbox)
    'Create variable for directory path, uneccessary but can later be used to collect user input if variable paths are required, or to set in app settings.
    Dim pathString As String: pathString = "\\HOLDENDATA\Common\Cases\"
    'Create comparison array with folder keys and names of Open cases
    Dim openKeys As Variant: openKeys = findFolders(pathString)
    'Use base 'Closed File' folder to pull yearly folders, for each year folder, pull 2D array of keys and folder names to add to the year folder's position in the 2D parent array. parentString = year folder, parentString + clsdKeys(1) = sub folders 2D array of path and key.
    Dim parentString As String: parentString = "\\HOLDENDATA\archive\Closed Files\"
    Dim clsdTemp As Variant: clsdTemp = findFolders(parentString)
    'folders will be returned with [key][fullName], then overwrite [key] with 2D array of subfolder[key][fullName]
    'for each parent folder, clsdKeys(1) loop through clsdKeys(0)(0)(i) for keys clsdKeys(0)(1)(i) for fullNames
    'also setup the final array for clsdKeys
    Dim clsdKeys As Variant: clsdKeys = findSubFolders(clsdTemp, parentString)
    'clear temp, just cause
    Set clsdTemp = Nothing
    'Variable to hold mail subject line
    Dim itemSubj As String
    'Variable for full folder path string
    Dim fullPath As String
    'For each email in the folder
    Dim item As Object
    For Each item In olFolder.Items
        'Make sure the item is an email
        If TypeOf item Is Outlook.MailItem Then
            'Pull Subject to compare with directory keys from findFolders()
            itemSubj = item.Subject
            'Loop through the keys vs the subject to find a match, at match use iterator to access associated matches folder path for Open Cases.
            Dim keyI As Integer: keyI = 1

            For Each key In openKeys(0)
                If InStr(itemSubj, key) > 0 Then
                    'How to access the correct path = openKeys(1)(keyI)
                    fullPath = "\\HOLDENDATA\Common\Cases\" & openKeys(1)(keyI) & "\2 Correspondence\E-mail Correspondence\"
                    'Uncomment below to confirm file path that is being used for save location
                    MsgBox (fullPath & "  :  " & itemSubj)
                    
                    'Call pdf creation sub
                    Call saveAsPDF(item, fullPath, wrdApp)
                    'Call sub to move email to archive
                    Call moveEmail(item, olFolder)
                Else
                    'Figure out what we're going to do with the leftovers, most likely do nothing and they'll live in the inbox, instead automate moving those that have been processed into the archive.
                End If
                keyI = keyI + 1
            Next key
            
            'reset counter for closed key looping
            keyI = 1
            'run through keys in clsdKeys HERE
            'extra iterator for clsdKeys loop
            Dim keyJ As Integer
            'Loop for accessing subfolder keys while keeping track of parent folders for use in saveAsPDF() for archived cases
            'unable to access the folder path for whatever Fed up reason, also Ubound for array length
            'Think I finally got the right setup on the loops for accessing the subfolder values in the 2D array in
            'clsdKeys(0) = parent folder array
            'clsdKeys(1)(# for nested key & name array)(0) = key value array
            'clsdKeys(1)(# for nested key & name array)(1)(# for array entry) = parent folder text
            For Each prntFldr In clsdKeys(0)
                keyJ = 1
                For Each key In clsdKeys(1)(keyI)(0)
                    'clsdKeys(1)/(array value)/(0=key array) (1=folderName object)/(1=folderName)
                    'MsgBox (prntFldr & "  " & key & "  " & clsdKeys(1)(keyI)(1)(keyJ))
                    If InStr(itemSubj, key) > 0 Then
                        fullPath = "\\HOLDENDATA\archive\Closed Files\" & prntFldr & "\" & clsdKeys(1)(keyI)(1)(keyJ) & "\2 Correspondence\E-mail Correspondence\"
                        MsgBox (fullPath)
                        'add check to see if the destination document is read only
                        If checkFolder(fullPath) = 0 Then
                            'Call pdf creation sub, that also calls the import function for attachments and the set values tab for the pdf portfolio fields
                            Call saveAsPDF(item, fullPath, wrdApp)
                            'Call sub to move email to archive
                            Call moveEmail(item, olFolder)
                        Else
                            MsgBox (fullPath & "  Does not exist")
                        End If
                    End If
                    keyJ = keyJ + 1
                Next key
                keyI = keyI + 1
            Next prntFldr
            
        End If
    Next item
    wrdApp.Quit
    olApp.Quit
    acroApp.Quit
End Sub

'Sub for saving email as pdf through word's functionality requires Folder Path String, File Name String build in checks for filename so it doesn't overwrite. Requires "Microsoft Scripting Runtime" & "Microsoft Word ### Object Library". Could check for directory to ensure it's there before proceeding, but we're pulling these folders already so we know they're there.
'  CHECK  "If Dir(dPath, vbDirectory) = vbNullString Then"
Sub saveAsPDF(email As MailItem, dPath As String, wrdApp As Word.Application)
    'use ChDir to set folder path to eliminate issues with 255 char name length
    setUNCPath ("C:\Windows\Temp\")
    'Use a format of "<received date & Time>, <subject line>, may need a check for if new file name exists, email.subject should be subject line, clean out erroneous chars from string with CleanName
    Dim olMHTML_Name As String: olMHTML_Name = Format(email.ReceivedTime, "mmddyy hhmmss ") & CleanName(email.Subject) & ".mht"
    Dim pdfName As String: pdfName = Format(email.ReceivedTime, "mmddyy hhmmss ") & CleanName(email.Subject) & ".pdf"
    'check if string lengths are longer than the maximum allowed in shared drive (160) and truncate if needed. Found that the move was failing with 160 dropped to 100 for more of a buffer.
    If Len(olMHTML_Name) > 100 Then
        olMHTML_Name = Left(olMHTML_Name, 100) & ".mht"
    End If
    If Len(pdfName) > 100 Then
        pdfName = Left(pdfName, 100) & ".pdf"
    End If
    'save email as olMHTML document in path folder
    email.SaveAs olMHTML_Name, olMHTML
    'Create document object to open file
    Dim wrdDoc As Word.Document
    'Save .mht document as pdf using pdfName as file name
    Set wrdDoc = wrdApp.Documents.Open(fileName:="C:\Windows\Temp\" & olMHTML_Name, Visible:=True)
        wrdApp.ActiveDocument.ExportAsFixedFormat OutputFileName:="C:\Windows\Temp\" & pdfName, ExportFormat:= _
        wdExportFormatPDF, OpenAfterExport:=False, OptimizeFor:=wdExportOptimizeForPrint, _
        Range:=wdExportAllDocument, From:=0, To:=0, item:=wdExportDocumentContent, _
        IncludeDocProps:=True, KeepIRM:=True, CreateBookmarks:=wdExportCreateNoBookmarks, _
        DocStructureTags:=True, BitmapMissingFonts:=True, UseISO19005_1:=False
    
    'Close document in word in order to clean up .mht file
    wrdDoc.Close

    'Clean up .mht file from folder
    With New FileSystemObject
        If .FileExists(olMHTML_Name) Then
            .DeleteFile olMHTML_Name
        End If
    End With
    
    Dim portfolioPath As String: portfolioPath = dPath & "E-Mail Corr.pdf"
    Dim emailPDFPath As String: emailPDFPath = "C:\Windows\Temp\" & pdfName
    'attachment counter for use later in setFieldValues
    Dim numAttached As Integer: numAttached = 0

    'For each attachment in the email copy to C:\Windows\Temp\ and then import into email pdf as attachments.
    For Each attached In email.Attachments
        'Use importDataObject() through JS library to import attachments from email to pdf before importing complete email&attachment pdf to pdf portfolio file.
        'Save attachments locally to use with import2PDF()
        Dim attachPath As String: attachPath = "C:\Windows\Temp\" & attached.fileName
        attached.SaveAsFile attachPath
        'Import the documents, objName = attached.FileName, importFilePath = attachPath, parentPDFPath = emailPDFPath
        Call import2PDF(attached.fileName, attachPath, emailPDFPath)
        'Clean up after the import is complete?
        With New FileSystemObject
            If .FileExists(attachPath) Then
                .DeleteFile attachPath
            End If
        End With
        numAttached = numAttached + 1
    Next attached

    'Variables for import2PDF() to import emailPDF into correct portfolio file
    'objName = pdfName, importFilePath = emailPDFPath, parentPDFPath = portfolioPath
    Call import2PDF(pdfName, emailPDFPath, portfolioPath)
    
    'Clean up the pdf file from the C drive
    With New FileSystemObject
        If .FileExists(emailPDFPath) Then
            .DeleteFile emailPDFPath
        End If
    End With
    'Now set the correct field values for the newly added dataObject in the portfolio file
    Call setFieldValues(pdfName, portfolioPath, email, numAttached)

End Sub

'Move email from inbox to Archive once converted to pdf and moved to correct case folder
Sub moveEmail(email As Outlook.MailItem, olFolder As Outlook.MAPIFolder)
    Set olDestFolder = olFolder.Parent.Folders("Archive")
    email.Move olDestFolder
End Sub

'Access acrobat's privileged javascript folder function safeImportDataObject through a jso object to add files to a pdf, pdf attachments, in a portfolio pdf these show up as the new items in the list and can be accessed through the dataObjects[#] or specifically by getDataO
'Acrobat JS privileged folder: C:\Users\<username>\AppData\Roaming\Adobe\Acrobat\Privileged\11.0\JavaScripts correct document is trustedFunc.js
'pass in 2 path strings to pdfs, open one, access jso object, likely the portfolio, then pass strings to safeImportDataObject for importing.

' should be called from the email to pdf conversion where it saved everything
' may not need to send in acroApp?
Sub import2PDF(objName As String, importFilePath As String, parentPDFPath As String)
    'Create Acrobat document to use for importing
    Dim acroDoc As Object
    Set acroDoc = CreateObject("AcroExch.PDDoc")
    
    If acroDoc.Open(parentPDFPath) = -1 Then
        'MsgBox ("Opened")
        Set jso = acroDoc.GetJSObject
        Call jso.safeImportDataObject(objName, importFilePath)
        'get a device-independent path to the document
        savePath = jso.Path
        'save through jso object using savePath string
        jso.SaveAs (savePath)
    Else
        MsgBox ("Unable to open: " + parentPDFPath)
    End If

    acroDoc.Close
End Sub

'Check if folderpath is available
Function checkFolder(dPath As String) As Integer
    If Dir(dPath, vbDirectory) <> vbNullString Then
        checkFolder = 0
    Else
        checkFolder = 1
    End If
End Function

Sub setFieldValues(pdfName As String, portfolioPath As String, email As Outlook.MailItem, numAttached As Integer)
    'Create Acrobat document to use for importing
    Dim acroDoc As Object
    Set acroDoc = CreateObject("AcroExch.PDDoc")
    
    If acroDoc.Open(portfolioPath) = -1 Then
        'MsgBox ("Opened")
        Set jso = acroDoc.GetJSObject
    Else
        MsgBox ("Unable to open: " + parentPDFPath)
    End If

    'coerce all into strings, except date, see if it'll take.
    Dim emailSender As String: emailSender = email.Sender
    Dim emailTo As String: emailTo = email.To
    Dim emailCC As String: emailCC = email.CC
    Dim emailSub As String: emailSub = email.Subject
    Dim emailRcvd As Date: emailRcvd = email.ReceivedTime
    
    'Call privileged field import JS function
    Call jso.setDataObjFieldValues(pdfName, emailSender, emailTo, emailCC, emailSub, emailRcvd, numAttached)
    
    'get a device-independent path to the document
    savePath = jso.Path
    'save through jso object using savePath string
    jso.SaveAs (savePath)
    
    acroDoc.Close
End Sub