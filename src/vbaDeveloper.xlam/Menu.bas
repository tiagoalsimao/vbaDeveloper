Attribute VB_Name = "Menu"
Option Explicit

Private Const MENU_TITLE = "VbaDeveloper"
Private Const XML_MENU_TITLE = "XML Import-Export"
Private Const MENU_REFRESH = "Refresh this menu"


Public Sub createMenu()

    'Add the top-level menu to the ribbon Add-ins section
    Dim rootMenu As CommandBarPopup
    Set rootMenu = Application.CommandBars(1).Controls.Add(Type:=msoControlPopup, _
            Before:=10, _
            Temporary:=True)
    rootMenu.caption = MENU_TITLE

    Dim exSubMenu As CommandBarPopup
    Dim imSubMenu As CommandBarPopup
    Dim formatSubMenu As CommandBarPopup
    Set exSubMenu = addSubmenu(rootMenu, 1, "Export code for ...")
    Set imSubMenu = addSubmenu(rootMenu, 2, "Import code for ...")
    Set formatSubMenu = addSubmenu(rootMenu, 3, "Format code for ...")
    
    addMenuSeparator rootMenu
    
    Dim refreshItem As CommandBarButton
    Set refreshItem = addMenuItem(rootMenu, "Menu.refreshMenu", MENU_REFRESH)
    refreshItem.FaceId = 37

    ' menuItem.FaceId = FaceId ' set a picture
    Dim vProject As Variant
    For Each vProject In Application.VBE.VBProjects
    
        ' We skip over unsaved projects where project.fileName throws error
        Dim project As VBProject
        Set project = vProject
        
        Dim projectName As String, caption As String
        projectName = project.name
        
        On Error Resume Next
            caption = projectName & " (" & Dir(project.fileName) & ")" '<- this can throw error
        If Err <> 0 Then GoTo nextProject

        Dim exCommand As String, imCommand As String, formatCommand As String
        exCommand = "'Menu.exportVbProject """ & project.fileName & """'"
        imCommand = "'Menu.importVbProject """ & project.fileName & """'"
        formatCommand = "'Menu.formatVbProject """ & project.fileName & """'"

        addMenuItem exSubMenu, exCommand, caption
        addMenuItem imSubMenu, imCommand, caption
        addMenuItem formatSubMenu, formatCommand, caption

nextProject:
        On Error GoTo 0 'reset the error handling
    Next vProject
    

    'Add menu items for creating and rebuilding XML files
    Dim xmlMenu As CommandBarPopup
    Set xmlMenu = Application.CommandBars(1).Controls.Add(Type:=msoControlPopup, _
            Before:=10, _
            Temporary:=True)
    xmlMenu.caption = XML_MENU_TITLE
    
    Dim exXmlSubMenu As CommandBarPopup
    Set exXmlSubMenu = addSubmenu(xmlMenu, 1, "Export XML for ...")
    
    Dim rebuildButton As CommandBarButton
    Set rebuildButton = addMenuItem(xmlMenu, "Menu.rebuildXML", "Rebuild a file")
    rebuildButton.FaceId = 35
    
    Set refreshItem = addMenuItem(xmlMenu, "Menu.refreshMenu", MENU_REFRESH)
    refreshItem.FaceId = 37

    'add menu items for all open files
    Dim openFile As Workbook
    For Each openFile In Application.Workbooks
        Dim fileName As String
        fileName = openFile.name
        Call addMenuItem(exXmlSubMenu, "'Menu.exportXML """ & fileName & """'", fileName)
    Next openFile

End Sub


Private Function addMenuItem(menu As CommandBarPopup, ByVal onAction As String, ByVal caption As String) As CommandBarButton
    Dim menuItem As CommandBarButton
    Set menuItem = menu.Controls.Add(Type:=msoControlButton)
    menuItem.onAction = onAction
    menuItem.caption = caption
    Set addMenuItem = menuItem
End Function


Private Function addSubmenu(menu As CommandBarPopup, ByVal position As Integer, ByVal caption As String) As CommandBarPopup
    Dim subMenu As CommandBarPopup
    Set subMenu = menu.Controls.Add(Type:=msoControlPopup)
    subMenu.onAction = position
    subMenu.caption = caption
    Set addSubmenu = subMenu
End Function


Private Sub addMenuSeparator(menuItem As CommandBarPopup)
    menuItem.BeginGroup = True
End Sub


'This sub should be executed when the workbook is closed
Public Sub deleteMenu()

    'For each control, check if its name matches the names of our custom menus - using this method deletes multiple instances of the menu in case duplicates are mistakenly created.
    Dim cbControl
    
    On Error Resume Next
        For Each cbControl In CommandBars(1).Controls               'TODO if more menus are added, should use a collection instead of multiple if statements (keep code DRY)
            If cbControl.caption = MENU_TITLE Then
                Debug.Print "Deleting" & MENU_TITLE
                cbControl.Delete
            End If
            If cbControl.caption = XML_MENU_TITLE Then
                Debug.Print "Deleting" & XML_MENU_TITLE
                cbControl.Delete
            End If
        Next cbControl
    On Error GoTo 0
    
End Sub

Public Sub refreshMenu()
    menu.deleteMenu
    menu.createMenu
End Sub

Public Sub exportVbProject(ByVal projectPath As String)

    On Error GoTo exportVbProject_Error

    Dim project As VBProject
    Set project = GetProjectByPath(projectPath)
    Build.exportVbaCode project
    
    Dim wb As Workbook
    Set wb = Build.openWorkbook(project.fileName)
    NamedRanges.exportNamedRanges wb
    MsgBox "Finished exporting code for: " & project.name

    Exit Sub
exportVbProject_Error:
    ErrorHandling.handleError "Menu.exportVbProject"
    
End Sub


Public Sub importVbProject(ByVal projectPath As String)

    On Error GoTo importVbProject_Error

    Dim project As VBProject
    Set project = GetProjectByPath(projectPath)
    Build.importVbaCode project
    
    Dim wb As Workbook
    Set wb = Build.openWorkbook(project.fileName)
    NamedRanges.importNamedRanges wb
    
    MsgBox "Finished importing code for: " & project.name

    On Error GoTo 0
    
    Exit Sub
importVbProject_Error:
    ErrorHandling.handleError "Menu.importVbProject"
    
End Sub


Public Sub formatVbProject(ByVal projectPath As String)

    On Error GoTo formatVbProject_Error

    Dim project As VBProject
    Set project = GetProjectByPath(projectPath)
    Formatter.formatProject project
    
    MsgBox "Finished formatting code for: " & project.name & vbNewLine _
    & vbNewLine _
    & "Did you know you can also format your code, while writing it, by typing 'application.Run ""formatActiveCodePane""' in the immediate window?"

    On Error GoTo 0
    Exit Sub
formatVbProject_Error:
    ErrorHandling.handleError "Menu.formatVbProject"
    
End Sub


Public Sub exportXML(ByVal fileShortName As String)

    'Ask them if they want to save the file first. Warn that existing files could be overwritten. Default to 'Cancel'
    Dim validateChoice As Integer, prompt As String, title As String
    prompt = "Are you sure you want to export " & fileShortName & " to XML? Any previously exported XML data for that file will be overwritten."
    title = "Overwrite existing XML?"
    validateChoice = MsgBox(prompt, vbYesNoCancel, title)
    If validateChoice = vbCancel Then Exit Sub

    prompt = "Do you want to save the file before exporting? If unsaved, the exported version will reflect only changes until your most recent save."
    title = "Save file first?"
    validateChoice = MsgBox(prompt, vbYesNoCancel, title)
    
    If validateChoice = vbCancel Then Exit Sub
    If validateChoice = vbYes Then
        Dim wkb As Workbook
        Set wkb = Workbooks(fileShortName)
        wkb.Save
    End If

    Call unpackXML(fileShortName)
    MsgBox ("File successfully exported to XML. Check the 'src' folder where the file is saved.")
    
End Sub


Public Sub rebuildXML()

    'This sub lets the user browse to a folder, sets the destination folder as two levels up the folder tree,
    'and then calls the 'rebuildXML' function to zip up the XML data into an Excel file
    Dim destinationFolder As String
    destinationFolder = "C:\"
    
    Dim containingFolderName As String
    containingFolderName = "C:\" ' TODO: Not used, to remove?
    
    'Select containing folder using file picker
    containingFolderName = GetFolder(destinationFolder) 'TODO: Not used, to remove?
    
    'Remove trailing slash if it exists
    containingFolderName = XMLexporter.removeSlash(containingFolderName)

    'destinationFolder is two levels up from the containing folder
    On Error GoTo folderError
        destinationFolder = containingFolderName
        destinationFolder = Left(destinationFolder, Len(destinationFolder) - (Len(destinationFolder) - InStrRev(destinationFolder, "\") + 1)) 'up one level
        destinationFolder = Left(destinationFolder, Len(destinationFolder) - (Len(destinationFolder) - InStrRev(destinationFolder, "\") + 1)) 'up another level
    On Error GoTo 0

    Dim errorFlag As Boolean
    errorFlag = False
    
    Dim errorMessage As String
    Call XMLexporter.rebuildXML(destinationFolder, containingFolderName, errorFlag, errorMessage)

folderError:
    
    If Err.Number <> 0 Then
        errorFlag = True
        errorMessage = "That's not a valid folder"
    End If

    'Report the status to the user
    If errorFlag = True Then
        MsgBox (errorMessage)
    Else
        MsgBox ("File succesfully rebuilt to here: " & vbCrLf & destinationFolder)
    End If

End Sub


Function GetFolder(InitDir As String) As String

    Dim sItem As String
    sItem = InitDir
    
    Dim fldr As FileDialog
    Set fldr = Application.FileDialog(msoFileDialogFolderPicker)
    
    With fldr
        .title = "Select a Folder"
        .AllowMultiSelect = False
        If Right(sItem, 1) <> "\" Then
            sItem = sItem & "\"
        End If
        .InitialFileName = sItem
        If .Show <> -1 Then
            sItem = InitDir
        Else
            sItem = .SelectedItems(1)
        End If
    End With
    
    GetFolder = sItem
    Set fldr = Nothing
    
End Function


Function GetProjectByPath(ByVal projectPath As String) As VBProject

    'Simple search to find project by file path
    Dim project As VBProject
    For Each project In Application.VBE.VBProjects
    
        On Error GoTo skipone
        If UCase(project.fileName) = UCase(projectPath) Then
            Set GetProjectByPath = project
            Exit Function
        End If
        
nextprj:
    Next project
    
    'If not found return nothing
    Exit Function
    
skipone:
    Resume nextprj
    
End Function

Public Sub teste()
'Purpose    :
'Author     :
'Description:
'Date       : 22out21
    Application.Cursor = xlWait
    
    Application.Cursor = xlDefault
End Sub
