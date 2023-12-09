VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pathLogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'' ********* Here's the CLASS MODULE '''
'' ********* call this pathLogger.cls

'VB5 rewrite (8/31/00)
' * PROGRAMMER  : Nicholas W. Saparoff
' Use collections instead of arrays
' Added a new feature which gives the full filename back too
'
'VB6 standard module [PathLog] 12-25-99 9:00 >1-19-00 15:09
' * PROGRAMMER  : John Almand
' * WEB SITE    : WWW.SHALLIZAR.COM
' * E-MAIL      : WEBMASTER@SHALLIZAR.COM
'
'This procedure has been excerpted from the PathLogger.bas
'module of the MVIEW project. It is the core algorithm of
'Pathlogger.bas. It does not have the functionality and
'features of the full PathLogger module. For more information
'on the complete PathLogger.bas module, contact
'www.shallizar.com
'
'PURPOSE:
' LOG THE FILES IN A BRANCH OF A DIRECTORY TREE
'
' Elucidation:
' Make an array of strings of the names of all the files
' in a folder and all of its subfolders.
' Make a seperate array of strings of the subfolder names,
' including the parent folder.
' Make an array of pointers linking the two string arrays.
'
'SPECIAL CHARACTERISTICS:
' NON-RECURSIVE
' This procedure iterates through the subfolders without
' repeatedly calling itself. It avoids the extra overhead
' and potential stack overflow of recursive procedures.
'
'REQUIRES:
' Reference to Microsoft Scripting Runtime
'
'#########################################################
Option Explicit

Private objFSO As FileSystemObject
Private objFiles As Files, objFile As File
Attribute objFile.VB_VarUserMemId = 1073938433
Private objFolders As Folders, objFolder As Folder
Attribute objFolders.VB_VarUserMemId = 1073938435
Attribute objFolder.VB_VarUserMemId = 1073938435
Private lngFNAMEScntr As Long
Attribute lngFNAMEScntr.VB_VarUserMemId = 1073938437
Private collFnames As New Collection
Attribute collFnames.VB_VarUserMemId = 1073938438
Private collLngPtr As New Collection
Attribute collLngPtr.VB_VarUserMemId = 1073938439
Private collStrPaths As New Collection
Attribute collStrPaths.VB_VarUserMemId = 1073938440
Private collPathFileNames As New Collection
Attribute collPathFileNames.VB_VarUserMemId = 1073938441
Private useFullOnly As Boolean
Attribute useFullOnly.VB_VarUserMemId = 1073938442

Public Sub PathLogInit()
'PURPOSE: Create & initialize objFSO as a
'FileSystemObject object.
'Call this procedure from the LOAD event of the Startup form.
    Set objFSO = New FileSystemObject

    useFullOnly = True

    Set collFnames = Nothing
    Set collLngPtr = Nothing
    Set collStrPaths = Nothing
    Set collPathFileNames = Nothing

End Sub

Public Sub LogPath(strPARENT As String)
'PURPOSE: Make a list of all the files in a Folder and all of
'its subfolders.

'Typically used with a DirListBox control on a form. A command
'button should be provided to run this procedure. After selecting
'a folder in the DirListBox, the user would click the button to
' log the folder and its subfolders (which is the Path).

'The calling form or module passes a string of the path,
'a variable array of strings for the subfolder names, a
'variable array of long integers for the pointers to the
'folder names, & a variable array of strings for the filenames.

'EXAMPLE:
' LogPath Dir1.Path, arrayFolders, arrayPointers, arrayFilenames

'The result in arrayFilenames can be used to fill a listbox.
'A filename selected in a listbox can be fully referenced by
'combining it with its parent folder name,
' thus: arrayFolders(arrayPointers(I)) & arrayFilenames(I)

    Dim lngTopIndex As Long
    Dim lngPathIndex As Long
    Dim strNextPath As String
    Dim tmpPathStr As String

    'if path is invalid, exit
    If Not objFSO.FolderExists(strPARENT$) Then Exit Sub

    ' "seed" the loop
    lngTopIndex& = 1
    lngPathIndex& = 1
    lngFNAMEScntr& = 1

    collStrPaths.Add IFBACKSLASH(strPARENT$), "K0"

    Do

        ' Add subfolders, if any, to array
        Set objFolders = _
        objFSO.GetFolder(collStrPaths(lngPathIndex&)).SubFolders
        For Each objFolder In objFolders
            lngTopIndex& = lngTopIndex& + 1
            collStrPaths.Add collStrPaths(lngPathIndex&) & objFolder.Name & "\", "K" & lngTopIndex&

        Next

        ' Add filenames, if any, to array
        tmpPathStr$ = collStrPaths(lngPathIndex&)

        Set objFiles = objFSO.GetFolder(tmpPathStr$).Files

        If useFullOnly = True Then

            For Each objFile In objFiles
                collPathFileNames.Add tmpPathStr$ & objFile.Name, "K" & lngFNAMEScntr&
                lngFNAMEScntr& = lngFNAMEScntr& + 1
            Next

        Else

            For Each objFile In objFiles
                collFnames.Add objFile.Name, "K" & lngFNAMEScntr&
                collPathFileNames.Add tmpPathStr$ & objFile.Name, "K" & lngFNAMEScntr&
                collLngPtr.Add lngPathIndex&, "K" & lngFNAMEScntr&
                lngFNAMEScntr& = lngFNAMEScntr& + 1
            Next

        End If



        ' Point to next entry in subfolder array
        lngPathIndex& = lngPathIndex& + 1

        ' If there are no more subfolders, exit
    Loop Until lngPathIndex& > lngTopIndex&

End Sub

Public Property Get Count() As Long
' Get the number of filenames found
    Count = lngFNAMEScntr
End Property

Public Sub Terminate()
' Clear all the objects from memory
    Set objFSO = Nothing
    Set objFiles = Nothing
    Set objFile = Nothing
    Set objFolders = Nothing
    Set objFolder = Nothing
    Set collFnames = Nothing
    Set collStrPaths = Nothing
    Set collLngPtr = Nothing
    Set collPathFileNames = Nothing

End Sub

Private Function IFBACKSLASH(strX As String) As String
' function for fixing the DOS path of a root directory
    IFBACKSLASH = IIf(Right(strX, 1) = "\", strX, strX & "\")
End Function
Public Property Let fullFileNamesOnly(ByVal fullPath As Boolean)

    useFullOnly = fullPath

End Property
Public Property Get fullFileNamesOnly() As Boolean

    fullFileNamesOnly = useFullOnly

End Property
Public Property Get fileNames() As Collection

    Set fileNames = collFnames

End Property

Public Property Get pathNames() As Collection

    Set pathNames = collStrPaths

End Property

Public Property Get pathPointers() As Collection

    Set pathPointers = collLngPtr

End Property

Public Property Get fullFileNames() As Collection

    Set fullFileNames = collPathFileNames

End Property



