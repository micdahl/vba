VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Const gitExecutable = "C:\Program Files (x86)\Git\bin\git.exe"
Const gitStatusParameters = "status --porcelain"
Const gitShowParameters = "log --oneline"
Const notVersionedString = "0"

' Run git to get documents version and store it to document variable "fileVersion"
' Only commited versions will be seen as revisioned
' notVersionedString will be used if document is not controlled by git
' or if document has changed after last commit
' Path to git executable is given in  Const gitExecutable
Private Sub actualizeVersion()
'could you please add actual object name instead of object
    Dim fileVersion As String
    Dim FileName As String
    Dim objShell As Object
    Dim objExec As Object
    Dim aVar As Variable
    Dim index As Long
    Dim status As String
    
    index = 0
    ' if document variable "fileVersion" not exists, add it,
    ' set aVar to this document variable
    For Each aVar In ActiveDocument.Variables
        If aVar.Name = "fileVersion" Then
            index = aVar.index
        End If
    Next aVar
            
    If index = 0 Then
        Set aVar = ActiveDocument.Variables.Add("fileVersion")
    Else
        Set aVar = ActiveDocument.Variables(index)
    End If
    
    If ActiveDocument.Path = "" Then
        ' document is not saved anywhere so even not versioned
        fileVersion = notVersionedString
    Else
        ' get status of file (no return if file is versioned and has not changed until last commit)
        FileName = ActiveDocument.Path & "\" & ActiveDocument.Name
        Set objShell = CreateObject("WScript.Shell")
        Set objExec = objShell.Exec(gitExecutable & " " & gitStatusParameters & " " & FileName)
        status = objExec.Stdout.ReadLine
        If status <> "" Then
            ' file has been modified until last commit or is not versioned at all
            fileVersion = notVersionedString
        Else
            ' get revision from git
            Set objExec = objShell.Exec(gitExecutable & " " & gitShowParameters & " " & FileName)
            fileVersion = objExec.Stdout.ReadLine
            If fileVersion = "" Then
                fileVersion = notVersionedString
            Else
                fileVersion = Left(fileVersion, InStr(1, fileVersion, " "))
            End If
        End If
    End If

    aVar.Value = fileVersion
    
    ' update all fields of the document
    ActiveDocument.Fields.Update
    ' trick to update fields in headline and footer
    ActiveDocument.PrintPreview
    ActiveDocument.ClosePrintPreview
    
End Sub

Private Sub Document_Open()

    Call actualizeVersion
    
End Sub


