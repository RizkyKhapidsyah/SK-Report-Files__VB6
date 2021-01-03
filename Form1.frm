VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Report Files on drive selected"
   ClientHeight    =   3540
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3540
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   315
      Left            =   150
      TabIndex        =   2
      Top             =   3150
      Width           =   4440
   End
   Begin VB.CommandButton cmdBuildReport 
      Caption         =   "&Build Report"
      Height          =   315
      Left            =   150
      TabIndex        =   1
      Top             =   975
      Width           =   4440
   End
   Begin VB.DriveListBox DriveDisk 
      Height          =   315
      Left            =   150
      TabIndex        =   0
      Top             =   450
      Width           =   2115
   End
   Begin VB.Label Label1 
      Caption         =   "Select a drive:"
      Height          =   165
      Left            =   150
      TabIndex        =   4
      Top             =   150
      Width           =   1215
   End
   Begin VB.Label lblPath 
      Height          =   765
      Left            =   150
      TabIndex        =   3
      Top             =   1425
      Width           =   4440
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ---------------------------------------------------------------------------
' Define Public variables
' ---------------------------------------------------------------------------
  Dim cnn              As ADODB.Connection
  Dim blnFirstRecord   As Boolean



  Private Sub cmdExit_Click()
  Set cnn = Nothing
  Unload Me

  End Sub


Private Sub cmdBuildReport_Click()
' ---------------------------------------------------------------------------
' Set mouse pointer
' ---------------------------------------------------------------------------
  Screen.MousePointer = vbHourglass
  
' ---------------------------------------------------------------------------
' Disable the exit command button
' ---------------------------------------------------------------------------
  cmdExit.Enabled = False
  
' ---------------------------------------------------------------------------
' Delete all records from database
' ---------------------------------------------------------------------------
  Call DeleteRecords
  
' ---------------------------------------------------------------------------
' Calls the routine
' ---------------------------------------------------------------------------
  Call Get_FileList("" & Left(DriveDisk.Drive, 2) & "\", True, "*.*")

  
' ---------------------------------------------------------------------------
' Define local variables
' ---------------------------------------------------------------------------
  Dim strSQL          As String
  Dim rsFiles         As ADODB.Recordset
  Dim strPath         As String
  Dim strOldPath      As String
  Dim SumSize         As Double
  On Error Resume Next
  

' ---------------------------------------------------------------------------
' Builds the select query, with a group by file extension clause
' ---------------------------------------------------------------------------
  strSQL = "Select  Path, Ext,count(Ext) as FilesNumber, Sum(Size) as CSize  From tblFiles"
  strSQL = strSQL & " Group By Path, Ext"
  
' ---------------------------------------------------------------------------
' Instantsiates recordset object, sets the active connection and source
' ---------------------------------------------------------------------------
  Set rsFiles = New ADODB.Recordset
  rsFiles.Open strSQL, cnn, adOpenDynamic, adLockOptimistic
  
  If rsFiles.BOF And rsFiles.EOF Then Exit Sub
  
  
' ---------------------------------------------------------------------------
' Opens file for writing
' ---------------------------------------------------------------------------
  
  Open Left(DriveDisk.Drive, 2) & "\ReportFiles.txt" For Output As #1
  Print #1,
  Print #1, Tab(10); "                Files  Statistics         "
  Print #1, Tab(10); "------------------------------------------"
  Print #1,
  Print #1, Tab(10); "Folder"; Spc(5); "Ext"; Spc(5); "No of Files"; Spc(5); "Size (KB)"
  Print #1, "   ==========================================================="
  
  Do While Not rsFiles.EOF
    
    
    DoEvents
    strOldPath = strPath
    strPath = rsFiles!Path
        
    If blnFirstRecord Then
        Print #1, Spc(20); "-----------------------------------"
        Print #1,
        Print #1, Tab(10); rsFiles!Path
        Print #1,
        Print #1, Spc(20); rsFiles!Ext; Tab(33); rsFiles!FilesNumber; Tab(45); Round(rsFiles!CSize / 1024, 2)
        SumSize = SumSize + rsFiles!CSize
        rsFiles.MoveNext
        strOldPath = strPath
        strPath = rsFiles!Path
        blnFirstRecord = False
    End If
    
    If strPath = strOldPath Then
        Print #1, Spc(20); rsFiles!Ext; Tab(33); rsFiles!FilesNumber; Tab(45); Round(rsFiles!CSize / 1024, 2)
        SumSize = SumSize + rsFiles!CSize
    Else
    
        If Not blnFirstRecord Then
            Print #1, Tab(45); "-----------"
            Print #1, Tab(45); Round(SumSize / 1024, 2) & " KB"
        End If
        SumSize = 0
        Print #1, Spc(20); "-----------------------------------"
        Print #1,
        If Not blnFirstRecord Then
            Print #1, Tab(10); rsFiles!Path
        End If
        Print #1,
        Print #1, Spc(20); rsFiles!Ext; Tab(33); rsFiles!FilesNumber; Tab(45); Round(rsFiles!CSize / 1024, 2)
        SumSize = SumSize + rsFiles!CSize
        
    End If
    rsFiles.MoveNext
  Loop
  
  Print #1, Tab(45); "-----------"
  Print #1, Tab(45); Round(SumSize / 1024, 2) & " KB"
  
  Print #1, "   ==========================================================="
  Close #1
  
' ---------------------------------------------------------------------------
' Re-set mouse pointer
' ---------------------------------------------------------------------------
  Screen.MousePointer = vbNormal
' ---------------------------------------------------------------------------

' ---------------------------------------------------------------------------
' Show the report's path
' ---------------------------------------------------------------------------
  lblPath.Caption = "Report Files at " & Left(DriveDisk.Drive, 2) & "\ReportFiles.txt"
' ---------------------------------------------------------------------------

' ---------------------------------------------------------------------------
' Enable the exit command button and hide the progress bar
' ---------------------------------------------------------------------------
  pBar.Visible = False: cmdExit.Enabled = True
  
End Sub


Private Sub Get_FileList(ByVal strFolder As String, _
                             Optional ByVal blnSearchSubfolders As Boolean = True, _
                             Optional strPattern As String = "*.*")
    
' ---------------------------------------------------------------------------
' Define local variables
' ---------------------------------------------------------------------------
  Dim objFile         As Scripting.File
  Dim objFolder       As Scripting.Folder
  Dim objSubFolder    As Scripting.Folder
  Dim strFileExt      As String
  Dim strFileName     As String
  Dim strTestExt      As String
  Dim strNewFolder    As String
  Dim blnGetAllFiles  As Boolean
  Dim strSQL          As String
  Dim comAppend       As ADODB.Command
  
  On Error Resume Next

' ---------------------------------------------------------------------------
' Instantsiate command object, set the active connection
' ---------------------------------------------------------------------------
  Set comAppend = New ADODB.Command
  comAppend.ActiveConnection = cnn
  
' ---------------------------------------------------------------------------
' Instantsiate Scripting object
' ---------------------------------------------------------------------------
  Set FSO = New Scripting.FileSystemObject
   
' ---------------------------------------------------------------------------
' get the starting folder
' ---------------------------------------------------------------------------
  
  Set objFolder = FSO.GetFolder(strFolder)
  
' ---------------------------------------------------------------------------
' Test the pattern we are looking for
' ---------------------------------------------------------------------------
  If strPattern = "*.*" Then
      blnGetAllFiles = True
  ElseIf InStr(strPattern, "*") = 0 Then
      blnGetAllFiles = True
  Else
     ' looking for a specific pattern
      blnGetAllFiles = False
      strTestExt = StrConv(FSO.GetExtensionName(strPattern), vbLowerCase)
  End If
  
' ---------------------------------------------------------------------------
' Check all the files in this directory
' ---------------------------------------------------------------------------
  For Each objFile In objFolder.Files
  
      ' see if the user cancelled processing
      
      If blnGetAllFiles Then
            strFileExt = StrConv(FSO.GetExtensionName(objFile.Path), vbLowerCase)

            ' Insert Into table tblFiles properties files
            ' build first the query
                    
            strSQL = "Insert Into tblFiles "
            strSQL = strSQL & "Values("
            strSQL = strSQL & "'" & objFile.Name & "', "
            strSQL = strSQL & "'" & strFileExt & "', "
            strSQL = strSQL & objFile.Size & ", "  ' size in B
            strSQL = strSQL & "'" & objFile.ParentFolder & "', "
            strSQL = strSQL & "'" & objFile.Path & "')"
            comAppend.CommandText = strSQL
            comAppend.Execute
      End If
  
  Next
        
      ' if requested, also search subdirectories.
  If blnSearchSubfolders Then

      For Each objSubFolder In objFolder.SubFolders
        
          strNewFolder = objSubFolder
        
          ' Do recursive calls from here.
          Call Get_FileList(strNewFolder, blnSearchSubfolders, strPattern)
      Next
  End If
    
Normal_Exit:
' ---------------------------------------------------------------------------
' Free objects from memory
' ---------------------------------------------------------------------------
  Set objFile = Nothing
  Set objFolder = Nothing
  Set objSubFolder = Nothing
  Set FSO = Nothing
  Set comAppend = Nothing
End Sub


Private Sub DeleteRecords()

' ---------------------------------------------------------------------------
' Define local variables
' ---------------------------------------------------------------------------
  Dim strSQL As String
  Dim comDeleteRecords As ADODB.Command
  
' ---------------------------------------------------------------------------
' Set command object connection and source
' ---------------------------------------------------------------------------
  Set comDeleteRecords = New ADODB.Command
  comDeleteRecords.ActiveConnection = cnn
  comDeleteRecords.CommandText = "Delete from tblFiles"
  comDeleteRecords.Execute
  
' ---------------------------------------------------------------------------
' Release memory resource
' ---------------------------------------------------------------------------
  Set comDeleteRecords = Nothing
  
End Sub

Private Sub Form_Load()
' ---------------------------------------------------------------------------
' Instantsiate connection object
' ---------------------------------------------------------------------------
  Set cnn = New ADODB.Connection

' ---------------------------------------------------------------------------
' Set connection string and open the database connection
' ---------------------------------------------------------------------------
  cnn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.3.51;Data Source=" & App.Path & "\DB.mdb"
  cnn.Open
  
  blnFirstRecord = True

End Sub
