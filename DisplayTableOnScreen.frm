VERSION 5.00
Begin VB.Form GetFileName 
   Caption         =   "Display a Table"
   ClientHeight    =   6555
   ClientLeft      =   1140
   ClientTop       =   1545
   ClientWidth     =   7395
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6555
   ScaleWidth      =   7395
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check1 
      Caption         =   "Use Selected Table Name below as the Table Name in my program"
      Height          =   375
      Left            =   4320
      TabIndex        =   9
      Top             =   1080
      Width           =   2895
   End
   Begin VB.TextBox txtText 
      Height          =   285
      Left            =   5040
      TabIndex        =   7
      Top             =   480
      Width           =   2175
   End
   Begin VB.CommandButton cmdEnd 
      Caption         =   "Exit"
      Height          =   495
      Left            =   3840
      TabIndex        =   5
      Top             =   5160
      Width           =   1215
   End
   Begin VB.ListBox lstList 
      Height          =   2595
      Left            =   3600
      TabIndex        =   3
      Top             =   2160
      Width           =   2055
   End
   Begin VB.FileListBox filFile 
      Height          =   2625
      Left            =   120
      Pattern         =   "*.mdb"
      TabIndex        =   2
      Top             =   480
      Width           =   2535
   End
   Begin VB.DirListBox dirDirectory 
      Height          =   1890
      Left            =   120
      TabIndex        =   1
      Top             =   3480
      Width           =   2535
   End
   Begin VB.DriveListBox drvDrive 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   5640
      Width           =   2535
   End
   Begin VB.Label Label3 
      Caption         =   "1.  Find a Database"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "2. Enter Table Name to be used in your program (optional)"
      Height          =   735
      Left            =   2880
      TabIndex        =   6
      Top             =   480
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "3. Select Tables to display"
      Height          =   255
      Left            =   3600
      TabIndex        =   4
      Top             =   1800
      Width           =   2055
   End
End
Attribute VB_Name = "GetFileName"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdEnd_Click()
    Label2.Visible = True
    End
End Sub
Private Sub dirDirectory_Change()
    filFile.Path = dirDirectory.Path
End Sub
Private Sub drvDrive_Change()
    dirDirectory.Path = drvDrive.Drive
End Sub
Private Sub filFile_Click()
    strWork = filFile.Path
    If Right(strWork, 1) <> "\" Then strWork = filFile.Path + "\"
    strDataBaseName = strWork + filFile.filename
    Set dbDatabase = Workspaces(0).OpenDatabase(strDataBaseName)
    intDatabaseCount = Workspaces(0).Databases.Count
    intOutSub = -1
    For J = 0 To intDatabaseCount - 1
        Set dbTempDataBase = Workspaces(0).Databases(J)
        For I = 0 To dbTempDataBase.TableDefs.Count - 1
            strInName = dbTempDataBase.TableDefs(I).Name
            If Mid$(strInName, 1, 4) = "MSys" Then GoTo GetNext1
            lstList.AddItem strInName
            intOutSub = intOutSub + 1
GetNext1:
            Next I
        Next J
End Sub
Private Sub Form_Load()
    txtText.Text = ""
    strAccessTableName = ""
End Sub
Private Sub lstList_Click()
    strTableName = lstList.List(lstList.ListIndex)
    If Check1.Value = 1 Then strAccessTableName = strTableName
    Unload GetFileName
    Load DisplayATable
End Sub
Private Sub txtText_LostFocus()
    If Not IsNull(txtText.Text) Then
        strAccessTableName = Trim(txtText.Text)
        I = InStr(1, strAccessTableName, " ")
        If I <> 0 Then
            strWork = "[" + strAccessTableName + "]"
            strAccessTableName = strWork
            End If
    End If
End Sub
