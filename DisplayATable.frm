VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form DisplayATable 
   Caption         =   "Form1"
   ClientHeight    =   11115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3225
   LinkTopic       =   "Form1"
   ScaleHeight     =   11115
   ScaleWidth      =   3225
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   10875
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   19182
      _Version        =   393216
      Rows            =   1
      Cols            =   3
      BackColorFixed  =   16777215
      ScrollBars      =   2
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Black"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FFFF&
      Caption         =   "Click on Field Name to Copy"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   2895
   End
End
Attribute VB_Name = "DisplayATable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Load()
    Show
    Set rsRecordSet = dbDatabase.OpenRecordset(strTableName, dbOpenSnapshot)
    DisplayATable.Caption = strTableName
    DisplayATable.Top = 0
    DisplayATable.Left = Screen.Width - 3200
    MSFlexGrid1.Cols = 3
    MSFlexGrid1.ColWidth(0) = 2000
    MSFlexGrid1.ColWidth(1) = 500
    MSFlexGrid1.ColWidth(2) = 450
        Set rsTempRecordSet = dbDatabase.Recordsets(0)
        FieldsCount = rsTempRecordSet.Fields.Count
        If FieldsCount > 40 Then FieldsCount = 40
        DisplayATable.Height = (FieldsCount * 270) + 800
        If FieldsCount >= 40 Then
            DisplayATable.Left = Screen.Width - 3400
            DisplayATable.Width = 3400
            End If
        Label1.Top = (270 * FieldsCount) + 40
        MSFlexGrid1.Height = (270 * FieldsCount) + 40
        MSFlexGrid1.Rows = rsTempRecordSet.Fields.Count
        DisplayATable.Refresh
        For I = 0 To rsTempRecordSet.Fields.Count - 1
            strWork = rsTempRecordSet.Fields(I).SourceField
            strWork1 = Space$(35)
            strWork2 = rsTempRecordSet.Fields(I).Name
            L = Len(strWork2)
            If L > 25 Then L = 25
            strWork1 = Trim(Mid$(strWork2, 1, L))
            strTypeOut = "Unk"
            strTypeIn = rsTempRecordSet.Fields(I).Type
            If strTypeIn = 0 Then strTypeOut = "Y/N"
            If strTypeIn = 1 Then strTypeOut = "Bln"
            If strTypeIn = 3 Then strTypeOut = "Int"
            If strTypeIn = 4 Then strTypeOut = "Lng"
            If strTypeIn = 5 Then strTypeOut = "Cur"
            If strTypeIn = 6 Then strTypeOut = "Sgl"
            If strTypeIn = 7 Then strTypeOut = "Dbl"
            If strTypeIn = 8 Then strTypeOut = "Date"
            If strTypeIn = 9 Then strTypeOut = "Tim"
            If strTypeIn = 10 Then strTypeOut = "Txt"
            If strTypeIn = 11 Then strTypeOut = "Lng"
            If strTypeIn = 12 Then strTypeOut = "Mem"
            intRowSub = intRowSub + 1
            MSFlexGrid1.Row = intRowSub - 1
            MSFlexGrid1.Col = 0
            MSFlexGrid1.Text = strWork1
            MSFlexGrid1.Col = 1
            MSFlexGrid1.Text = strTypeOut
            If rsTempRecordSet.Fields(I).Type = dbText Then
                MSFlexGrid1.Col = 2
                MSFlexGrid1.Text = rsTempRecordSet.Fields(I).Size
                End If
            Next I
End Sub
Private Sub MSFlexGrid1_Click()
    MSFlexGrid1.Col = 0
    Clipboard.Clear
    strWork1 = Trim(MSFlexGrid1.Text)
    I = InStr(1, strWork1, " ")
    If I <> 0 Then strWork1 = "[" + Trim(MSFlexGrid1.Text) + "]"
    strWork = strWork1
    If strAccessTableName <> "" Then strWork = strAccessTableName + "!" + strWork1
    Clipboard.SetText strWork
End Sub
