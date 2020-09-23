VERSION 5.00
Begin VB.Form structure 
   Caption         =   "Database Structure Report"
   ClientHeight    =   9225
   ClientLeft      =   2970
   ClientTop       =   1800
   ClientWidth     =   9375
   Icon            =   "structure.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   9225
   ScaleWidth      =   9375
   Begin VB.CheckBox chkIdx 
      Caption         =   "Show indexes"
      Height          =   255
      Left            =   7140
      TabIndex        =   14
      Top             =   4635
      Width           =   2100
   End
   Begin VB.TextBox txtPW 
      Height          =   330
      Left            =   6765
      TabIndex        =   10
      Top             =   3690
      Width           =   2505
   End
   Begin VB.DirListBox Dir1 
      Height          =   3465
      Left            =   30
      TabIndex        =   7
      Top             =   600
      Width           =   4560
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   30
      TabIndex        =   6
      Top             =   240
      Width           =   4560
   End
   Begin VB.FileListBox File1 
      Height          =   2625
      Left            =   4710
      Pattern         =   "*.mdb"
      TabIndex        =   5
      Top             =   480
      Width           =   4560
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   30
      TabIndex        =   3
      Top             =   4440
      Width           =   6795
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Print"
      Height          =   375
      Left            =   8295
      TabIndex        =   2
      Top             =   4170
      Width           =   945
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4095
      Left            =   30
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   5010
      Width           =   9240
   End
   Begin VB.CommandButton cmdList 
      Caption         =   "&List"
      Height          =   375
      Left            =   7155
      TabIndex        =   0
      Top             =   4170
      Width           =   945
   End
   Begin VB.Label Label6 
      Caption         =   "Databases found"
      Height          =   255
      Left            =   4725
      TabIndex        =   13
      Top             =   240
      Width           =   1800
   End
   Begin VB.Label Label5 
      Caption         =   "Structure"
      Height          =   210
      Left            =   30
      TabIndex        =   12
      Top             =   4800
      Width           =   750
   End
   Begin VB.Label Label4 
      Caption         =   "Database Password (if any)"
      Height          =   240
      Left            =   4725
      TabIndex        =   11
      Top             =   3735
      Width           =   1995
   End
   Begin VB.Label Label3 
      Caption         =   "Double click to list the structure"
      Height          =   255
      Left            =   4725
      TabIndex        =   9
      Top             =   3120
      Width           =   4005
   End
   Begin VB.Label Label2 
      Caption         =   "Select the database"
      Height          =   255
      Left            =   30
      TabIndex        =   8
      Top             =   30
      Width           =   1800
   End
   Begin VB.Label Label1 
      Caption         =   "Database Pathname"
      Height          =   255
      Left            =   30
      TabIndex        =   4
      Top             =   4155
      Width           =   1815
   End
End
Attribute VB_Name = "structure"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdList_Click()
On Error Resume Next
    Dim fld As Field
    Dim fld1 As String * 20
    Dim fld2 As String * 10
    Dim fld3 As String * 20
    Dim dbs As Database
    Dim tdf As TableDef
    Dim idx As Index
    Dim tbls As Integer
    Dim flds As Integer
    Dim PW As String
    
    Text1.Text = Text2.Text
    
    If txtPW.Text <> "" Then
        If UCase(Right(Text2.Text, 10)) = "CRMMGR.MDB" Then
            txtPW.Text = "123321"
        End If
    End If
    Err = 0
    
    If txtPW.Text <> "" Then
        PW = ";pwd=" & txtPW.Text & ""
        Set dbs = OpenDatabase(Text2.Text, False, False, PW)
    Else
        Set dbs = OpenDatabase(Text2.Text)
    End If
    If Err <> 0 Then
        MsgBox "DB problem. May be password protected"
        Exit Sub
    End If
    
    tbls = dbs.TableDefs.Count
    i = 0
    For Each tdf In dbs.TableDefs
        If Mid(tdf.Name, 1, 2) <> "MS" Then
            Text1.Text = Text1.Text & Chr(13) & Chr(10) & "Table : " & tdf.Name & Chr(13) & Chr(10)
            flds = tdf.Fields.Count
            If chkIdx.Value = vbChecked Then
                For Each idx In tdf.Indexes
                    Text1.Text = Text1.Text & "  Index Name : " & idx.Name & Chr(13) & Chr(10) & "  Index Fields : " & idx.Fields & Chr(13) & Chr(10)
                Next idx
            End If
            For n = 0 To flds - 1
                Set fld = dbs.TableDefs(i).Fields(n)
                fld1 = fld.Name
                fld2 = FieldType(fld.Type)
                If fld.AllowZeroLength = True Then
                    fld3 = "0Lth OK"
                Else
                    fld3 = ""
                End If
                
                Text1.Text = Text1.Text & "          " & fld1 & "   " & fld2 & "  " & Format(fld.Size, "#00") & "   " & fld3 & Chr(13) & Chr(10)
            Next
        End If
        i = i + 1
    Next tdf
  
    dbs.Close
    Set dbs = Nothing
     
End Sub

Private Sub Command2_Click()

Printer.Font = "courier new"
Printer.FontSize = 10
Printer.Print Text1.Text
Printer.EndDoc


End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path

End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
File1.Path = Dir1.Path



End Sub

Private Sub File1_Click()
Dim sPath As String
sPath = File1.Path
If Right(sPath, 1) <> "\" Then sPath = sPath & "\"
sPath = sPath & File1.FileName
Text2.Text = sPath



End Sub

Private Sub File1_DblClick()
Text2.Text = File1.Path & "\" & File1.FileName
Call cmdList_Click ' list
End Sub

Private Sub Form_Load()

Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2

Me.Caption = App.Title & " ver " & App.Major & "." & App.Minor

File1.Pattern = "*.mdb"

If Command$ <> "" Then
    Text2.Text = Command$
    Call cmdList_Click
End If

End Sub

Function FieldType(intType As Integer) As String

    Select Case intType
        Case dbBoolean
            FieldType = "Boolean"
        Case dbByte
            FieldType = "Byte"
        Case dbInteger
            FieldType = "Integer"
        Case dbLong
            FieldType = "Long"
        Case dbCurrency
            FieldType = "Currency"
        Case dbSingle
            FieldType = "Single"
        Case dbDouble
            FieldType = "Double"
        Case dbDate
            FieldType = "Date"
        Case dbText
            FieldType = "Text"
        Case dbLongBinary
            FieldType = "LongBinary"
        Case dbMemo
            FieldType = "Memo"
        Case dbGUID
            FieldType = "GUID"
    End Select

End Function

Private Sub Form_Resize()
Text1.Width = Me.Width - 300
End Sub


