VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Penguin Reg Fix"
   ClientHeight    =   4320
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7125
   ForeColor       =   &H00C0C0C0&
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   7125
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   1320
      Width           =   855
   End
   Begin VB.CommandButton cmdDeleteInvalidKey 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Fix"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   1320
      Width           =   855
   End
   Begin VB.CommandButton cmdStartStop 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Start"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1320
      Width           =   855
   End
   Begin VB.Frame frameTool 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Caption         =   "Last Update :"
      Height          =   4575
      Left            =   -120
      TabIndex        =   0
      Top             =   -240
      Width           =   9015
      Begin VB.Frame frameTool_ScanReg 
         BackColor       =   &H00FFC0C0&
         Height          =   5895
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   9255
         Begin VB.TextBox txtCurKey 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   615
            Left            =   240
            MultiLine       =   -1  'True
            TabIndex        =   2
            Top             =   600
            Width           =   6495
         End
         Begin MSComctlLib.ListView lvErrorRegKey 
            Height          =   2055
            Left            =   120
            TabIndex        =   3
            Top             =   1920
            Width           =   6735
            _ExtentX        =   11880
            _ExtentY        =   3625
            View            =   3
            MultiSelect     =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            Checkboxes      =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            Icons           =   "img16x16"
            SmallIcons      =   "img16x16"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            Appearance      =   1
            NumItems        =   4
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Found Error in"
               Object.Width           =   2293
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "RootKey"
               Object.Width           =   2647
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "SubKey"
               Object.Width           =   3881
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Value"
               Object.Width           =   3087
            EndProperty
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "HELP"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000C&
            Height          =   255
            Left            =   6240
            TabIndex        =   9
            Top             =   120
            Width           =   615
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFC0C0&
            Caption         =   "Credit : Mithun Raj.P"
            BeginProperty Font 
               Name            =   "System"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   255
            Left            =   4680
            TabIndex        =   8
            Top             =   4080
            Width           =   2175
         End
         Begin VB.Label Label1 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Penguin Reg Fix"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   495
            Left            =   120
            TabIndex        =   7
            Top             =   1440
            Width           =   3015
         End
         Begin VB.Shape Shape1 
            FillColor       =   &H00FFFFFF&
            FillStyle       =   0  'Solid
            Height          =   855
            Left            =   120
            Shape           =   4  'Rounded Rectangle
            Top             =   480
            Width           =   6735
         End
         Begin VB.Label lblScanRegError 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   1440
            TabIndex        =   6
            Top             =   4080
            Width           =   1035
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label9 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Errors Found :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   255
            Left            =   120
            TabIndex        =   5
            Top             =   4080
            Width           =   1395
            WordWrap        =   -1  'True
         End
         Begin VB.Label lblScanRegStatus 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Scan :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   255
            Left            =   120
            TabIndex        =   4
            Top             =   240
            Width           =   2235
            WordWrap        =   -1  'True
         End
      End
   End
   Begin MSComctlLib.ImageList img16x16 
      Left            =   3120
      Top             =   1200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0ECA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":116B
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuop 
      Caption         =   "Op"
      Visible         =   0   'False
      Begin VB.Menu mnuOpSelAll 
         Caption         =   "Select All"
      End
      Begin VB.Menu mnuopUnSelAll 
         Caption         =   "Unselect All"
      End
   End
   Begin VB.Menu mnuTray 
      Caption         =   "tray"
      Visible         =   0   'False
      Begin VB.Menu mnuOpen 
         Caption         =   "Open"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "Close"
      End
   End
   Begin VB.Menu mnuProMenu 
      Caption         =   "pro_menu"
      Visible         =   0   'False
      Begin VB.Menu mnuProMenu_Ban 
         Caption         =   "&This Process is NOT Safe"
      End
      Begin VB.Menu mnuProMenu_Safe 
         Caption         =   "&This Process Is Safe"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' This is actualy a part of BG antivirus program coded by Yun Bunchhay..,
' some of the functions(which has not used in this code) in this programs are only for your refference....
Option Explicit
'scan registry variable with EVENTS
Dim WithEvents cReg As cRegSearch
Attribute cReg.VB_VarHelpID = -1


Private Sub Command1_Click()
Unload help
End
End Sub

'========================================'
' FORM EVENTS                            '
'========================================'
Private Sub Form_Load()
    
    'check if application is already loaded
    If App.PrevInstance = True Then
        MsgBox "Application is already running."
        'Call ShowWindow(app., 1)
        End
    End If
   help.Visible = False
    'scan reg content
   Set cReg = New cRegSearch
    intSettingRegOption = 1
    CreateDwordValue GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", "ScanRegOption", 1
End Sub


Private Sub Form_Unload(Cancel As Integer)
Unload help
End
End Sub

Private Sub frameTool_ScanReg_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label3.ForeColor = &H8000000C
End Sub

Private Sub Label3_Click()
help.Visible = True
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label3.ForeColor = &HFF&
End Sub

'========================================'
' SCAN                                   '
'========================================'







'========================================'


'Scan Registry
'-------------

Private Sub lvErrorRegKey_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If Item.Checked Then
        Item.Checked = False
    Else
        Item.Checked = True
    End If
End Sub


Private Sub lvErrorRegKey_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'right click
    If Button = 2 Then
        'select item first
        'Call lvErrorRegKey_ItemClick
        PopupMenu mnuop, , Me.lvErrorRegKey.Left + Me.frameTool.Left + Me.frameTool_ScanReg.Left + X, Me.lvErrorRegKey.Top + Me.frameTool.Top + Me.frameTool_ScanReg.Top + Y
    End If
End Sub

Private Sub mnuOpSelAll_Click()
    Dim i As Long
    With Me.lvErrorRegKey
    For i = 1 To .ListItems.Count
        'checked
        .ListItems.Item(i).Checked = True
    Next
    End With
End Sub

Private Sub mnuopUnSelAll_Click()
    Dim i As Long
    With Me.lvErrorRegKey
    For i = 1 To .ListItems.Count
        'unchecked
        .ListItems.Item(i).Checked = False
    Next
    End With
End Sub


Private Sub cmdStartStop_Click()
    
    'button CAPTION
    If cmdStartStop.Caption = "&Start" Then 'start => stop
        cmdStartStop.Caption = "&Stop"
    Else
        cmdStartStop.Caption = "&Start"     'stop => start
        cReg.StopSearch
        txtCurKey.Text = ""
        Exit Sub
    End If
    
    'clear items
    Me.lvErrorRegKey.ListItems.Clear
    txtCurKey.Text = ""
    lblScanRegStatus.Caption = "Scanning :"
    lblScanRegError.Caption = 0
    
    'SEARCH START
    '============
    '0=HKEY_ALL
    cReg.RootKey = 0
    'Don't search in any specific subkey (Search in all subkeys)
    cReg.SubKey = ""
    'Only find errors in value names and value values
    cReg.SearchFlags = KEY_NAME * 0 + VALUE_NAME * 1 + VALUE_VALUE * 1 + WHOLE_STRING * 0
    'Search for registry values with the suffix "C:\"
    cReg.SearchString = "C:\"
    'Start searching for invalid registry values
    cReg.DoSearch
    '=============
    'SEARCH FINISH
    
    txtCurKey.Text = ""
End Sub

Private Sub cmdDeleteInvalidKey_Click()

    Dim removed As Long, i As Integer
    'I don't think this is necessary, but if the registry backup takes a while, this program tells the user to wait.
    'txtCurKey.FontSize = 12
    'txtCurKey.FontBold = True
    'txtCurKey.Text = "Creating Registry Backup..."
    BackupReg
    'change status
    'txtCurKey.Text = "Registry Backup completed. Cleaning Errors..."
        
    'Loop through every item in lvwRegErrors
    With Me.lvErrorRegKey
    For i = 1 To .ListItems.Count
        'checked to be deleted
        If .ListItems.Item(i).Checked = True Then
            'Delete the registry error and mark the item as removed
            DeleteRegKey GetClassKey(.ListItems.Item(i).SubItems(1)), .ListItems.Item(i).SubItems(2), .ListItems.Item(i).SubItems(3)
            '.ListItems.Item(i).Text = "Cleaned"
            .ListItems.Item(i).Icon = 1
            .ListItems.Item(i).SmallIcon = 1
             removed = removed + 1
        End If
    Next
    End With
    'no deletion
    If removed = 0 Then GoTo endSub
    'change last status
    txtCurKey.Text = "Cleaning Errors completed."
    
endSub:
    'txtCurKey.FontSize = 8
    'txtCurKey.FontBold = False
    txtCurKey.Text = ""
    
End Sub

'Create a backup of the registry, using the "regedit.exe /e" command takes too long.
Public Sub BackupReg()

    Dim i As Integer
    Dim TheKey As String
    Dim TheValue As String
    Dim DefaultValue As Boolean
    Dim BackupFilename As String
    Dim f As Long
    
    'check folder backup
    If FileorFolderExists(App.Path & "\RegBak") = False Then MkDir App.Path & "\RegBak"
    
    BackupFilename = App.Path & "\RegBak\Backup_" & Format(Date, "dd-mm-yyyy") & "_" & Format(Time, "hh-nn-ss") & ".reg"
    'MsgBox BackupFilename
    
    'open file to write
    f = FreeFile
    Open BackupFilename For Output As #f
    Print #f, "REGEDIT4" & vbCrLf
    'Loops through all the checked items and saves the values reg file
    With lvErrorRegKey
    For i = 1 To .ListItems.Count
        If .ListItems.Item(i).Checked = True Then
        
            TheKey = ReverseString(.ListItems.Item(i).SubItems(1) & "\" & .ListItems.Item(i).SubItems(2))
            'the value might ends with a "\", then it's the default value for that key
            If Right$(TheKey, 1) = "\" Then DefaultValue = True: TheKey = Mid(TheKey, 2)
            TheValue = Chr(34) & Replace(ReverseString(Mid(TheKey, 1, InStr(1, TheKey, "\") - 1)), "\", "\\") & Chr(34)
            TheKey = ReverseString(Mid(TheKey, InStr(1, TheKey, "\") + 1))
            If DefaultValue = True Then TheValue = "@"
            'add key to .reg file
            Print #f, "[" & TheKey & "]" '& vbCrLf
            Print #f, TheValue & "=" & Chr(34) & .ListItems.Item(i).SubItems(3) & Chr(34) '& vbCrLf
            
        End If
    Next
    Close #f
    End With
    
End Sub


'class cRegSearch event
Private Sub cReg_SearchFound(ByVal sRootKey As String, ByVal sKey As String, ByVal sValue As Variant, ByVal lFound As FOUND_WHERE)
    
    Dim KN As String    'KeyName
    Dim FileorPath As String  'File Path
    Dim X As ListItem
    
    'WHERE
    Select Case lFound
    Case FOUND_IN_KEY_NAME
        KN = "KEY NAME"
    Case FOUND_IN_VALUE_NAME
        KN = "VALUE NAME"
    Case FOUND_IN_VALUE_VALUE
        KN = "DATA"
    End Select

    FileorPath = sValue
    
    'Condition !
    'If Right$(FileorPath, 4) = ".EXE" Or Right$(FileorPath, 4) = ".exe" Or Right$(FileorPath, 4) = ".DLL" Or Right$(FileorPath, 4) = ".dll" Or Right$(FileorPath, 4) = ".OCX" Or Right$(FileorPath, 4) = ".ocx" Or Right$(FileorPath, 4) = ".SYS" Or Right$(FileorPath, 4) = ".sys" Or Right$(FileorPath, 4) = ".VXD" Or Right$(FileorPath, 4) = ".vxd" Or Right$(FileorPath, 3) = ".AX" Or Right$(FileorPath, 3) = ".ax" Then
    
    'check if actual file exist as in registry
    If FileorFolderExists(FormatValue(FileorPath)) = False Then 'not exist => invalid key
        
        If intSettingRegOption = 1 Then 'scan all
            'add to list for any key
            With Me.lvErrorRegKey
                Set X = .ListItems.Add(, , KN, 2, 2)
                X.SubItems(1) = sRootKey
                X.SubItems(2) = sKey
                X.SubItems(3) = sValue
            End With
            Set X = Nothing
            'add to counter
            Me.lblScanRegError.Caption = Int(Me.lblScanRegError.Caption) + 1
        Else    'scan specific extension
            'MsgBox FileorPath
            'MsgBox Mid(FileorPath, InStr(1, FileorPath, ".") + 1, 3)
            'MsgBox InStr(1, LCase(Mid(FileorPath, InStr(1, FileorPath, ".") + 1, 3)), LCase(strScanRegExt))
            'If InStr(1, Right$(FileorPath, 3), strScanRegExt, vbTextCompare) > 0 Then    'found in extension
            If InStr(1, strScanRegExt, Mid(FileorPath, InStr(1, FileorPath, ".") + 1, 3), vbTextCompare) > 0 Then    'found in extension
                With Me.lvErrorRegKey
                    Set X = .ListItems.Add(, , KN, 2, 2)
                    X.SubItems(1) = sRootKey
                    X.SubItems(2) = sKey
                    X.SubItems(3) = sValue
                End With
                Set X = Nothing
                'add to counter
                Me.lblScanRegError.Caption = Int(Me.lblScanRegError.Caption) + 1
            End If
        End If
    End If
    
End Sub

'class cRegSearch event
Private Sub cReg_SearchFinished(ByVal lReason As Long)
    
    If lReason = 0 Then
        Me.lblScanRegStatus.Caption = "Scan Completed"
    ElseIf lReason = 1 Then
        Me.lblScanRegStatus.Caption = "Scan Cancelled"
    Else
        Me.lblScanRegStatus.Caption = "Scan Error"
    End If
    cmdStartStop.Caption = "&Start"
End Sub

'class cRegSearch event, when change key to search
Private Sub cReg_SearchKeyChanged(ByVal sFullKeyName As String)
    txtCurKey.Text = sFullKeyName
End Sub


