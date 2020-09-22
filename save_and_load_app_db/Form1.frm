VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "LOAD AND SAVE APPLICATION FROM/INTO DB"
   ClientHeight    =   4035
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7125
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4035
   ScaleWidth      =   7125
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin TabDlg.SSTab SSTab1 
      Height          =   3915
      Left            =   90
      TabIndex        =   0
      Top             =   60
      Width           =   6915
      _ExtentX        =   12197
      _ExtentY        =   6906
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "Load App from DB"
      TabPicture(0)   =   "Form1.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cboApp"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtDesc"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtOutput"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdOutput"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmdload"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "Save App into DB"
      TabPicture(1)   =   "Form1.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label4(0)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label4(1)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label4(2)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label4(3)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "txtAppName"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "TxtDescription"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "txtFileName"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "TxtArg"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "cmdopen"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "cmdSave"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "CommonDialog1"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).ControlCount=   11
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   -70650
         Top             =   3030
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         DefaultExt      =   "EXE"
         Filter          =   "Application (*.EXE)|*.EXE|All Files (*.*)|*.*"
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Height          =   645
         Left            =   -71940
         Picture         =   "Form1.frx":0038
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   3030
         Width           =   885
      End
      Begin VB.CommandButton cmdopen 
         Height          =   375
         Left            =   -68790
         Picture         =   "Form1.frx":05C2
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   2010
         Width           =   465
      End
      Begin VB.TextBox TxtArg 
         Height          =   345
         Left            =   -73200
         TabIndex        =   16
         Top             =   2490
         Width           =   4875
      End
      Begin VB.TextBox txtFileName 
         Height          =   345
         Left            =   -73200
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   2010
         Width           =   4335
      End
      Begin VB.TextBox TxtDescription 
         Height          =   915
         Left            =   -73200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   12
         Top             =   990
         Width           =   4875
      End
      Begin VB.TextBox txtAppName 
         Height          =   345
         Left            =   -73200
         TabIndex        =   10
         Top             =   510
         Width           =   4875
      End
      Begin VB.CommandButton cmdload 
         Caption         =   "&Load"
         Height          =   705
         Left            =   2970
         Picture         =   "Form1.frx":0B4C
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   2970
         Width           =   1065
      End
      Begin VB.CommandButton cmdOutput 
         Height          =   375
         Left            =   6270
         Picture         =   "Form1.frx":10D6
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   2430
         Width           =   495
      End
      Begin VB.TextBox txtOutput 
         Height          =   345
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   6
         Text            =   "c:\"
         Top             =   2430
         Width           =   4395
      End
      Begin VB.TextBox txtDesc 
         Height          =   1365
         Left            =   1800
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   960
         Width           =   4935
      End
      Begin VB.ComboBox cboApp 
         Height          =   315
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   510
         Width           =   4935
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Argument:"
         Height          =   195
         Index           =   3
         Left            =   -74790
         TabIndex        =   15
         Top             =   2550
         Width           =   720
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "FileName:"
         Height          =   195
         Index           =   2
         Left            =   -74790
         TabIndex        =   13
         Top             =   2070
         Width           =   705
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Description:"
         Height          =   195
         Index           =   1
         Left            =   -74790
         TabIndex        =   11
         Top             =   1050
         Width           =   840
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Application Name:"
         Height          =   195
         Index           =   0
         Left            =   -74790
         TabIndex        =   9
         Top             =   570
         Width           =   1290
      End
      Begin VB.Label Label3 
         Caption         =   "Output directory:"
         Height          =   255
         Left            =   210
         TabIndex        =   5
         Top             =   2490
         Width           =   1485
      End
      Begin VB.Label Label2 
         Caption         =   "Description:"
         Height          =   255
         Left            =   210
         TabIndex        =   3
         Top             =   1080
         Width           =   1365
      End
      Begin VB.Label Label1 
         Caption         =   "Application Name:"
         Height          =   315
         Left            =   180
         TabIndex        =   1
         Top             =   570
         Width           =   2025
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Rs As ADODB.Recordset
Private sConn As String

Private Sub cboApp_Click()
ReadData
End Sub

Private Sub cmdload_Click()
On Error Resume Next
    
    ' check recordset state
    If Not (Rs.EOF And Rs.BOF) Then
    
        Dim Dfile As String
        
        Dfile = txtOutput.Text & Rs.Fields(2).Value
        
        ' save binary data to file
        LoadFileFromRecordset Rs.Fields(4), Dfile
        
        If Not IsNull(Rs.Fields(3).Value) Then
            Dfile = Dfile & Rs.Fields(3).Value
        End If
        
        ' open executeable file
        CloseHandle Shell(Dfile)
        
    End If
    
End Sub

Private Sub cmdopen_Click()
    On Error Resume Next
    With Me.CommonDialog1
        .FileName = ""
        .ShowOpen
        If .FileName <> "" Then
        txtFileName = .FileName
        End If
    End With
    
End Sub

Private Sub cmdOutput_Click()
On Error Resume Next
    Dim sFolder As String
    sFolder = GetFolder(hWnd)
    
    If sFolder <> "" Then
        txtOutput = sFolder & "\"
    Else
        txtOutput = "C:\"
    End If
End Sub

Private Sub cmdSave_Click()

    SaveApp
    MsgBox "save Finished !!!", vbInformation
    txtAppName.Text = ""
    TxtDescription.Text = ""
    txtFileName.Text = ""
    TxtArg.Text = ""
    
End Sub

Private Sub Form_Load()
    On Error Resume Next
    
    ' prepare connection string
    sConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\myappdb.mdb"
    
    ' create new recordset
    Set Rs = New ADODB.Recordset
    
    ' set cursor location
    Rs.CursorLocation = adUseClient
    
    ' open recordset
    Rs.Open "select * from tbl_app", sConn, adOpenStatic, adLockOptimistic
    
    FillAppNameIntoCombo
    
End Sub


Private Sub SaveApp()
    On Error Resume Next
    With Rs
    
        'Add new Record
        .AddNew
        
        ' Application Name
        .Fields(0).Value = txtAppName.Text
        
        ' Description
        If TxtDescription.Text <> "" Then
            .Fields(1).Value = TxtDescription.Text
        End If
        
        ' File Title
        If txtFileName.Text <> "" Then
            .Fields(2).Value = Mid(txtFileName.Text, InStrRev(txtFileName.Text, "\") + 1)
        End If
        
        ' Arguments
        If TxtArg.Text <> "" Then
            .Fields(3).Value = TxtArg.Text
        End If
        
        
        
        Dim vtData() As Byte
        Dim nfile As Long
        nfile = FreeFile
        
        ' Open file as binary data
        Open txtFileName.Text For Binary As nfile
        
        ' get length of file
            ReDim vtData(LOF(nfile))
            
            'get binary data
            Get nfile, , vtData
            
            ' save binary torecordset
            .Fields(4).AppendChunk vtData
        
        Close nfile
        
        'update recordset
        .Update
        
        'update combo list
        FillAppNameIntoCombo
        
    End With
End Sub



Private Sub FillAppNameIntoCombo()
On Error Resume Next

    Dim R As New ADODB.Recordset
    R.CursorLocation = adUseClient
    
    'open recordset
    R.Open "select appname from tbl_app order by appname", sConn, adOpenStatic, adLockOptimistic
    
    ' check recordset status, have records or not
    If Not (R.EOF And R.BOF) Then
        'if have records
        cboApp.Clear
        Do While Not R.EOF
            ' insert appname into cboapp list
            cboApp.AddItem R(0).Value
            R.MoveNext
        Loop
    End If
    
    ' clean up
    R.Close
    Set R = Nothing
    
End Sub


Private Sub LoadFileFromRecordset(MyField As Field, StrFilename As String)
    On Error GoTo err_x

        Dim file_num As Long
        Dim file_length As Long
        Dim bytes() As Byte
        file_num = FreeFile
        
        Open StrFilename For Binary Access Write As #file_num
            file_length = LenB(MyField.Value)
            bytes() = MyField.GetChunk(file_length)
            Put #file_num, , bytes()
            
        Close #file_num

    Exit Sub
err_x:
End Sub


Private Sub ReadData()
On Error Resume Next

    ' find record
    
    Rs.Find "appname='" & cboApp.Text & "'", , , 1
    
    If Not (Rs.EOF And Rs.BOF) Then
        If Not IsNull(Rs.Fields(1).Value) Then txtDesc.Text = Rs.Fields(1).Value
    Else
        Rs.MoveFirst
    End If
End Sub












