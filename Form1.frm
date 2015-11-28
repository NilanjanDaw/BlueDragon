VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "BlueDragon"
   ClientHeight    =   9000
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9000
   ScaleWidth      =   15000
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Frame option 
      BackColor       =   &H8000000E&
      BorderStyle     =   0  'None
      Height          =   3135
      Index           =   0
      Left            =   7440
      TabIndex        =   24
      Top             =   4920
      Visible         =   0   'False
      Width           =   5535
      Begin VB.PictureBox deletePic 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1575
         Index           =   0
         Left            =   3720
         Picture         =   "Form1.frx":0000
         ScaleHeight     =   1575
         ScaleWidth      =   1455
         TabIndex        =   26
         Top             =   840
         Width           =   1455
      End
      Begin VB.PictureBox playPic 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1455
         Index           =   0
         Left            =   600
         Picture         =   "Form1.frx":0B1B
         ScaleHeight     =   1455
         ScaleWidth      =   1455
         TabIndex        =   25
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label title 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Bauhaus 93"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   1080
         TabIndex        =   29
         Top             =   120
         Width           =   3255
      End
      Begin VB.Label delete 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Delete"
         BeginProperty Font 
            Name            =   "Bauhaus 93"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   3840
         TabIndex        =   28
         Top             =   2400
         Width           =   1335
      End
      Begin VB.Label playbutton 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Play"
         BeginProperty Font 
            Name            =   "Bauhaus 93"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   720
         TabIndex        =   27
         Top             =   2400
         Width           =   1335
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H8000000C&
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   9975
      Left            =   13680
      TabIndex        =   2
      Top             =   480
      Width           =   6375
      Begin VB.Frame option 
         BackColor       =   &H8000000E&
         BorderStyle     =   0  'None
         Height          =   3135
         Index           =   2
         Left            =   360
         TabIndex        =   36
         Top             =   4440
         Visible         =   0   'False
         Width           =   5535
         Begin VB.PictureBox playPic 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1455
            Index           =   2
            Left            =   600
            Picture         =   "Form1.frx":242C
            ScaleHeight     =   1455
            ScaleWidth      =   1455
            TabIndex        =   38
            Top             =   960
            Width           =   1455
         End
         Begin VB.PictureBox deletePic 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1575
            Index           =   2
            Left            =   3720
            Picture         =   "Form1.frx":3D3D
            ScaleHeight     =   1575
            ScaleWidth      =   1455
            TabIndex        =   37
            Top             =   840
            Width           =   1455
         End
         Begin VB.Label playbutton 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Play"
            BeginProperty Font 
               Name            =   "Bauhaus 93"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   2
            Left            =   720
            TabIndex        =   41
            Top             =   2400
            Width           =   1335
         End
         Begin VB.Label delete 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Delete"
            BeginProperty Font 
               Name            =   "Bauhaus 93"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   2
            Left            =   3840
            TabIndex        =   40
            Top             =   2400
            Width           =   1335
         End
         Begin VB.Label title 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Bauhaus 93"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   2
            Left            =   1080
            TabIndex        =   39
            Top             =   240
            Width           =   3255
         End
      End
      Begin VB.CommandButton Search 
         Caption         =   "Search"
         BeginProperty Font 
            Name            =   "Bauhaus 93"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   4440
         TabIndex        =   23
         Top             =   8880
         Width           =   1695
      End
      Begin VB.TextBox searchBox 
         BeginProperty Font 
            Name            =   "Bauhaus 93"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   480
         TabIndex        =   21
         Top             =   8880
         Width           =   3855
      End
      Begin VB.ListBox ListVideo 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000C&
         BeginProperty Font 
            Name            =   "Bauhaus 93"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4890
         Left            =   480
         Sorted          =   -1  'True
         TabIndex        =   19
         Top             =   3720
         Width           =   5295
      End
      Begin VB.CommandButton uploadVideo 
         Caption         =   "Upload"
         BeginProperty Font 
            Name            =   "Bauhaus 93"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1920
         TabIndex        =   14
         Top             =   2640
         Width           =   1695
      End
      Begin VB.TextBox pathText 
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   2
         Left            =   2520
         TabIndex        =   11
         Top             =   1800
         Width           =   3495
      End
      Begin VB.CommandButton browseVideo 
         Caption         =   "Browse"
         BeginProperty Font 
            Name            =   "Bauhaus 93"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   480
         TabIndex        =   8
         Top             =   1800
         Width           =   1695
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Video Library"
         BeginProperty Font 
            Name            =   "Bauhaus 93"
            Size            =   21.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   615
         Index           =   2
         Left            =   960
         TabIndex        =   5
         Top             =   480
         Width           =   4455
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H8000000C&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   9975
      Left            =   7080
      TabIndex        =   1
      Top             =   480
      Width           =   6255
      Begin VB.CommandButton Search 
         Caption         =   "Search"
         BeginProperty Font 
            Name            =   "Bauhaus 93"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   4320
         TabIndex        =   22
         Top             =   8880
         Width           =   1695
      End
      Begin VB.TextBox searchBox 
         BeginProperty Font 
            Name            =   "Bauhaus 93"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   360
         TabIndex        =   20
         Top             =   8880
         Width           =   3855
      End
      Begin VB.ListBox ListAudio 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000C&
         BeginProperty Font 
            Name            =   "Bauhaus 93"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4890
         Left            =   480
         Sorted          =   -1  'True
         TabIndex        =   18
         Top             =   3720
         Width           =   5175
      End
      Begin VB.CommandButton uploadAudio 
         Caption         =   "Upload"
         BeginProperty Font 
            Name            =   "Bauhaus 93"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2040
         TabIndex        =   13
         Top             =   2640
         Width           =   1695
      End
      Begin VB.TextBox pathText 
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   1
         Left            =   2520
         TabIndex        =   10
         Top             =   1800
         Width           =   3495
      End
      Begin VB.CommandButton browseAudio 
         Caption         =   "Browse"
         BeginProperty Font 
            Name            =   "Bauhaus 93"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   480
         TabIndex        =   7
         Top             =   1800
         Width           =   1695
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Audio Library"
         BeginProperty Font 
            Name            =   "Bauhaus 93"
            Size            =   21.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   615
         Index           =   1
         Left            =   840
         TabIndex        =   4
         Top             =   480
         Width           =   4455
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000C&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   9975
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   6495
      Begin VB.Frame option 
         BackColor       =   &H8000000E&
         BorderStyle     =   0  'None
         Height          =   3135
         Index           =   1
         Left            =   480
         TabIndex        =   30
         Top             =   4440
         Visible         =   0   'False
         Width           =   5535
         Begin VB.PictureBox playPic 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1455
            Index           =   1
            Left            =   600
            Picture         =   "Form1.frx":4858
            ScaleHeight     =   1455
            ScaleWidth      =   1455
            TabIndex        =   32
            Top             =   840
            Width           =   1455
         End
         Begin VB.PictureBox deletePic 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1575
            Index           =   1
            Left            =   3720
            Picture         =   "Form1.frx":6169
            ScaleHeight     =   1575
            ScaleWidth      =   1455
            TabIndex        =   31
            Top             =   840
            Width           =   1455
         End
         Begin VB.Label playbutton 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Show"
            BeginProperty Font 
               Name            =   "Bauhaus 93"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   720
            TabIndex        =   35
            Top             =   2520
            Width           =   1335
         End
         Begin VB.Label delete 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Delete"
            BeginProperty Font 
               Name            =   "Bauhaus 93"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   3840
            TabIndex        =   34
            Top             =   2520
            Width           =   1335
         End
         Begin VB.Label title 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Bauhaus 93"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   1080
            TabIndex        =   33
            Top             =   120
            Width           =   3255
         End
      End
      Begin VB.CommandButton Search 
         Caption         =   "Search"
         BeginProperty Font 
            Name            =   "Bauhaus 93"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   4440
         TabIndex        =   17
         Top             =   8880
         Width           =   1695
      End
      Begin VB.TextBox searchBox 
         BeginProperty Font 
            Name            =   "Bauhaus 93"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   480
         TabIndex        =   16
         Top             =   8880
         Width           =   3855
      End
      Begin VB.ListBox ListImages 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000C&
         BeginProperty Font 
            Name            =   "Bauhaus 93"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4890
         Left            =   480
         Sorted          =   -1  'True
         TabIndex        =   15
         Top             =   3720
         Width           =   5535
      End
      Begin VB.CommandButton uploadImage 
         Caption         =   "Upload"
         BeginProperty Font 
            Name            =   "Bauhaus 93"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1920
         TabIndex        =   12
         Top             =   2640
         Width           =   1695
      End
      Begin VB.TextBox pathText 
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   2520
         TabIndex        =   9
         Top             =   1800
         Width           =   3495
      End
      Begin VB.CommandButton BrowseImage 
         Caption         =   "Browse"
         BeginProperty Font 
            Name            =   "Bauhaus 93"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   480
         TabIndex        =   6
         Top             =   1800
         Width           =   1695
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Image Library"
         BeginProperty Font 
            Name            =   "Bauhaus 93"
            Size            =   21.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   615
         Index           =   0
         Left            =   840
         TabIndex        =   3
         Top             =   480
         Width           =   4455
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3000
      Top             =   8520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   0
      Top             =   8520
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   873
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Nilanjan\Documents\db1.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Nilanjan\Documents\db1.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Image Image1 
      Height          =   16200
      Left            =   0
      Picture         =   "Form1.frx":6C84
      Top             =   0
      Width           =   28800
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim recordCon As ADODB.Connection
Dim recordSet As ADODB.recordSet
Dim filetitle As String
Const BlockSize = 32768
Private Declare Function GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Const MAX_PATH = 260

Private Sub browseAudio_Click()
    With CommonDialog1
    .Filter = "All Files|*.*|MP3s|*.mp3|Flacs|*.flc|WAVs|*.wav"
    .DialogTitle = "Choose a file"
    .ShowOpen
    End With
    pathText(1).Text = CommonDialog1.filename
    filetitle = CommonDialog1.filetitle
End Sub

Private Sub browseImage_Click()
    With CommonDialog1
    .Filter = "All Files|*.*|Bitmaps|*.bmp|GIFs|*.gif|JPEGs|*.jpg"
    .DialogTitle = "Choose a file"
    .ShowOpen
    End With
    pathText(0).Text = CommonDialog1.filename
    filetitle = CommonDialog1.filetitle
End Sub

Private Sub browseVideo_Click()
     With CommonDialog1
    .Filter = "All Files|*.*|Mpeg4|*.mp4|GIFs|*.gif|JPEGs|*.jpg"
    .DialogTitle = "Choose a file"
    .ShowOpen
    End With
    pathText(2).Text = CommonDialog1.filename
    filetitle = CommonDialog1.filetitle
End Sub

Private Sub delete_Click(Index As Integer)
    If Index = 0 Then
    recordCon.Execute "Delete * from table1 where filename='" & ListAudio.Text & "' and type='audio';"
    ElseIf Index = 1 Then
    recordCon.Execute "Delete * from table1 where filename='" & ListImages.Text & "' and type='image';"
    ElseIf Index = 2 Then
    recordCon.Execute "Delete * from table1 where filename='" & ListVideo.Text & "' and type='video';"
    End If
    List
End Sub

Private Sub deletePic_GotFocus(Index As Integer)
    'recordSet.Open "Select * from table1", recordCon, adOpenDynamic, adLockOptimistic
    If Index = 0 Then
    recordCon.Execute "Delete * from table1 where filename='" & ListAudio.Text & "' and type='audio';"
    ElseIf Index = 1 Then
    recordCon.Execute "Delete * from table1 where filename='" & ListImages.Text & "' and type='image';"
    ElseIf Index = 2 Then
    recordCon.Execute "Delete * from table1 where filename='" & ListVideo.Text & "' and type='video';"
    End If
    List
    'recordSet.Close
End Sub

Private Sub Form_Load()
    Dim a As Variant
    On Error GoTo HandleError
    Set recordCon = New ADODB.Connection
    Set recordSet = New ADODB.recordSet
    recordCon.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.path & "\db1.mdb;Persist Security Info=False"
    recordCon.Open
    List
    'progress_upload.Show
    Exit Sub
HandleError:
    MsgBox Err.Description
End Sub

Private Sub List()
    recordSet.Open "Select * from table1", recordCon, adOpenDynamic, adLockOptimistic
    'rs.Open "Select * from table1", recordCon, adOpenDynamic, adLockOptimistic
    Set recordSet = recordCon.Execute("Select * from table1 where type='image';")
    ListImages.Clear
    ListAudio.Clear
    ListVideo.Clear
    
    While Not recordSet.EOF
    ListImages.AddItem (recordSet("filename"))
    recordSet.MoveNext
    Wend
    recordSet.Close
    Set recordSet = recordCon.Execute("Select * from table1 where type='audio';")
    While Not recordSet.EOF
    ListAudio.AddItem (recordSet("filename"))
    recordSet.MoveNext
    Wend
    recordSet.Close
    Set recordSet = recordCon.Execute("Select * from table1 where type='video';")
    While Not recordSet.EOF
    ListVideo.AddItem (recordSet("filename"))
    recordSet.MoveNext
    Wend
    recordSet.Close
End Sub

Private Sub play(Text As String)
    
    recordSet.Open "Select * from table1", recordCon, adOpenDynamic, adLockOptimistic
    Dim NumBlocks As Integer, DestFile, fileNum, block_num As Integer, i As Integer
              Dim fileLength As Long, LeftOver As Long
              Dim FileData, filename As String
              Dim RetVal As Variant
              Dim bytes() As Byte
    Set recordSet = recordCon.Execute("Select * from table1 where filename='" & Text & "';")
    
    fileLength = recordSet("filesize")
    
    If fileLength = 0 Then
    MsgBox "Data Unavailable"
    
    Else
    filename = CreateTempFile("djvu")
    'Print filename
    fileNum = FreeFile
    Open filename For Binary As #fileNum
    fileLength = recordSet("file").ActualSize
    NumBlocks = fileLength \ BlockSize
    LeftOver = fileLength Mod BlockSize
        progress_upload.Show
        progress_upload.Caption = "Getting Files Ready"
    For block_num = 1 To NumBlocks
        bytes() = recordSet!file.GetChunk(BlockSize)
        Put #fileNum, , bytes()
            progress_upload.ProgressBar1.Value = (block_num + 1) / (NumBlocks + 1) * 100
           progress_upload.Label1.Caption = (block_num + 1) / (NumBlocks + 1) * 100
           progress_upload.Label1.Caption = progress_upload.Label1.Caption & "%"
    Next block_num
        progress_upload.Hide
    If LeftOver > 0 Then
        bytes() = recordSet!file.GetChunk(LeftOver)
        Put #fileNum, , bytes()
    End If
    progress_upload.Hide
    Close #fileNum
    'Picture1.Picture = LoadPicture(FileName)
    Form2.Show
    Form2.WindowsMediaPlayer1.URL = filename
    'Form2.WindowsMediaPlayer1.Controls.play
    Form2.WindowsMediaPlayer1.settings.autoStart = True
    End If
    'Kill FileName
    recordSet.Close
    End Sub
    Private Sub ImageShow()
    'MsgBox "In"
    recordSet.Open "Select * from table1", recordCon, adOpenDynamic, adLockOptimistic
    Dim NumBlocks As Integer, DestFile, fileNum, block_num As Integer, i As Integer
              Dim fileLength As Long, LeftOver As Long
              Dim FileData, filename As String
              Dim RetVal As Variant
              Dim bytes() As Byte
    Set recordSet = recordCon.Execute("Select * from table1 where filename='" & ListImages.Text & "';")
    
    fileLength = recordSet("filesize")
    
    If fileLength = 0 Then
    MsgBox "Data Unavailable"
    
    Else
    filename = CreateTempFile("djvu")
    'Print filename
    fileNum = FreeFile
    Open filename For Binary As #fileNum
    fileLength = recordSet("file").ActualSize
    NumBlocks = fileLength \ BlockSize
    LeftOver = fileLength Mod BlockSize
        
    For block_num = 1 To NumBlocks
        bytes() = recordSet!file.GetChunk(BlockSize)
        Put #fileNum, , bytes()
    Next block_num

    If LeftOver > 0 Then
        bytes() = recordSet!file.GetChunk(LeftOver)
        Put #fileNum, , bytes()
    End If
    End If
    Close #fileNum
    Form4.Show
    Form4.Image1.Picture = LoadPicture(filename)
    Kill filename
    recordSet.Close
End Sub

    
Private Function CreateTempFile(sPrefix As String) As String
         Dim sTmpPath As String * 512
         Dim sTmpName As String * 576
         Dim nRet As Long

         nRet = GetTempPath(512, sTmpPath)
         If (nRet > 0 And nRet < 512) Then
            nRet = GetTempFileName(sTmpPath, sPrefix, 0, sTmpName)
            If nRet <> 0 Then
               CreateTempFile = Left$(sTmpName, InStr(sTmpName, vbNullChar) - 1)
            End If
         End If
      End Function

Private Sub Form1_Unload(Cancel As Integer)
    On Error GoTo HandleError
    If MsgBox("Are you sure you want to quit", vbYesNoCancel) = vbYes Then
    recordSet.Close
    recordCon.Close
    Set recordSet = Nothing
    Set recordCon = Nothing
    Exit Sub
    Else
    Cancel = True
    End If
HandleError:
    MsgBox Err
End Sub


Private Sub ListAudio_Click()
    Me.title(0).Caption = ListAudio.Text
    Me.option(0).Visible = True
End Sub


Private Sub ListAudio_LostFocus()
    Me.option(0).Visible = False
End Sub
Private Sub ListImages_Click()
    Me.title(1).Caption = ListImages.Text
    Me.option(1).Visible = True
End Sub

Private Sub ListImages_LostFocus()
    Me.option(1).Visible = False
End Sub

Private Sub ListVideo_Click()
    Me.title(2).Caption = ListVideo.Text
    Me.option(2).Visible = True
End Sub
Private Sub ListVideo_LostFocus()
    Me.option(2).Visible = False
End Sub

Private Sub option_DblClick(Index As Integer)
    Me.option(Index).Visible = False
End Sub

Private Sub playbutton_Click(Index As Integer)
    If Index = 0 Then
    play (ListAudio.Text)
    ElseIf Index = 1 Then
    ImageShow
    ElseIf Index = 2 Then
    play (ListVideo.Text)
    End If
End Sub

Private Sub playPic_GotFocus(Index As Integer)
    If Index = 0 Then
    play (ListAudio.Text)
    ElseIf Index = 1 Then
    ImageShow
    ElseIf Index = 2 Then
    play (ListVideo.Text)
    End If
End Sub

Private Sub Search_Click(Index As Integer)
    recordSet.Open "Select * from table1", recordCon, adOpenDynamic, adLockOptimistic
    If Index = 0 Then
    
        Set recordSet = recordCon.Execute("Select * from table1 where type='image' and filename='" & searchBox(0).Text & "';")
        If recordSet.EOF Then
            MsgBox "No record found", vbCritical, "Search Failed"
            recordSet.Close
            List
            Exit Sub
        Else
            ListImages.Clear
            While Not recordSet.EOF
            ListImages.AddItem (recordSet("filename"))
            recordSet.MoveNext
            Wend
        End If
    End If
    If Index = 1 Then
    
        Set recordSet = recordCon.Execute("Select * from table1 where type='audio' and filename='" & searchBox(1).Text & "';")
        If recordSet.EOF Then
            MsgBox "No record found", vbCritical, "Search Failed"
            recordSet.Close
            List
            Exit Sub
        Else
            ListAudio.Clear
            While Not recordSet.EOF
            ListAudio.AddItem (recordSet("filename"))
            recordSet.MoveNext
            Wend
        End If
    End If
    If Index = 2 Then
    
        Set recordSet = recordCon.Execute("Select * from table1 where type='video' and filename='" & searchBox(2).Text & "';")
        If recordSet.EOF Then
            MsgBox "No record found", vbCritical, "Search Failed"
            recordSet.Close
            List
            Exit Sub
        Else
            ListVideo.Clear
            While Not recordSet.EOF
            ListVideo.AddItem (recordSet("filename"))
            recordSet.MoveNext
            Wend
            End If
    End If
    recordSet.Close
End Sub

Private Sub uploadAudio_Click()
recordSet.Open "Select * from table1", recordCon, adOpenDynamic, adLockOptimistic
Dim path As String, SourceFile As Integer, NumBlocks As Integer, i As Integer
    Dim a As Variant
    Dim fileLength As Long, LeftOver As Long
              Dim FileData As String
              Dim RetVal As Variant
              Dim bytes() As Byte
    If pathText(1).Text = "" Then
        a = MsgBox("Error Empty Path", vbCritical, "Error")
    Else
        path = pathText(1).Text
        SourceFile = FreeFile
        Open path For Binary Access Read As #SourceFile
        fileLength = LOF(SourceFile)
        If fileLength = 0 Then
            MsgBox "File Length zero", vbCritical, "Error"
        Else
            progress_upload.Show
            
            'ProgressBar1.Visible = True
            recordSet.AddNew
            NumBlocks = fileLength \ BlockSize
            LeftOver = fileLength Mod BlockSize
            
            ReDim bytes(LeftOver)
            Get #SourceFile, , bytes()
            recordSet("file").AppendChunk bytes()
    '        RetVal = SysCmd(acSysCmdInitMeter, "Reading BLOB", fileLength \ 1000)
            progress_upload.ProgressBar1.Value = 1 / (NumBlocks + 1) * 100
            progress_upload.Label1.Caption = 1 / (NumBlocks + 1) * 100
            progress_upload.Label1.Caption = progress_upload.Label1.Caption & "%"
            ReDim bytes(BlockSize)
            For i = 1 To NumBlocks
            Get SourceFile, , bytes()
            recordSet("file").AppendChunk bytes()
    '        RetVal = SysCmd(acSysCmdInitMeter, "Reading BLOB", BlockSize * i \ 1000)
            progress_upload.ProgressBar1.Value = (i + 1) / (NumBlocks + 1) * 100
            progress_upload.Label1.Caption = (i + 1) / (NumBlocks + 1) * 100
            progress_upload.Label1.Caption = progress_upload.Label1.Caption & "%"
            Next i
            recordSet.Fields("filename") = filetitle
            recordSet.Fields("extension") = Right$(filetitle, 3)
            recordSet.Fields("filesize") = fileLength
            recordSet.Fields("type") = "audio"
            recordSet.Update
            progress_upload.Hide
     '       RetVal = SysCmd(acSysCmdRemoveMeter)
            Close SourceFile
        End If
    End If
    Set recordSet = recordCon.Execute("Select * from table1 where type='audio';")
    ListAudio.Clear
    While Not recordSet.EOF
    ListAudio.AddItem (recordSet("filename"))
    recordSet.MoveNext
    Wend
    recordSet.Close
End Sub


Private Sub uploadImage_Click()
    recordSet.Open "Select * from table1", recordCon, adOpenDynamic, adLockOptimistic
    Dim path As String, SourceFile As Integer, NumBlocks As Integer, i As Integer
    Dim a As Variant
    Dim fileLength As Long, LeftOver As Long
              Dim FileData As String
              Dim RetVal As Variant
              Dim bytes() As Byte
    If pathText(0).Text = "" Then
        a = MsgBox("Error Empty Path", vbCritical, "Error")
    Else
        path = pathText(0).Text
        SourceFile = FreeFile
        Open path For Binary Access Read As #SourceFile
        fileLength = LOF(SourceFile)
        If fileLength = 0 Then
            MsgBox "File Length zero", vbCritical, "Error"
        Else
            progress_upload.Show
            
            'ProgressBar1.Visible = True
            recordSet.AddNew
            NumBlocks = fileLength \ BlockSize
            LeftOver = fileLength Mod BlockSize
            
            ReDim bytes(LeftOver)
            Get #SourceFile, , bytes()
            recordSet("file").AppendChunk bytes()
    '        RetVal = SysCmd(acSysCmdInitMeter, "Reading BLOB", fileLength \ 1000)
            progress_upload.Label1.Caption = 1 / (NumBlocks + 1) * 100
            progress_upload.Label1.Caption = progress_upload.Label1.Caption & "%"
            ReDim bytes(BlockSize)
            For i = 1 To NumBlocks
            Get SourceFile, , bytes()
            recordSet("file").AppendChunk bytes()
    '        RetVal = SysCmd(acSysCmdInitMeter, "Reading BLOB", BlockSize * i \ 1000)
            progress_upload.ProgressBar1.Value = (i + 1) / (NumBlocks + 1) * 100
            progress_upload.Label1.Caption = (i + 1) / (NumBlocks + 1) * 100
            progress_upload.Label1.Caption = progress_upload.Label1.Caption & "%"
            Next i
            recordSet.Fields("filename") = filetitle
            recordSet.Fields("extension") = Right$(filetitle, 3)
            recordSet.Fields("filesize") = fileLength
            recordSet.Fields("type") = "image"
            recordSet.Update
            progress_upload.Hide
     '       RetVal = SysCmd(acSysCmdRemoveMeter)
            Close SourceFile
        End If
    End If
    Set recordSet = recordCon.Execute("Select * from table1 where type='image';")
    ListImages.Clear
    While Not recordSet.EOF
    ListImages.AddItem (recordSet("filename"))
    recordSet.MoveNext
    Wend
    recordSet.Close
End Sub

Private Sub uploadVideo_Click()
    recordSet.Open "Select * from table1", recordCon, adOpenDynamic, adLockOptimistic
Dim path As String, SourceFile As Integer, NumBlocks As Integer, i As Integer
    Dim a As Variant
    Dim fileLength As Long, LeftOver As Long
              Dim FileData As String
              Dim RetVal As Variant
              Dim bytes() As Byte
    If pathText(2).Text = "" Then
        a = MsgBox("Error Empty Path", vbCritical, "Error")
    Else
        path = pathText(2).Text
        SourceFile = FreeFile
        Open path For Binary Access Read As #SourceFile
        fileLength = LOF(SourceFile)
        If fileLength = 0 Then
            MsgBox "File Length zero", vbCritical, "Error"
        Else
        
            progress_upload.Show
            
            'ProgressBar1.Visible = True
            recordSet.AddNew
            NumBlocks = fileLength \ BlockSize
            LeftOver = fileLength Mod BlockSize
            
            ReDim bytes(LeftOver)
            Get #SourceFile, , bytes()
            recordSet("file").AppendChunk bytes()
    '        RetVal = SysCmd(acSysCmdInitMeter, "Reading BLOB", fileLength \ 1000)
            progress_upload.Label1.Caption = 1 / (NumBlocks + 1) * 100
            progress_upload.Label1.Caption = progress_upload.Label1.Caption & "%"
            ReDim bytes(BlockSize)
            For i = 1 To NumBlocks
            Get SourceFile, , bytes()
            recordSet("file").AppendChunk bytes()
    '        RetVal = SysCmd(acSysCmdInitMeter, "Reading BLOB", BlockSize * i \ 1000)
            progress_upload.ProgressBar1.Value = (i + 1) / (NumBlocks + 1) * 100
            progress_upload.Label1.Caption = (i + 1) / (NumBlocks + 1) * 100
            progress_upload.Label1.Caption = progress_upload.Label1.Caption & "%"
            Next i
            
            recordSet.Fields("filename") = filetitle
            recordSet.Fields("extension") = Right$(filetitle, 3)
            recordSet.Fields("filesize") = fileLength
            recordSet.Fields("type") = "video"
            recordSet.Update
            progress_upload.Hide
     '       RetVal = SysCmd(acSysCmdRemoveMeter)
            Close SourceFile
        End If
    End If
    Set recordSet = recordCon.Execute("Select * from table1 where type='video';")
    ListVideo.Clear
    While Not recordSet.EOF
    ListVideo.AddItem (recordSet("filename"))
    recordSet.MoveNext
    Wend
    recordSet.Close
End Sub



