VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form progress_upload 
   Caption         =   "Uploading"
   ClientHeight    =   2835
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14895
   LinkTopic       =   "Form2"
   ScaleHeight     =   2835
   ScaleWidth      =   14895
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Width           =   14535
      _ExtentX        =   25638
      _ExtentY        =   873
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Please Wait..."
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   615
      Left            =   4560
      TabIndex        =   1
      Top             =   360
      Width           =   5295
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   735
      Left            =   6240
      TabIndex        =   0
      Top             =   1920
      Width           =   2415
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   16200
      Left            =   0
      Picture         =   "progress_upload.frx":0000
      Top             =   0
      Width           =   28800
   End
End
Attribute VB_Name = "progress_upload"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
