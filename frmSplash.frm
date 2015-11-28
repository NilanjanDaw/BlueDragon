VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   9315
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   15870
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9315
   ScaleWidth      =   15870
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   1920
      ScaleHeight     =   375
      ScaleWidth      =   12615
      TabIndex        =   3
      Top             =   8400
      Width           =   12615
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   3120
      Top             =   5880
   End
   Begin VB.Label percentage 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   7560
      TabIndex        =   4
      Top             =   7920
      Width           =   1575
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Loading please wait"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   1215
      Left            =   5280
      TabIndex        =   2
      Top             =   4920
      Width           =   4095
   End
   Begin VB.Label lblVersion 
      BackStyle       =   0  'Transparent
      Caption         =   "Version 1.0.0"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   11280
      TabIndex        =   1
      Top             =   2640
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "BlueDragon"
      BeginProperty Font 
         Name            =   "Bauhaus 93"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   1095
      Left            =   4080
      TabIndex        =   0
      Top             =   1440
      Width           =   8295
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   16200
      Left            =   0
      Picture         =   "frmSplash.frx":000C
      Top             =   0
      Width           =   28800
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ctr As Integer
Option Explicit

Private Sub Form_KeyPress(KeyAscii As Integer)
    Unload Me
End Sub

Private Sub Form_Load()
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    ctr = 0
End Sub


Private Sub Timer1_Timer()
    ctr = ctr + 1
    If ctr = 20 Then
    Unload Me
    Form1.Show
    Else
    Picture1.Width = ctr / 20 * 12615
    percentage.Caption = ctr / 20 * 100
    percentage.Caption = percentage.Caption & "%"
    End If
End Sub
