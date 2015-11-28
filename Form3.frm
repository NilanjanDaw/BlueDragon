VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "BlueDragon"
   ClientHeight    =   1740
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7725
   LinkTopic       =   "Form3"
   Picture         =   "Form3.frx":0000
   ScaleHeight     =   1740
   ScaleWidth      =   7725
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture3 
      Height          =   855
      Left            =   6000
      ScaleHeight     =   795
      ScaleWidth      =   795
      TabIndex        =   2
      Top             =   120
      Width           =   855
   End
   Begin VB.PictureBox Picture2 
      Height          =   855
      Left            =   3480
      ScaleHeight     =   795
      ScaleWidth      =   795
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
   Begin VB.PictureBox Picture1 
      Height          =   855
      Left            =   840
      ScaleHeight     =   795
      ScaleWidth      =   795
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Delete"
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
      Left            =   5760
      TabIndex        =   5
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Modify"
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
      Left            =   3240
      TabIndex        =   4
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Play"
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
      Left            =   600
      TabIndex        =   3
      Top             =   1200
      Width           =   1335
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
