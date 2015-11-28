VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BlueDragon"
   ClientHeight    =   8325
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15030
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   8325
   ScaleWidth      =   15030
   StartUpPosition =   2  'CenterScreen
   Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer1 
      CausesValidation=   0   'False
      Height          =   8295
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15015
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   -1  'True
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   26485
      _cy             =   14631
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Dim file As String
 

Private Sub Form_Unload(Cancel As Integer)
Kill WindowsMediaPlayer1.URL
End Sub
