VERSION 5.00
Begin VB.Form frmConfig 
   BackColor       =   &H00C0C000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configration  -As You Wish 2.1"
   ClientHeight    =   2190
   ClientLeft      =   2325
   ClientTop       =   1230
   ClientWidth     =   3480
   FillColor       =   &H00C0C0C0&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2190
   ScaleWidth      =   3480
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtMicName 
      Height          =   285
      Left            =   1800
      MaxLength       =   15
      TabIndex        =   11
      Text            =   "Panasonic"
      Top             =   480
      Width           =   1575
   End
   Begin VB.TextBox txtThreshold 
      Height          =   285
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   8
      Text            =   "60"
      Top             =   840
      Width           =   615
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   285
      Left            =   2280
      Max             =   100
      SmallChange     =   5
      TabIndex        =   7
      Top             =   840
      Value           =   60
      Width           =   255
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   5
      Top             =   1680
      Width           =   975
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   4
      Top             =   1680
      Width           =   975
   End
   Begin VB.TextBox txtPortAddress 
      Height          =   285
      Left            =   2520
      MaxLength       =   4
      TabIndex        =   2
      Text            =   "378"
      Top             =   1200
      Width           =   615
   End
   Begin VB.TextBox txtSpeaker 
      Height          =   285
      Left            =   1440
      MaxLength       =   20
      TabIndex        =   1
      Text            =   "Asim "
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label lblMicName 
      BackStyle       =   0  'Transparent
      Caption         =   "MicroPhone Name:"
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
      TabIndex        =   10
      Top             =   480
      Width           =   1695
   End
   Begin VB.Label lblThreshold 
      BackStyle       =   0  'Transparent
      Caption         =   "Threshold Level:"
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
      TabIndex        =   9
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label lblHexNotation 
      BackStyle       =   0  'Transparent
      Caption         =   "h"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3240
      TabIndex        =   6
      Top             =   1200
      Width           =   135
   End
   Begin VB.Label lblParallelPortAdd 
      BackStyle       =   0  'Transparent
      Caption         =   "Parallel Port Base Address:"
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
      TabIndex        =   3
      Top             =   1200
      Width           =   2415
   End
   Begin VB.Label lblSpeaker 
      BackStyle       =   0  'Transparent
      Caption         =   "Speaker Name:"
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
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "frmConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim FileNo As Integer
Dim ConfigInfo As ConfigData

Private Sub cmdCancel_Click()
frmMain.Visible = True
Me.Visible = False
End Sub





Private Sub cmdOk_Click()

'Saving the changed values
Call SaveConfigData

frmMain.Visible = True
Me.Visible = False
End Sub

Private Sub Form_Load()

'Retreiving the saved values.
Call RestoreConfigData

frmConfig.Icon = frmMain.Icon
End Sub


Private Sub SaveConfigData()
    FileNo = FreeFile
    Open "AYWConfig.txt" For Random As FileNo Len = 42
        ConfigInfo.SpeakerName = Trim(txtSpeaker.Text)
        ConfigInfo.PportAddress = Trim(txtPortAddress.Text)
        ConfigInfo.MicName = Trim(txtMicName.Text)
        ConfigInfo.ThresholdLevel = Trim(txtThreshold.Text)
        Put #FileNo, 1, ConfigInfo
    Close #FileNo
End Sub

Private Sub RestoreConfigData()
    FileNo = FreeFile
    Open "AYWConfig.txt" For Random As FileNo Len = 42
    Get #FileNo, 1, ConfigInfo
        txtSpeaker.Text = ConfigInfo.SpeakerName
        txtPortAddress.Text = ConfigInfo.PportAddress
    Close #FileNo

End Sub

Private Sub VScroll1_Change()
txtThreshold.Text = VScroll1.Value
End Sub
