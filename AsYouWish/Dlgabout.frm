VERSION 5.00
Begin VB.Form dlgAbout 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Asim SoftTech Inc."
   ClientHeight    =   2250
   ClientLeft      =   2970
   ClientTop       =   1140
   ClientWidth     =   3690
   Icon            =   "Dlgabout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2250
   ScaleWidth      =   3690
   ShowInTaskbar   =   0   'False
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "asim_electro@hotmail.com"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   1080
      TabIndex        =   7
      Top             =   1920
      Width           =   2055
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Email: "
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   600
      TabIndex        =   6
      Top             =   1920
      Width           =   495
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "http://i.am/asim_electro/"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   1320
      TabIndex        =   5
      Top             =   1680
      Width           =   1935
   End
   Begin VB.Label Label5 
      BackColor       =   &H00000000&
      Caption         =   "Website :"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   600
      TabIndex        =   4
      Top             =   1680
      Width           =   735
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Microsoft Certified Professional (M.C.P.) Student of B.Sc. Elect. Engg. (Electronics), U.E.T. Lahore, Pakistan. "
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   600
      TabIndex        =   3
      Top             =   1080
      Width           =   3135
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "SYED ASIM HUSSAIN ABBASI"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   3255
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Written By"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Left            =   1200
      TabIndex        =   1
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "AS YOU WISH 2.1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3735
   End
End
Attribute VB_Name = "dlgAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Unload(Cancel As Integer)
frmMain.Visible = True
End Sub
