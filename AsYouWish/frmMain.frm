VERSION 5.00
Object = "{60462311-3373-11D1-8C43-0060081841DE}#1.0#0"; "Xcommand.dll"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "As You Wish 2.1      - Asim SoftTech Inc."
   ClientHeight    =   4890
   ClientLeft      =   285
   ClientTop       =   720
   ClientWidth     =   9015
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4890
   ScaleWidth      =   9015
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   2160
      Top             =   3840
   End
   Begin VB.CommandButton cmdConfig 
      Caption         =   "&Config"
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
      Left            =   2520
      TabIndex        =   68
      Top             =   4080
      Width           =   1215
   End
   Begin HSRLibCtl.Vcommand Vcommand1 
      Height          =   615
      Left            =   240
      OleObjectBlob   =   "frmMain.frx":030A
      TabIndex        =   67
      Top             =   3480
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
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
      TabIndex        =   66
      Top             =   4080
      Width           =   1215
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   65
      Top             =   4560
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   582
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   2
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   6068
            MinWidth        =   6068
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   9703
            MinWidth        =   9703
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdDeactivate 
      Caption         =   "&Deactivate"
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
      Left            =   5160
      TabIndex        =   64
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton cmdAbout 
      Caption         =   "&About"
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
      Left            =   7800
      TabIndex        =   63
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
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
      Left            =   6600
      TabIndex        =   62
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton cmdActivate 
      Caption         =   "Ac&tivate"
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
      Left            =   3960
      TabIndex        =   61
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton cmdRestore 
      Caption         =   "&Restore"
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
      Left            =   120
      MaskColor       =   &H000000C0&
      TabIndex        =   60
      Top             =   4080
      Width           =   1215
   End
   Begin VB.TextBox txtDevice24 
      Height          =   285
      Left            =   6960
      MaxLength       =   15
      TabIndex        =   51
      Top             =   3240
      Width           =   975
   End
   Begin VB.TextBox txtDevice23 
      Height          =   285
      Left            =   6960
      MaxLength       =   15
      TabIndex        =   50
      Top             =   2880
      Width           =   975
   End
   Begin VB.TextBox txtDevice22 
      Height          =   285
      Left            =   6960
      MaxLength       =   15
      TabIndex        =   49
      Top             =   2520
      Width           =   975
   End
   Begin VB.TextBox txtDevice21 
      Height          =   285
      Left            =   6960
      MaxLength       =   15
      TabIndex        =   48
      Top             =   2160
      Width           =   975
   End
   Begin VB.TextBox txtDevice20 
      Height          =   285
      Left            =   6960
      MaxLength       =   15
      TabIndex        =   47
      Top             =   1800
      Width           =   975
   End
   Begin VB.TextBox txtDevice19 
      Height          =   285
      Left            =   6960
      MaxLength       =   15
      TabIndex        =   46
      Top             =   1440
      Width           =   975
   End
   Begin VB.TextBox txtDevice18 
      Height          =   285
      Left            =   6960
      MaxLength       =   15
      TabIndex        =   45
      Top             =   1080
      Width           =   975
   End
   Begin VB.TextBox txtDevice17 
      Height          =   285
      Left            =   6960
      MaxLength       =   15
      TabIndex        =   44
      Top             =   720
      Width           =   975
   End
   Begin VB.TextBox txtDevice16 
      Height          =   285
      Left            =   3960
      MaxLength       =   15
      TabIndex        =   31
      Top             =   3240
      Width           =   975
   End
   Begin VB.TextBox txtDevice15 
      Height          =   285
      Left            =   3960
      MaxLength       =   15
      TabIndex        =   30
      Top             =   2880
      Width           =   975
   End
   Begin VB.TextBox txtDevice14 
      Height          =   285
      Left            =   3960
      MaxLength       =   15
      TabIndex        =   29
      Top             =   2520
      Width           =   975
   End
   Begin VB.TextBox txtDevice13 
      Height          =   285
      Left            =   3960
      MaxLength       =   15
      TabIndex        =   28
      Top             =   2160
      Width           =   975
   End
   Begin VB.TextBox txtDevice12 
      Height          =   285
      Left            =   3960
      MaxLength       =   15
      TabIndex        =   27
      Top             =   1800
      Width           =   975
   End
   Begin VB.TextBox txtDevice11 
      Height          =   285
      Left            =   3960
      MaxLength       =   15
      TabIndex        =   26
      Top             =   1440
      Width           =   975
   End
   Begin VB.TextBox txtDevice10 
      Height          =   285
      Left            =   3960
      MaxLength       =   15
      TabIndex        =   25
      Top             =   1080
      Width           =   975
   End
   Begin VB.TextBox txtDevice9 
      Height          =   285
      Left            =   3960
      MaxLength       =   15
      TabIndex        =   24
      Top             =   720
      Width           =   975
   End
   Begin VB.TextBox txtDevice8 
      Height          =   285
      Left            =   960
      MaxLength       =   15
      TabIndex        =   8
      Top             =   3240
      Width           =   975
   End
   Begin VB.TextBox txtDevice7 
      Height          =   285
      Left            =   960
      MaxLength       =   15
      TabIndex        =   7
      Top             =   2880
      Width           =   975
   End
   Begin VB.TextBox txtDevice6 
      Height          =   285
      Left            =   960
      MaxLength       =   15
      TabIndex        =   6
      Top             =   2520
      Width           =   975
   End
   Begin VB.TextBox txtDevice5 
      Height          =   285
      Left            =   960
      MaxLength       =   15
      TabIndex        =   5
      Top             =   2160
      Width           =   975
   End
   Begin VB.TextBox txtDevice4 
      Height          =   285
      Left            =   960
      MaxLength       =   15
      TabIndex        =   4
      Top             =   1800
      Width           =   975
   End
   Begin VB.TextBox txtDevice3 
      Height          =   285
      Left            =   960
      MaxLength       =   15
      TabIndex        =   3
      Top             =   1440
      Width           =   975
   End
   Begin VB.TextBox txtDevice2 
      Height          =   285
      Left            =   960
      MaxLength       =   15
      TabIndex        =   2
      Top             =   1080
      Width           =   975
   End
   Begin VB.TextBox txtDevice1 
      Height          =   285
      Left            =   960
      MaxLength       =   15
      TabIndex        =   1
      Top             =   720
      Width           =   975
   End
   Begin VB.Shape shpDevice1 
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   23
      Left            =   8280
      Shape           =   3  'Circle
      Top             =   3240
      Width           =   375
   End
   Begin VB.Shape shpDevice1 
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   22
      Left            =   8280
      Shape           =   3  'Circle
      Top             =   2880
      Width           =   375
   End
   Begin VB.Shape shpDevice1 
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   21
      Left            =   8280
      Shape           =   3  'Circle
      Top             =   2520
      Width           =   375
   End
   Begin VB.Shape shpDevice1 
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   20
      Left            =   8280
      Shape           =   3  'Circle
      Top             =   2160
      Width           =   375
   End
   Begin VB.Shape shpDevice1 
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   19
      Left            =   8280
      Shape           =   3  'Circle
      Top             =   1800
      Width           =   375
   End
   Begin VB.Shape shpDevice1 
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   18
      Left            =   8280
      Shape           =   3  'Circle
      Top             =   1440
      Width           =   375
   End
   Begin VB.Shape shpDevice1 
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   17
      Left            =   8280
      Shape           =   3  'Circle
      Top             =   1080
      Width           =   375
   End
   Begin VB.Shape shpDevice1 
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   16
      Left            =   8280
      Shape           =   3  'Circle
      Top             =   720
      Width           =   375
   End
   Begin VB.Shape shpDevice1 
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   15
      Left            =   5280
      Shape           =   3  'Circle
      Top             =   3240
      Width           =   375
   End
   Begin VB.Shape shpDevice1 
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   14
      Left            =   5280
      Shape           =   3  'Circle
      Top             =   2880
      Width           =   375
   End
   Begin VB.Shape shpDevice1 
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   13
      Left            =   5280
      Shape           =   3  'Circle
      Top             =   2520
      Width           =   375
   End
   Begin VB.Shape shpDevice1 
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   12
      Left            =   5280
      Shape           =   3  'Circle
      Top             =   2160
      Width           =   375
   End
   Begin VB.Shape shpDevice1 
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   11
      Left            =   5280
      Shape           =   3  'Circle
      Top             =   1800
      Width           =   375
   End
   Begin VB.Shape shpDevice1 
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   10
      Left            =   5280
      Shape           =   3  'Circle
      Top             =   1440
      Width           =   375
   End
   Begin VB.Shape shpDevice1 
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   9
      Left            =   5280
      Shape           =   3  'Circle
      Top             =   1080
      Width           =   375
   End
   Begin VB.Shape shpDevice1 
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   8
      Left            =   5280
      Shape           =   3  'Circle
      Top             =   720
      Width           =   375
   End
   Begin VB.Shape shpDevice1 
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   7
      Left            =   2160
      Shape           =   3  'Circle
      Top             =   3240
      Width           =   375
   End
   Begin VB.Shape shpDevice1 
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   6
      Left            =   2160
      Shape           =   3  'Circle
      Top             =   2880
      Width           =   375
   End
   Begin VB.Shape shpDevice1 
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   5
      Left            =   2160
      Shape           =   3  'Circle
      Top             =   2520
      Width           =   375
   End
   Begin VB.Shape shpDevice1 
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   4
      Left            =   2160
      Shape           =   3  'Circle
      Top             =   2160
      Width           =   375
   End
   Begin VB.Shape shpDevice1 
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   3
      Left            =   2160
      Shape           =   3  'Circle
      Top             =   1800
      Width           =   375
   End
   Begin VB.Shape shpDevice1 
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   2
      Left            =   2160
      Shape           =   3  'Circle
      Top             =   1440
      Width           =   375
   End
   Begin VB.Shape shpDevice1 
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   1
      Left            =   2160
      Shape           =   3  'Circle
      Top             =   1080
      Width           =   375
   End
   Begin VB.Shape shpDevice1 
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   0
      Left            =   2160
      Shape           =   3  'Circle
      Top             =   720
      Width           =   375
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00E0E0E0&
      X1              =   -105
      X2              =   9015
      Y1              =   3855
      Y2              =   3855
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00808080&
      X1              =   0
      X2              =   9000
      Y1              =   3840
      Y2              =   3840
   End
   Begin VB.Label lblDNo24 
      BackStyle       =   0  'Transparent
      Caption         =   "24."
      Height          =   255
      Left            =   6120
      TabIndex        =   59
      Top             =   3240
      Width           =   255
   End
   Begin VB.Label lblDNo23 
      BackStyle       =   0  'Transparent
      Caption         =   "23."
      Height          =   255
      Left            =   6120
      TabIndex        =   58
      Top             =   2880
      Width           =   255
   End
   Begin VB.Label lblDNo22 
      BackStyle       =   0  'Transparent
      Caption         =   "22."
      Height          =   255
      Left            =   6120
      TabIndex        =   57
      Top             =   2520
      Width           =   255
   End
   Begin VB.Label lblDNo21 
      BackStyle       =   0  'Transparent
      Caption         =   "21."
      Height          =   255
      Left            =   6120
      TabIndex        =   56
      Top             =   2160
      Width           =   255
   End
   Begin VB.Label lblDNo20 
      BackStyle       =   0  'Transparent
      Caption         =   "20."
      Height          =   255
      Left            =   6120
      TabIndex        =   55
      Top             =   1800
      Width           =   255
   End
   Begin VB.Label lblDNo19 
      BackStyle       =   0  'Transparent
      Caption         =   "19."
      Height          =   255
      Left            =   6120
      TabIndex        =   54
      Top             =   1440
      Width           =   255
   End
   Begin VB.Label lblDNo18 
      BackStyle       =   0  'Transparent
      Caption         =   "18."
      Height          =   255
      Left            =   6120
      TabIndex        =   53
      Top             =   1080
      Width           =   255
   End
   Begin VB.Label lblDNo17 
      BackStyle       =   0  'Transparent
      Caption         =   "17."
      Height          =   255
      Left            =   6120
      TabIndex        =   52
      Top             =   720
      Width           =   255
   End
   Begin VB.Label lblStatus3 
      BackStyle       =   0  'Transparent
      Caption         =   "Status"
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
      Left            =   8160
      TabIndex        =   43
      Top             =   240
      Width           =   615
   End
   Begin VB.Label lblDeviceName3 
      BackStyle       =   0  'Transparent
      Caption         =   "Device Name"
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
      Left            =   6960
      TabIndex        =   42
      Top             =   120
      Width           =   615
   End
   Begin VB.Label lblDeviceN3 
      BackStyle       =   0  'Transparent
      Caption         =   "Device No."
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
      Left            =   6120
      TabIndex        =   41
      Top             =   120
      Width           =   615
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00808080&
      X1              =   5970
      X2              =   5970
      Y1              =   -30
      Y2              =   3855
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00E0E0E0&
      X1              =   5985
      X2              =   5985
      Y1              =   0
      Y2              =   3855
   End
   Begin VB.Label Label25 
      BackColor       =   &H00C0C000&
      Height          =   3975
      Left            =   5880
      TabIndex        =   40
      Top             =   -120
      Width           =   3135
   End
   Begin VB.Label lblDNo16 
      BackStyle       =   0  'Transparent
      Caption         =   "16."
      Height          =   255
      Left            =   3120
      TabIndex        =   39
      Top             =   3240
      Width           =   255
   End
   Begin VB.Label lblDNo15 
      BackStyle       =   0  'Transparent
      Caption         =   "15."
      Height          =   255
      Left            =   3120
      TabIndex        =   38
      Top             =   2880
      Width           =   255
   End
   Begin VB.Label lblDNo14 
      BackStyle       =   0  'Transparent
      Caption         =   "14."
      Height          =   255
      Left            =   3120
      TabIndex        =   37
      Top             =   2520
      Width           =   255
   End
   Begin VB.Label lblDNo13 
      BackStyle       =   0  'Transparent
      Caption         =   "13."
      Height          =   255
      Left            =   3120
      TabIndex        =   36
      Top             =   2160
      Width           =   255
   End
   Begin VB.Label lblDNo12 
      BackStyle       =   0  'Transparent
      Caption         =   "12."
      Height          =   255
      Left            =   3120
      TabIndex        =   35
      Top             =   1800
      Width           =   255
   End
   Begin VB.Label lblDNo11 
      BackStyle       =   0  'Transparent
      Caption         =   "11."
      Height          =   255
      Left            =   3120
      TabIndex        =   34
      Top             =   1440
      Width           =   255
   End
   Begin VB.Label lblDNo10 
      BackStyle       =   0  'Transparent
      Caption         =   "10."
      Height          =   255
      Left            =   3120
      TabIndex        =   33
      Top             =   1080
      Width           =   255
   End
   Begin VB.Label lblDNo9 
      BackStyle       =   0  'Transparent
      Caption         =   "9."
      Height          =   255
      Left            =   3120
      TabIndex        =   32
      Top             =   720
      Width           =   255
   End
   Begin VB.Label lblStatus2 
      BackStyle       =   0  'Transparent
      Caption         =   "Status"
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
      Left            =   5160
      TabIndex        =   23
      Top             =   240
      Width           =   615
   End
   Begin VB.Label lblDeviceName2 
      BackStyle       =   0  'Transparent
      Caption         =   "Device Name"
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
      Left            =   3960
      TabIndex        =   22
      Top             =   120
      Width           =   615
   End
   Begin VB.Label lblDeviceN2 
      BackStyle       =   0  'Transparent
      Caption         =   "Device No."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   3120
      TabIndex        =   21
      Top             =   135
      Width           =   645
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00E0E0E0&
      X1              =   2880
      X2              =   2880
      Y1              =   -60
      Y2              =   3855
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   2865
      X2              =   2865
      Y1              =   -30
      Y2              =   3840
   End
   Begin VB.Label Label13 
      BackColor       =   &H00C0C000&
      Height          =   4095
      Left            =   2880
      TabIndex        =   20
      Top             =   -240
      Width           =   3015
   End
   Begin VB.Label lblStatus1 
      BackStyle       =   0  'Transparent
      Caption         =   "Status"
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
      Left            =   2115
      TabIndex        =   19
      Top             =   225
      Width           =   615
   End
   Begin VB.Label lblDNo8 
      BackStyle       =   0  'Transparent
      Caption         =   "8."
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   3240
      Width           =   255
   End
   Begin VB.Label lblDNo7 
      BackStyle       =   0  'Transparent
      Caption         =   "7."
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   2880
      Width           =   255
   End
   Begin VB.Label lblDNo6 
      BackStyle       =   0  'Transparent
      Caption         =   "6."
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   2520
      Width           =   255
   End
   Begin VB.Label lblDNo5 
      BackStyle       =   0  'Transparent
      Caption         =   "5."
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   2160
      Width           =   255
   End
   Begin VB.Label lblDNo4 
      BackStyle       =   0  'Transparent
      Caption         =   "4."
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   1800
      Width           =   255
   End
   Begin VB.Label lblDNo3 
      BackStyle       =   0  'Transparent
      Caption         =   "3."
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   1440
      Width           =   255
   End
   Begin VB.Label lblDNo2 
      BackStyle       =   0  'Transparent
      Caption         =   "2."
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   1080
      Width           =   255
   End
   Begin VB.Label lblDNo1 
      BackStyle       =   0  'Transparent
      Caption         =   "1."
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   720
      Width           =   255
   End
   Begin VB.Label lblDeviceN1 
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Device No."
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
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   615
   End
   Begin VB.Label lblDeviceName1 
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Device Name"
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
      Left            =   960
      TabIndex        =   9
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C000&
      Height          =   4095
      Left            =   -120
      TabIndex        =   0
      Top             =   -240
      Width           =   3015
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim DeviceNames As DeviceData
Dim ConfigValues As ConfigData

Dim recordlen As Integer
Dim FileNo As Integer
Public MainCmdMenu As Long

Dim PortAPre As Byte
Dim PortBPre As Byte
Dim PortCPre As Byte

Dim DeviceState(23) As Boolean
Dim PositionNo As Integer

Dim pPortBaddress As Integer


Private Sub cmdAbout_Click()
dlgAbout.Show 1
End Sub

Private Sub cmdActivate_Click()
        FileNo = FreeFile
    Open "AYWConfig.txt" For Random As FileNo Len = 42
    Get #FileNo, 1, ConfigValues
        pPortBaddress = "&H" + Trim(ConfigValues.PportAddress)
    Close #FileNo

    cmdRestore.Enabled = False
    cmdConfig.Enabled = False
    cmdSave.Enabled = False
    cmdExit.Enabled = False
    cmdActivate.Enabled = False
    cmdDeactivate.Enabled = True
    
    Vcommand1.Initialized = 1
    MainCmdMenu = Vcommand1.MenuCreate(App.EXEName, "Menu1", 4)
    Vcommand1.Enabled = 1
    Vcommand1.Threshold = 20
    If txtDevice1.Text <> "" Then
    Vcommand1.AddCommand MainCmdMenu, 0, (txtDevice1.Text), "", "catagory1", 0, ""
    End If
    If txtDevice2.Text <> "" Then
    Vcommand1.AddCommand MainCmdMenu, 1, (txtDevice2.Text), "", "catagory1", 0, ""
    End If
    If txtDevice3.Text <> "" Then
    Vcommand1.AddCommand MainCmdMenu, 2, (txtDevice3.Text), "", "catagory1", 0, ""
    End If
    If txtDevice4.Text <> "" Then
    Vcommand1.AddCommand MainCmdMenu, 3, (txtDevice4.Text), "", "catagory1", 0, ""
    End If
    If txtDevice5.Text <> "" Then
    Vcommand1.AddCommand MainCmdMenu, 4, (txtDevice5.Text), "", "catagory1", 0, ""
    End If
    If txtDevice6.Text <> "" Then
    Vcommand1.AddCommand MainCmdMenu, 5, (txtDevice6.Text), "", "catagory1", 0, ""
    End If
    If txtDevice7.Text <> "" Then
    Vcommand1.AddCommand MainCmdMenu, 6, (txtDevice7.Text), "", "catagory1", 0, ""
    End If
    If txtDevice8.Text <> "" Then
    Vcommand1.AddCommand MainCmdMenu, 7, (txtDevice8.Text), "", "catagory1", 0, ""
    End If
    If txtDevice9.Text <> "" Then
    Vcommand1.AddCommand MainCmdMenu, 8, (txtDevice9.Text), "", "catagory1", 0, ""
    End If
    If txtDevice10.Text <> "" Then
    Vcommand1.AddCommand MainCmdMenu, 9, (txtDevice10.Text), "", "catagory1", 0, ""
    End If
    If txtDevice11.Text <> "" Then
    Vcommand1.AddCommand MainCmdMenu, 10, (txtDevice11.Text), "", "catagory1", 0, ""
    End If
    If txtDevice12.Text <> "" Then
    Vcommand1.AddCommand MainCmdMenu, 11, (txtDevice12.Text), "", "catagory1", 0, ""
    End If
    If txtDevice13.Text <> "" Then
    Vcommand1.AddCommand MainCmdMenu, 12, (txtDevice13.Text), "", "catagory1", 0, ""
    End If
    If txtDevice14.Text <> "" Then
    Vcommand1.AddCommand MainCmdMenu, 13, (txtDevice14.Text), "", "catagory1", 0, ""
    End If
    If txtDevice15.Text <> "" Then
    Vcommand1.AddCommand MainCmdMenu, 14, (txtDevice15.Text), "", "catagory1", 0, ""
    End If
    If txtDevice16.Text <> "" Then
    Vcommand1.AddCommand MainCmdMenu, 15, (txtDevice16.Text), "", "catagory1", 0, ""
    End If
    If txtDevice17.Text <> "" Then
    Vcommand1.AddCommand MainCmdMenu, 16, (txtDevice17.Text), "", "catagory1", 0, ""
    End If
    If txtDevice18.Text <> "" Then
    Vcommand1.AddCommand MainCmdMenu, 17, (txtDevice18.Text), "", "catagory1", 0, ""
    End If
    If txtDevice19.Text <> "" Then
    Vcommand1.AddCommand MainCmdMenu, 18, (txtDevice19.Text), "", "catagory1", 0, ""
    End If
    If txtDevice20.Text <> "" Then
    Vcommand1.AddCommand MainCmdMenu, 19, (txtDevice20.Text), "", "catagory1", 0, ""
    End If
    If txtDevice21.Text <> "" Then
    Vcommand1.AddCommand MainCmdMenu, 20, (txtDevice21.Text), "", "catagory1", 0, ""
    End If
    If txtDevice22.Text <> "" Then
    Vcommand1.AddCommand MainCmdMenu, 21, (txtDevice22.Text), "", "catagory1", 0, ""
    End If
    If txtDevice23.Text <> "" Then
    Vcommand1.AddCommand MainCmdMenu, 22, (txtDevice23.Text), "", "catagory1", 0, ""
    End If
    If txtDevice24.Text <> "" Then
    Vcommand1.AddCommand MainCmdMenu, 23, (txtDevice24.Text), "", "catagory1", 0, ""
    End If
    
    Vcommand1.Activate MainCmdMenu
    StatusBar1.Panels(1).Text = "Listening..."
    
    Timer1.Enabled = True
End Sub

Private Sub cmdConfig_Click()
Me.Visible = False
    
'Loading Form Congiguration
frmConfig.Visible = True

End Sub

Private Sub cmdDeactivate_Click()
    
    Vcommand1.Deactivate MainCmdMenu

    cmdRestore.Enabled = True
    cmdConfig.Enabled = True
    cmdSave.Enabled = True
    cmdExit.Enabled = True
    cmdActivate.Enabled = True
    cmdDeactivate.Enabled = False

    StatusBar1.Panels(1).Text = "Edit Mode"
    StatusBar1.Panels(2).Text = " "
      
    Call OFFall
    Timer1.Enabled = False
End Sub


Private Sub cmdExit_Click()
    Vcommand1.ReleaseMenu MainCmdMenu
    End
End Sub

Private Sub cmdRestore_Click()
    FileNo = FreeFile
    Open "Devices.txt" For Random As FileNo Len = recordlen
         Get #FileNo, 1, DeviceNames
        'Retriving the Names of the Devices from the saved file.
        txtDevice1.Text = Trim(DeviceNames.Device1)
        txtDevice2.Text = Trim(DeviceNames.Device2)
        txtDevice3.Text = Trim(DeviceNames.Device3)
        txtDevice4.Text = Trim(DeviceNames.Device4)
        txtDevice5.Text = Trim(DeviceNames.Device5)
        txtDevice6.Text = Trim(DeviceNames.Device6)
        txtDevice7.Text = Trim(DeviceNames.Device7)
        txtDevice8.Text = Trim(DeviceNames.Device8)
        txtDevice9.Text = Trim(DeviceNames.Device9)
        txtDevice10.Text = Trim(DeviceNames.Device10)
        txtDevice11.Text = Trim(DeviceNames.Device11)
        txtDevice12.Text = Trim(DeviceNames.Device12)
        txtDevice13.Text = Trim(DeviceNames.Device13)
        txtDevice14.Text = Trim(DeviceNames.Device14)
        txtDevice15.Text = Trim(DeviceNames.Device15)
        txtDevice16.Text = Trim(DeviceNames.Device16)
        txtDevice17.Text = Trim(DeviceNames.Device17)
        txtDevice18.Text = Trim(DeviceNames.Device18)
        txtDevice19.Text = Trim(DeviceNames.Device19)
        txtDevice20.Text = Trim(DeviceNames.Device20)
        txtDevice21.Text = Trim(DeviceNames.Device21)
        txtDevice22.Text = Trim(DeviceNames.Device22)
        txtDevice23.Text = Trim(DeviceNames.Device23)
        txtDevice24.Text = Trim(DeviceNames.Device24)
        
     Close #FileNo
    
End Sub

Private Sub cmdSave_Click()
    FileNo = FreeFile
    Open "Devices.txt" For Random As FileNo Len = recordlen
        DeviceNames.Device1 = Trim(txtDevice1.Text)
        DeviceNames.Device2 = Trim(txtDevice2.Text)
        DeviceNames.Device3 = Trim(txtDevice3.Text)
        DeviceNames.Device4 = Trim(txtDevice4.Text)
        DeviceNames.Device5 = Trim(txtDevice5.Text)
        DeviceNames.Device6 = Trim(txtDevice6.Text)
        DeviceNames.Device7 = Trim(txtDevice7.Text)
        DeviceNames.Device8 = Trim(txtDevice8.Text)
        DeviceNames.Device9 = Trim(txtDevice9.Text)
        DeviceNames.Device10 = Trim(txtDevice10.Text)
        DeviceNames.Device11 = Trim(txtDevice11.Text)
        DeviceNames.Device12 = Trim(txtDevice12.Text)
        DeviceNames.Device13 = Trim(txtDevice13.Text)
        DeviceNames.Device14 = Trim(txtDevice14.Text)
        DeviceNames.Device15 = Trim(txtDevice15.Text)
        DeviceNames.Device16 = Trim(txtDevice16.Text)
        DeviceNames.Device17 = Trim(txtDevice17.Text)
        DeviceNames.Device18 = Trim(txtDevice18.Text)
        DeviceNames.Device19 = Trim(txtDevice19.Text)
        DeviceNames.Device20 = Trim(txtDevice20.Text)
        DeviceNames.Device21 = Trim(txtDevice21.Text)
        DeviceNames.Device22 = Trim(txtDevice22.Text)
        DeviceNames.Device23 = Trim(txtDevice23.Text)
        DeviceNames.Device24 = Trim(txtDevice24.Text)
        
        Put #FileNo, 1, DeviceNames
    Close #FileNo

End Sub


Private Sub Form_Load()
    
    'Loading frm config
    Load frmConfig
    
    FileNo = FreeFile
    Open "AYWConfig.txt" For Random As FileNo Len = 42
    Get #FileNo, 1, ConfigValues
    pPortBaddress = "&H" + Trim(ConfigValues.PportAddress)
    Close #FileNo

    'Controlling the states of the Buttons at the start.
    cmdRestore.Enabled = True
    cmdConfig.Enabled = True
    cmdSave.Enabled = True
    cmdExit.Enabled = True
    cmdActivate.Enabled = True
    cmdDeactivate.Enabled = False
    
    recordlen = Len(DeviceNames)     'Finding the length of record
    Call cmdRestore_Click
    
    'Status Bar Displays the staus.
    StatusBar1.Panels(1).Text = "Edit Mode"
    
    'Initializing the PPI for Mode 0 operation
    'with all Ports for Output purposes.
    DlPortWritePortUchar (pPortBaddress + 2), &HD
    DlPortWritePortUchar pPortBaddress, &H80
    DlPortWritePortUchar (pPortBaddress + 2), &H5
  
    Call OFFall
    
    
End Sub


Private Sub Form_Unload(Cancel As Integer)
Call OFFall
Call cmdExit_Click

End Sub


Private Sub Timer1_Timer()
StatusBar1.Panels(2).Text = " "
End Sub

Private Sub Vcommand1_CommandRecognize(ByVal id As Long, ByVal CmdName As String, ByVal Flags As Long, ByVal Action As String, ByVal NumLists As Long, ByVal ListValues As String, ByVal command As String)
StatusBar1.Panels(2).Text = "Command Heard: " + command

If shpDevice1(id).FillColor = vbBlack Then
 shpDevice1(id).FillColor = vbYellow
Else
 shpDevice1(id).FillColor = vbBlack
End If

If id >= 0 And id <= 7 Then
Call ppiPortA(CInt(id))

ElseIf id >= 8 And id <= 15 Then
Call ppiPortB(CInt(id))

ElseIf id >= 16 And id <= 23 Then
Call ppiPortC(CInt(id))
End If

End Sub
Private Sub ppiPortA(DeviceID As Integer)


PositionNo = 2 ^ DeviceID

Select Case DeviceID
    Case 0
        If DeviceState(0) = True Then
            PositionNo = -PositionNo
            DeviceState(0) = False
        Else
            DeviceState(0) = True
        End If
    Case 1
        If DeviceState(1) = True Then
            PositionNo = -PositionNo
            DeviceState(1) = False
        Else
            DeviceState(1) = True
        End If
    Case 2
        If DeviceState(2) = True Then
            PositionNo = -PositionNo
            DeviceState(2) = False
        Else
            DeviceState(2) = True
        End If
    Case 3
        If DeviceState(3) = True Then
            PositionNo = -PositionNo
            DeviceState(3) = False
        Else
            DeviceState(3) = True
        End If
    Case 4
        If DeviceState(4) = True Then
            PositionNo = -PositionNo
            DeviceState(4) = False
        Else
            DeviceState(4) = True
        End If
    Case 5
        If DeviceState(5) = True Then
            PositionNo = -PositionNo
            DeviceState(5) = False
        Else
            DeviceState(5) = True
        End If
    Case 6
        If DeviceState(6) = True Then
            PositionNo = -PositionNo
            DeviceState(6) = False
        Else
            DeviceState(6) = True
        End If
    Case 7
        If DeviceState(7) = True Then
            PositionNo = -PositionNo
            DeviceState(7) = False
        Else
            DeviceState(7) = True
        End If
End Select

PortAPre = PortAPre + PositionNo
    'Initializing the PPI for Mode 0 operation
    'with all Ports for Output purposes.
    DlPortWritePortUchar (pPortBaddress + 2), &HD
    DlPortWritePortUchar pPortBaddress, &H80
    DlPortWritePortUchar (pPortBaddress + 2), &H5


DlPortWritePortUchar (pPortBaddress + 2), &HB
DlPortWritePortUchar pPortBaddress, PortAPre
DlPortWritePortUchar (pPortBaddress + 2), &H3



End Sub

Private Sub ppiPortB(DeviceID As Integer)

PositionNo = 2 ^ (DeviceID - 8)

Select Case DeviceID
    Case 8
        If DeviceState(8) = True Then
            PositionNo = -PositionNo
            DeviceState(8) = False
        Else
            DeviceState(8) = True
        End If
    Case 9
        If DeviceState(9) = True Then
            PositionNo = -PositionNo
            DeviceState(9) = False
        Else
            DeviceState(9) = True
        End If
    Case 10
        If DeviceState(10) = True Then
            PositionNo = -PositionNo
            DeviceState(10) = False
        Else
            DeviceState(10) = True
        End If
    Case 11
        If DeviceState(11) = True Then
            PositionNo = -PositionNo
            DeviceState(11) = False
        Else
            DeviceState(11) = True
        End If
    Case 12
        If DeviceState(12) = True Then
            PositionNo = -PositionNo
            DeviceState(12) = False
        Else
            DeviceState(12) = True
        End If
    Case 13
        If DeviceState(13) = True Then
            PositionNo = -PositionNo
            DeviceState(13) = False
        Else
            DeviceState(13) = True
        End If
    Case 14
        If DeviceState(14) = True Then
            PositionNo = -PositionNo
            DeviceState(14) = False
        Else
            DeviceState(14) = True
        End If
    Case 15
        If DeviceState(15) = True Then
            PositionNo = -PositionNo
            DeviceState(15) = False
        Else
            DeviceState(15) = True
        End If
End Select

PortBPre = PortBPre + PositionNo


    'Initializing the PPI for Mode 0 operation
    'with all Ports for Output purposes.
    DlPortWritePortUchar (pPortBaddress + 2), &HD
    DlPortWritePortUchar pPortBaddress, &H80
    DlPortWritePortUchar (pPortBaddress + 2), &H5




DlPortWritePortUchar (pPortBaddress + 2), &H9
DlPortWritePortUchar pPortBaddress, PortBPre
DlPortWritePortUchar (pPortBaddress + 2), &H1

End Sub

Private Sub ppiPortC(DeviceID As Integer)

PositionNo = 2 ^ (DeviceID - 16)

Select Case DeviceID
    Case 16
        If DeviceState(16) = True Then
            PositionNo = -PositionNo
            DeviceState(16) = False
        Else
            DeviceState(16) = True
        End If
    Case 17
        If DeviceState(17) = True Then
            PositionNo = -PositionNo
            DeviceState(17) = False
        Else
            DeviceState(17) = True
        End If
    Case 18
        If DeviceState(18) = True Then
            PositionNo = -PositionNo
            DeviceState(18) = False
        Else
            DeviceState(18) = True
        End If
    Case 19
        If DeviceState(19) = True Then
            PositionNo = -PositionNo
            DeviceState(19) = False
        Else
            DeviceState(19) = True
        End If
    Case 20
        If DeviceState(20) = True Then
            PositionNo = -PositionNo
            DeviceState(20) = False
        Else
            DeviceState(20) = True
        End If
    Case 21
        If DeviceState(21) = True Then
            PositionNo = -PositionNo
            DeviceState(21) = False
        Else
            DeviceState(21) = True
        End If
    Case 22
        If DeviceState(22) = True Then
            PositionNo = -PositionNo
            DeviceState(22) = False
        Else
            DeviceState(22) = True
        End If
    Case 23
        If DeviceState(23) = True Then
            PositionNo = -PositionNo
            DeviceState(23) = False
        Else
            DeviceState(23) = True
        End If
End Select

PortCPre = PortCPre + PositionNo

    'Initializing the PPI for Mode 0 operation
    'with all Ports for Output purposes.
    DlPortWritePortUchar (pPortBaddress + 2), &HD
    DlPortWritePortUchar pPortBaddress, &H80
    DlPortWritePortUchar (pPortBaddress + 2), &H5


DlPortWritePortUchar (pPortBaddress + 2), &HF
DlPortWritePortUchar pPortBaddress, PortCPre
DlPortWritePortUchar (pPortBaddress + 2), &H7

End Sub


Private Sub OFFall()
    Dim Dummy As Integer
    
    'Making all outputs LOW

    DlPortWritePortUchar (pPortBaddress + 2), &HB
    DlPortWritePortUchar pPortBaddress, &H0
    DlPortWritePortUchar (pPortBaddress + 2), &H3

    DlPortWritePortUchar (pPortBaddress + 2), &H9
    DlPortWritePortUchar pPortBaddress, &H0
    DlPortWritePortUchar (pPortBaddress + 2), &H1


    DlPortWritePortUchar (pPortBaddress + 2), &HF
    DlPortWritePortUchar pPortBaddress, &H0
    DlPortWritePortUchar (pPortBaddress + 2), &H7

    For Dummy = 0 To 23
    shpDevice1(Dummy).FillColor = vbBlack
    Next
End Sub
