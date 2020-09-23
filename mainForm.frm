VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form mainForm 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Sniper"
   ClientHeight    =   4575
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5700
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   Picture         =   "mainForm.frx":0000
   ScaleHeight     =   4575
   ScaleMode       =   0  'User
   ScaleWidth      =   5685.04
   StartUpPosition =   2  'CenterScreen
   Begin InetCtlsObjects.Inet iNet 
      Left            =   4110
      Top             =   15
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.TextBox threads 
      Appearance      =   0  'Flat
      BackColor       =   &H008C6456&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   5895
      MaxLength       =   3
      TabIndex        =   35
      Text            =   "200"
      Top             =   3750
      Width           =   480
   End
   Begin VB.TextBox timeout 
      Appearance      =   0  'Flat
      BackColor       =   &H008C6456&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   5880
      MaxLength       =   5
      TabIndex        =   33
      Text            =   "4"
      Top             =   2400
      Width           =   705
   End
   Begin VB.CheckBox verifyCon 
      Appearance      =   0  'Flat
      BackColor       =   &H00312118&
      Caption         =   "Verify"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   5880
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   31
      Top             =   1005
      Width           =   1320
   End
   Begin VB.ListBox results 
      Appearance      =   0  'Flat
      BackColor       =   &H00312118&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1290
      ItemData        =   "mainForm.frx":B018A
      Left            =   210
      List            =   "mainForm.frx":B018C
      TabIndex        =   25
      Top             =   2235
      Width           =   5235
   End
   Begin MSWinsockLib.Winsock wSock 
      Index           =   0
      Left            =   4695
      Top             =   -15
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox portEnd 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H008C6456&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   4440
      MaxLength       =   5
      TabIndex        =   17
      Text            =   "65535"
      Top             =   1455
      Width           =   840
   End
   Begin VB.TextBox portStart 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H008C6456&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   4440
      MaxLength       =   5
      TabIndex        =   16
      Text            =   "1"
      Top             =   975
      Width           =   840
   End
   Begin VB.TextBox ipEnd 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H008C6456&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   4
      Left            =   3030
      MaxLength       =   3
      TabIndex        =   7
      Text            =   "1"
      Top             =   1455
      Width           =   480
   End
   Begin VB.TextBox ipEnd 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H008C6456&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   3
      Left            =   2310
      MaxLength       =   3
      TabIndex        =   6
      Text            =   "0"
      Top             =   1455
      Width           =   480
   End
   Begin VB.TextBox ipEnd 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H008C6456&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   2
      Left            =   1590
      MaxLength       =   3
      TabIndex        =   5
      Text            =   "0"
      Top             =   1455
      Width           =   480
   End
   Begin VB.TextBox ipEnd 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H008C6456&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   1
      Left            =   870
      MaxLength       =   3
      TabIndex        =   4
      Text            =   "127"
      Top             =   1455
      Width           =   480
   End
   Begin VB.TextBox ipStart 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H008C6456&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   4
      Left            =   3030
      MaxLength       =   3
      TabIndex        =   3
      Text            =   "1"
      Top             =   975
      Width           =   480
   End
   Begin VB.TextBox ipStart 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H008C6456&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   3
      Left            =   2325
      MaxLength       =   3
      TabIndex        =   2
      Text            =   "0"
      Top             =   975
      Width           =   480
   End
   Begin VB.TextBox ipStart 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H008C6456&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   2
      Left            =   1590
      MaxLength       =   3
      TabIndex        =   1
      Text            =   "0"
      Top             =   975
      Width           =   480
   End
   Begin VB.TextBox ipStart 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H008C6456&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   1
      Left            =   870
      MaxLength       =   3
      TabIndex        =   0
      Text            =   "127"
      Top             =   975
      Width           =   480
   End
   Begin VB.Label clear 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "[ Clear ]"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4725
      TabIndex        =   42
      Top             =   1935
      Width           =   750
   End
   Begin VB.Label description 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Version 1.1     Author: Luke Huxley"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   435
      Index           =   5
      Left            =   120
      TabIndex        =   41
      Top             =   4680
      Width           =   1740
   End
   Begin VB.Line Line1 
      BorderColor     =   &H004F3830&
      X1              =   194.488
      X2              =   5415.749
      Y1              =   3675
      Y2              =   3675
   End
   Begin VB.Label commandButton 
      Alignment       =   2  'Center
      BackColor       =   &H004F3830&
      BackStyle       =   0  'Transparent
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   195
      TabIndex        =   40
      Top             =   3825
      Width           =   1125
   End
   Begin VB.Label commandButton 
      Alignment       =   2  'Center
      BackColor       =   &H004F3830&
      BackStyle       =   0  'Transparent
      Caption         =   "About"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   1440
      TabIndex        =   39
      Top             =   3825
      Width           =   1125
   End
   Begin VB.Label commandButton 
      Alignment       =   2  'Center
      BackColor       =   &H004F3830&
      BackStyle       =   0  'Transparent
      Caption         =   "Apply"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   5835
      TabIndex        =   38
      Top             =   4215
      Width           =   1470
   End
   Begin VB.Label description 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "ports"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   4
      Left            =   6465
      TabIndex        =   37
      Top             =   3750
      Width           =   540
   End
   Begin VB.Label description 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Number of ports to connect to:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   435
      Index           =   3
      Left            =   5865
      TabIndex        =   36
      Top             =   3225
      Width           =   1365
   End
   Begin VB.Shape textBorder 
      BorderColor     =   &H006C4D42&
      Height          =   255
      Index           =   17
      Left            =   5880
      Top             =   3735
      Width           =   510
   End
   Begin VB.Label description 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "ms"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   2
      Left            =   6675
      TabIndex        =   34
      Top             =   2400
      Width           =   270
   End
   Begin VB.Shape textBorder 
      BorderColor     =   &H006C4D42&
      Height          =   255
      Index           =   16
      Left            =   5865
      Top             =   2385
      Width           =   735
   End
   Begin VB.Label description 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Time given to connect:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   480
      Index           =   1
      Left            =   5880
      TabIndex        =   32
      Top             =   1830
      Width           =   1335
   End
   Begin VB.Label description 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Verify unable to connect:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   450
      Index           =   0
      Left            =   5880
      TabIndex        =   30
      Top             =   540
      Width           =   1380
   End
   Begin VB.Shape textBorder 
      BackColor       =   &H00312118&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H006C4D42&
      Height          =   930
      Index           =   15
      Left            =   5790
      Top             =   3165
      Width           =   1500
   End
   Begin VB.Shape textBorder 
      BackColor       =   &H00312118&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H006C4D42&
      Height          =   975
      Index           =   14
      Left            =   5790
      Top             =   1770
      Width           =   1500
   End
   Begin VB.Shape textBorder 
      BackColor       =   &H00312118&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H006C4D42&
      Height          =   900
      Index           =   13
      Left            =   5790
      Top             =   465
      Width           =   1500
   End
   Begin VB.Label title 
      Alignment       =   2  'Center
      BackColor       =   &H008C6456&
      Caption         =   "Timeout"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   5
      Left            =   5805
      TabIndex        =   29
      Top             =   1500
      Width           =   1470
   End
   Begin VB.Label title 
      Alignment       =   2  'Center
      BackColor       =   &H008C6456&
      Caption         =   "Threads"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   4
      Left            =   5805
      TabIndex        =   28
      Top             =   2865
      Width           =   1470
   End
   Begin VB.Label title 
      Alignment       =   2  'Center
      BackColor       =   &H008C6456&
      Caption         =   "Connection"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   3
      Left            =   5805
      TabIndex        =   27
      Top             =   165
      Width           =   1470
   End
   Begin VB.Shape titleBar 
      BackColor       =   &H008C6456&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H006C4D42&
      Height          =   360
      Index           =   3
      Left            =   5790
      Top             =   120
      Width           =   1500
   End
   Begin VB.Label title 
      Alignment       =   2  'Center
      BackColor       =   &H008C6456&
      Caption         =   "Open Ports"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Index           =   1
      Left            =   330
      TabIndex        =   26
      Top             =   1920
      Width           =   5070
   End
   Begin VB.Shape textBorder 
      BorderColor     =   &H006C4D42&
      Height          =   1320
      Index           =   12
      Left            =   195
      Top             =   2220
      Width           =   5265
   End
   Begin VB.Shape progressBar 
      BackColor       =   &H006C4D42&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H006C4D42&
      FillColor       =   &H00FFFFFF&
      Height          =   180
      Left            =   3450
      Top             =   4305
      Width           =   15
   End
   Begin VB.Label currentPortLabel 
      BackColor       =   &H004F3830&
      BackStyle       =   0  'Transparent
      Caption         =   "Not scanning"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   105
      TabIndex        =   24
      Top             =   4275
      Width           =   2850
   End
   Begin VB.Shape progressBarBackground 
      BackColor       =   &H008C6456&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H006C4D42&
      FillColor       =   &H00312118&
      Height          =   180
      Left            =   3450
      Top             =   4305
      Width           =   2100
   End
   Begin VB.Label commandButton 
      Alignment       =   2  'Center
      BackColor       =   &H004F3830&
      BackStyle       =   0  'Transparent
      Caption         =   "Options"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   2670
      TabIndex        =   23
      Top             =   3825
      Width           =   1125
   End
   Begin VB.Image closeButtonOver 
      Height          =   300
      Left            =   5145
      Picture         =   "mainForm.frx":B018E
      Top             =   90
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Line Line3 
      BorderColor     =   &H004F3830&
      X1              =   5355.906
      X2              =   3785.04
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Line Line2 
      BorderColor     =   &H004F3830&
      X1              =   254.331
      X2              =   3560.63
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Label commandButton 
      Alignment       =   2  'Center
      BackColor       =   &H004F3830&
      BackStyle       =   0  'Transparent
      Caption         =   "Scan"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   4335
      TabIndex        =   22
      Top             =   3840
      Width           =   1125
   End
   Begin VB.Shape commandButtonBack 
      BackColor       =   &H008C6456&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H006C4D42&
      Height          =   300
      Index           =   0
      Left            =   4320
      Top             =   3810
      Width           =   1125
   End
   Begin VB.Image closeButton 
      Height          =   300
      Left            =   5145
      Picture         =   "mainForm.frx":B08B0
      Top             =   90
      Width           =   435
   End
   Begin VB.Image mainTitleBar 
      Height          =   375
      Left            =   15
      Picture         =   "mainForm.frx":B0FD2
      Top             =   15
      Width           =   5655
   End
   Begin VB.Label title 
      Alignment       =   2  'Center
      BackColor       =   &H008C6456&
      Caption         =   "Port Range"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   0
      Left            =   3825
      TabIndex        =   21
      Top             =   540
      Width           =   1575
   End
   Begin VB.Label title 
      Alignment       =   2  'Center
      BackColor       =   &H008C6456&
      Caption         =   "IP Range"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   2
      Left            =   255
      TabIndex        =   20
      Top             =   525
      Width           =   3375
   End
   Begin VB.Label end 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "End:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   3900
      TabIndex        =   19
      Top             =   1470
      Width           =   510
   End
   Begin VB.Label start 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Start:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   3900
      TabIndex        =   18
      Top             =   1020
      Width           =   510
   End
   Begin VB.Shape textBorder 
      BorderColor     =   &H006C4D42&
      Height          =   255
      Index           =   9
      Left            =   4425
      Top             =   1440
      Width           =   870
   End
   Begin VB.Shape textBorder 
      BorderColor     =   &H006C4D42&
      Height          =   255
      Index           =   0
      Left            =   4425
      Top             =   960
      Width           =   870
   End
   Begin VB.Label end 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "End:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   315
      TabIndex        =   15
      Top             =   1470
      Width           =   510
   End
   Begin VB.Label start 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Start:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   315
      TabIndex        =   14
      Top             =   1005
      Width           =   510
   End
   Begin VB.Shape textBorder 
      BorderColor     =   &H006C4D42&
      Height          =   255
      Index           =   8
      Left            =   3015
      Top             =   1440
      Width           =   510
   End
   Begin VB.Shape textBorder 
      BorderColor     =   &H006C4D42&
      Height          =   255
      Index           =   7
      Left            =   3015
      Top             =   960
      Width           =   510
   End
   Begin VB.Shape textBorder 
      BorderColor     =   &H006C4D42&
      Height          =   255
      Index           =   6
      Left            =   2310
      Top             =   960
      Width           =   510
   End
   Begin VB.Shape textBorder 
      BorderColor     =   &H006C4D42&
      Height          =   255
      Index           =   5
      Left            =   1575
      Top             =   960
      Width           =   510
   End
   Begin VB.Shape textBorder 
      BorderColor     =   &H006C4D42&
      Height          =   255
      Index           =   4
      Left            =   855
      Top             =   960
      Width           =   510
   End
   Begin VB.Shape textBorder 
      BorderColor     =   &H006C4D42&
      Height          =   255
      Index           =   3
      Left            =   2295
      Top             =   1440
      Width           =   510
   End
   Begin VB.Shape textBorder 
      BorderColor     =   &H006C4D42&
      Height          =   255
      Index           =   2
      Left            =   1575
      Top             =   1440
      Width           =   510
   End
   Begin VB.Shape textBorder 
      BorderColor     =   &H006C4D42&
      Height          =   255
      Index           =   1
      Left            =   855
      Top             =   1440
      Width           =   510
   End
   Begin VB.Label dot 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   8
      Left            =   2820
      TabIndex        =   13
      Top             =   1410
      Width           =   255
   End
   Begin VB.Label dot 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   7
      Left            =   2100
      TabIndex        =   12
      Top             =   1410
      Width           =   255
   End
   Begin VB.Label dot 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   6
      Left            =   1380
      TabIndex        =   11
      Top             =   1410
      Width           =   255
   End
   Begin VB.Label dot 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   2820
      TabIndex        =   10
      Top             =   930
      Width           =   255
   End
   Begin VB.Label dot 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   2100
      TabIndex        =   9
      Top             =   930
      Width           =   255
   End
   Begin VB.Label dot 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   1380
      TabIndex        =   8
      Top             =   930
      Width           =   255
   End
   Begin VB.Shape titleBar 
      BackColor       =   &H008C6456&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H006C4D42&
      Height          =   360
      Index           =   0
      Left            =   3735
      Top             =   480
      Width           =   1740
   End
   Begin VB.Shape titleBar 
      BackColor       =   &H008C6456&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H006C4D42&
      Height          =   360
      Index           =   1
      Left            =   195
      Top             =   480
      Width           =   3465
   End
   Begin VB.Shape commandButtonBack 
      BackColor       =   &H008C6456&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H006C4D42&
      Height          =   300
      Index           =   1
      Left            =   2670
      Top             =   3795
      Width           =   1125
   End
   Begin VB.Shape titleBar 
      BackColor       =   &H008C6456&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H006C4D42&
      Height          =   360
      Index           =   2
      Left            =   195
      Top             =   1875
      Width           =   5265
   End
   Begin VB.Shape textBorder 
      BackColor       =   &H00312118&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H006C4D42&
      Height          =   990
      Index           =   11
      Left            =   3735
      Top             =   825
      Width           =   1725
   End
   Begin VB.Shape textBorder 
      BackColor       =   &H00312118&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H006C4D42&
      Height          =   990
      Index           =   10
      Left            =   195
      Top             =   825
      Width           =   3465
   End
   Begin VB.Shape titleBar 
      BackColor       =   &H008C6456&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H006C4D42&
      Height          =   360
      Index           =   4
      Left            =   5790
      Top             =   2835
      Width           =   1500
   End
   Begin VB.Shape titleBar 
      BackColor       =   &H008C6456&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H006C4D42&
      Height          =   360
      Index           =   5
      Left            =   5790
      Top             =   1455
      Width           =   1500
   End
   Begin VB.Shape commandButtonBack 
      BackColor       =   &H008C6456&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H006C4D42&
      Height          =   300
      Index           =   2
      Left            =   5805
      Top             =   4185
      Width           =   1500
   End
   Begin VB.Shape commandButtonBack 
      BackColor       =   &H008C6456&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H006C4D42&
      Height          =   300
      Index           =   4
      Left            =   195
      Top             =   3795
      Width           =   1125
   End
   Begin VB.Shape commandButtonBack 
      BackColor       =   &H008C6456&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H006C4D42&
      Height          =   300
      Index           =   3
      Left            =   1440
      Top             =   3795
      Width           =   1125
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   3855
      Left            =   120
      Top             =   360
      Width           =   5445
   End
End
Attribute VB_Name = "mainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim originalStartPort, i, ii, threadAmount, currentPort, connectionTimeOut, previousThread, IPArrayValue, portListIndex(999999) As Integer
Dim currentIP, portIndex(65535), IPIndex(999999) As String
Dim waitForConnect, ranOnce As Boolean
Dim mb As VbMsgBoxResult

'   windows api's for the window movement
Dim lFormTopMouseDown, lFormLeftMouseDown, jumpGap As Long
Dim CursorLoc As POINTAPI
Dim lpwndpl As WINDOWPLACEMENT
Const SW_SHOWNORMAL = 1
Private Type POINTAPI
    X As Long
    Y As Long
End Type
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Type WINDOWPLACEMENT
    Length As Long
    flags As Long
    ShowCmd As Long
    ptMinPosition As POINTAPI
    ptMaxPosition As POINTAPI
    rcNormalPosition As RECT
End Type
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function SetWindowPlacement Lib "user32" (ByVal hwnd As Long, lpwndpl As WINDOWPLACEMENT) As Long
Const HTCAPTION = 2
Const WM_NCLBUTTONDOWN = &HA1
Const HTLEFT = 10
Const HTRight = 11
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

'   clear results
Private Sub clear_Click()
    IPArrayValue = 0
    results.clear
End Sub

'   close button animations
Private Sub closeButton_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    closeButtonOver.Visible = True
End Sub

Private Sub closeButton_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    closeButtonOver.Visible = False
    If X < 405 And Y < 405 And X > 0 And Y > 0 Then End
End Sub

Private Sub commandButton_Click(Index As Integer)
        
    '   options button
    If Index = 1 And mainForm.Width = 7410 Then
        mainForm.Width = 5685
        mainForm.Height = 4575
    ElseIf Index = 1 Then
        mainForm.Width = 7410
        mainForm.Height = 4575
    End If
    
    '   exit button
    If Index = 4 Then
        End
    End If
    
    '   about button
    If Index = 3 And mainForm.Height = 4575 Then
        mainForm.Width = 5685
        mainForm.Height = 5235
    ElseIf Index = 3 Then
        mainForm.Width = 5685
        mainForm.Height = 4575
    End If
    
    
    '   scan button
    If Index = 0 Then
        If commandButton(0).Caption = "Scan" Then
            startHunting
        Else
            PauseHunting
        End If
    End If
    
    '   apply button
    If Index = 2 Then
        PauseHunting
        For i = 1 To threadAmount
            Unload wSock(i)
        Next i
        threadAmount = Val(threads.Text)
        previousThread = threadAmount
        For i = 1 To threadAmount
            Load wSock(i)
        Next i
        connectionTimeOut = Val(timeout.Text)
        If verifyCon.Value = 1 Then waitForConnect = True Else waitForConnect = False
        mainForm.Width = 5700
        mainForm.Height = 4575
    End If
    
End Sub

'   button animations
Private Sub commandButton_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    commandButtonBack(Index).BackColor = &H6C4D42
End Sub

Private Sub commandButton_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    commandButtonBack(Index).BackColor = &H8C6456
End Sub

'   open internet url
Private Sub description_Click(Index As Integer)
    If Index = 6 Then
        ShellExecute hwnd, "open", "http://www.hackboxed.d2g.com/", vbNullString, vbNullString, 0
    End If
End Sub

'   internet url animations
Private Sub description_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Index = 6 Then description(6).ForeColor = &H6C4D42
End Sub

Private Sub description_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Index = 6 Then description(6).ForeColor = &HFFFFFF
End Sub

'   form loading
Private Sub Form_Load()
    
    '   port index
    
    portIndex(21) = "FTP Service, Dolly Trojan"
    portIndex(23) = "TELNET Service"
    portIndex(25) = "SMTP, AntiGen"
    portIndex(31) = "Agent 31, Hacker's Paradise"
    portIndex(41) = "Deep Throat"
    portIndex(53) = "DNS"
    portIndex(58) = "DM Setup"
    portIndex(79) = "Firehotcker"
    portIndex(80) = "HTTP Service, Executor"
    portIndex(90) = "Hidden Port 2.o"
    portIndex(110) = "ProMail Trojan"
    portIndex(113) = "Kazimas"
    portIndex(119) = "Happy99"
    portIndex(121) = "Jammer Killah"
    portIndex(129) = "Password Generator Protocol"
    portIndex(135) = "Netbios Remote Procedure Call"
    portIndex(137) = "Netbios Name (DoS attacks)"
    portIndex(138) = "Netbios Datagram"
    portIndex(139) = "Netbios Session (DoS attacks)"
    portIndex(146) = "Infector 1.3"
    portIndex(421) = "TCP Wrappers"
    portIndex(456) = "Hacker's Paradise"
    portIndex(531) = "Rasmin"
    portIndex(555) = "Stealth Spy, Phaze, 7-11 Trojan"
    portIndex(666) = "Attack FTP"
    portIndex(777) = "AIM Spy Application"
    portIndex(901) = "Backdoor.Devil"
    portIndex(902) = "Backdoor.Devil"
    portIndex(911) = "Dark Shadow"
    portIndex(999) = "DeepThroat"
    portIndex(9400) = "InCommand"
    portIndex(9999) = "The prayer 1.2 -1.3"
    portIndex(1000) = "Der Spaeher"
    portIndex(1001) = "Silencer, WebEx"
    portIndex(1011) = "Doly Trojan"
    portIndex(1012) = "Doly Trojan"
    portIndex(1015) = "Doly Trojan"
    portIndex(1024) = "NetSpy"
    portIndex(1027) = "ICQ"
    portIndex(1029) = "ICQ"
    portIndex(1032) = "ICQ"
    portIndex(1033) = "NetSpy"
    portIndex(1034) = "Backdoor.Systec"
    portIndex(1042) = "Bla"
    portIndex(1045) = "Rasmin"
    portIndex(1090) = "Xtreme"
    portIndex(1170) = "Voice Streaming Audio"
    portIndex(1207) = "SoftWar"
    portIndex(1214) = "KaZaa File Sharing"
    portIndex(1234) = "Ultors Trojan"
    portIndex(1243) = "Sub Seven"
    portIndex(1245) = "VooDoo Doll"
    portIndex(1269) = "Maverick's Matrix"
    portIndex(12631) = "WhackJob"
    portIndex(1394) = "GoFriller, Backdoor G-1"
    portIndex(1492) = "FTP99CMP"
    portIndex(1505) = "FunkProxy "
    portIndex(1509) = "Psyber Streaming server"
    portIndex(1600) = "Shivka-Burka"
    portIndex(1604) = "ICA Browser"
    portIndex(1722) = "Backdoor.NetControle"
    portIndex(1807) = "SpySender"
    portIndex(1981) = "Shockrave"
    portIndex(1999) = "BackDoor"
    portIndex(2000) = "Remote Explorer"
    portIndex(2001) = "Trojan Cow"
    portIndex(2002) = "TransScout"
    portIndex(2003) = "TransScout"
    portIndex(2004) = "TransScout"
    portIndex(2005) = "TransScout"
    portIndex(2023) = "Ripper"
    portIndex(2115) = "Bugs"
    portIndex(2140) = "Deep Throat"
    portIndex(2155) = "Illusion Mailer"
    portIndex(2283) = "HLV Rat5"
    portIndex(2565) = "Striker"
    portIndex(2583) = "WinCrash"
    portIndex(2716) = "The Prayer 1.2 -1.3"
    portIndex(2721) = "Phase Zero"
    portIndex(2801) = "Phineas Phucker"
    portIndex(3017) = "MSN Messenger"
    portIndex(3024) = "WinCrash"
    portIndex(3028) = "Ring Zero"
    portIndex(3129) = "Master's Paradise"
    portIndex(3150) = "Deep Throat"
    portIndex(3332) = "Q0 BackDoor"
    portIndex(3459) = "Eclipse 2000"
    portIndex(3700) = "Portal of Doom"
    portIndex(3791) = "Eclypse"
    portIndex(4100) = "Watchguard Firebox  admin DoS Expl"
    portIndex(4092) = "WinCrash"
    portIndex(4567) = "File Nail"
    portIndex(4590) = "ICQ Trojan"
    portIndex(5000) = "Sokets de Trois v1./Bubbel"
    portIndex(5001) = "Sokets de Trois v1./Bubbel"
    portIndex(5011) = "Ootlt"
    portIndex(5031) = "Net Metropolitan 1.0"
    portIndex(5032) = "Net Metropolitan 1.04"
    portIndex(5321) = "Firehotcker"
    portIndex(5400) = "Blade Runner"
    portIndex(5401) = "Blade Runner"
    portIndex(5402) = "Blade Runner"
    portIndex(5521) = "Illusion Mailer"
    portIndex(5550) = "Xtcp"
    portIndex(5512) = "Xtcp"
    portIndex(5555) = "ServeMe"
    portIndex(5556) = "BO Facil"
    portIndex(5557) = "BO Facil"
    portIndex(5569) = "Robo-Hack"
    portIndex(5637) = "PC Crasher"
    portIndex(5638) = "PC Crasher"
    portIndex(5714) = "WinCrash"
    portIndex(5741) = "WinCrash"
    portIndex(5742) = "WinCrash"
    portIndex(6000) = "The Thing 1.6"
    portIndex(6112) = "Battle.net Game"
    portIndex(6346) = "Gnutella clone"
    portIndex(6400) = "The Thing"
    portIndex(6667) = "Sub-7 Trojan"
    portIndex(6669) = "Vampyre"
    portIndex(6670) = "Deep Throat"
    portIndex(6671) = "Deep Throat"
    portIndex(6711) = "Sub Seven, Backdoor.G"
    portIndex(6712) = "Sub Seven"
    portIndex(6713) = "Sub Seven"
    portIndex(6723) = "Mstream attack-handler"
    portIndex(6771) = "Deep Throat"
    portIndex(6776) = "Sub Seven, Backdoor.G"
    portIndex(6912) = "Sh*t Heap "
    portIndex(6939) = "Indoctrination"
    portIndex(6969) = "Gate Crasher, Priority"
    portIndex(6970) = "Gate Crasher"
    portIndex(7000) = "Remote Grab"
    portIndex(7028) = "Unknown Trojan"
    portIndex(7300) = "Net Monitor"
    portIndex(7301) = "Net Monitor"
    portIndex(7306) = "Net Monitor"
    portIndex(7307) = "Net Monitor"
    portIndex(7308) = "Net Monitor"
    portIndex(7597) = "QaZ"
    portIndex(7614) = "Backdoor.GRM"
    portIndex(7789) = "ICKiller"
    portIndex(8012) = "BackDoor-KL"
    portIndex(8080) = "Ring Zero"
    portIndex(8787) = "BackOrifice 2000"
    portIndex(8879) = "BackOrifice 2000"
    portIndex(9872) = "Portal of Doom"
    portIndex(9873) = "Portal of Doom"
    portIndex(9874) = "Portal of Doom"
    portIndex(9875) = "Portal of Doom"
    portIndex(9876) = "Cyber Attacker"
    portIndex(9878) = "Trans Scout"
    portIndex(9989) = "iNi-Killer"
    portIndex(10008) = "Cheese worm"
    portIndex(10067) = "Portal of Doom"
    portIndex(10167) = "Portal of Doom"
    portIndex(10520) = "Acid Shivers"
    portIndex(10607) = "Coma"
    portIndex(10666) = "Ambush"
    portIndex(11000) = "Senna Spy"
    portIndex(11050) = "Host Control"
    portIndex(11223) = "Progenic Trojan"
    portIndex(11831) = "Latinus Server"
    portIndex(12076) = "GJamer"
    portIndex(12223) = "Hack'99, KeyLogger"
    portIndex(12345) = "Netbus, Ultor's Trojan"
    portIndex(12346) = "Netbus"
    portIndex(12456) = "NetBus"
    portIndex(12361) = "Whack-a-Mole"
    portIndex(12362) = "Whack-a-Mole"
    portIndex(12631) = "Whack Job"
    portIndex(12701) = "Eclypse 2000"
    portIndex(12754) = "Mstream attack-handler"
    portIndex(13000) = "Senna Spy"
    portIndex(13700) = "Kuang2 the Virus"
    portIndex(15104) = "Mstream attack-handler"
    portIndex(16484) = "Mosucker"
    portIndex(16959) = "SubSeven DEFCON8 2.1 Backdoor"
    portIndex(16969) = "Priority"
    portIndex(17300) = "Kuang2 The Virus"
    portIndex(20000) = "Millennium"
    portIndex(20001) = "Millennium"
    portIndex(20034) = "NetBus 2 Pro"
    portIndex(20203) = "Logged!"
    portIndex(20331) = "Bla Trojan"
    portIndex(20432) = "Shaft Client to handlers"
    portIndex(20433) = "Shaft Agent to handlers"
    portIndex(21554) = "GirlFriend"
    portIndex(22222) = "Prosiak"
    portIndex(22784) = "Backdoor-ADM"
    portIndex(23476) = "Donald Dick"
    portIndex(23477) = "Donald Dick"
    portIndex(26274) = "Delta Source"
    portIndex(27573) = "Sub-7 2.1 "
    portIndex(27665) = "Trin00 DoS Attack"
    portIndex(29559) = "Latinus Server"
    portIndex(29891) = "The Unexplained"
    portIndex(30029) = "AOL Trojan"
    portIndex(30100) = "NetSphere"
    portIndex(30101) = "NetSphere"
    portIndex(30102) = "NetSphere"
    portIndex(30133) = "NetSphere Final"
    portIndex(30303) = "Sockets de Troie"
    portIndex(30999) = "Kuang2 "
    portIndex(31336) = "BO-Whack"
    portIndex(31337) = "Netpatch"
    portIndex(31338) = "NetSpy DK"
    portIndex(31339) = "NetSpy DK"
    portIndex(31666) = "BOWhack"
    portIndex(31785) = "Hack'a'Tack"
    portIndex(32418) = "Acid Battery"
    portIndex(33270) = "Trinity Trojan "
    portIndex(33333) = "Prosiak"
    portIndex(33911) = "Spirit 2001 a"
    portIndex(34324) = "BigGluck, TN"
    portIndex(37651) = "Yet Another Trojan"
    portIndex(40421) = "Master's Paradise"
    portIndex(40412) = "The Spy"
    portIndex(40421) = "Agent, Master's of Paradise"
    portIndex(40422) = "Master's Paradise"
    portIndex(40423) = "Master's Paradise"
    portIndex(40425) = "Master's Paradise"
    portIndex(40426) = "Master's Paradise"
    portIndex(43210) = "Master's Paradise"
    portIndex(47252) = "Delta Source"
    portIndex(50505) = "Sokets de Trois v2"
    portIndex(50776) = "Fore"
    portIndex(53001) = "Remote Windows Shutdown"
    portIndex(54320) = "Back Orifice 2000"
    portIndex(54321) = "School Bus, Back Orifice"
    portIndex(57341) = "NetRaider Trojan"
    portIndex(58008) = "BackDoor.Tron"
    portIndex(58009) = "BackDoor.Tron"
    portIndex(60000) = "Deep Throat"
    portIndex(61466) = "Telecommando"
    portIndex(61348) = "Bunker-Hill Trojan"
    portIndex(61603) = "Bunker-Hill Trojan"
    portIndex(63485) = "Bunker-Hill Trojan"
    portIndex(65000) = "Stacheldraht,  Devil"
    portIndex(65535) = "Adore Worm/Linux"

    
    threadAmount = 200
    previousThread = threadAmount
    For i = 1 To threadAmount
        Load wSock(i)
    Next i
    
    waitForConnect = False
    connectionTimeOut = 5
    IPArrayValue = 0
    
End Sub

'   select all in textboxes
Private Sub ipStart_GotFocus(Index As Integer)
    ipStart(Index).SelStart = 0
    ipStart(Index).SelLength = 3
End Sub

Private Sub ipEnd_GotFocus(Index As Integer)
    ipEnd(Index).SelStart = 0
    ipEnd(Index).SelLength = 3
End Sub

Private Sub results_DblClick()
    
    '   port uses
    Select Case portListIndex(results.ListIndex)
    Case 80
        ShellExecute hwnd, "open", "http://" & IPIndex(results.ListIndex) & "/", vbNullString, vbNullString, 0
    End Select
    
End Sub

Private Sub threads_GotFocus()
    threads.SelStart = 0
    threads.SelLength = 3
End Sub

Private Sub timeout_GotFocus()
    timeout.SelStart = 0
    timeout.SelLength = 5
End Sub

'   filter keys
Private Sub timeout_KeyPress(KeyAscii As Integer)
    If KeyAscii < 48 And KeyAscii <> 8 Or KeyAscii > 57 Then
        KeyAscii = 0
    End If
End Sub

Private Sub threads_KeyPress(KeyAscii As Integer)
    If KeyAscii < 48 And KeyAscii <> 8 Or KeyAscii > 57 Then
        KeyAscii = 0
    End If
End Sub

Private Sub ipStart_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii < 48 And KeyAscii <> 8 Or KeyAscii > 57 Then
        If KeyAscii = 46 And Index < 4 Then
            ipStart(Index + 1).SetFocus
        ElseIf KeyAscii = 46 Then
            ipEnd(1).SetFocus
        End If
        KeyAscii = 0
    End If
End Sub

Private Sub ipEnd_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii < 48 And KeyAscii <> 8 Or KeyAscii > 57 Then
        If KeyAscii = 46 And Index < 4 Then
            ipEnd(Index + 1).SetFocus
        ElseIf KeyAscii = 46 Then
            portStart.SetFocus
        End If
        KeyAscii = 0
    End If
End Sub

'   move window
Private Sub mainTitleBar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ReleaseCapture
    Call SendMessage(hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
End Sub

'   more selections
Private Sub portEnd_GotFocus()
    portEnd.SelStart = 0
    portEnd.SelLength = 5
End Sub

Private Sub portStart_GotFocus()
    portStart.SelStart = 0
    portStart.SelLength = 5
End Sub

'   more key filtering
Private Sub portStart_KeyPress(KeyAscii As Integer)
    If KeyAscii < 48 And KeyAscii <> 8 Or KeyAscii > 57 Then
        If KeyAscii = 46 Then
            portEnd.SetFocus
        End If
        KeyAscii = 0
    End If
End Sub

Private Sub portEnd_KeyPress(KeyAscii As Integer)
    If KeyAscii < 48 And KeyAscii <> 8 Or KeyAscii > 57 Then
        KeyAscii = 0
    End If
End Sub

'   port scanning
Private Sub startHunting()
    
    ranOnce = False
    originalStartPort = Val(portStart.Text)
    
    ipStart(1).Locked = True
    ipStart(2).Locked = True
    ipStart(3).Locked = True
    ipStart(4).Locked = True
    
    ipEnd(1).Locked = True
    ipEnd(2).Locked = True
    ipEnd(3).Locked = True
    ipEnd(4).Locked = True
    
    portStart.Locked = True
    portEnd.Locked = True
    
    commandButton(0).Caption = "Pause"
    
    nextIP
        
End Sub

Private Sub PauseHunting()
    
    ipStart(1).Locked = False
    ipStart(2).Locked = False
    ipStart(3).Locked = False
    ipStart(4).Locked = False
    
    ipEnd(1).Locked = False
    ipEnd(2).Locked = False
    ipEnd(3).Locked = False
    ipEnd(4).Locked = False
    
    portStart.Locked = False
    portEnd.Locked = False
    
    commandButton(0).Caption = "Scan"
    
    currentPortLabel.Caption = "Scanning Halted"
    
    For i = 1 To threadAmount
        wSock(i).Close
    Next i
    
    progressBar.Width = 0
    
End Sub

Private Sub scanPorts()

    If (Val(portEnd.Text) - originalStartPort) < threadAmount Then
        For i = 1 To threadAmount
            Unload wSock(i)
        Next i
        threadAmount = Val(portEnd.Text) - originalStartPort
        If threadAmount = 0 Then threadAmount = 1
        threads.Text = threadAmount
        For i = 1 To threadAmount
            Load wSock(i)
        Next i
    End If
    
    If waitForConnect = True Then
        For i = 1 To threadAmount
            If currentPort < Val(portEnd.Text) + 1 Then
                wSock(i).Connect currentIP, currentPort
                If currentPort = Val(portEnd.Text) Then progressBar.Width = 2100 Else progressBar.Width = ((2100 / (Val(portEnd.Text) - originalStartPort)) * (currentPort - originalStartPort))
                currentPortLabel.Caption = "Scanning " & currentIP & " : " & currentPort
                portStart.Text = currentPort
                currentPort = currentPort + 1
            Else
                endPorts
            End If
        Next i
    Else
        For ii = originalStartPort To Val(portEnd.Text)
            For i = 1 To threadAmount
                If commandButton(0).Caption = "Pause" Then
                    If currentPort < Val(portEnd.Text) + 1 Then
                        wSock(i).Close
                        wSock(i).Connect currentIP, currentPort
                        If currentPort = Val(portEnd.Text) Then progressBar.Width = 2100 Else progressBar.Width = ((2100 / (Val(portEnd.Text) - originalStartPort)) * (currentPort - originalStartPort))
                        currentPortLabel.Caption = "Scanning " & currentIP & " : " & currentPort
                        portStart.Text = currentPort
                        currentPort = currentPort + 1
                    Else
                        endPorts
                    End If
                End If
            Next i
        pause ((connectionTimeOut / 10))
        Next ii
    End If
    
End Sub

Public Sub pause(Duration As Single)
    Dim Current As Single
    Current = Timer
    Do Until Timer - Current >= Duration
        DoEvents
    Loop
End Sub

Private Sub nextIP()
    
        If ipStart(1).Text = ipEnd(1).Text Then
        If ipStart(2).Text = ipEnd(2).Text Then
            If ipStart(3).Text = ipEnd(3).Text Then
                If ipStart(4).Text = ipEnd(4).Text Then
                    If ranOnce = False Then ranOnce = True Else PauseHunting
                Else
                    ipStart(4).Text = ipStart(4).Text + 1
                End If
            Else
                If ipStart(4).Text = 255 Then
                    ipStart(4).Text = 0
                    ipStart(3).Text = ipStart(3).Text + 1
                Else
                    ipStart(4).Text = ipStart(4).Text + 1
                End If
            End If
        Else
            If ipStart(4).Text = 255 Then
                If ipStart(3).Text = 255 Then
                    ipStart(4).Text = 0
                    ipStart(3).Text = 0
                    ipStart(2).Text = ipStart(2).Text + 1
                Else
                    ipStart(4).Text = 0
                    ipStart(3).Text = ipStart(3).Text + 1
                End If
            Else
                ipStart(4).Text = ipStart(4).Text + 1
            End If
        End If
    Else
        If ipStart(4).Text = 255 Then
            If ipStart(3).Text = 255 Then
                If ipStart(2).Text = 255 Then
                    ipStart(4).Text = 0
                    ipStart(3).Text = 0
                    ipStart(2).Text = 0
                    ipStart(1).Text = ipStart(1).Text + 1
                Else
                    ipStart(4).Text = 0
                    ipStart(3).Text = 0
                    ipStart(2).Text = ipStart(2).Text + 1
                End If
            Else
                ipStart(4).Text = 0
                ipStart(3).Text = ipStart(3).Text + 1
            End If
        Else
            ipStart(4).Text = ipStart(4).Text + 1
        End If
    End If
    currentIP = ipStart(1).Text & "." & ipStart(2).Text & "." & ipStart(3).Text & "." & ipStart(4).Text
    currentPort = originalStartPort
    scanPorts
    
End Sub


Private Sub wSock_Connect(Index As Integer)

    If portIndex(wSock(Index).RemotePort) = "" Then results.AddItem (wSock(Index).RemoteHost & ": " & wSock(Index).RemotePort) Else results.AddItem (wSock(Index).RemoteHost & ": " & wSock(Index).RemotePort & "   [ " & portIndex(wSock(Index).RemotePort) & " ]")
    IPIndex(IPArrayValue) = wSock(Index).RemoteHost
    portListIndex(IPArrayValue) = wSock(Index).RemotePort
    IPArrayValue = IPArrayValue + 1
    If waitForConnect = False Then
        If commandButton(0).Caption = "Pause" Then
            If currentPort < Val(portEnd.Text) + 1 Then
                wSock(Index).Close
                wSock(Index).Connect currentIP, currentPort
                currentPort = currentPort + 1
                progressBar.Width = ((2100 / (Val(portEnd.Text) - originalStartPort)) * (currentPort - originalStartPort))
                portStart.Text = currentPort
            Else
                endPorts
            End If
        End If
    End If
End Sub

Private Sub endPorts()
    portStart.Text = originalStartPort
    For i = 1 To threadAmount
        Unload wSock(i)
    Next i
    threadAmount = previousThread
    threads.Text = threadAmount
    For i = 1 To threadAmount
        Load wSock(i)
    Next i
    nextIP
End Sub

Private Sub wSock_Error(Index As Integer, ByVal Number As Integer, description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    If waitForConnect = False Then
        If commandButton(0).Caption = "Pause" Then
            If currentPort < Val(portEnd.Text) + 1 Then
                wSock(Index).Close
                wSock(Index).Connect currentIP, currentPort
                currentPort = currentPort + 1
                progressBar.Width = ((2100 / (Val(portEnd.Text) - originalStartPort)) * (currentPort - originalStartPort))
                currentPortLabel.Caption = "Scanning " & currentIP & " : " & currentPort
                portStart.Text = currentPort
            Else
                endPorts
            End If
        End If
    End If
End Sub



