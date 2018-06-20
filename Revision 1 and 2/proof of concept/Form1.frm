VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   Caption         =   "UnHexaLock Part I"
   ClientHeight    =   3465
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8775
   LinkTopic       =   "Form1"
   ScaleHeight     =   3465
   ScaleWidth      =   8775
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Choose Folder"
      Height          =   375
      Left            =   600
      TabIndex        =   7
      Top             =   960
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6960
      Top             =   2520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   6480
      Top             =   2520
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   611
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Load File"
      Height          =   375
      Left            =   600
      TabIndex        =   2
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label lblStatus 
      Caption         =   "Idle"
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
      Left            =   960
      TabIndex        =   10
      Top             =   3000
      Width           =   1695
   End
   Begin VB.Label Label8 
      Caption         =   "Status:"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   3000
      Width           =   615
   End
   Begin VB.Label Label7 
      Caption         =   "Note: As soon as file is loaded, it will be sent off to Part II to be written to disk"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   2640
      Width           =   5535
   End
   Begin VB.Label Label6 
      Caption         =   "Label6"
      Height          =   255
      Left            =   1920
      TabIndex        =   6
      Top             =   600
      Width           =   6735
   End
   Begin VB.Label Label5 
      Caption         =   "Output Directory:"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label Label4 
      Caption         =   "0KB"
      Height          =   255
      Left            =   1920
      TabIndex        =   4
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Size:"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Not Loaded"
      Height          =   255
      Left            =   1920
      TabIndex        =   1
      Top             =   1440
      Width           =   6735
   End
   Begin VB.Label Label1 
      Caption         =   "File:"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   1440
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'A big thanks goes out to this website so I didn't have to create my own code.  Thanks for the awsome
'file splitting application guys.
'http://www.ostrosoft.com/vb/projects/split.asp

Public TheFile As String
Public TheOutput As String
Public WinsockStatus As Boolean
Public SendComplete As Boolean

Private Sub Command1_Click()
CommonDialog1.ShowOpen
Label2.Caption = CommonDialog1.FileName

Winsock1.SendData "FLNM" & TheOutput & "\" & CommonDialog1.FileTitle

    DoEvents

Winsock1.SendData "OPEN"

    DoEvents


    Open CommonDialog1.FileName For Binary As #1
        Label4.Caption = Int(Val(LOF(1) / 1024)) & "KB"
        Dim b() As Byte
        Dim ss As Double 'split size
        ss = 10240 ' this is 10K, a nominal size to send via winsock.
        nLen = FileLen(CommonDialog1.FileName)
        lblStatus.ForeColor = vbRed
        While nLen > ss
            ReDim b(ss - 1)
            lblStatus.Caption = "Reading Data"
            Get #1, ss * i + 1, b()
            SendComplete = False
            lblStatus.Caption = "Sending Data"
            Winsock1.SendData b()
            lblStatus.Caption = "Busy"
            While SendComplete = False
                DoEvents ' wait for this to be sent
            Wend
            i = i + 1
            nLen = nLen - ss

        Wend
        ReDim b(nLen - 1)
        Get #1, ss * i + 1, b()
        Winsock1.SendData b()
    Close #1
    lblStatus.Caption = "Done"
    lblStatus.ForeColor = vbGreen
    DoEvents
Winsock1.SendData "CLSE"
End Sub

Private Sub Command2_Click()
TheOutput = BrowseForFolder(Me.hWnd, "Choose Output Directory...")
Label6.Caption = TheOutput
End Sub

Private Sub Form_Load()
Winsock1.Listen
Shell "part2.exe /part2"
End Sub

Private Sub Form_Unload(Cancel As Integer)
Winsock1.Close
End Sub

Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
Winsock1.Close
DoEvents
Winsock1.Accept requestID
End Sub

Private Sub Winsock1_SendComplete()
SendComplete = True
End Sub
