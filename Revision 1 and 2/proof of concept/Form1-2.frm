VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Part II"
   ClientHeight    =   435
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5160
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   435
   ScaleWidth      =   5160
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   611
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private FileName As String
Private Status As Integer

Private Sub Form_Load()
Winsock1.Connect "127.0.0.1"
End Sub

Private Sub Winsock1_Close()
Unload Me
End
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Winsock1.GetData dat$
If Left(dat, 4) = "FLNM" Then 'Incomming FileName
    FileName = Mid(dat, 5, Len(dat) - 4)
ElseIf Left(dat, 4) = "OPEN" Then 'Create new file
    Close #1
    Open FileName For Binary Access Write As #1
ElseIf Left(dat, 4) = "CLSE" Then 'Close File and PrePare for New one
    Close #1
Else 'Data to be put to file
    Put #1, , dat
End If
End Sub
