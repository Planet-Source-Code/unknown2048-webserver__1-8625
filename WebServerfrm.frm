VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form WebServerfrm 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Web Server V 1.0"
   ClientHeight    =   2145
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4080
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2145
   ScaleWidth      =   4080
   StartUpPosition =   3  'Windows-Standard
   Begin VB.TextBox WebPfad 
      Height          =   285
      Left            =   840
      TabIndex        =   6
      Top             =   120
      Width           =   3015
   End
   Begin VB.ListBox List1 
      BackColor       =   &H80000004&
      Height          =   840
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   2655
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Close"
      Height          =   375
      Left            =   3120
      TabIndex        =   1
      Top             =   1200
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Start"
      Height          =   375
      Left            =   3120
      TabIndex        =   0
      Top             =   600
      Width           =   855
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Index           =   0
      Left            =   4080
      Top             =   360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label2 
      Caption         =   "Web dir"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   735
   End
   Begin VB.Label conlab 
      Caption         =   "0"
      Height          =   255
      Left            =   1080
      TabIndex        =   4
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Connects:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1680
      Width           =   855
   End
End
Attribute VB_Name = "WebServerfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Connections As Integer


Private Sub Command1_Click()
Connections = 1
Me.Winsock1(0).Close
Me.Winsock1(0).LocalPort = 80
Me.Winsock1(0).Listen
Me.List1.AddItem Time & " Server started"
End Sub

Private Sub Ip(GetD, Index, ConnectD)
If ConnectD = "Connect" Then
    Me.List1.AddItem ConnectD & " " & Time & " " & Winsock1(Index).RemoteHostIP
Else
    Me.List1.AddItem ConnectD & " " & Time & " " & Winsock1(Index).RemoteHostIP
End If
End Sub

Private Sub Command2_Click()
End
End Sub

Private Sub Form_Load()
Me.Show
Me.WebPfad = App.Path & "\"
End Sub

Private Sub Winsock1_ConnectionRequest(Index As Integer, ByVal requestID As Long)
Ip strdata$, Index, "Connect"
  If Index = 0 Then
      Connections = Connections + 1
      conlab = conlab + 1
      Load Winsock1(Connections)
      Winsock1(Connections).LocalPort = 0
      Winsock1(Connections).Accept requestID
      
  End If
End Sub

Private Sub Winsock1_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Dim strdata As String
Winsock1(Index).GetData strdata$
If Mid$(strdata$, 1, 3) = "GET" Then
    findget = InStr(strdata$, "GET ")
    spc2 = InStr(findget + 5, strdata$, " ")
    Page = Mid$(strdata$, findget + 5, spc2 - (findget + 4))
    Ip strdata$, Index, "Aksed for " & Page
    SendPage Page, Index
End If
End Sub

Private Sub Winsock1_SendComplete(Index As Integer)
    Winsock1(Index).Close
    conlab = conlab - 1
End Sub
Public Sub SendPage(Page, Index)
On Error GoTo Fehler
If Page = " " Then Page = "index.html"
  Nr = FreeFile
  Tx$ = " "
  Lg = FileLen(WebPfad & Page)
  Open WebPfad & Page For Binary As Nr
    Tx1$ = ""
    For m = 1 To Lg
      Get #Nr, , Tx$
      Tx1$ = Tx1$ + Tx$
    Next
  Close Nr
  Winsock1(Index).SendData Tx1$
Exit Sub
Fehler:
If Err.Number = 53 Then Winsock1(Index).SendData "The URL you asked for does not exist on this website "
End Sub

