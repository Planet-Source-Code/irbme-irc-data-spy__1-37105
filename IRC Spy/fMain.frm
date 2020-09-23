VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "IRC packet sniffer"
   ClientHeight    =   9855
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9510
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9855
   ScaleWidth      =   9510
   StartUpPosition =   3  'Windows Default
   Begin RichTextLib.RichTextBox txtHelp 
      Height          =   7575
      Left            =   120
      TabIndex        =   15
      Top             =   1680
      Visible         =   0   'False
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   13361
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      Appearance      =   0
      TextRTF         =   $"fMain.frx":0000
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "What do I do?"
      Height          =   375
      Left            =   7920
      TabIndex        =   16
      Top             =   9360
      Width           =   1455
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Sed Command to Server"
      Height          =   375
      Left            =   6960
      TabIndex        =   14
      Top             =   1200
      Width           =   2415
   End
   Begin VB.TextBox txtSay 
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Text            =   "Send Commands to server"
      Top             =   1200
      Width           =   6615
   End
   Begin RichTextLib.RichTextBox txtData 
      Height          =   7575
      Left            =   120
      TabIndex        =   9
      Top             =   1680
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   13361
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      Appearance      =   0
      TextRTF         =   $"fMain.frx":05F2
   End
   Begin VB.OptionButton ChatMode 
      Caption         =   "Chat Text"
      Height          =   375
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   9360
      Width           =   1935
   End
   Begin VB.OptionButton RawData 
      Caption         =   "Raw Data"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   9360
      Value           =   -1  'True
      Width           =   1935
   End
   Begin VB.Frame clnt 
      Caption         =   "Client"
      Height          =   975
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   4575
      Begin VB.CommandButton cmdListen 
         Caption         =   "Listen"
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   3015
      End
      Begin VB.TextBox txtClientPort 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3840
         TabIndex        =   6
         Text            =   "6667"
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "Port:"
         Height          =   255
         Left            =   3360
         TabIndex        =   7
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame srvr 
      Caption         =   "Server"
      Height          =   975
      Left            =   4800
      TabIndex        =   0
      Top             =   120
      Width           =   4575
      Begin VB.TextBox txtServerPort 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3840
         TabIndex        =   3
         Text            =   "6667"
         Top             =   360
         Width           =   495
      End
      Begin VB.TextBox txtServer 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   840
         TabIndex        =   1
         Text            =   "us.undernet.org"
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label Label2 
         Caption         =   "Port:"
         Height          =   255
         Left            =   3360
         TabIndex        =   4
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Address:"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   855
      End
   End
   Begin MSWinsockLib.Winsock Srv 
      Left            =   5880
      Top             =   9360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock IRC 
      Left            =   5400
      Top             =   9360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox txtChat 
      Height          =   7575
      Left            =   120
      TabIndex        =   12
      Top             =   1680
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   13361
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      Appearance      =   0
      TextRTF         =   $"fMain.frx":0674
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'HOW THIS WORKS:
'--------------------------------------------

'We listen for a connection from an IRC client

'Once we get one, we have to fool it into thinking we are a server
'To do this, we connect to a real server
'Now we accept the clients connection

'Whenever the client sends data, we relay it on to the server
'Whenever the server sends data, we relay it on to the client

'So the server thinks we are the client and the client thinks we are the server
'In actual fact we are sitting right in the middle them


'HOW ITS USEFULL
'--------------------------------------------

'If you want to write an IRC client or server then this is a valuable program
'It lets you see all the data!

'If you want to write oyur own Chat program, its useful to know how others work

'THE CODE
'--------------------------------------------

Private Sub ChatMode_Click()
    
    'Display the correct RTF boxes
    If RawData.Value = True Then
        txtData.Visible = True
        txtChat.Visible = False
    Else
        txtData.Visible = False
        txtChat.Visible = True
    End If

End Sub


Private Sub cmdHelp_Click()

Static state As Boolean
    
    'Display the correct RTF boxes
    If state = False Then
        txtHelp.Visible = True
        txtChat.Visible = False
        txtData.Visible = False
        
        cmdHelp.Caption = "OK"
    Else
        txtHelp.Visible = False
        txtChat.Visible = True
        txtData.Visible = True
        
        cmdHelp.Caption = "What do I do?"
    End If
    
    state = Not state
    
End Sub

Private Sub cmdListen_Click()

Static state As Boolean
    
    'Set up RTF text
    txtData.SelStart = Len(txtData.Text)
    txtData.SelColor = RGB(0, 0, 180)
    
    'If listen
    If state = False Then
        'Set local port and start listening
        IRC.LocalPort = txtClientPort.Text
        IRC.Listen
        
        'change caption and display status in RTF box
        cmdListen.Caption = "Cancel"
        
        txtData.SelText = txtData.SelText & vbCrLf & "Listening for client connections on port " & txtClientPort.Text & "..."
        txtData.SelStart = Len(txtData.Text)
    'If cancel
    ElseIf state = True Then
        
        'Stop listening
        IRC.Close
        
        'change caption and display status in RTF box
        cmdListen.Caption = "Listen"
        
        txtData.SelText = txtData.SelText & vbCrLf & "Not listening..."
        txtData.SelStart = Len(txtData.Text)
    End If
    
    state = Not state
        
End Sub


Private Sub cmdSend_Click()
    
    'Setup RTF box
    txtData.SelStart = Len(txtData.Text)
    txtData.SelColor = RGB(0, 0, 180)

    'If connected
    If Srv.state = sckConnected Then
        
        'Send data
        Srv.SendData txtSay.Text
        
        'Display status
        txtData.SelText = txtData.SelText & vbCrLf & "Sending data: " & txtSay.Text
        txtData.SelStart = Len(txtData.Text)
    Else
        'Display error
        txtData.SelText = txtData.SelText & vbCrLf & "Cannot send data. Not connected"
        txtData.SelStart = Len(txtData.Text)
    End If
    
End Sub


Private Sub Form_Unload(Cancel As Integer)
    
    'Close connections and end
    IRC.Close
    Srv.Close
    End

End Sub


Private Sub IRC_ConnectionRequest(ByVal requestID As Long)
   
  Dim I As Long
    
    'Setup RTF box
    txtData.SelStart = Len(txtData.Text)
    txtData.SelColor = RGB(0, 0, 180)
    
    'Close any previous connections
    IRC.Close
    Srv.Close
    
    'Display status
    txtData.SelText = txtData.SelText & vbCrLf & "Connecting to " & txtServer.Text & " (" & txtServerPort.Text & ")" & "..."
    txtData.SelStart = Len(txtData.Text)
    
    'Connect to server
    Srv.Connect txtServer.Text, txtServerPort.Text
    
    'Wait until connected
    Do Until Srv.state = sckConnected
        DoEvents
    Loop

    'Display status
    txtData.SelText = txtData.SelText & vbCrLf & "Connecting to client (" & txtClientPort.Text & ")" & "..."
    txtData.SelStart = Len(txtData.Text)
    
    'Accept client connection request
    IRC.Accept requestID
    
    'Disable connection boxes
    clnt.Enabled = False
    srvr.Enabled = False
    
    'Display status
    txtData.SelText = txtData.SelText & vbCrLf & "Connected" & vbCrLf & "---------------------------------------------------------------" & vbCrLf & vbCrLf
    txtData.SelStart = Len(txtData.Text)
    
End Sub


Private Sub IRC_DataArrival(ByVal bytesTotal As Long)

  Dim data As String

    IRC.GetData data
    
    'If connected
    If Srv.state = sckConnected Then
        'Relay data on to server
        Srv.SendData data
    Else
        'Display error
        txtData.SelStart = Len(txtData.Text)
        txtData.SelColor = RGB(0, 0, 180)
        txtData.SelText = txtData.SelText & vbCrLf & "Disconnected from server"
        txtData.SelStart = Len(txtData.Text)
    End If
    
    'Display recieved data
    txtData.SelStart = Len(txtData.Text)
    txtData.SelColor = RGB(180, 0, 0)
    txtData.SelText = txtData.SelText & vbCrLf & data
    txtData.SelStart = Len(txtData.Text)
    
End Sub


Private Sub RawData_Click()
    
    'Display correct RTF boxes
    If RawData.Value = True Then
        txtData.Visible = True
        txtChat.Visible = False
    Else
        txtData.Visible = False
        txtChat.Visible = True
    End If
    
End Sub


Private Sub Srv_DataArrival(ByVal bytesTotal As Long)
  
  Dim data As String

    Srv.GetData data
    
    'If connected
    If IRC.state = sckConnected Then
        'Relay data to client
        IRC.SendData data
    Else
        'Display error
        txtData.SelStart = Len(txtData.Text)
        txtData.SelColor = RGB(0, 0, 180)
        txtData.SelText = txtData.SelText & vbCrLf & "Disconnected from client"
        txtData.SelStart = Len(txtData.Text)
    End If
    
    'Display recieved data
    txtData.SelStart = Len(txtData.Text)
    txtData.SelColor = RGB(0, 180, 0)
    txtData.SelText = txtData.SelText & vbCrLf & data
    txtData.SelStart = Len(txtData.Text)

End Sub
