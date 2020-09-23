VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmRCV 
   BackColor       =   &H8000000A&
   Caption         =   "Receive Data Form"
   ClientHeight    =   6000
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4830
   LinkTopic       =   "Form1"
   ScaleHeight     =   6000
   ScaleWidth      =   4830
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text_HostIP 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   360
      TabIndex        =   7
      TabStop         =   0   'False
      Text            =   "Text4"
      Top             =   4560
      Width           =   1935
   End
   Begin VB.TextBox Text_HostPort 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2520
      TabIndex        =   6
      TabStop         =   0   'False
      Text            =   "Text5"
      Top             =   4560
      Width           =   1935
   End
   Begin VB.CommandButton Command_Clear 
      Caption         =   "Clear Data"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2640
      MouseIcon       =   "frmRCV.frx":0000
      MousePointer    =   99  'Custom
      Picture         =   "frmRCV.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   240
      Width           =   1815
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   2055
      Left            =   360
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   2040
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   3625
      _Version        =   393217
      ReadOnly        =   -1  'True
      TextRTF         =   $"frmRCV.frx":074C
   End
   Begin MSWinsockLib.Winsock WinsockDataRX 
      Left            =   2640
      Top             =   1320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   360
      Locked          =   -1  'True
      TabIndex        =   2
      TabStop         =   0   'False
      Text            =   "Text3"
      Top             =   1440
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   360
      Locked          =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Text            =   "Text2"
      Top             =   840
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   360
      Locked          =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Text            =   "Text1"
      Top             =   240
      Width           =   1935
   End
   Begin RichTextLib.RichTextBox RichTextBox_Messages 
      Height          =   615
      Left            =   360
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   5280
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   1085
      _Version        =   393217
      BackColor       =   -2147483633
      ReadOnly        =   -1  'True
      TextRTF         =   $"frmRCV.frx":07D7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      Caption         =   "Status Messages"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   360
      TabIndex        =   10
      Top             =   5040
      Width           =   4095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      Caption         =   "This Host's IP"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   360
      TabIndex        =   9
      Top             =   4320
      Width           =   1935
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      Caption         =   """Listen to"" Port"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   2520
      TabIndex        =   8
      Top             =   4320
      Width           =   1935
   End
End
Attribute VB_Name = "frmRCV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    
    '**********************************************************************
    '*
    '*  Mark Mokoski
    '*  19-FEB-2006
    '*  markm@cmtelephone.com
    '*  www.rjillc.com
    '*
    '*  Winsock inter-application Receive or "Host" App
    '*
    '*  Simple application to demonstrate a method of inter-application  communication with the use of the Winsock control.  Project is in two parts.
    '*
    '*  Part one is the receive or "Host" application.
    '*  The Host listens in it's IP address on a port that is set from the Host Form (default is port 8000).
    '*
    '*  Part two is the transmit or "Client" application.
    '*  When the "send" button is pressed, the Client connects to the Host, then sends data contained in the text box controls to the Host.
    '*
    '*  The Host receives data one "control" at a time in a Ping-Pong fashion.  First the control ID is sent, then acknowledged to the Client ("ACK"), then the Client sends the controls data and the Host sends a acknowledgement.  This is repeated until all control data is sent to the Host.  The Client then disconnects from the host.
    '*
    '*  The "Change" event for the controls on the Host can then be used to process the newly arrived data.
    '*
    '*  Both applications can be on the same computer (use the loopback address 127.0.0.1),
    '*  On different computers on your network, or even used over the public internet (not withstanding router, firewall and NAT translation settings).
    '*
    '*  I hope you fine this usefull.
    '*
    '***********************************************************************
    
    Option Explicit
    

Private Sub Command_Clear_Click()

    'Clear text boxes
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    RichTextBox1.Text = ""
    'Close (or make shure) currnt connection
    WinsockDataRX.Close
    'Get ready to accept a TPC/IP connection on port 8000
    WinsockDataRX.Listen

End Sub

Private Sub Form_Load()

    'Clear text boxes
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    RichTextBox1.Text = ""
    RichTextBox_Messages.Text = ""
    
    'Set default IP Port
    WinsockDataRX.Protocol = sckTCPProtocol
    'Find local IP address
    Text_HostIP.Text = WinsockDataRX.LocalIP
    Text_HostPort.Text = "8000"
    
    'Close any open connection
    WinsockDataRX.Close

    'Wait until it's close

        Do Until WinsockDataRX.State = sckClosed
            DoEvents
        Loop

    'Set socket to listen on
    WinsockDataRX.LocalPort = Val(Text_HostPort.Text)
    'Get ready to accept a TPC/IP connection on port 8000(default)
    WinsockDataRX.Listen

End Sub

Private Sub Text_HostPort_Change()

    'Close any opnen connection
    WinsockDataRX.Close

    'Wait until it's close

        Do Until WinsockDataRX.State = sckClosed
            DoEvents
        Loop

    'Change the port number to listen to
    WinsockDataRX.LocalPort = Val(Text_HostPort.Text)
    'Now listen on the new port
    WinsockDataRX.Listen

End Sub

Private Sub WinsockDataRX_Close()

    'Set up to listen for a new connection
    WinsockDataRX.Close

        Do Until WinsockDataRX.State = sckClosed
            DoEvents
        Loop

    'Get ready to accept a TPC/IP connection on port 8000
    WinsockDataRX.Listen
    Command_Clear.Enabled = True
    
End Sub

Private Sub WinsockDataRX_Connect()

    RichTextBox_Messages.Text = ""

End Sub

Private Sub WinsockDataRX_ConnectionRequest(ByVal requestID As Long)

    'Get the new connection info and accept it to start getting data

        If WinsockDataRX.State <> sckClosed Then WinsockDataRX.Close
    WinsockDataRX.Accept requestID
    Command_Clear.Enabled = False

End Sub

Private Sub WinsockDataRX_DataArrival(ByVal bytesTotal As Long)

    Dim controlID            As Byte
    Dim rcvText              As String

    'OK, here we get the data from the other app
    'We get the first byte and that is the ID for the control
    'that gets the remaining data.
    WinsockDataRX.GetData controlID, vbByte, 1

        Select Case controlID
        
            Case &H1    'Text1
                'Get the remaining Data
                WinsockDataRX.GetData rcvText, vbString
                Text1.Text = rcvText
                
            Case &H2    'Text2
                'Get the remaining Data
                WinsockDataRX.GetData rcvText, vbString
                Text2.Text = rcvText

            Case &H3    'Text3
                'Get the remaining Data
                WinsockDataRX.GetData rcvText, vbString
                Text3.Text = rcvText

            Case &H4    'RichTextBox1
                'Get the remaining Data
                WinsockDataRX.GetData rcvText, vbString
                RichTextBox1.Text = rcvText
        End Select

    'Send acknowlage to sening app
    WinsockDataRX.SendData "ACK"


End Sub

Private Sub WinsockDataRX_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)

    'Put an error message on the form

        If Description = "" Then
            RichTextBox_Messages.Text = "ERROR - Unknown Winsock Error, Connection Closed"
        Else
            RichTextBox_Messages.Text = "ERROR " & Number & " - " & Description
        End If

    'Close the connection
    WinsockDataRX.Close
    'Clear text boxes
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    RichTextBox1.Text = ""
    Command_Clear.Enabled = True

End Sub
