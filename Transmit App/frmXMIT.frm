VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmXMIT 
   BackColor       =   &H8000000A&
   Caption         =   "Transmit Data Form"
   ClientHeight    =   6000
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4830
   LinkTopic       =   "Form1"
   ScaleHeight     =   6000
   ScaleWidth      =   4830
   StartUpPosition =   3  'Windows Default
   Begin RichTextLib.RichTextBox RichTextBox_Messages 
      Height          =   615
      Left            =   360
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   5280
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   1085
      _Version        =   393217
      BackColor       =   -2147483633
      ReadOnly        =   -1  'True
      TextRTF         =   $"frmXMIT.frx":0000
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
      TabIndex        =   7
      TabStop         =   0   'False
      Text            =   "Text5"
      Top             =   4560
      Width           =   1935
   End
   Begin VB.TextBox Text_HostIP 
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
      Left            =   360
      TabIndex        =   6
      TabStop         =   0   'False
      Text            =   "Text4"
      Top             =   4560
      Width           =   1935
   End
   Begin VB.Timer Timer_TimeOut 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   3240
      Top             =   1320
   End
   Begin VB.CommandButton Command_Send 
      Caption         =   "Send Data"
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
      MouseIcon       =   "frmXMIT.frx":0082
      MousePointer    =   99  'Custom
      Picture         =   "frmXMIT.frx":038C
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   240
      Width           =   1815
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   2055
      Left            =   360
      TabIndex        =   3
      Top             =   2040
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   3625
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmXMIT.frx":07CE
   End
   Begin MSWinsockLib.Winsock WinsockDataTX 
      Left            =   2640
      Top             =   1320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Text            =   "Text3"
      Top             =   1440
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   840
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   240
      Width           =   1935
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      Caption         =   "Host Application Port"
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
      TabIndex        =   9
      Top             =   4320
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      Caption         =   "Host Application IP"
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
      TabIndex        =   8
      Top             =   4320
      Width           =   1935
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
      TabIndex        =   5
      Top             =   5040
      Width           =   4095
   End
End
Attribute VB_Name = "frmXMIT"
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
    '*  Winsock inter-application Transmit or "Client" App
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
    Dim dataACK                As Boolean
    Dim Error_wsock            As Boolean
   
Private Sub Command_Send_Click()

    '/Purpose: Send Data to remote Application
    '/Created: 2/21/2006 03:15 PM
    '/Created By: Mark Mokoski

    'Disable the "Send" Button

    On Error GoTo Err_Command_Send_Click

    Command_Send.Enabled = False
    'Clear any old error messages
    RichTextBox_Messages.Text = ""

    'Open connection to remote Socket (loopback IP 127.0.0.1)
    WinsockDataTX.Connect Text_HostIP.Text, Val(Text_HostPort.Text)

    'Set connection time-out timer
    Timer_TimeOut.Enabled = True

    'Wait until Connection is complete

        Do While WinsockDataTX.State <> sckConnected
            DoEvents

                If Error_wsock = True Then Exit Do
        Loop

    'Stop connection time-out timer
    Timer_TimeOut.Enabled = False

    Command_Send.Enabled = False
    'Set data ACK flag to start state
    dataACK = False

    Dim x                   As Integer
    Dim dataSlot            As Byte
    Dim dataText            As String
    
    'With "for / next" loop, run thru all the controls and send data od each
    'in turn as they are acknowlaged (ACK)from the companion app

        For x = 1 To 4

                Select Case x
                
                    Case 1  'Send Text1 Text

                        If Text1.Text <> "" Then    'If nothing in textbox, don't sens anything
                            dataSlot = &H1          '"Control ID code" for textbox 1
                            dataText = Text1.Text
                            Call TX_SendData(dataSlot, dataText)
                        End If

                    Case 2  'Send Text2 Text

                        If Text2.Text <> "" Then    'If nothing in textbox, don't sens anything
                            dataSlot = &H2          '"Control ID code" for textbox 2
                            dataText = Text2.Text
                            Call TX_SendData(dataSlot, dataText)
                        End If

                    Case 3  'Send Text13 Text

                        If Text3.Text <> "" Then    'If nothing in textbox, don't sens anything
                            dataSlot = &H3          '"Control ID code" for textbox 3
                            dataText = Text3.Text
                            Call TX_SendData(dataSlot, dataText)
                        End If

                    Case 4  'Send RichTextBox1 Text

                        If RichTextBox1.Text <> "" Then     'If nothing in textbox, don't sens anything
                            dataSlot = &H4                  '"Control ID code" for richtextbox 1
                            dataText = RichTextBox1.Text
                            Call TX_SendData(dataSlot, dataText)
                        End If

                End Select

        Next x
        
    'When we are all done sending any data, clear the text boxes and close the connection
    'Clear text boxes
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    RichTextBox1.Text = ""
    
    'Close connection
    WinsockDataTX.Close

    'Wait for connection to close

        Do Until WinsockDataTX.State = sckClosed
            DoEvents
        Loop

    Command_Send.Enabled = True

        If Error_wsock = False Then
            RichTextBox_Messages.Text = "Data Sent Successfully!"
        End If

    Error_wsock = False
    Text1.SetFocus


Exit_Command_Send_Click:

    On Error Resume Next
    Exit Sub

Err_Command_Send_Click:

    MsgBox Err.Number & ": " & Err.Description, vbCritical, "frmXMIT" & ": " & "Command_Send_Click"
    Resume Exit_Command_Send_Click

End Sub

Private Sub Form_Load()

    'Clear text boxes
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    RichTextBox1.Text = ""
    RichTextBox_Messages.Text = ""
    'Set default host IP and Port
    Text_HostIP.Text = "127.0.0.1"
    Text_HostPort = "8000"
    'Set Winsock to TCP Protocol
    WinsockDataTX.Protocol = sckTCPProtocol
    Error_wsock = False

End Sub

Private Sub Form_Terminate()

    'Force connection closed on form terminate
    WinsockDataTX.Close
    Unload Me

End Sub


Private Sub Form_Unload(Cancel As Integer)

    'Force connection closed on form terminate
    WinsockDataTX.Close

End Sub

Private Sub Timer_TimeOut_Timer()

    'Stop the timer
    Timer_TimeOut.Enabled = False
    'Put error message on the form
    RichTextBox_Messages.Text = "ERROR - Connection Timed Out, Connection Closed"
    'Clear text boxes
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    RichTextBox1.Text = ""
    
    'Close connection
    WinsockDataTX.Close

    'Wait for connection to close

        Do Until WinsockDataTX.State = sckClosed
            DoEvents
        Loop

    Command_Send.Enabled = True


End Sub

Private Sub WinsockDataTX_DataArrival(ByVal bytesTotal As Long)

    Dim controlID            As String

    'OK, here we get data from the other app.
    'We get the data and and see if returned as an "ACK"
    'then send more data or close socket if done

    WinsockDataTX.GetData controlID, vbString

        If controlID = "ACK" Then
            dataACK = True
        Else
            dataACK = False
        End If

End Sub

Private Sub WinsockDataTX_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)

    'Beep to get your attention
    Beep

    'Diable time-out timer
    Timer_TimeOut.Enabled = False
    'Put an error message on the form
    Error_wsock = True

        If Description = "" Then
            RichTextBox_Messages.Text = "ERROR - Unknown Winsock Error, Connection Closed"
        Else
            RichTextBox_Messages.Text = "ERROR " & Number & " - " & Description
        End If

    'Close the connection
    WinsockDataTX.Close
    'Clear text boxes
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    RichTextBox1.Text = ""
    Command_Send.Enabled = True

End Sub

Private Sub TX_SendData(Cdata As Byte, Tdata As String)

    'Send control data
    WinsockDataTX.SendData (Cdata)  'Control "ID" code
    WinsockDataTX.SendData (Tdata)  'Control Text
    
    'Set connection time-out timer
    Timer_TimeOut.Enabled = True
    
    'Wait for Host App Acknowlage

        Do Until dataACK = True
            DoEvents
        Loop

    'Stop connection time-out timer
    Timer_TimeOut.Enabled = False

    dataACK = False

End Sub
