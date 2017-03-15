VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Time waster"
   ClientHeight    =   4140
   ClientLeft      =   7620
   ClientTop       =   1215
   ClientWidth     =   8415
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4140
   ScaleWidth      =   8415
   Begin VB.CommandButton cmdSend 
      Caption         =   "[S]"
      Height          =   255
      Left            =   8040
      TabIndex        =   2
      Top             =   3840
      Width           =   375
   End
   Begin VB.TextBox txtServer 
      Height          =   285
      Left            =   0
      TabIndex        =   1
      Text            =   "ANNOUNCEHello"
      Top             =   3840
      Width           =   8055
   End
   Begin MSWinsockLib.Winsock sckServer 
      Index           =   0
      Left            =   6480
      Top             =   3240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   40334
   End
   Begin VB.TextBox txtLog 
      Height          =   3855
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   0
      Width           =   8415
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdSend_Click()
SendToAll txtServer.Text & vbCrLf
txtServer.Text = ""
End Sub

Private Sub Form_Load()
InitServer
AddLog Now & " - InitServer Completed."
End Sub

Private Sub sckServer_Close(Index As Integer)
On Error Resume Next
sckServer(Index).Close
RemoveUser (Index)
End Sub

Private Sub sckServer_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Dim strData As String
sckServer(Index).GetData strData
'Debug.Print strData

'Debug.Print StringToHex(strData)
If Users(Index).Nick = "" Then AddLog Now & " - Parsing new packet from " & Users(Index).IP ' Else AddLog Now & " - Parsing new packet from " & Users(Index).Nick
Parse strData, Index

End Sub

Private Sub sckServer_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
sckServer(Index).Close
RemoveUser (Index)
AddLog Time & " - sckServer Err Index: " & Index & " - " & Number & ": " & Description
End Sub

Private Sub sckServer_ConnectionRequest(Index As Integer, ByVal requestID As Long)
Dim i As Integer
AddLog Now & " - Connection request from " & sckServer(Index).RemoteHostIP & " on index " & Index
For i = 1 To MaxUsers
    If Users(i).IP = "" Then
        Users(i).IP = sckServer(Index).RemoteHostIP
        sckServer(i).Close
        sckServer(i).Accept requestID
        AddLog "Accepted connection on socket #" & i
        Exit For
    End If
Next
End Sub

Public Sub InitServer()
    On Error Resume Next
    Dim i As Integer

    For i = 1 To MaxUsers
        Load sckServer(i)
    Next

With sckServer(0)
    .Close
    '.LocalIP = ""
    .LocalPort = "9876"
    .Listen
End With

If sckServer(0).State = sckListening Then AddLog Now & " - Server listening on port " & sckServer(0).LocalPort _
Else AddLog Now & " - Server init failed, not listening for incoming connections!"

End Sub

Private Sub txtLog_Change()
On Error Resume Next
txtLog.SelStart = Len(txtLog.Text)
End Sub

Private Sub txtServer_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cmdSend_Click
End Sub
