Attribute VB_Name = "modServer"
Option Explicit


' TODO: Verify admin level for various commands.
' Fix Games?
' Fix Host/Join Games
' Implement /away
' Implement /stats
' Implement /time
' Test login/account creation
' Prevent one person from opening/joining multiple games
' Remove games from the Game list after removal

Public Const MaxUsers As Integer = 50

Public Sub AddLog(strEvent As String)
    With frmMain.txtLog
    .Text = .Text & strEvent & vbCrLf
    End With
End Sub

Public Sub RemoveUser(Index As Integer)
    With Users(Index)

    AddLog Now & " - RemoveUser - Removing user " & .Nick & " with access level " & .AdminLevel & " with IP " & .IP
    If .LoggedIn Then SendToAll "ADMINTXTUser " & .Nick & " has left the server."
        .AdminLevel = 0
        .CurrentGame = vbNull
        '.Email = vbNull
        .Index = 0
        .IP = ""
        .Nick = vbNull
        .Status = vbNull
        .LoggedIn = False
        'TODO Add rest later
    End With
    DoEvents
    SendUserList
    
End Sub

Public Sub CloseIndex(Index As Integer, Optional Kick As Boolean)

    If Kick Then
        Send "KICKYou've been kicked.", Index
        AddLog Now & " - Kicking " & Users(Index).Nick & " (" & Users(Index).IP & ")"
    Else
        AddLog Now & " - Closing Index " & Index & "."
    End If
    DoEvents
    frmMain.sckServer(Index).Close
    RemoveUser (Index)
End Sub

Public Sub Parse(Data As String, Index As Integer)
Debug.Print Data
With frmMain.sckServer(Index)
Dim strLeft As String
Dim strLeftData As String
Dim strTemp() As String
Dim i As Integer
Dim ii As Integer
Dim GTFO As Boolean
strTemp = Split(Data, vbCrLf)

For i = 0 To UBound(strTemp)

    If Left(strTemp(i), 2) = "GD" Then
        If Not InStr(1, strTemp(i), "CHECKUNIT") >= 1 Then Debug.Print strTemp(i)
        'SendToAll strTemp(i) & vbCrLf
        If frmMain.sckServer(Users(Index).GameOpponentIndex).State = sckConnected Then frmMain.sckServer(Users(Index).GameOpponentIndex).SendData strTemp(i) & vbCrLf
        Exit For
    End If


Select Case Left(strTemp(i), 4)

    Case Is = "NAME"
        If Users(Index).LoggedIn Then CloseIndex Index, True: Exit Sub 'They're trying to haxxor
        DoEvents
        If Not IsUserLoggedIn(Right(strTemp(i), Len(strTemp(i)) - 4)) Then
            AddLog Now & " - Parsing NAME from " & Users(Index).IP
            Users(Index).Nick = Right(strTemp(i), Len(strTemp(i)) - 4)
            If VerifyKeyAsString(Users(Index).Nick) = False Then AddLog Now & " - Malformed Nick; removing user..": _
                CloseIndex Index, True: Exit Sub
        Else
            Send "ANNOUNCE" & "That account is already logged in!", Index
            DoEvents
            CloseIndex Index, True
            Exit Sub
        End If
        
    Case Is = "PASS"
        If Users(Index).LoggedIn Then CloseIndex Index, True: Exit Sub 'They're trying to haxxor
        DoEvents
        AddLog Now & " - Parsing PASS from " & Users(Index).Nick
        Users(Index).Password = Right(strTemp(i), Len(strTemp(i)) - 4)
        
        If Not DoesUserExist(Index) Then
            'create the account
            CreateAccount Index
        End If
        
        If IsPasswordCorrect(Index) Then
            If Not IsUserBanned(Index) Then
                .SendData ("GOODPASS" & vbCrLf): DoEvents
                Users(Index).LoggedIn = True
                WriteLastIP Index
                Users(Index).Index = Index
                SendUserList 'obvious
                SendMOTD Index
            Else
                Send "TEMPBAN1", Index
                DoEvents
                CloseIndex Index
            End If
        Else
            .SendData ("BADPASS" & vbCrLf)
            DoEvents
            CloseIndex (Index)
            Exit Sub
        End If
        
    'TODO SEND ANNOUNCEMENTS?
    
    Case Is = "VERS"
        Users(Index).Version = Right(strTemp(i), Len(strTemp(i)) - 7)
        DoEvents
        AddLog Now & " - Parsing VERS from " & Users(Index).Nick
    '.SendData ("BADVERS" & vbCrLf & "Test" & vbCrLf)
    'AddLog Now & " - SENDING BADVER: " & Index
    
    Case Is = "CHAT"
    strLeft = Right(strTemp(i), Len(strTemp(i)) - 4)
    AddLog Now & " - Parsing " & strTemp(i) & " from " & Users(Index).Nick
    If Left(strLeft, 3) = "TXT" Then
        SendToAll ("CHATTXT" & Right(strLeft, Len(strLeft) - 3)) & vbCrLf
    Else
        MsgBox "CHATMSG?!"
    End If
    
    Case Is = "MSGT" 'MSGTXT
        strLeft = Right(strTemp(i), Len(strTemp(i)) - 4)
        SendToAll ("CHATMSG" & Users(Index).Nick & " says to " & Users(Index).Status & ": " & Right(strLeft, Len(strLeft) - 3)) & vbCrLf
    
    Case Is = "MSGN" 'MSGNAME
        strLeft = Right(strTemp(i), Len(strTemp(i)) - 4)
    'SendToAll ("CHATMSG" & Users(Index).Nick & " " & Right(strLeft, Len(strLeft) - 3))
        Users(Index).Status = Right(strLeft, Len(strLeft) - 3)
    
    Case Is = "METX" 'METXT
        strLeft = Right(strTemp(i), Len(strTemp(i)) - 1)
        SendToAll ("METXT" & Right(strLeft, Len(strLeft) - 3)) & vbCrLf
    
    
    'GAME IMPLEMENTATION?
    
    Case Is = "JOIN"    'Sent to server, server must respond something.
                        'If you send JOINGAME, it opens an empty
                        'Join Game window, but doesn't connect.
                        'I suspect the Join Game window connects
                        'to it's opponent's IP Address, but I'm
                        'Not sure what data to send it. Perhaps GD something.
                        'Sending JOINGAME with an IP, or Gamenumber,
                        'or Gamename doesn't have any other outcome.
        'RELATED::
        'JOINNAME
        'JOINPASS
        'JOINBAD
        'GAMENUMBER
        'Not sure what to send back....
        
        strLeft = Right(strTemp(i), Len(strTemp(i)) - 4) 'NAME/PASS/BAD...
        strLeftData = Right(strLeft, Len(strLeft) - 4) 'Actual name/pass
        AddLog Now & " - Parsing " & strLeft & " from " & Users(Index).Nick
        AddLog Now & " Data: " & strLeftData
    Select Case Left(strLeft, 4)
        Case Is = "NAME" 'GAMENAME received
            Users(Index).GameName = strLeftData
            For ii = 1 To UBound(Users) 'Loop through all users
                If Users(ii).Hosting = True Then 'If one is hosting
                    If Users(ii).GameName = strLeftData Then 'Check to see if it's his game. Does not check for dupes.
                        Users(Index).GameOpponentNick = Users(ii).Nick 'set the joiner's opponent
                        Users(Index).GameOpponentIndex = ii 'set the joiner's opponent's winsock index
                        Users(ii).GameOpponentIndex = Index
                        Users(ii).GameOpponentNick = Users(Index).Nick
                        
                        Exit For
                    End If
                End If
            Next ii
            'SendToAll "ADDGAME" & Users(Index).GameName
        Case Is = "PASS"
        DoEvents
            '.SendData "GAMENUMBER" & 1 & vbCrLf 'Not sure..
            .SendData "JOINGAME" & "1" & vbCrLf
            DoEvents
        Case Is = "HOST"
            Users(Index).Hosting = False
    End Select
    
    Case Is = "SETU" 'SETUPGAME
    'Need to add game to list for everyone
    'Game Name isn't available yet.
    'Let's announce new game is starting, tho
        AddLog Now & " - Parsing SETU from " & Users(Index).Nick
        SendToAll "MODTXT" & Users(Index).Nick & " has started hosting a game!" & vbCrLf
        DoEvents
        .SendData "GAMENUMBER" & 1 & vbCrLf 'Not sure..
    
    Case Is = "REMO" 'REMOVEGAME
        AddLog Now & " - Parsing REMO from " & Users(Index).Nick
        Users(Index).Hosting = False
        'SEND REMOVEGAME to all
        
    Case Is = "STAR" 'STARTGAME, send to opponent
        AddLog Now & " - Parsing STAR from " & Users(Index).Nick
        If frmMain.sckServer(Users(Index).GameOpponentIndex).State = sckConnected Then _
        frmMain.sckServer(Users(Index).GameOpponentIndex).SendData strTemp(i) & vbCrLf
    
    Case Is = "GAME"
        AddLog Now & " - Parsing GAME from " & Users(Index).Nick
        strLeft = Right(strTemp(i), Len(strTemp(i)) - 4)
        strLeftData = Right(strLeft, Len(strLeft) - 4)
        AddLog Now & " Data: " & strLeft & " " & strLeftData
    Select Case Left(strLeft, 4)
    
        Case Is = "NAME"
            Users(Index).GameName = strLeftData
            SendToAll "ADDGAME" & Users(Index).GameName & vbCrLf
        Case Is = "PASS"
            Users(Index).GamePass = strLeftData
        Case Is = "HOST"
            Users(Index).Hosting = True
    End Select
      
    Case Else
    On Error Resume Next
    If Not GTFO Then
        If Not strTemp(i) = "" Then
            AddLog Now & " - ELSE: " & strTemp(i) & vbCrLf
                For ii = i To UBound(strTemp)
                    If Not strTemp(ii) = "" Then SendToAll "ADMINTXT" & Time & " - Unknown: " & strTemp(ii) & vbCrLf
                Next ii
        End If
    End If
    DoEvents
End Select

'When using the HOST or JOIN windows on client side,
'typing in the chat causes it to send a message to the server:
'GD<#>@@COMMANDdata
'So: GD1@@CHATbob:stuff
'the # is the Gamenumber/room number?

'Perhaps I have to populate the client Joingame list using listadd.. or a GD command?
'There are a lot of GD commands in the code..

'REMOVEGAME
'ADDGAME STARTED..?
'CHATTXT DONE!
'METXT DONE!
'ANNOUNCE CLIENTSIDE
'CHATMSG
'ADMINTXT CLIENTSIDE
'MODTXT CLIENTSIDE
'LISTKILL DONE?
'LISTADD DONE
'LISTDONE DONE
'BADVERS NOTIMPLEMENTINGYET
'BADPASS NOTIMPLEMENTINGYET
'KICK CLIENTSIDE
'TEMPBAN
'IPBAN
'COMPBAN
'Press '9' on the numberpad, type in docent6891 as the password for moderator/admin panel
'JOINGAME UNSURE.. WIP
'GAMENUMBER UNSURE.. WIP
'JOINBAD WIP
'GD I think = GameData?
'GETMOTD In client, not seen yet
'MSGNAME DONE
'MSGTXT DONE

'GETTIME /time
'AWAY /away
'GETSTATS
'SENDBUG /bug

Next i

End With

End Sub

Public Function Send(Data As String, Index As Integer) As Boolean
On Error GoTo sendFail
    frmMain.sckServer(Index).SendData Data & vbCrLf: Send = True
Exit Function

sendFail:
    Send = False
    Exit Function
End Function

Public Function SendToAll(Data As String) As Boolean
On Error Resume Next
Dim i As Integer

For i = 0 To UBound(Users)
DoEvents
If Users(i).LoggedIn = True Then
    If frmMain.sckServer(i) = sckConnected Then
    frmMain.sckServer(i).SendData Data
    DoEvents
    End If
End If

Next
DoEvents
End Function

Public Function SendUserList() As Boolean
On Error Resume Next
Dim i As Integer
SendToAll "LISTKILL" & vbCrLf
DoEvents
For i = 0 To UBound(Users)

If Users(i).LoggedIn = True Then
    SendToAll "LISTADD" & Users(i).Nick & vbCrLf
    DoEvents
End If

Next
DoEvents
SendToAll "LISTDONE" & vbCrLf
End Function

Public Function SendMOTD(Index As Integer) As Boolean

    Send "MODTXT" & ReadIni("Data/MOTD.ini", "MOTD", "1"), Index

End Function


