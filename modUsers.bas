Attribute VB_Name = "modUsers"
Option Explicit

Public Type UserData
    IP As String
    Index As Integer
    AdminLevel As Integer '0 = normal user, 1 = mod, 2 = full admin
    Nick As String
    CurrentGame As String
    Status As String
    LoggedIn As Boolean
    Version As String
    Password As String
    GameName As String
    GamePass As String
    Hosting As Boolean
    GameStart As Boolean
    GameOpponentNick As String
    GameOpponentIndex As Integer
    GameMap As String
    GameTime As Integer
End Type

Public Const AccountPath = "Data\Accounts.ini"

Public Users(MaxUsers) As UserData

Public Function IsUserLoggedIn(Name As String) As Boolean
Dim i As Integer

    For i = 0 To UBound(Users)
        If LCase$(Name) = LCase$(Users(i).Nick) Then
            IsUserLoggedIn = True
            Exit Function
        End If
    Next i

End Function

Public Function IsPasswordCorrect(ByVal Index As Integer) As Boolean
'use readini to read password and verify it matches
Dim strTempPW As String
Dim strRealPW As String

IsPasswordCorrect = False
strRealPW = ReadIni(AccountPath, Users(Index).Nick, "Password")
strTempPW = Users(Index).Password

'filter through the password to make certain it doesn't contain strange characters
If VerifyKeyAsString(strTempPW) Then
    If strTempPW = strRealPW Then IsPasswordCorrect = True
End If

End Function

Public Function DoesUserExist(ByVal Index As Integer) As Boolean
'check if the user exists with readini
    DoesUserExist = False

    If Len(ReadIni(AccountPath, Users(Index).Nick, "Exists")) > 0 Then DoesUserExist = True

End Function

Public Sub CreateAccount(ByVal Index As Integer)
'pull the name/pass from the passed users() index
'use writeini to save the user
    
    If DoesUserExist(Index) Then AddLog ("Unable to create user! Already exists."): Exit Sub
    
    WriteIni AccountPath, Users(Index).Nick, "Exists", 1
    WriteIni AccountPath, Users(Index).Nick, "Banned", 0
    WriteIni AccountPath, Users(Index).Nick, "Admin", 0
    WriteIni AccountPath, Users(Index).Nick, "Win", 0
    WriteIni AccountPath, Users(Index).Nick, "Lose", 0
    WriteIni AccountPath, Users(Index).Nick, "Password", Users(Index).Password
    WriteIni AccountPath, Users(Index).Nick, "LastIP", ""
    
    
    SendToAll "MODTXT" & "Welcome new user " & Users(Index).Nick & "!"
    
End Sub

Public Sub WriteLastIP(ByVal Index As Integer)
'write their last login IP
    WriteIni AccountPath, Users(Index).Nick, "LastIP", Users(Index).IP
End Sub

Public Function IsUserBanned(ByVal Index As Integer) As Boolean
'verify IP and/or if account is banned

    If Val(ReadIni(AccountPath, Users(Index).Nick, "Banned")) > 0 Then IsUserBanned = True

End Function
