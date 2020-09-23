Attribute VB_Name = "deltree"
Option Explicit 'Makes sure all the variables are declared

'**********************************************************************************
'This module contains all the code for the deltree command. ie "deltree documents" etc.
'I put it in here to clean up the code on the frmGame a bit :-)
'**********************************************************************************

'They want to delete the Documents folder and all its contents
Public Sub DeltreeHomeDocuments()
    GetLevel 'Where are they?
    If Disconnected = False Then
        Select Case Level
            Case "home" 'If they are in the C:\ folder
                frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & "deltree documents" & vbCrLf
                frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Access is Denied" & vbCrLf
                frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Documents is a System Directory and Cannot be Deleted" & vbCrLf
                frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & vbCrLf
                Exit Sub

        End Select
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & "deltree documents" & vbCrLf
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Could not Find " & Lvl & "Documents" & vbCrLf
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & vbCrLf
        Exit Sub

    Else 'The Computer is disconnected
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & "Computer is Disconnected" & vbCrLf
    End If
    Exit Sub 'Exits the Sub :)
End Sub


'They want to delete the Images folder in the Documents folder and all its contents
Public Sub DeltreeHomeDocImages()
    GetLevel 'Where are they?
    If Disconnected = False Then
        Select Case Level
            Case "documents" 'If they are in the C:\Documents folder
                frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & "deltree images" & vbCrLf
                frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Access is Denied" & vbCrLf
                frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Images is a System Directory and Cannot be Deleted" & vbCrLf
                frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & vbCrLf
                Exit Sub

        End Select
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & "deltree images" & vbCrLf
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Could not Find " & Lvl & "Images" & vbCrLf
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & vbCrLf
        Exit Sub

    Else 'The Computer is disconnected
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & "Computer is Disconnected" & vbCrLf
    End If
    Exit Sub 'Exits the Sub :)
End Sub


'They want to delete the Recieved folder in the Documents folder and all its contents
Public Sub DeltreeHomeDocRecieved()
    GetLevel 'Where are they?
    If Disconnected = False Then
        Select Case Level
            Case "documents" 'If they are in the C:\Documents\ folder
                frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & "deltree recieved" & vbCrLf
                frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Access is Denied" & vbCrLf
                frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Recieved is a System Directory and Cannot be Deleted" & vbCrLf
                frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & vbCrLf
                Exit Sub

        End Select
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & "deltree recieved" & vbCrLf
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Could not Find " & Lvl & "Recieved" & vbCrLf
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & vbCrLf
        Exit Sub

    Else 'The Computer is disconnected
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & "Computer is Disconnected" & vbCrLf
    End If
    Exit Sub 'Exits the Sub :)
End Sub


'They want to delete the Help folder and all its contents
Public Sub DeltreeHomeHelp()
    GetLevel 'Where are they?
    If Disconnected = False Then
        Select Case Level
            Case "home" 'If they are in the C:\ folder
                frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & "deltree help" & vbCrLf
                frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Access is Denied" & vbCrLf
                frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Help is a System Directory and Cannot be Deleted" & vbCrLf
                frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & vbCrLf
                Exit Sub

        End Select
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & "deltree help" & vbCrLf
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Could not Find " & Lvl & "Help" & vbCrLf
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & vbCrLf
        Exit Sub

    Else 'The Computer is disconnected
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & "Computer is Disconnected" & vbCrLf
    End If
    Exit Sub 'Exits the Sub :)
End Sub


'They want to delete the Software folder and all its contents
Public Sub DeltreeHomeSoftware()
    GetLevel 'Where are they?
    If Disconnected = False Then
        Select Case Level
            Case "home" 'If they are in the C:\ folder
                frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & "deltree software" & vbCrLf
                frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Access is Denied" & vbCrLf
                frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Software is a System Directory and Cannot be Deleted" & vbCrLf
                frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & vbCrLf
                Exit Sub

        End Select
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & "deltree software" & vbCrLf
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Could not Find " & Lvl & "Software" & vbCrLf
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & vbCrLf
        Exit Sub

    Else 'The Computer is disconnected
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & "Computer is Disconnected" & vbCrLf
    End If
    Exit Sub 'Exits the Sub :)
End Sub


'They want to delete the Downloads folder and all its contents
Public Sub DeltreeHomeDownloads()
    GetLevel 'Where are they?
    If Disconnected = False Then
        Select Case Level
            Case "home" 'If they are in the C:\ folder
                frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & "deltree downloads" & vbCrLf
                frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Access is Denied" & vbCrLf
                frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Downloads is a System Directory and Cannot be Deleted" & vbCrLf
                frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & vbCrLf
                Exit Sub

        End Select
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & "deltree downloads" & vbCrLf
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Could not Find " & Lvl & "Downloads" & vbCrLf
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & vbCrLf
        Exit Sub

    Else 'The Computer is disconnected
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & "Computer is Disconnected" & vbCrLf
    End If
    Exit Sub 'Exits the Sub :)
End Sub


'They want to delete the System folder and all its contents
Public Sub DeltreeHomeSystem()
    GetLevel 'Where are they?
    If Disconnected = False Then
        Select Case Level
            Case "home" 'If they are in the C:\ folder
                frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & "deltree system" & vbCrLf
                frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Access is Denied" & vbCrLf
                frmGame.txtConsole.Text = frmGame.txtConsole.Text & "System is a System Directory and Cannot be Deleted" & vbCrLf
                frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & vbCrLf
                Exit Sub

        End Select
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & "deltree system" & vbCrLf
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Could not Find " & Lvl & "System" & vbCrLf
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & vbCrLf
        Exit Sub

    Else 'The Computer is disconnected
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & "Computer is Disconnected" & vbCrLf
    End If
    Exit Sub 'Exits the Sub :)
End Sub


'They want to delete the Boot folder in the System Directory and all its contents
Public Sub DeltreeHomeSysBoot()
    GetLevel 'Where are they?
    If Disconnected = False Then
        Select Case Level
            Case "homesystem" 'If they are in the C:\System folder
                frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & "deltree boot" & vbCrLf
                frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Access is Denied" & vbCrLf
                frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Boot is a System Directory and Cannot be Deleted" & vbCrLf
                frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & vbCrLf
                Exit Sub

        End Select
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & "deltree boot" & vbCrLf
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Could not Find " & Lvl & "Boot" & vbCrLf
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & vbCrLf
        Exit Sub

    Else 'The Computer is disconnected
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & "Computer is Disconnected" & vbCrLf
    End If
    Exit Sub 'Exits the Sub :)
End Sub


'They want to delete the Kernel folder in the System Directory and all its contents
Public Sub DeltreeHomeSysKernel()
    GetLevel 'Where are they?
    If Disconnected = False Then
        Select Case Level
            Case "homesystem" 'If they are in the C:\System\ folder
                frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & "deltree kernel" & vbCrLf
                frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Access is Denied" & vbCrLf
                frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Kernel is a System Directory and Cannot be Deleted" & vbCrLf
                frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & vbCrLf
                Exit Sub

        End Select
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & "deltree kernel" & vbCrLf
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Could not Find " & Lvl & "Kernel" & vbCrLf
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & vbCrLf
        Exit Sub

    Else 'The Computer is disconnected
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & "Computer is Disconnected" & vbCrLf
    End If
    Exit Sub 'Exits the Sub :)
End Sub
