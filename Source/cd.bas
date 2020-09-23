Attribute VB_Name = "cd"
Option Explicit 'Makes sure all the variables are declared

'**********************************************************************************
'This module contains all the code for the del command. ie "del readme.txt" etc.
'I put it in here to clean up the code on the frmGame a bit :-)
'**********************************************************************************


'They want to change directory to the Documents folder
Public Sub CdDocuments()
    GetLevel 'Where are we? Loads the GetLevel Sub in (main.bas)
    If Disconnected = False Then
        Select Case Server
            Case "Server X10 - Home Computer"
                Select Case Level
                    Case "home"
                        Level = "documents" 'Makes it so we are in the Documents Folder
                        frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & "C:\Documents\" & vbCrLf
                        Exit Sub
                        
                End Select 'Level Select

            'Case "Somwhere Else"
                'Blah blah blah

            frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & "cd Recieved" & vbCrLf
            frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Could not Find " & Lvl & "Recieved" & vbCrLf
            frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & vbCrLf
            Exit Sub
        End Select 'Server Select
    Else
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & "Computer is Disconnected" & vbCrLf
    End If
    Exit Sub 'Ends the Current Sub
End Sub


'They want to change directory to the Recieved folder
Public Sub CdRecieved()
    GetLevel 'Where are we? Loads the GetLevel Sub in (main.bas)
    If Disconnected = False Then
        Select Case Server 'Are we home or at another server?
            Case "Server X10 - Home Computer"
                Select Case Level
                    Case "documents"
                        Level = "homerecieved"
                        frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & "C:\Documents\Recieved\" & vbCrLf
                        Exit Sub
                End Select 'Level Select
                
            'Case "Somwhere Else"
                'Blah blah blah
                
            frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & "cd Recieved" & vbCrLf
            frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Could not Find " & Lvl & "Recieved" & vbCrLf
            frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & vbCrLf
            Exit Sub
        End Select 'Server Select
    Else
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & "Computer is Disconnected" & vbCrLf
    End If
    Exit Sub
End Sub


'They want to change directory to the Images folder
Public Sub CdImages()
    GetLevel 'Where are we? Loads the GetLevel Sub in (main.bas)
    If Disconnected = False Then
        Select Case Server 'Are we home or at another server?
            Case "Server X10 - Home Computer"
                Select Case Level
                    Case "documents"
                        Level = "homedocimages"
                        frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & "C:\Documents\Images\" & vbCrLf
                        Exit Sub
                End Select 'Level Select
                
            'Case "Somwhere Else"
                'Blah blah blah
                
            frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & "cd Images" & vbCrLf
            frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Could not Find " & Lvl & "Images" & vbCrLf
            frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & vbCrLf
            Exit Sub
        End Select 'Server Select
    Else
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & "Computer is Disconnected" & vbCrLf
    End If
    Exit Sub
End Sub


'They want to change directory to the Downloads folder
Public Sub CdDownloads()
    GetLevel 'Where are we? Loads the GetLevel Sub in (main.bas)
    If Disconnected = False Then
        Select Case Server 'Are we home or at another server?
            Case "Server X10 - Home Computer"
                Select Case Level
                    Case "home"
                        Level = "homedownloads"
                        frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & "C:\Downloads\" & vbCrLf
                        Exit Sub
                End Select 'Level Select
                
            'Case "Somwhere Else"
                'Blah blah blah
                
            frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & "cd Downloads" & vbCrLf
            frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Could not Find " & Lvl & "Downloads" & vbCrLf
            frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & vbCrLf
            Exit Sub
        End Select 'Server Select
    Else
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & "Computer is Disconnected" & vbCrLf
    End If
    Exit Sub
End Sub


'They want to change directory to the Software folder
Public Sub CdSoftware()
    GetLevel 'Where are we? Loads the GetLevel Sub in (main.bas)
    If Disconnected = False Then
        Select Case Server 'Are we home or at another server?
            Case "Server X10 - Home Computer"
                Select Case Level
                    Case "home"
                        Level = "homesoftware"
                        frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & "C:\Software\" & vbCrLf
                        Exit Sub
                End Select 'Level Select
                
            'Case "Somwhere Else"
                'Blah blah blah
                
            frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & "cd Software" & vbCrLf
            frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Could not Find " & Lvl & "Software" & vbCrLf
            frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & vbCrLf
            Exit Sub
        End Select 'Server Select
    Else
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & "Computer is Disconnected" & vbCrLf
    End If
    Exit Sub
End Sub


'They want to change directory to the System folder
Public Sub CdSystem()
    GetLevel 'Where are we? Loads the GetLevel Sub in (main.bas)
    If Disconnected = False Then
        Select Case Server 'Are we home or at another server?
            Case "Server X10 - Home Computer"
                Select Case Level
                    Case "home"
                        Level = "homesystem"
                        frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & "C:\System\" & vbCrLf
                        Exit Sub
                End Select 'Level Select
                
            'Case "Somwhere Else"
                'Blah blah blah
                
            frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & "cd System" & vbCrLf
            frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Could not Find " & Lvl & "System" & vbCrLf
            frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & vbCrLf
            Exit Sub
        End Select 'Server Select
    Else
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & "Computer is Disconnected" & vbCrLf
    End If
    Exit Sub
End Sub


'They want to change directory to the Boot folder in the System folder
Public Sub CdSysBoot()
    GetLevel 'Where are we? Loads the GetLevel Sub in (main.bas)
    If Disconnected = False Then
        Select Case Server 'Are we home or at another server?
            Case "Server X10 - Home Computer"
                Select Case Level
                    Case "homesystem"
                        Level = "homesysboot"
                        frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & "C:\System\Boot\" & vbCrLf
                        Exit Sub
                End Select 'Level Select
                
            'Case "Somwhere Else"
                'Blah blah blah
                
            frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & "cd Boot" & vbCrLf
            frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Could not Find " & Lvl & "Boot" & vbCrLf
            frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & vbCrLf
            Exit Sub
        End Select 'Server Select
    Else
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & "Computer is Disconnected" & vbCrLf
    End If
    Exit Sub
End Sub


'They want to change directory to the Kernel folder in the System folder
Public Sub CdSysKernel()
    GetLevel 'Where are we? Loads the GetLevel Sub in (main.bas)
    If Disconnected = False Then
        Select Case Server 'Are we home or at another server?
            Case "Server X10 - Home Computer"
                Select Case Level
                    Case "homesystem"
                        Level = "homesyskernel"
                        frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & "C:\System\Kernel\" & vbCrLf
                        Exit Sub
                End Select 'Level Select
                
            'Case "Somwhere Else"
                'Blah blah blah
                
            frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & "cd Kernel" & vbCrLf
            frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Could not Find " & Lvl & "Kernel" & vbCrLf
            frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & vbCrLf
            Exit Sub
        End Select 'Server Select
    Else
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & "Computer is Disconnected" & vbCrLf
    End If
    Exit Sub
End Sub


'They want to change directory to the System folder
Public Sub CdHomeHelp()
    GetLevel 'Where are we? Loads the GetLevel Sub in (main.bas)
    If Disconnected = False Then
        Select Case Server 'Are we home or at another server?
            Case "Server X10 - Home Computer"
                Select Case Level
                    Case "home"
                        Level = "homehelp"
                        frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & "C:\Help\" & vbCrLf
                        Exit Sub
                End Select 'Level Select
                
            'Case "Somwhere Else"
                'Blah blah blah
                
            frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & "cd Help" & vbCrLf
            frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Could not Find " & Lvl & "Help" & vbCrLf
            frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & vbCrLf
            Exit Sub
        End Select 'Server Select
    Else
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & "Computer is Disconnected" & vbCrLf
    End If
    Exit Sub
End Sub
