Attribute VB_Name = "del"
Option Explicit 'Makes sure all the variables are declared

'**********************************************************************************
'This module contains all the code for the del command. ie "del readme.txt" etc.
'I put it in here to clean up the code on the frmGame a bit :-)
'**********************************************************************************


'They want to Delete C:\Documents\Images\Test.jpg
Public Sub DelHomeDocImgTest()
    GetLevel 'Where are we? Loads the GetLevel Sub in (main.bas)
    If Disconnected = False Then 'Are we disconnected
        If Level = "homedocimages" Then
            If IsHomeDocImgTestDel = False Then 'Has it been deleted already?
                frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & "del test.jpg" & vbCrLf
                frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Successfully Deleted" & vbCrLf
                IsHomeDocImgTestDel = True 'Set it so that the file has been deleted (Sort Of)
                frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & vbCrLf
            Else 'Yes it has :[
                frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & "del Test.jpg" & vbCrLf
                frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Could not Find " & Lvl & "Test.jpg" & vbCrLf
                frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & vbCrLf
            End If
            Exit Sub 'Ends the Sub :)
        End If
        
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & "del Test.jpg" & vbCrLf
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Could not Find " & Lvl & "Test.jpg" & vbCrLf
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & vbCrLf
        Exit Sub
        
    Else 'The Computer is disconnected
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & "Computer is Disconnected" & vbCrLf
    End If
    Exit Sub 'Ends the Sub :)
End Sub


'They want to Delete a Readme.txt File
Public Sub DelReadmetxt()
    GetLevel  'Where are we? Loads the GetLevel Sub in (main.bas)
    If Disconnected = False Then
        Select Case Level
            'They want to delete C:\Readme.txt
            Case "home" 'Are we Home (C:\)?
                If IsHomeReadmeDel = False Then 'Has it been deleted already?
                    GetLevel 'Get's the current level
                    frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & "del readme.txt" & vbCrLf
                    frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Successfully Deleted" & vbCrLf
                    IsHomeReadmeDel = True 'Deletes the File. (Sort-of)
                    frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & vbCrLf
                Else 'Oh No It has been deleted :-)
                    frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & "del readme.txt" & vbCrLf
                    frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Could not Find " & Lvl & "Readme.txt" & vbCrLf
                    frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & vbCrLf
                End If
                Exit Sub 'Ends the Sub :)
                
            'They want to delete C:\Documents\Recieved\Readme.txt
            Case "homerecieved" 'Are we In the recieved folder (C:\Documents\Recieved\)?
                If IsHomeDocRecReadmeDel = False Then  'Has it been deleted already?
                    GetLevel 'Get's the current level
                    frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & "del readme.txt" & vbCrLf
                    frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Successfully Deleted" & vbCrLf
                    IsHomeDocRecReadmeDel = True 'Tell the program that is has been deleted
                    frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & vbCrLf
                Else 'It has been deleted
                    frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & "del readme.txt" & vbCrLf
                    frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Could not Find " & Lvl & "Readme.txt" & vbCrLf
                    frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & vbCrLf
                End If
                Exit Sub 'Ends the Sub
        End Select
        
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & "del readme.txt" & vbCrLf
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Could not Find " & Lvl & "Readme.txt" & vbCrLf
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & vbCrLf
        Exit Sub
    Else 'The Computer is disconnected
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & "Computer is Disconnected" & vbCrLf
    End If
    Exit Sub 'Exits the Sub :)
End Sub


'They want to Delete C:\System\View.exe
Public Sub DelSysViewexe()
    GetLevel 'Where are they?
    If Disconnected = False Then 'Are we disconnected??
        If Level = "homesystem" Then
            Select Case Level 'Where are we
                Case "homesystem" 'If they are in the C:\System\ folder
                    frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & "del view.exe" & vbCrLf
                    frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Access is Denied" & vbCrLf
                    frmGame.txtConsole.Text = frmGame.txtConsole.Text & "View.exe is a System File and Cannot be Deleted" & vbCrLf
                    frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & vbCrLf
                    Exit Sub 'Exits the sub
            End Select
            frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & "del view.exe" & vbCrLf
            frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Could not Find " & Lvl & "View.exe" & vbCrLf
            frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & vbCrLf
            Exit Sub 'Exits the Sub :)
        End If
        
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & "del view.exe" & vbCrLf
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Could not Find " & Lvl & "View.exe" & vbCrLf
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & vbCrLf
        Exit Sub
        
    Else 'The Computer is disconnected
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & "Computer is Disconnected" & vbCrLf
    End If
    Exit Sub 'Exits the Sub :)
End Sub


'They typed del and something that isn't there. e.g.(del gklfj)
Public Sub DelSomething()
    If Disconnected = False Then
        cGetCharCount = Len(Text) 'Here is where we count the number of characters
        'This will tell up what they typed after del
        'Mid(What we want searched, The Starting Character, Then the total Characters/The Last Char)
        Result = Mid(Text, 5, cGetCharCount)

        frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & "del " & Result & vbCrLf
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Could not Find " & Lvl & Result & vbCrLf
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & vbCrLf
        Exit Sub 'Exits the Sub :)
    Else 'The Computer is disconnected
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & "Computer is Disconnected" & vbCrLf
    End If
    Exit Sub 'Exits the Sub :)
End Sub


'They want to Delete C:\System\Deltree.exe
Public Sub DelSysDeltree()
    GetLevel 'Where are they?
    If Disconnected = False Then
        Select Case Level
            Case "homesystem" 'If they are in the C:\System\ folder
                frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & "del deltree.exe" & vbCrLf
                frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Access is Denied" & vbCrLf
                frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Deltree.exe is a System File and Cannot be Deleted" & vbCrLf
                frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & vbCrLf
                Exit Sub
        End Select
                
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & "del deltree.exe" & vbCrLf
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Could not Find " & Lvl & "Deltree.exe" & vbCrLf
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & vbCrLf
        Exit Sub
                
    Else 'The Computer is disconnected
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & "Computer is Disconnected" & vbCrLf
    End If
    Exit Sub 'Exits the Sub :)
End Sub


'They want to Delete C:\System\Del.exe
Public Sub DelSysDel()
    GetLevel 'Where are they?
    If Disconnected = False Then 'Are they disconnected?
        Select Case Level 'Where are they
            Case "homesystem" 'If they are in the C:\System\ folder
                frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & "del del.exe" & vbCrLf
                frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Access is Denied" & vbCrLf
                frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Del.exe is a System File and Cannot be Deleted" & vbCrLf
                frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & vbCrLf
                Exit Sub
                
        End Select
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & "del del.exe" & vbCrLf
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Could not Find " & Lvl & "Del.exe" & vbCrLf
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & vbCrLf
        Exit Sub
                
    Else 'The Computer is disconnected
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & "Computer is Disconnected" & vbCrLf
    End If
    Exit Sub 'Exits the Sub :)
End Sub


'They want to Delete C:\System\Commands.exe
Public Sub DelSysCommands()
    GetLevel 'Where are they?
    If Disconnected = False Then 'Is the Computer Disconnected
        Select Case Level 'Where are They?
            Case "homesystem" 'If they are in the C:\System\ folder
                frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & "del commands.exe" & vbCrLf
                frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Access is Denied" & vbCrLf
                frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Commands.exe is a System File and Cannot be Deleted" & vbCrLf
                frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & vbCrLf
                Exit Sub
                    
        End Select
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & "del commands.exe" & vbCrLf
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Could not Find " & Lvl & "Commands.exe" & vbCrLf
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & vbCrLf
        Exit Sub
                
    Else 'The Computer is disconnected
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & "Computer is Disconnected" & vbCrLf
    End If
    Exit Sub 'Exits the Sub :)
End Sub


'They want to Delete C:\System\Cls.exe
Public Sub DelSysCls()
    GetLevel 'Where are they?
    If Disconnected = False Then
        Select Case Level
            Case "homesystem" 'If they are in the C:\System\ folder
                frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & "del cls.exe" & vbCrLf
                frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Access is Denied" & vbCrLf
                frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Cls.exe is a System File and Cannot be Deleted" & vbCrLf
                frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & vbCrLf
                Exit Sub
                    
        End Select
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & "del cls.exe" & vbCrLf
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Could not Find " & Lvl & "Cls.exe" & vbCrLf
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & vbCrLf
        Exit Sub

    Else 'The Computer is disconnected
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & "Computer is Disconnected" & vbCrLf
    End If
    Exit Sub 'Exits the Sub :)
End Sub


'They want to Delete C:\System\Boot\Boot.ini
Public Sub DelSysbootboot()
    GetLevel 'Where are they?
    If Disconnected = False Then
        Select Case Level
            Case "homesysboot" 'If they are in the C:\System\Boot folder
                frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & "del boot.ini" & vbCrLf
                frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Access is Denied" & vbCrLf
                frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Boot.ini is a System File and Cannot be Deleted" & vbCrLf
                frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & vbCrLf
                Exit Sub
                    
        End Select
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & "del boot.ini" & vbCrLf
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Could not Find " & Lvl & "Boot.ini" & vbCrLf
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & vbCrLf
        Exit Sub

    Else 'The Computer is disconnected
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & "Computer is Disconnected" & vbCrLf
    End If
    Exit Sub 'Exits the Sub :)
End Sub


'They want to Delete C:\System\Boot\User.dat
Public Sub DelSysbootuser()
    GetLevel 'Where are they?
    If Disconnected = False Then
        Select Case Level
            Case "homesysboot" 'If they are in the C:\System\Boot folder
                frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & "del user.dat" & vbCrLf
                frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Access is Denied" & vbCrLf
                frmGame.txtConsole.Text = frmGame.txtConsole.Text & "User.dat is a System File and Cannot be Deleted" & vbCrLf
                frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & vbCrLf
                Exit Sub
                    
        End Select
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & "del user.dat" & vbCrLf
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Could not Find " & Lvl & "User.dat" & vbCrLf
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & vbCrLf
        Exit Sub

    Else 'The Computer is disconnected
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & "Computer is Disconnected" & vbCrLf
    End If
    Exit Sub 'Exits the Sub :)
End Sub

'They want to Delete C:\System\Boot\System.dat
Public Sub DelSysbootsystem()
    GetLevel 'Where are they?
    If Disconnected = False Then
        Select Case Level
            Case "homesysboot" 'If they are in the C:\System\Boot folder
                frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & "del system.dat" & vbCrLf
                frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Access is Denied" & vbCrLf
                frmGame.txtConsole.Text = frmGame.txtConsole.Text & "System.dat is a System File and Cannot be Deleted" & vbCrLf
                frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & vbCrLf
                Exit Sub
                    
        End Select
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & "del system.dat" & vbCrLf
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Could not Find " & Lvl & "System.dat" & vbCrLf
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & vbCrLf
        Exit Sub

    Else 'The Computer is disconnected
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & "Computer is Disconnected" & vbCrLf
    End If
    Exit Sub 'Exits the Sub :)
End Sub


'They want to Delete C:\System\Kernel\Kernel.sys
Public Sub DelSyskernelkernel()
    GetLevel 'Where are they?
    If Disconnected = False Then
        Select Case Level
            Case "homesyskernel" 'If they are in the C:\System\Kernel folder
                frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & "del kernel.sys" & vbCrLf
                frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Access is Denied" & vbCrLf
                frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Kernel.sys is a System File and Cannot be Deleted" & vbCrLf
                frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & vbCrLf
                Exit Sub
                    
        End Select
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & "del kernel.sys" & vbCrLf
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Could not Find " & Lvl & "Kernel.sys" & vbCrLf
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & vbCrLf
        Exit Sub

    Else 'The Computer is disconnected
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & "Computer is Disconnected" & vbCrLf
    End If
    Exit Sub 'Exits the Sub :)
End Sub


'They want to Delete C:\Help\Help.hlp
Public Sub DelHomeHelpHelp()
    GetLevel 'Where are they?
    If Disconnected = False Then
        Select Case Level
            Case "homehelp" 'If they are in the C:\Help\ folder
                frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & "del help.hlp" & vbCrLf
                frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Access is Denied" & vbCrLf
                frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Help.hlp is a System File and Cannot be Deleted" & vbCrLf
                frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & vbCrLf
                Exit Sub
                    
        End Select
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & "del help.hlp" & vbCrLf
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Could not Find " & Lvl & "Help.hlp" & vbCrLf
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & vbCrLf
        Exit Sub

    Else 'The Computer is disconnected
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & "Computer is Disconnected" & vbCrLf
    End If
    Exit Sub 'Exits the Sub :)
End Sub
