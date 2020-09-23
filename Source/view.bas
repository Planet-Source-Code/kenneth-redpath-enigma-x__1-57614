Attribute VB_Name = "view"
Option Explicit 'Makes sure all the variables are declared

'**********************************************************************************
'This module contains all the code for the view command. ie "view readme.txt" etc.
'I put it in here to clean up the code on the frmGame a bit :-)
'**********************************************************************************


'They want to View the "C:\Documents\Images\Test.jpg" file
Public Sub ViewTestjpg()
    GetLevel 'Where are we? Loads the GetLevel Sub in (main.bas)
    If Disconnected = False Then 'Are we disconnected?
        If Level = "homedocimages" Then 'Are we at C:\Documents\Images\?
            If IsHomeDocImgTestDel = False Then 'Has it been Deleted? :[
                frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & "view Test.jpg" & vbCrLf
                frmView.imgImage.Visible = True 'Shows the Picture box -- Where we will show the picture
                frmView.Caption = "View.exe - Test.jpg" 'Changes the caption of the View.exe form to "View.exe - Test.jpg"
                frmView.imgImage.Picture = LoadResPicture(101, 0) 'Loads the Test.jpg picture from the Resource File
                frmView.Width = 2340 'Changes the Width of the View.exe form so it is the same width as the Picture
                frmView.Height = 2610 'Changes the Height of the View.exe form so it is the same height as the Picture
                frmView.Show vbModal, frmGame 'Shows the form once the width, height and other aspects are set-up
                frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & vbCrLf
            Else 'Yes it has been deleted
                frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & "view Test.jpg" & vbCrLf
                frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Could not Find " & Lvl & "Test.jpg" & vbCrLf
                frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & vbCrLf
            End If 'Has it been Deleted
            Exit Sub 'Ends the Sub :]
        End If 'Are we at C:\Documents\Recieved\?
        
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & "view test.jpg" & vbCrLf
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Could not Find " & Lvl & "Test.jpg" & vbCrLf
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & vbCrLf
        Exit Sub
        
    Else 'We have been disconnected
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & "Computer is Disconnected" & vbCrLf
    End If
    Exit Sub 'Ends the Current Sub
End Sub


'They want to View the "C:\Readme.txt" file
Public Sub ViewHomeReadmetxt()
    GetLevel 'Where are we? Loads the GetLevel Sub in (main.bas)
    If Disconnected = False Then 'Are we disconnected?
        If Level = "home" Then
            If IsHomeReadmeDel = False Then 'Has the file been deleted? :-0
                frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & "view readme.txt" & vbCrLf 'This shows view readme.txt in the console (more of a perfection thing) :-)
                frmView.imgImage.Visible = False 'Hides the picture box as it isn't a picture
                frmView.txtView.Visible = True  'Shows the Text box -- Where we will show the readme file
                frmView.Height = 2160 'Changes the Height of the View.exe form so it is just bigger than the text box
                frmView.Width = 3210 'Changes the Width of the View.exe form so it is just bigger than the text box
                frmView.txtView.Height = 1815 'Changes the Height of the Text box to just over the contents' height.
                frmView.txtView.Width = 3330 'Changes the Width of the Text box to just over the contents' width.
                frmView.txtView.Left = 0 'This makes the text box appear at the left of the screen
                frmView.txtView.Top = 0 'This makes the text box appear at the top of the screen
                frmView.Caption = "View.exe - Readme.txt" 'This just changes the caption of the form
                frmView.txtView.Text = "" 'Clears the text box ready for the readme file
                            
                'Now we enter the text - Use vbCrLf for a new line(very useful) :-)
                frmView.txtView.Text = "This is the Read-Me File for Enigma X" & vbCrLf & vbCrLf & _
                                       "If you Have and questions about this game" & vbCrLf & _
                                       "or have any ideas for improving this game" & vbCrLf & _
                                       "email me @ kjredpath@dodo.com.au" & vbCrLf & vbCrLf & _
                                       "If you need help with the commands" & vbCrLf & _
                                       "type command or press F1"
    
                frmView.Show vbModal, frmGame 'Shows the View.exe form, once all the things are set up
                frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & vbCrLf
    
            Else 'Oh No It has been deleted :-O
                frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & "view readme.txt" & vbCrLf
                frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Could not Find " & Lvl & "Readme.txt" & vbCrLf
                frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & vbCrLf
            End If
            Exit Sub 'Ends the Sub :)
        End If
        
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & "view readme.txt" & vbCrLf
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Could not Find " & Lvl & "Readme.txt" & vbCrLf
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & vbCrLf
        Exit Sub

    Else 'We have been disconnected
        'This just shows the User that we are disconnected
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & "Computer is Disconnected" & vbCrLf
    End If

    Exit Sub 'Ends the Current Sub :-(
End Sub


'They want to View the "C:\Documents\Recieved\Readme.txt" file
Public Sub ViewDocRecReadmetxt()
    GetLevel
    If Disconnected = False Then 'Are we disconnected
        If Level = "homerecieved" Then
            If IsHomeDocRecReadmeDel = False Then 'Has the file been deleted? :-0
                frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & "view readme.txt" & vbCrLf 'This shows view readme.txt in the console (more of a perfection thing) :-)
                frmView.imgImage.Visible = False 'Hides the picture box as it isn't a picture
                frmView.txtView.Visible = True  'Shows the Text box -- Where we will show the readme file
                frmView.Height = 2160 'Changes the Height of the View.exe form so it is just bigger than the text box
                frmView.Width = 3585 'Changes the Width of the View.exe form so it is just bigger than the text box
                frmView.txtView.Height = 1815 'Changes the Height of the Text box to just over the contents' height.
                frmView.txtView.Width = 3495 'Changes the Width of the Text box to just over the contents' width.
                frmView.txtView.Left = 0 'This makes the text box appear at the left of the screen
                frmView.txtView.Top = 0 'This makes the text box appear at the top of the screen
                frmView.Caption = "View.exe - Readme.txt" 'This just changes the caption of the form
                frmView.txtView.Text = "" 'Clears the text box ready for the readme file
                                
                'Now we enter the text - Use vbCrLf for a new line(very useful) :-)
                frmView.txtView.Text = "This folder is where all the documents that are" & vbCrLf & _
                                        "recieved will go. You will get documents from" & vbCrLf & _
                                        "employers etc. These may contain useful" & vbCrLf & _
                                        "information so remember to check frequently." & vbCrLf & vbCrLf & _
                                        "If you accept a mission a message will pop-up" & vbCrLf & _
                                        "or will show up in the Console saying that" & vbCrLf & _
                                        "there is a new message to be read."
        
                frmView.Show vbModal, frmGame 'Shows the View.exe form, once all the things are set up
                frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & vbCrLf
            Else 'Oh No It has been deleted :-)
                frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & "view readme.txt" & vbCrLf
                frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Could not Find " & Lvl & "Readme.txt" & vbCrLf
                frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & vbCrLf
            End If
            Exit Sub 'Ends the Sub :)
        End If
        
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & "view readme.txt" & vbCrLf
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Could not Find " & Lvl & "Readme.txt" & vbCrLf
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & vbCrLf
        Exit Sub

    Else 'The Computer has been disconnected
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & "Computer is Disconnected" & vbCrLf
    End If
End Sub


'They want to View the "C:\Help\Help.hlp" file
Public Sub ViewHelpHelp()
    GetLevel
    If Disconnected = False Then 'Are we disconnected
        If Level = "homehelp" Then
            frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & "view help.hlp" & vbCrLf 'This shows view help.hlp in the console (more of a perfection thing) :-)
            frmHelp.Show vbModal, frmGame
            frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & vbCrLf
            Exit Sub 'Ends the Sub :)
        End If
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & "view help.hlp" & vbCrLf
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Could not Find " & Lvl & "Help.hlp" & vbCrLf
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & vbCrLf
        Exit Sub
    Else 'The Computer has been disconnected
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & "Computer is Disconnected" & vbCrLf
    End If
End Sub


'They want to View the "C:\System\Boot\Boot.ini" file
Public Sub ViewBootini()
    GetLevel
    If Disconnected = False Then 'Are we disconnected
        If Level = "homesysboot" Then
            frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & "view boot.ini" & vbCrLf 'This shows view help.hlp in the console (more of a perfection thing) :-)
            frmGame.txtConsole.Text = frmGame.txtConsole.Text & "You cannot view this file." & vbCrLf
            frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & vbCrLf
            Exit Sub 'Ends the Sub :)
        End If
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & "view boot.ini" & vbCrLf
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Could not Find " & Lvl & "Boot.ini" & vbCrLf
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & vbCrLf
        Exit Sub
    Else 'The Computer has been disconnected
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & "Computer is Disconnected" & vbCrLf
    End If
End Sub


'They want to View the "C:\System\Boot\User.dat" file
Public Sub ViewSysBootUser()
    GetLevel
    If Disconnected = False Then 'Are we disconnected
        If Level = "homesysboot" Then
            frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & "view user.dat" & vbCrLf 'This shows view help.hlp in the console (more of a perfection thing) :-)
            frmGame.txtConsole.Text = frmGame.txtConsole.Text & "You cannot view this file." & vbCrLf
            frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & vbCrLf
            Exit Sub 'Ends the Sub :)
        End If
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & "view user.dat" & vbCrLf
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Could not Find " & Lvl & "User.dat" & vbCrLf
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & vbCrLf
        Exit Sub
    Else 'The Computer has been disconnected
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & "Computer is Disconnected" & vbCrLf
    End If
End Sub


'They want to View the "C:\System\Boot\System.dat" file
Public Sub ViewSysBootSystem()
    GetLevel
    If Disconnected = False Then 'Are we disconnected
        If Level = "homesysboot" Then
            frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & "view system.dat" & vbCrLf 'This shows view help.hlp in the console (more of a perfection thing) :-)
            frmGame.txtConsole.Text = frmGame.txtConsole.Text & "You cannot view this file." & vbCrLf
            frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & vbCrLf
            Exit Sub 'Ends the Sub :)
        End If
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & "view system.dat" & vbCrLf
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Could not Find " & Lvl & "System.dat" & vbCrLf
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & vbCrLf
        Exit Sub
    Else 'The Computer has been disconnected
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & "Computer is Disconnected" & vbCrLf
    End If
End Sub


'They want to View the "C:\System\Cls.exe" file
Public Sub ViewSysCls()
    GetLevel
    If Disconnected = False Then 'Are we disconnected
        If Level = "homesystem" Then
            frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & "view cls.exe" & vbCrLf 'This shows view help.hlp in the console (more of a perfection thing) :-)
            frmGame.txtConsole.Text = frmGame.txtConsole.Text & "You cannot view this file." & vbCrLf
            frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & vbCrLf
            Exit Sub 'Ends the Sub :)
        End If
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & "view cls.exe" & vbCrLf
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Could not Find " & Lvl & "Cls.exe" & vbCrLf
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & vbCrLf
        Exit Sub
    Else 'The Computer has been disconnected
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & "Computer is Disconnected" & vbCrLf
    End If
End Sub


'They want to View the "C:\System\Commands.exe" file
Public Sub ViewSysCommands()
    GetLevel
    If Disconnected = False Then 'Are we disconnected
        If Level = "homesystem" Then
            frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & "view commands.exe" & vbCrLf 'This shows view help.hlp in the console (more of a perfection thing) :-)
            frmGame.txtConsole.Text = frmGame.txtConsole.Text & "You cannot view this file." & vbCrLf
            frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & vbCrLf
            Exit Sub 'Ends the Sub :)
        End If
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & "view commands.exe" & vbCrLf
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Could not Find " & Lvl & "Commands.exe" & vbCrLf
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & vbCrLf
        Exit Sub
    Else 'The Computer has been disconnected
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & "Computer is Disconnected" & vbCrLf
    End If
End Sub


'They want to View the "C:\System\Del.exe" file
Public Sub ViewSysDel()
    GetLevel
    If Disconnected = False Then 'Are we disconnected
        If Level = "homesystem" Then
            frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & "view del.exe" & vbCrLf 'This shows view help.hlp in the console (more of a perfection thing) :-)
            frmGame.txtConsole.Text = frmGame.txtConsole.Text & "You cannot view this file." & vbCrLf
            frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & vbCrLf
            Exit Sub 'Ends the Sub :)
        End If
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & "view del.exe" & vbCrLf
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Could not Find " & Lvl & "Del.exe" & vbCrLf
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & vbCrLf
        Exit Sub
    Else 'The Computer has been disconnected
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & "Computer is Disconnected" & vbCrLf
    End If
End Sub


'They want to View the "C:\System\Deltree.exe" file
Public Sub ViewSysDeltree()
    GetLevel
    If Disconnected = False Then 'Are we disconnected
        If Level = "homesystem" Then
            frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & "view deltree.exe" & vbCrLf 'This shows view help.hlp in the console (more of a perfection thing) :-)
            frmGame.txtConsole.Text = frmGame.txtConsole.Text & "You cannot view this file." & vbCrLf
            frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & vbCrLf
            Exit Sub 'Ends the Sub :)
        End If
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & "view deltree.exe" & vbCrLf
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Could not Find " & Lvl & "Deltree.exe" & vbCrLf
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & vbCrLf
        Exit Sub
    Else 'The Computer has been disconnected
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & "Computer is Disconnected" & vbCrLf
    End If
End Sub


'They want to View the "C:\System\View.exe" file
Public Sub ViewSysView()
    GetLevel
    If Disconnected = False Then 'Are we disconnected
        If Level = "homesystem" Then
            frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & "view view.exe" & vbCrLf 'This shows view help.hlp in the console (more of a perfection thing) :-)
            frmGame.txtConsole.Text = frmGame.txtConsole.Text & "You cannot view this file." & vbCrLf
            frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & vbCrLf
            Exit Sub 'Ends the Sub :)
        End If
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & "view view.exe" & vbCrLf
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Could not Find " & Lvl & "View.exe" & vbCrLf
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & vbCrLf
        Exit Sub
    Else 'The Computer has been disconnected
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & "Computer is Disconnected" & vbCrLf
    End If
End Sub


'They want to View the "C:\System\Kernel\Kernel.sys" file
Public Sub ViewSysKerKernel()
    GetLevel
    If Disconnected = False Then 'Are we disconnected
        If Level = "homesyskernel" Then
            frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & "view kernel.sys" & vbCrLf 'This shows view help.hlp in the console (more of a perfection thing) :-)
            frmGame.txtConsole.Text = frmGame.txtConsole.Text & "You cannot view this file." & vbCrLf
            frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & vbCrLf
            Exit Sub 'Ends the Sub :)
        End If
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & "view kernel.sys" & vbCrLf
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Could not Find " & Lvl & "Kernel.sys" & vbCrLf
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & vbCrLf
        Exit Sub
    Else 'The Computer has been disconnected
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & "Computer is Disconnected" & vbCrLf
    End If
End Sub
