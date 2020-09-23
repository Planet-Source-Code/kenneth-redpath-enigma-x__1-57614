Attribute VB_Name = "conn"
Option Explicit 'Makes sure all the variables are declared

'**********************************************************************************
'This module contains all the code for the conn command. ie "conn home" etc.
'I put it in here to clean up the code on the frmGame a bit :-)
'**********************************************************************************


'This is connecting to the Home Computer
Public Sub ConHomeComputer()

    Connecting = True 'This tells the program that it is connecting. So _
                       if the user enters code it does nothing
    Disconnected = False 'This tells the program that it is not disconnected
    
    Level = "home" 'This sets the Level to the home directory (C:\)
    GetLevel
    Server = "Server X10 - Home Computer" 'Sets the server / Maybe used in missions ??
    
    frmGame.txtConsole.SelText = "Connecting to Home Computer ."
    
    'Checks to see if they want to play the dial.mp3 file.
    If PlayFile = False Then
    Else
        'Plays the dial.mp3 file in the Sounds directory
        PlayFile = FileExists(App.Path & "\Sounds\dial.mp3")
        If PlayFile = True Then
            frmGame.PlayMp3.FileName = App.Path & "\Sounds\dial.mp3"
            frmGame.PlayMp3.Play
            Pause (5)
        End If
    End If
    
    Pause (0.5)
    frmGame.txtConsole.SelText = "."
    Pause (0.5)
    frmGame.txtConsole.SelText = "."
    Pause (0.5)
    frmGame.txtConsole.SelText = "." & vbCrLf
    frmGame.txtConsole.Text = ""
    frmGame.txtConsole.SelText = "Connected to Home Computer"
    Pause (0.5)
    
    frmGame.txtConsole.Text = ""
    frmGame.txtConsole.Text = frmGame.txtConsole.Text & "************************Login************************" & vbCrLf & vbCrLf
    frmGame.txtConsole.Text = frmGame.txtConsole.Text & "User Name: " & s_UserName & vbCrLf
    
    'This Looks Cool :-) Basically it counts the number of characters in the password, then enters that as (*)'s
    frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Password: "
    Dim TempNum
    TempNum = 0 'Set it so it starts at zero
    While TempNum < Len(s_Password) 'Keep adding the (*) until the Variable is the same size as the Password
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & "*"
        Pause (0.09) 'Pause for effect
        TempNum = TempNum + 1 'This makes sure that it keeps going
    Wend
    'Pause (0.5) 'Another Pause for Effect
    
    frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & vbCrLf & "*****************************************************" & vbCrLf
    frmGame.txtConsole.Text = frmGame.txtConsole.Text & "."
    Pause (0.5)
    frmGame.txtConsole.Text = frmGame.txtConsole.Text & "."
    Pause (0.5)
    frmGame.txtConsole.Text = frmGame.txtConsole.Text & "."
    Pause (0.5)
    frmGame.txtConsole.Text = frmGame.txtConsole.Text & "."
    Pause (0.5)
    frmGame.txtConsole.Text = frmGame.txtConsole.Text & "." & vbCrLf
    frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Verification Accepted." & vbCrLf
    Pause (0.5)
    
    If NewServer = True Then
        SetupNewServer 'This loads SetupNewServer from (main.bas)
    End If
    
    
    frmGame.txtConsole.Text = "" 'Clear the Text in the console
    frmGame.txtConsole.SelText = "Welcome " & s_UserName & " To Server X10 (Home Computer)" & vbCrLf & vbCrLf
    
    'Plays the welcome.mp3 file in the Sounds directory
    PlayFile = FileExists(App.Path & "\Sounds\welhome.mp3")
    If PlayFile = True Then
        frmGame.PlayMp3.FileName = App.Path & "\Sounds\welhome.mp3"
        frmGame.PlayMp3.Play
    End If
    
    frmGame.txtConsole.Text = frmGame.txtConsole.Text & "From here you can access your Computer's Files etc." & vbCrLf
    frmGame.txtConsole.Text = frmGame.txtConsole.Text & "If you Need Help Press F1 or type Help." & vbCrLf & vbCrLf

    frmGame.txtConsole.Text = frmGame.txtConsole.Text & "The commands are:" & vbCrLf
    frmGame.txtConsole.Text = frmGame.txtConsole.Text & "dir - To see what is in a directory" & vbCrLf
    frmGame.txtConsole.Text = frmGame.txtConsole.Text & "cd - To Change directory i.e. (cd home)" & vbCrLf
    frmGame.txtConsole.Text = frmGame.txtConsole.Text & "cls - Clears the Screen" & vbCrLf
    frmGame.txtConsole.Text = frmGame.txtConsole.Text & "del - Deletes a file you specify i.e. (del home)" & vbCrLf
    frmGame.txtConsole.Text = frmGame.txtConsole.Text & "view - Viewes a file you specify i.e. (view file.jpg)" & vbCrLf
    frmGame.txtConsole.Text = frmGame.txtConsole.Text & "deltree - Removes a whole directory and its contents." & vbCrLf
    frmGame.txtConsole.Text = frmGame.txtConsole.Text & "disc - Disconnects from the current computer" & vbCrLf
    frmGame.txtConsole.Text = frmGame.txtConsole.Text & "conn - Connects to a Computer IP (i.e. home, 312.324.1.5)" & vbCrLf
    frmGame.txtConsole.Text = frmGame.txtConsole.Text & "For a Full List of Commands see the Help File" & vbCrLf
    frmGame.txtConsole.Text = frmGame.txtConsole.Text & "or Type commands" & vbCrLf & vbCrLf
    
    frmGame.txtConsole.Text = frmGame.txtConsole.Text & Lvl & vbCrLf
    
    Connecting = False 'We have finished connecting so allow user input
    NewServer = False 'Because this server has been set-up (it isn't new anymore)
End Sub

