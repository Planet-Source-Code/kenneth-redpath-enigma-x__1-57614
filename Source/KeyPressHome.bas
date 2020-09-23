Attribute VB_Name = "KeyPressHome"
'This will store all the stuff to do when something is typed then pressed _
    Enter. Only at Home Server

Public Sub KeysHome()
    'This is so we can determine if the user has typed something that isn't there like (del fjkdhsjkl;df or view kdfh.but)
    lCaseText = LCase$(Text) 'Makes what the user has typed lower case
    cGetCharCount = Len(lCaseText) 'Here is where we count the number of characters
    'Mid(What we want searched, The Starting Character, Then the total Characters/The Last Char)
    Result = Mid(lCaseText, 5, cGetCharCount) 'This is for three letter programs. (del)
    Result1 = Mid(lCaseText, 4, cGetCharCount) 'This is for two letter programs. (cd)
    Result2 = Mid(lCaseText, 6, cGetCharCount) 'This is for four letter programs. (view)
    Result3 = Mid(lCaseText, 9, cGetCharCount) 'This is for seven letter programs. (deltree)

    Select Case LCase$(Text) 'Changes what the user has typed and makes it lower case
    
        Case "cd" ' Change Directory with no directory file on end
            GetLevel 'Where are we? Loads the GetLevel Sub in (main.bas)
            If Disconnected = False Then 'If we arnen't disconnected
                frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & "cd" & vbCrLf
                frmGame.txtConsole.Text = frmGame.txtConsole.Text & "You need to add a folder or drive to the end. i.e. (cd home)" & vbCrLf
                Exit Sub
            Else
                frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & "Computer is Disconnected" & vbCrLf
            End If
            Exit Sub 'Ends the Current Sub
            
         'Change directory to the Documents Folder
        Case "cd documents"
            CdDocuments 'Calls the CdDocuments Sub from (cd.bas)
            Exit Sub
            
         'Change directory to the Documents\Recieved Folder
        Case "cd recieved"
            CdRecieved 'Calls the CdRecieved Sub from (cd.bas)
            Exit Sub
        
         'Change directory to the Documents\Images Folder
        Case "cd images"
            CdImages 'Calls the CdImages Sub from (cd.bas)
            Exit Sub
            
        'Change directory to the Downloads Folder
        Case "cd downloads"
            CdDownloads 'Calls the CdDownloads Sub from (cd.bas)
            Exit Sub
            
        'Change directory to the Software Folder
        Case "cd software"
            CdSoftware 'Calls the CdSoftware Sub from (cd.bas)
            Exit Sub
            
        'Change directory to the System Folder
        Case "cd system"
            CdSystem 'Calls the CdSystem Sub from (cd.bas)
            Exit Sub
            
        'Change directory to the Boot Folder in the System Folder
        Case "cd boot"
            CdSysBoot 'Calls the CdSysBoot Sub from (cd.bas)
            Exit Sub
            
        'Change directory to the Kernel Folder in the System Folder
        Case "cd kernel"
            CdSysKernel 'Calls the CdSysKernel Sub from (cd.bas)
            Exit Sub
            
        'Change directory to the Help Folder
        Case "cd help"
            CdHomeHelp 'Calls the CdHomeHelp Sub from (cd.bas)
            Exit Sub
            
        Case "cd .."
            GetLevel 'Where are we? Loads the GetLevel Sub in (main.bas)
            If Disconnected = False Then
            Select Case Level
                Case "home" 'We are at home (C:\)
                    frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & "cd .." & vbCrLf
                    frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & "C:\" & vbCrLf
                    Level = "home"
                    Exit Sub
                
                Case "documents" 'We are in the Documents Folder (C:\Documents)"
                    frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & "cd .." & vbCrLf
                    frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & "C:\" & vbCrLf
                    Level = "home"
                    Exit Sub
                    
                Case "homerecieved" 'We are in the Recieved folder under my docs (C:\Documents\Recieved)
                    frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & "cd .." & vbCrLf
                    Level = "documents"
                    GetLevel
                    frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & vbCrLf
                    Exit Sub
                    
                Case "homedocimages" 'We are in the Images Folder in the Documents Folder At Home
                    frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & "cd .." & vbCrLf
                    Level = "documents"
                    GetLevel
                    frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & vbCrLf
                    Exit Sub
                
                Case "homedownloads" 'We are in the Downloads Folder At Home
                    frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & "cd .." & vbCrLf
                    Level = "home"
                    GetLevel
                    frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & vbCrLf
                    Exit Sub
                    
                Case "homesoftware" 'We are in the Software Folder At Home
                    frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & "cd .." & vbCrLf
                    Level = "home"
                    GetLevel
                    frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & vbCrLf
                    Exit Sub
                    
                Case "homesystem" 'We are in the Software Folder At Home
                    frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & "cd .." & vbCrLf
                    Level = "home"
                    GetLevel
                    frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & vbCrLf
                    Exit Sub
                    
                Case "homesysboot" 'We are in the Boot Folder in the System folder At Home
                    frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & "cd .." & vbCrLf
                    Level = "homesystem"
                    GetLevel
                    frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & vbCrLf
                    Exit Sub
                    
                Case "homesyskernel" 'We are in the Kernel Folder in the System folder At Home
                    frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & "cd .." & vbCrLf
                    Level = "homesystem"
                    GetLevel
                    frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & vbCrLf
                    Exit Sub
                    
                Case "homehelp" 'We are in the Help Folder At Home
                    frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & "cd .." & vbCrLf
                    Level = "home"
                    GetLevel
                    frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & vbCrLf
                    Exit Sub
                    
                    
                    
            End Select 'Level Select
            
            Else
                frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & "Computer is Disconnected" & vbCrLf
            End If
            
            Exit Sub

        Case "dir" 'They want to view the current Directory's contents
        GetLevel 'Where are we? Loads the GetLevel Sub in (main.bas)
            If Disconnected = False Then
            Select Case LCase$(Level) 'Where are they??
                
                'They want to view the contents of C:\ drive on the Home Computer
                Case "home" 'If they are at the home Directory (C:\)
                    DirHome 'Loads the DirHome sub from (dir.bas)
                    Exit Sub 'Exits the Sub

                'They want to view the contents of the Folder C:\Documents\ on the Home Computer
                Case "documents" 'If they are at the documents Directory (C:\Documents)
                    DirHomeDocuments 'Loads the DirHomeDocuments sub from (dir.bas)
                    Exit Sub 'Exits the Sub

                'They want to view the contents of the Folder C:\Documents\Recieved\ on the Home Computer
                Case "homerecieved"
                    DirHomeDocRecieved 'Loads the DirHomeDocRecieved sub from (dir.bas)
                    Exit Sub 'Exits the Sub Duh! :-)

                'They want to view the contents of the Folder C:\Documents\Images\ on the Home Computer
                Case "homedocimages"
                    DirHomeDocImages 'Loads the DirHomeDocImages sub from (dir.bas)
                    Exit Sub 'I WONDER WHAT THIS DOES??????? :-)
                
                'They want to view the contents of the Folder C:\Downloads\ on the Home Computer
                Case "homedownloads"
                    DirHomeDownloads 'Loads the DirHomeDownloads sub from (dir.bas)
                    Exit Sub 'Exits the Sub
                    
                'They want to view the contents of the Folder C:\System\ on the Home Computer
                Case "homesystem"
                    DirHomeSystem 'Loads the DirHomeSystem sub from (dir.bas)
                    Exit Sub 'Exits the Sub
                    
                'They want to view the contents of the Folder C:\System\ on the Home Computer
                Case "homesoftware"
                    DirHomeSoftware 'Loads the DirHomeSoftware sub from (dir.bas)
                    Exit Sub 'Exits the Sub
                    
                'They want to view the contents of the Folder C:\System\Boot\ on the Home Computer
                Case "homesysboot"
                    DirHomeSysBoot 'Loads the DirHomeBoot sub from (dir.bas)
                    Exit Sub 'Exits the Sub
                    
                'They want to view the contents of the Folder C:\System\Kernel\ on the Home Computer
                Case "homesyskernel"
                    DirHomeSysKernel 'Loads the DirHomeSysKernel sub from (dir.bas)
                    Exit Sub 'Exits the Sub
                
                'They want to view the contents of the Folder C:\Help\ on the Home Computer
                Case "homehelp"
                    DirHomeHelp 'Loads the DirHomeHelp sub from (dir.bas)
                    Exit Sub
                    

            End Select
            
            Else 'It has been disconnected
                frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & "Computer is Disconnected" & vbCrLf
            End If
            
            Exit Sub 'Ends the Current Sub :-0
            
        Case "view readme.txt" 'They want to view the Readme.txt File
            GetLevel 'Where are we? Loads the GetLevel Sub in (main.bas)
            If Disconnected = False Then 'Are we disconnected?
                Select Case Level
            
                    'View the readme.txt file in C:\ at the Home Computer
                    Case "home" 'Its at Home C:\
                        ViewHomeReadmetxt 'Calls the ViewHomeReadmetxt from (view.bas)
                        Exit Sub 'Exits the Sub
                    
                    'View the readme.txt file in C:\Documents\Recieved\ at the Home Computer
                    Case "homerecieved" 'Its at Home C:\Documents\Recieved
                        ViewDocRecReadmetxt 'Calls the ViewDocRecReadmetxt from (view.bas)
                        Exit Sub 'Ends the Current Sub :-)
                    
                End Select 'Level Select
                
                frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & "view readme.txt" & vbCrLf
                frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Could not Find " & Lvl & "Readme.txt" & vbCrLf
                frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & vbCrLf
                Exit Sub 'Ends the Sub (as if you didn't know already) :-)
                
            Else 'The Computer is disconnected
                frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & "Computer is Disconnected" & vbCrLf
            End If
            
        'View the test image in the Images folder on the Home Computer (C:\Documents\Images\)
        Case "view test.jpg"
            ViewTestjpg 'Calls the ViewTestjpg from (view.bas)
            Exit Sub 'Exits the Sub
            
        'View the Boot.ini file in the Boot folder in System on the Home Computer (C:\System\Boot\)
        Case "view boot.ini"
            ViewBootini 'Calls the ViewBootini from (view.bas)
            Exit Sub 'Exits the Sub
            
        'View the User.dat file in the Boot folder in System on the Home Computer (C:\System\Boot\)
        Case "view user.dat"
            ViewSysBootUser 'Calls the ViewSysBootUser from (view.bas)
            Exit Sub 'Exits the Sub
            
        'View the System.dat file in the Boot folder in System on the Home Computer (C:\System\Boot\)
        Case "view system.dat"
            ViewSysBootSystem 'Calls the ViewSysBootSystem from (view.bas)
            Exit Sub 'Exits the Sub
        
        'View the Cls.exe file in the System folder on the Home Computer (C:\System\)
        Case "view cls.exe"
            ViewSysCls 'Calls the ViewSysCls from (view.bas)
            Exit Sub 'Exits the Sub
            
        'View the Commands.exe file in the System folder on the Home Computer (C:\System\)
        Case "view commands.exe"
            ViewSysCommands 'Calls the ViewSysCommands from (view.bas)
            Exit Sub 'Exits the Sub
            
        'View the Del.exe file in the System folder on the Home Computer (C:\System\)
        Case "view del.exe"
            ViewSysDel 'Calls the ViewSysDel from (view.bas)
            Exit Sub 'Exits the Sub
            
        'View the Deltree.exe file in the System folder on the Home Computer (C:\System\)
        Case "view deltree.exe"
            ViewSysDeltree 'Calls the ViewSysDeltree from (view.bas)
            Exit Sub 'Exits the Sub
            
        'View the View.exe file in the System folder on the Home Computer (C:\System\)
        Case "view view.exe"
            ViewSysView 'Calls the ViewSysView from (view.bas)
            Exit Sub 'Exits the Sub
            
        'View the Kernel.sys file in the Kernel folder in the System folder on the Home Computer (C:\System\Kernel\)
        Case "view kernel.sys"
            ViewSysKerKernel 'Calls the ViewSysKerKernel from (view.bas)
            Exit Sub 'Exits the Sub
            
        'View the Help file in the Help folder on the Home Computer (C:\Help\)
        Case "view help.hlp"
            ViewHelpHelp 'Calls the ViewHelpHelp from (view.bas)
            Exit Sub 'Exits the Sub
            
        Case "cls" 'Clears the Screen
            GetLevel 'Where are we? Loads the GetLevel Sub in (main.bas)
            If Disconnected = False Then 'Are we disconnected?
                frmGame.txtConsole.Text = "" 'Clear the Console
                frmGame.txtConsole.Text = Lvl & vbCrLf 'Clears the Console Except for the Current Level C:\, C:\Documents etc.
            
            Else 'The Computer is disconnected
                frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & "Computer is Disconnected" & vbCrLf
            End If
            
            Exit Sub 'Ends the Current Sub
        
        Case "disc" 'Disconnecting
            GetLevel 'Where are we? Loads the GetLevel Sub in (main.bas)
            If Disconnected = False Then
            frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & "disc" & vbCrLf
            frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Disconnected from Computer" & vbCrLf
            Disconnected = True
            
            Else
                frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & "You Cannot Disconnect Twice??" & vbCrLf
            End If
            Exit Sub
            
        Case "conn" 'Connecting
            GetLevel 'Where are we? Loads the GetLevel Sub in (main.bas)
            If Disconnected = False Then
                frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & "conn" & vbCrLf
                frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Already Connected" & vbCrLf
                frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & vbCrLf
            Else
                frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & "conn" & vbCrLf
                frmGame.txtConsole.Text = frmGame.txtConsole.Text & "You must add what you want to connect to:" & vbCrLf
                frmGame.txtConsole.Text = frmGame.txtConsole.Text & "i.e. home, 323.568.5.9" & vbCrLf
            End If
            Exit Sub
        
        Case "conn home" 'Connecting to home Computer
            GetLevel 'Where are we? Loads the GetLevel Sub in (main.bas)
            If Disconnected = False Then 'Are we disconnected?
                'If we are then tell them that they have to disconnect first
                frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & "conn home" & vbCrLf
                frmGame.txtConsole.Text = frmGame.txtConsole.Text & "You have to Disconnect from this Server first." & vbCrLf
                frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & vbCrLf
            Else
                frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & "conn home" & vbCrLf
                ConHomeComputer 'Loads the ConHomeComputer Sub from (main.bas)
            End If
            Exit Sub
            
        Case "view" 'Viewing a File
            GetLevel 'Where are we? Loads the GetLevel Sub in (main.bas)
            If Disconnected = False Then 'Are we disconnected?
                    frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & "view" & vbCrLf
                    frmGame.txtConsole.Text = frmGame.txtConsole.Text & "You have to add what you want to View i.e.(view log.txt)" & vbCrLf
                    frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & vbCrLf
            Else 'We are disconnected
                frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & "Computer is Disconnected" & vbCrLf
            End If
            Exit Sub 'Ends the Current Sub
            
        Case "del" 'Deleting a File
            GetLevel 'Where are we? Loads the GetLevel Sub in (main.bas)
            If Disconnected = False Then 'Are we disconnected?
                    frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & "del" & vbCrLf
                    frmGame.txtConsole.Text = frmGame.txtConsole.Text & "You have to add what you want to Delete i.e.(del log.txt)" & vbCrLf
                    frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & vbCrLf
            Else 'We are disconnected
                frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & "Computer is Disconnected" & vbCrLf
            End If
            Exit Sub 'Ends the Current Sub
            
        Case "del readme.txt" 'WHAT you want to delete the readme file.
            DelReadmetxt
            Exit Sub
            
        Case "del test.jpg" 'They want to delete C:\Documents\Images\Test.jpg
            DelHomeDocImgTest 'Calls the DelHomeDocImgTest from (del.bas)
            Exit Sub 'Exits the Sub
            
        Case "del cls.exe" 'They want to delete C:\System\Cls.exe
            DelSysCls 'Calls the DelSysCls Sub from (del.bas)
            Exit Sub
            
        Case "del commands.exe" 'They want to delete C:\System\Commands.exe
            DelSysCommands 'Calls the DelSysCommands Sub from (del.bas)
            Exit Sub
            
        Case "del del.exe" 'They want to delete C:\System\Del.exe
            DelSysDel 'Calls the DelSysDel Sub from (del.bas)
            Exit Sub
            
        Case "del deltree.exe" 'They want to delete C:\System\Deltree.exe
            DelSysDeltree 'Calls the DelSysDeltree Sub from (del.bas)
            Exit Sub
            
        Case "del view.exe" 'They want to delete C:\System\View.exe
            DelSysViewexe 'Calls the DelSysViewexe Sub from (del.bas)
            Exit Sub
            
        Case "del boot.ini" 'They want to delete C:\System\Boot\Boot.ini
            DelSysbootboot 'Calls the DelSysbootboot Sub from (del.bas)
            Exit Sub
            
        Case "del user.dat" 'They want to delete C:\System\Boot\User.dat
            DelSysbootuser 'Calls the DelSysbootuser Sub from (del.bas)
            Exit Sub
            
        Case "del system.dat" 'They want to delete C:\System\Boot\System.dat
            DelSysbootsystem 'Calls the DelSysbootsystem Sub from (del.bas)
            Exit Sub
            
        Case "del kernel.sys" 'They want to delete C:\System\Kernel\Kernel.sys
            DelSyskernelkernel 'Calls the DelSyskernelkernel Sub from (del.bas)
            Exit Sub
            
        Case "del help.hlp" 'They want to delete C:\Help\Help.hlp
            DelHomeHelpHelp 'Calls the DelHomeHelpHelp Sub from (del.bas)
            Exit Sub
            
        'They want to delete something that isn't there
        Case "del " & Result
            DelSomething 'Calls the DelSomething Sub from (del.bas)
            Exit Sub
        
        
        'They want to Change Directory to something that isn't allowed (cd klgjfklg)
        Case "cd " & Result1
            If Disconnected = False Then
                cGetCharCount = Len(Text) 'Here is where we count the number of characters
                'This will tell up what they typed after del
                'Mid(What we want searched, The Starting Character, Then the total Characters/The Last Char)
                Result = Mid(Text, 4, cGetCharCount)
                
                frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & "cd " & Result1 & vbCrLf
                frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Could not Find " & Lvl & Result1 & vbCrLf
                frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & vbCrLf
                Exit Sub
            
            Else 'The Computer is disconnected
                frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & "Computer is Disconnected" & vbCrLf
            End If
            Exit Sub 'Exits the Sub :)
            
        'They want to View a file that isn't allowed (view klgjfklg.fdjk)
        Case "view " & Result2
            If Disconnected = False Then
                cGetCharCount = Len(Text) 'Here is where we count the number of characters
                'This will tell up what they typed after del
                'Mid(What we want searched, The Starting Character, Then the total Characters/The Last Char)
                Result = Mid(Text, 6, cGetCharCount)
                
                frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & "view " & Result2 & vbCrLf
                frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Could not Find " & Lvl & Result2 & vbCrLf
                frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & vbCrLf
                Exit Sub
            
            Else 'The Computer is disconnected
                frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & "Computer is Disconnected" & vbCrLf
            End If
            Exit Sub 'Exits the Sub :)
            
        Case "deltree" 'Deleting a Directory and its contents
            GetLevel 'Where are we? Loads the GetLevel Sub in (main.bas)
            If Disconnected = False Then 'Are we disconnected?
                    frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & "deltree" & vbCrLf
                    frmGame.txtConsole.Text = frmGame.txtConsole.Text & "You have to add what you want to Delete i.e.(deltree system)" & vbCrLf
                    frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & vbCrLf
            Else 'We are disconnected
                frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & "Computer is Disconnected" & vbCrLf
            End If
            Exit Sub 'Ends the Current Sub
            
        Case "deltree documents" 'Deleting the Documents Directory and its contents
            DeltreeHomeDocuments 'Calls the DeltreeHomeDocuments Sub from (deltree.bas)
            Exit Sub
            
        Case "deltree images" 'Deleting the Images folder in the Documents Directory and its contents
            DeltreeHomeDocImages 'Calls the DeltreeHomeDocImages Sub from (deltree.bas)
            Exit Sub
            
        Case "deltree recieved" 'Deleting the Images folder in the Documents Directory and its contents
            DeltreeHomeDocRecieved 'Calls the DeltreeHomeDocRecieved Sub from (deltree.bas)
            Exit Sub
            
        Case "deltree downloads" 'Deleting the Downloads Directory and its contents
            DeltreeHomeDownloads 'Calls the DeltreeHomeDownloads Sub from (deltree.bas)
            Exit Sub
            
        Case "deltree help" 'Deleting the Help Directory and its contents
            DeltreeHomeHelp 'Calls the DeltreeHomeHelp Sub from (deltree.bas)
            Exit Sub
            
        Case "deltree software" 'Deleting the Software Directory and its contents
            DeltreeHomeSoftware 'Calls the DeltreeHomeSoftware Sub from (deltree.bas)
            Exit Sub
            
        Case "deltree system" 'Deleting the System Directory and its contents
            DeltreeHomeSystem 'Calls the DeltreeHomeSystem Sub from (deltree.bas)
            Exit Sub
            
        Case "deltree boot" 'Deleting the Boot Directory in System and its contents
            DeltreeHomeSysBoot 'Calls the DeltreeHomeSystem Sub from (deltree.bas)
            Exit Sub
            
        Case "deltree kernel" 'Deleting the Kernel Directory in System and its contents
            DeltreeHomeSysKernel 'Calls the DeltreeHomeSysKernel Sub from (deltree.bas)
            Exit Sub
            
        'They want to Delete a directory that isn't allowed (deltree klgjfklg.fdjk)
        Case "deltree " & Result3
            If Disconnected = False Then
                cGetCharCount = Len(Text) 'Here is where we count the number of characters
                'This will tell up what they typed after del
                'Mid(What we want searched, The Starting Character, Then the total Characters/The Last Char)
                Result = Mid(Text, 6, cGetCharCount)
                
                frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & "deltree " & Result3 & vbCrLf
                frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Could not Find " & Lvl & Result3 & vbCrLf
                frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & vbCrLf
                Exit Sub
            
            Else 'The Computer is disconnected
                frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & "Computer is Disconnected" & vbCrLf
            End If
            Exit Sub 'Exits the Sub :)
            
        
    End Select 'Select what the user has typed
    
    If Disconnected = False Then 'Have we been disconnected
    'These are shown when their was an unrecognized command typed in
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & "'" & Text & "' is not recognized as an internal or " & _
                                                                 "external command, operable program or batch file." & vbCrLf
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & "If you Need help on Commands type command or Press F1" & vbCrLf
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & vbCrLf
    Else 'The Computer is disconnected
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & "Computer is Disconnected" & vbCrLf
    End If
    
    Exit Sub
    
End Sub
