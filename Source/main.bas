Attribute VB_Name = "main"
Option Explicit 'Makes sure all the variables are declared

'**********************************************************************************
'This module contains all the variables etc. for the Game
'Also the Sub to calculate the free space on the hard drive, the Pause and GetLevel
'**********************************************************************************

Public s_LoadGameFileName As String  'This is the Loaded Game's File Name
Public b_GameIsLoaded As Boolean     'This stores the True / False value _
                                      if the Game has been Loaded
Public NewServer As Boolean          'This will determine if it is a new game so need to install folders and files.
Public ServerBeingCreated As Boolean 'This will tell the program that the server is still being created so do nothing


'This is where we store the number of characters in a string (Mainly what the user has typed)
Public cGetCharCount
Public lCaseText 'This is where we store the Lower case version of what the _
                  User has typed

Public Result 'This is a variable so that if i need a result _
               of say the Hard drive size etc i can use this.
Public Result1 'This is  to store the result of the two letter programs (cd klfdj)
Public Result2 'This is  to store the result of the four letter programs (view klfdj)
Public Result3 'This is  to store the result of the seven letter programs (deltree klfdj)

'The Current User's Information
Public s_UserName As String 'The User's Name
Public s_Email As String    'Their Email Address
Public s_Password As String 'Their Password


'Where are they information
Public Status As String     'The Status of the Server / Connected / Dis-Connected
Public Level As String      'The Current Directory Level i.e. (home)
Public Lvl As String        'The Current Directory i.e. (C:\)
Public Server As String     'The Server Name i.e. Home (Server X10 - Home Computer) _
                             You Might be able to change this somewhere?? maybe under _
                             Computer Form??

                         
Public Text As String 'This stores what the User has typed in the User Text Box

'These are just in case the user types something while it is connecting _
    Disconnecting or Doing Something (Running a program or something)
Public Connecting As Boolean
Public Disconnecting As Boolean
Public DoingSomething As Boolean


'Has the Files been deleted or not.
Public IsHomeReadmeDel As Boolean 'Has C:\Readme.txt been deleted?? :-]
Public IsHomeDocRecReadmeDel As Boolean 'Has C:\Documents\Recieved\Readme.txt been deleted??
Public IsHomeDocImgTestDel As Boolean 'Has C:\Documents\Images\Test.jpg been deleted?? :-(


Public Disconnected As Boolean 'Sees if it has been disconnected


'User's Computer Information
Public MotherBoard As String 'The Computer's Motherboard
Public CPU As String         'The Computer's CPU
Public CPUSize As String     'The Computer's CPU Size in MHz
Public Memory As String      'The Computer's Memory
Public MemorySize As String  'The Computer's Memory Size in MBs
Public Modem As String       'The Computer's Modem
Public HDDSize As String     'This is the Hard Drive Size
Public HDDName As String     'This is the Name of the Hard Drive
Public HDDSerialNo As String 'This is the Hard Drives Serial No -- I dont know if _
                              i should keep this a constant or make it so when you buy HDD _
                              upgrades this changes with the different hard drives. :)

'This is to work out how much space is left on the hard drive.
Public TotalSize ' As Long 'This is how much all the files on the Server are _
                              taking up.
Public FreeSpace    'This is the total free space on the Hard-Drive

Public PlayFile As Boolean 'This will store the True/False value for if they want to play an mp3 file.



'This is the Pause Code
Sub Pause(interval)
    Dim Current
    Current = Timer
    Do While Timer - Current < Val(interval)
        DoEvents
    Loop
End Sub

'This is to find out where the User is ??
Public Sub GetLevel()
    Select Case Level 'Selects the Level (C:\, C:\Documents etc)
    
        Case "home" 'At Home Computer in the C:\ Drive
            Lvl = "home"
            Lvl = "C:\"
            Exit Sub
            
        Case "documents" 'At Home Computer in the C:\Documents Folder
            Level = "documents"
            Lvl = "C:\Documents\"
            Exit Sub
            
        Case "homerecieved" 'At Home Computer in the C:\Documents\Recieved Folder
            Level = "homerecieved"
            Lvl = "C:\Documents\Recieved\"
            Exit Sub
            
        Case "homedocimages" 'At Home Computer in the Images folder under Documents (C:\Documents\Images)
            Level = "homedocimages"
            Lvl = "C:\Documents\Images\"
            Exit Sub
            
        Case "homedownloads" 'At Home Computer in the Downloads folder (C:\Downloads\)
            Level = "homedownloads"
            Lvl = "C:\Downloads\"
            Exit Sub 'Three guesses what this does! :)
            
        Case "homesystem" 'At Home Computer in the Systems folder (C:\System\)
            Level = "homesystem"
            Lvl = "C:\System\"
            Exit Sub
            
        Case "homesoftware" 'At Home Computer in the Software folder (C:\Software\)
            Level = "homesoftware"
            Lvl = "C:\Software\"
            
        Case "homesysboot" 'At Home Computer in the Boot folder under System (C:\System\Boot\)
            Level = "homesysboot"
            Lvl = "C:\System\Boot\"
            Exit Sub
            
        Case "homesyskernel" 'At Home Computer in the Kernel folder under System (C:\System\Kernel\)
            Level = "homesyskernel"
            Lvl = "C:\System\Kernel\"
            Exit Sub
            
        Case "homehelp" 'At Home Computer in the Help folder (C:\Help\)
            Level = "homehelp"
            Lvl = "C:\Help\"
            Exit Sub
            
    End Select
End Sub
    
Public Function CalcFreeSpace() 'This Calculates the Total Free Space on the Hard Drive
    TotalFileSize
    FreeSpace = HDDSize - TotalSize
    Result = Format(FreeSpace, "##,##0")
End Function

Public Function TotalFileSize() 'This is where it gets the total file size from (this will change with different servers)
    Select Case Server
    
        'We are home
        Case "Server X10 - Home Computer"
            'The Different Files, Readme's and images etc.
            Dim HomeReadme 'C:\Readme.txt Size: 233 bytes
            Dim HomeDocRecReadMe 'C:\Documents\Recieved\Readme.txt Size: 256 bytes
            Dim HomeDocImgTest 'C:\Documents\Images\Test.jpg Size: 3,520 bytes
            Dim HomeSysCls 'C:\System\Cls.exe Size: 6,594 bytes
            Dim HomeSysCommands 'C:\System\Commands.exe Size: 45,379 bytes
            Dim HomeSysDel 'C:\System\Del.exe Size: 14,144 bytes
            Dim HomeSysDeltree 'C:\System\Deltree.exe Size: 19,083 bytes
            Dim HomeSysView 'C:\System\View.exe Size: 45,056 bytes
            Dim HomeSysBootBoot 'C:\System\Boot\Boot.ini Size: 211 bytes
            Dim HomeSysBootUser 'C:\System\Boot\User.dat Size: 225,312 bytes
            Dim HomeSysBootSystem 'C:\System\Boot\System.dat Size: 2,932,768 bytes
            Dim HomeSysKernelKernel 'C:\System\Kernel\Kernel.sys Size: 79,691,776 bytes
            
            
            'The Values for the files that cannot be deleted i.e. system files cls, del etc.
            HomeSysCls = 6594 'C:\System\Cls.exe Size: 6594 bytes
            HomeSysCommands = 45379 'C:\System\Commands.exe Size: 45,379 bytes
            HomeSysDel = 14144 'C:\System\Del.exe Size: 14,144 bytes
            HomeSysDeltree = 19083 'C:\System\Deltree.exe Size: 19,083 bytes
            HomeSysView = 45056 'C:\System\View.exe Size: 45,056
            HomeSysBootBoot = 211 'C:\System\Boot\Boot.ini Size: 211 bytes
            HomeSysBootUser = 225312 'C:\System\Boot\User.dat Size: 225,312 bytes
            HomeSysBootSystem = 932768 'C:\System\Boot\System.dat Size: 2,932,768 bytes
            HomeSysKernelKernel = 5691776 'C:\System\Kernel\Kernel.sys Size: 79,691,776 bytes
            
            
            'Does the C:\Readme.txt exist?
            If IsHomeReadmeDel = True Then 'If the file has been deleted then dont subtract its size from hard drive total size
            HomeReadme = 0 'It isn't there so it is 0 bytes
            Else 'It has not been deleted
            HomeReadme = 233 'It is there so it is 233 bytes
            End If
            
            'Does the C:\Documents\Recieved\Readme.txt exist?
            If IsHomeDocRecReadmeDel = True Then 'If the file has been deleted then dont subtract its size from hard drive total size
            HomeDocRecReadMe = 0 'It isn't there so it is 0 bytes
            Else 'It has not been deleted
            HomeDocRecReadMe = 256 'It is there so it is 233 bytes
            End If
            
            'Does the C:\Documents\Images\Test.jpg exist?
            If IsHomeDocImgTestDel = True Then 'If the file has been deleted then dont subtract its size from hard drive total size
            HomeDocImgTest = 0 'It isn't there so it is 0 bytes
            Else 'It has not been deleted
            HomeDocImgTest = 3520 'It is there so it is 3520 bytes
            End If
            
            'This is the total size of all the documents etc. on the Home Comp
            TotalSize = (HomeReadme + HomeDocRecReadMe + HomeDocImgTest + HomeSysCls + _
                         HomeSysCommands + HomeSysDel + HomeSysDeltree + HomeSysView + _
                         HomeSysBootBoot + HomeSysBootUser + HomeSysBootSystem + _
                         HomeSysKernelKernel)
            
        Exit Function 'Exits the Function :-)
        
    End Select 'Server Select
End Function


'This checks to see if the file exists
Public Function FileExists(FullFileName As String) As Boolean
    On Error GoTo MakeF
        'If file does Not exist, there will be an Error
        Open FullFileName For Input As #1
        Close #1
        'No Error, File Exists
        FileExists = True
    Exit Function
MakeF:
        'Error, The File does not exist
        FileExists = False
    Exit Function
End Function


'This will setup the new server
Public Sub SetupNewServer()
    ServerBeingCreated = True

    'Displays the Formatting text (more for looks than anything)
    frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & "Formatting Hard-Drive" & vbCrLf
    Pause (0.2)
    frmGame.txtConsole.SelStart = InStr(frmGame.txtConsole.Text, "Formatting Hard-Drive") - 1
    frmGame.txtConsole.SelLength = Len("Formatting Hard-Drive")
    frmGame.txtConsole.SelText = "Formatting Hard-Drive 1%"
    Pause (0.03)
    
    Dim TempNum
    TempNum = 1
    While TempNum < 100
        frmGame.txtConsole.SelStart = InStr(frmGame.txtConsole.Text, "Formatting Hard-Drive " & TempNum & "%") - 1
        frmGame.txtConsole.SelLength = Len("Formatting Hard-Drive " & TempNum & "%")
        TempNum = TempNum + 1
        frmGame.txtConsole.SelText = "Formatting Hard-Drive " & TempNum & "%"
        Pause (0.03)
    Wend
    Pause (1)
    
    'As you might have guessed Formatting complete :-)
    frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Format Complete!" & vbCrLf
    
    
    'Copying the System Files
    frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & "Copying System Files" & vbCrLf
    Pause (0.2)
    frmGame.txtConsole.SelStart = InStr(frmGame.txtConsole.Text, "Copying System Files") - 1
    frmGame.txtConsole.SelLength = Len("Copying System Files")
    frmGame.txtConsole.SelText = "Copying System Files 1%"
    Pause (0.01)
    
    TempNum = 1
    While TempNum < 100
        frmGame.txtConsole.SelStart = InStr(frmGame.txtConsole.Text, "Copying System Files " & TempNum & "%") - 1
        frmGame.txtConsole.SelLength = Len("Copying System Files " & TempNum & "%")
        TempNum = TempNum + 1
        frmGame.txtConsole.SelText = "Copying System Files " & TempNum & "%"
        Pause (0.01)
    Wend
    Pause (1)
    
    'Copying of files is complete
    frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Inflating C:\System\Cls.exe" & vbCrLf
    Pause (0.5)
    frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Inflating C:\System\Commands.exe" & vbCrLf
    Pause (0.5)
    frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Inflating C:\System\Del.exe" & vbCrLf
    Pause (0.5)
    frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Inflating C:\System\Deltree.exe" & vbCrLf
    Pause (0.5)
    frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Inflating C:\System\View.exe" & vbCrLf
    Pause (0.5)
    frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Inflating C:\Readme.txt" & vbCrLf
    Pause (0.5)
    frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Inflating C:\Documents\Recieved\Readme.txt" & vbCrLf
    Pause (0.5)
    frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Inflating C:\Documents\Images\Test.jpg" & vbCrLf
    Pause (0.5)
    frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Inflating C:\System\Boot\User.dat" & vbCrLf
    Pause (0.5)
    frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Inflating C:\System\Boot\System.dat" & vbCrLf
    Pause (0.5)
    frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Inflating C:\System\Kernel\Kernel.sys" & vbCrLf
    Pause (1)
    frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Inflating C:\Help\Help.hlp" & vbCrLf
    Pause (0.9)
    frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Copying Completed!" & vbCrLf
    Pause (0.1)

    frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & "Server Set-Up is Completed." & vbCrLf
    frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & "Thank-you for using B4W" & vbCrLf
    Pause (1)

    ServerBeingCreated = False
End Sub
