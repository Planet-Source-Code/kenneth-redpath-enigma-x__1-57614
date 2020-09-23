Attribute VB_Name = "dir"
Option Explicit 'Makes sure all the variables are declared

'**********************************************************************************
'This module contains all the code for the dir command. ie "dir documents" etc.
'I put it in here to clean up the code on the frmGame a bit :-)
'**********************************************************************************


'They want to browse the C:\ of the Home Computer
Public Sub DirHome()
    CalcFreeSpace 'Get the free space on the hard drive
    frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & "dir" & vbCrLf
    If HDDName = "" Then
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Drive C has no name" & vbCrLf
    Else
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Drive C is " & HDDName & vbCrLf
    End If
    frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Volume Serial Number is " & HDDSerialNo & vbCrLf & vbCrLf
    frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Directory of " & Lvl & vbCrLf & vbCrLf
    frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Documents     <DIR>" & vbCrLf
    frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Downloads     <DIR>" & vbCrLf
    frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Help          <DIR>" & vbCrLf
    If IsHomeReadmeDel = False Then frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Readme.txt    223 bytes" & vbCrLf
    frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Software      <DIR>" & vbCrLf
    frmGame.txtConsole.Text = frmGame.txtConsole.Text & "System        <DIR>" & vbCrLf
    If IsHomeReadmeDel = False Then
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & "     1 File(s)     223 bytes" & vbCrLf
    Else
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & "     0 File(s)     0 bytes" & vbCrLf
    End If
    frmGame.txtConsole.Text = frmGame.txtConsole.Text & "     5 Dir(s)      " & Result & " bytes free" & vbCrLf
    frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & vbCrLf
                    
    Exit Sub 'Ends the Current Sub
End Sub


'They want to browse the C:\Documents\ folder on the Home Computer
Public Sub DirHomeDocuments()
    CalcFreeSpace 'Get the free space on the hard drive
    frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & "dir" & vbCrLf
    If HDDName = "" Then
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Drive C has no name" & vbCrLf
    Else
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Drive C is " & HDDName & vbCrLf
    End If
    frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Volume Serial Number is " & HDDSerialNo & vbCrLf & vbCrLf
    frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Directory of " & Lvl & vbCrLf & vbCrLf
    frmGame.txtConsole.Text = frmGame.txtConsole.Text & "..            <DIR>" & vbCrLf
    frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Images        <DIR>" & vbCrLf
    frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Recieved      <DIR>" & vbCrLf
    frmGame.txtConsole.Text = frmGame.txtConsole.Text & "     0 File(s)     0 bytes" & vbCrLf
    frmGame.txtConsole.Text = frmGame.txtConsole.Text & "     3 Dir(s)      " & Result & " bytes free" & vbCrLf
    frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & vbCrLf
    Exit Sub 'Ends the Current Sub
End Sub


'They want to browse the C:\Documents\Recieved\ folder on the Home Computer
Public Sub DirHomeDocRecieved()
    CalcFreeSpace 'Get the free space on the hard drive
    frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & "dir" & vbCrLf
    If HDDName = "" Then
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Drive C has no name" & vbCrLf
    Else
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Drive C is " & HDDName & vbCrLf
    End If
    frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Volume Serial Number is " & HDDSerialNo & vbCrLf & vbCrLf
    frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Directory of " & Lvl & vbCrLf & vbCrLf
    frmGame.txtConsole.Text = frmGame.txtConsole.Text & "..            <DIR>" & vbCrLf
    If IsHomeDocRecReadmeDel = False Then frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Readme.txt    256 bytes" & vbCrLf
    If IsHomeDocRecReadmeDel = False Then
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & "     1 File(s)     256 bytes" & vbCrLf
    Else
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & "     0 File(s)     0 bytes" & vbCrLf
    End If
    frmGame.txtConsole.Text = frmGame.txtConsole.Text & "     1 Dir(s)      " & Result & " bytes free" & vbCrLf
    frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & vbCrLf
    Exit Sub 'Exits the Sub
End Sub


'They want to browse the C:\Documents\Images\ folder on the Home Computer
Public Sub DirHomeDocImages()
    CalcFreeSpace
    frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & "dir" & vbCrLf
    If HDDName = "" Then
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Drive C has no name" & vbCrLf
    Else
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Drive C is " & HDDName & vbCrLf
    End If
    frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Volume Serial Number is " & HDDSerialNo & vbCrLf & vbCrLf
    frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Directory of " & Lvl & vbCrLf & vbCrLf
    frmGame.txtConsole.Text = frmGame.txtConsole.Text & "..            <DIR>" & vbCrLf
    If IsHomeDocImgTestDel = False Then frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Test.jpg      3,520 bytes" & vbCrLf
    If IsHomeDocImgTestDel = False Then
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & "     1 File(s)     3,520 bytes" & vbCrLf
    Else
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & "     0 File(s)     0 bytes" & vbCrLf
    End If
    frmGame.txtConsole.Text = frmGame.txtConsole.Text & "     1 Dir(s)      " & Result & " bytes free" & vbCrLf
    frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & vbCrLf
    Exit Sub 'Exits the Sub
End Sub


'They want to browse the C:\Downloads\ folder on the Home Computer
Public Sub DirHomeDownloads()
    CalcFreeSpace
    frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & "dir" & vbCrLf
    If HDDName = "" Then
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Drive C has no name" & vbCrLf
    Else
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Drive C is " & HDDName & vbCrLf
    End If
    frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Volume Serial Number is " & HDDSerialNo & vbCrLf & vbCrLf
    frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Directory of " & Lvl & vbCrLf & vbCrLf
    frmGame.txtConsole.Text = frmGame.txtConsole.Text & "..            <DIR>" & vbCrLf
    frmGame.txtConsole.Text = frmGame.txtConsole.Text & "     1 Dir(s)      " & Result & " bytes free" & vbCrLf
    frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & vbCrLf
    Exit Sub 'Exits the Sub
End Sub


'They want to browse the C:\System\ folder on the Home Computer
Public Sub DirHomeSystem()
    CalcFreeSpace
    frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & "dir" & vbCrLf
    If HDDName = "" Then
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Drive C has no name" & vbCrLf
    Else
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Drive C is " & HDDName & vbCrLf
    End If
    frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Volume Serial Number is " & HDDSerialNo & vbCrLf & vbCrLf
    frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Directory of " & Lvl & vbCrLf & vbCrLf
    frmGame.txtConsole.Text = frmGame.txtConsole.Text & "..            <DIR>" & vbCrLf
    frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Boot          <DIR>" & vbCrLf
    frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Cls.exe       6,594 bytes" & vbCrLf
    frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Commands.exe  45,379 bytes" & vbCrLf
    frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Del.exe       14,144 bytes" & vbCrLf
    frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Deltree.exe   19,083 bytes" & vbCrLf
    frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Kernel        <DIR>" & vbCrLf
    frmGame.txtConsole.Text = frmGame.txtConsole.Text & "View.exe      45,056 bytes" & vbCrLf
    frmGame.txtConsole.Text = frmGame.txtConsole.Text & "     5 File(s)     130,256 bytes" & vbCrLf
    frmGame.txtConsole.Text = frmGame.txtConsole.Text & "     2 Dir(s)      " & Result & " bytes free" & vbCrLf
    frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & vbCrLf
    Exit Sub 'Exits the Sub
End Sub


'They want to browse the C:\System\Boot\ folder on the Home Computer
Public Sub DirHomeSysBoot()
    CalcFreeSpace
    frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & "dir" & vbCrLf
    If HDDName = "" Then
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Drive C has no name" & vbCrLf
    Else
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Drive C is " & HDDName & vbCrLf
    End If
    frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Volume Serial Number is " & HDDSerialNo & vbCrLf & vbCrLf
    frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Directory of " & Lvl & vbCrLf & vbCrLf
    frmGame.txtConsole.Text = frmGame.txtConsole.Text & "..            <DIR>" & vbCrLf
    frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Boot.ini      211 bytes" & vbCrLf
    frmGame.txtConsole.Text = frmGame.txtConsole.Text & "User.dat      225,312 bytes" & vbCrLf
    frmGame.txtConsole.Text = frmGame.txtConsole.Text & "System.dat    2,932,768 bytes" & vbCrLf
    frmGame.txtConsole.Text = frmGame.txtConsole.Text & "     3 File(s)     3,158,291 bytes" & vbCrLf
    frmGame.txtConsole.Text = frmGame.txtConsole.Text & "     1 Dir(s)      " & Result & " bytes free" & vbCrLf
    frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & vbCrLf
    
    Exit Sub 'Exits the Sub
End Sub


'They want to browse the C:\System\Kernel\ folder on the Home Computer
Public Sub DirHomeSysKernel()
    CalcFreeSpace
    frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & "dir" & vbCrLf
    If HDDName = "" Then
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Drive C has no name" & vbCrLf
    Else
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Drive C is " & HDDName & vbCrLf
    End If
    frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Volume Serial Number is " & HDDSerialNo & vbCrLf & vbCrLf
    frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Directory of " & Lvl & vbCrLf & vbCrLf
    frmGame.txtConsole.Text = frmGame.txtConsole.Text & "..            <DIR>" & vbCrLf
    frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Kernel.sys    79,691,776 bytes" & vbCrLf
    frmGame.txtConsole.Text = frmGame.txtConsole.Text & "     1 File(s)     79,691,776 bytes" & vbCrLf
    frmGame.txtConsole.Text = frmGame.txtConsole.Text & "     1 Dir(s)      " & Result & " bytes free" & vbCrLf
    frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & vbCrLf
    
    Exit Sub 'Exits the Sub
End Sub


'They want to browse the C:\Help\ folder on the Home Computer
Public Sub DirHomeHelp()
    CalcFreeSpace
    frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & "dir" & vbCrLf
    If HDDName = "" Then
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Drive C has no name" & vbCrLf
    Else
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Drive C is " & HDDName & vbCrLf
    End If
    frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Volume Serial Number is " & HDDSerialNo & vbCrLf & vbCrLf
    frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Directory of " & Lvl & vbCrLf & vbCrLf
    frmGame.txtConsole.Text = frmGame.txtConsole.Text & "..            <DIR>" & vbCrLf
    frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Help.hlp    1,125,888 bytes" & vbCrLf
    frmGame.txtConsole.Text = frmGame.txtConsole.Text & "     1 File(s)     1,125,888 bytes" & vbCrLf
    frmGame.txtConsole.Text = frmGame.txtConsole.Text & "     1 Dir(s)      " & Result & " bytes free" & vbCrLf
    frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & vbCrLf
    
    Exit Sub 'Exits the Sub
End Sub


'They want to browse the C:\Software\ folder on the Home Computer
Public Sub DirHomeSoftware()
    CalcFreeSpace
    frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & "dir" & vbCrLf
    If HDDName = "" Then
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Drive C has no name" & vbCrLf
    Else
        frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Drive C is " & HDDName & vbCrLf
    End If
    frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Volume Serial Number is " & HDDSerialNo & vbCrLf & vbCrLf
    frmGame.txtConsole.Text = frmGame.txtConsole.Text & "Directory of " & Lvl & vbCrLf & vbCrLf
    frmGame.txtConsole.Text = frmGame.txtConsole.Text & "..            <DIR>" & vbCrLf
    frmGame.txtConsole.Text = frmGame.txtConsole.Text & "     0 File(s)     0 bytes" & vbCrLf
    frmGame.txtConsole.Text = frmGame.txtConsole.Text & "     1 Dir(s)      " & Result & " bytes free" & vbCrLf
    frmGame.txtConsole.Text = frmGame.txtConsole.Text & vbCrLf & Lvl & vbCrLf
    
    Exit Sub 'Exits the Sub
End Sub
