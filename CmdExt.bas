Attribute VB_Name = "CmdExt"
Public Enum CharmelionButtonTypes
    [STYLE Windows 16-bit] = 1    'the old-fashioned Win16 button
    [STYLE Windows 32-bit] = 2    'the classic windows button
    [STYLE Windows XP] = 3        'the new brand XP button totally owner-drawn
    [STYLE MacOS 9] = 4           'I suppose it looks exactly as a Mac button... i took the style from a GetRight skin!!!
    [STYLE Java metal] = 5        'there are also other styles but not so different from windows one
    [STYLE Netscape 6] = 6        'this is the button displayed in web-pages, it also appears in some java apps
    [STYLE Simple Flat] = 7       'the standard flat button seen on toolbars
    [STYLE Flat Highlight] = 8    'again the flat button but this one has no border until the mouse is over it
    [STYLE Office XP] = 9         'the new Office XP button
    [STYLE Transparent] = 11      'suggested from a user...
    [STYLE 3D Hover] = 12         'took this one from "Noteworthy Composer" toolbal
    [STYLE Oval Flat] = 13        'a simple Oval Button
    [STYLE KDE 2] = 14            'the great standard KDE2 button!
End Enum

Function ExtendCommands() ' Extend standard voice commands
    mainfrm.Agent1.Characters("Genie").Commands.Item("SayTime").Voice = "(...Time?|...Time...)"                                                     ' EXTENDED COMMANDS, Allows non-standard input
    mainfrm.Agent1.Characters("Genie").Commands.Item("SayDate").Voice = "(...Date?|...Date...)"
    mainfrm.Agent1.Characters("Genie").Commands.Item("ReadClip").Voice = "(Read Clipboard|...Clipboard?)"
    mainfrm.Agent1.Characters("Genie").Commands.Item("SS").Voice = "(...Say Somthing|Say Somthing)"
    mainfrm.Agent1.Characters("Genie").Commands.Item("CC").Voice = "(Change Character|...Character)"
    mainfrm.Agent1.Characters("Genie").Commands.Item("Options").Voice = "(Options|...Options)"
    mainfrm.Agent1.Characters("Genie").Commands.Item("Actions").Voice = "(...Actions?|Actions)"
    mainfrm.Agent1.Characters("Genie").Commands.Item("VoiceA").Voice = "(...Voice Actions|Voice Actions)"
    mainfrm.Agent1.Characters("Genie").Commands.Item("notepad").Voice = "(...NotePad|NotePad)"
    mainfrm.Agent1.Characters("Genie").Commands.Item("calc").Voice = "(...Calculator|Calculator)"
    mainfrm.Agent1.Characters("Genie").Commands.Item("defrag").Voice = "(...Defragmentor|Defragmentor)"
    mainfrm.Agent1.Characters("Genie").Commands.Item("OCD").Voice = "(...Open CD (Drive|Door)|Open CD (Drive|Door))"
    mainfrm.Agent1.Characters("Genie").Commands.Item("CCD").Voice = "(...Close CD (Drive|Door)|Close CD (Drive|Door))"
    mainfrm.Agent1.Characters("Genie").Commands.Item("explorer").Voice = "(...Explorer|Explorer)"
    mainfrm.Agent1.Characters("Genie").Commands.Item("Close").Voice = "(...(Quit|Close) [ComTalk]|(Quit|Close) [ComTalk])"
    mainfrm.Agent1.Characters("Genie").Commands.Item("ShutDown").Voice = "(...Shut Down [Computer]|Shut Down [Computer])"
    mainfrm.Agent1.Characters("Genie").Commands.Item("Restart").Voice = "(...Restart [Computer]|Restart [Computer])"
    mainfrm.Agent1.Characters("Genie").Commands.Item("LogOff").Voice = "(...LogOff [Computer]|LogOff [Computer])"
    If NoLockDown = False Then ' Don't load command if command line parameter to disable it is used
    mainfrm.Agent1.Characters("Genie").Commands.Item("LockWS").Voice = "(Lock Workstation|Lock Workstation)"
    End If
    mainfrm.Agent1.Characters("Genie").Commands.Item("SWRun").Voice = "(...(Windows Been Running?|Windows Runtime)|...Windows Runtime...)"
    mainfrm.Agent1.Characters("Genie").Commands.Item("About").Voice = "(About...|About)"
    mainfrm.Agent1.Characters("Genie").Commands.Item("SAC").Voice = "(...Clock|Sync Clock...)"
    mainfrm.Agent1.Characters("Genie").Commands.Item("MSAgentOptions").Voice = "(...Agent Options|...MSAgent Options)"
End Function

Function ExtendedMenu(GetCA As Boolean) ' Loads the popup menu -
                                        ' Previously used to select between in-built and extended poputp, but now
                                        ' only the extended popup menu.
    CustomMenu.MNUCusEmpty.Visible = True
    CustomMenu.MNUONE.Visible = False
    CustomMenu.MNUTWO.Visible = False
    CustomMenu.MNUTHREE.Visible = False
    CustomMenu.MNUFOUR.Visible = False
    CustomMenu.MNUFIVE.Visible = False
    CustomMenu.MNUSIX.Visible = False
    CustomMenu.MNUSEVEN.Visible = False
    CustomMenu.MNUEIGHT.Visible = False
    CustomMenu.MNUNINE.Visible = False
    CustomMenu.MNUTEN.Visible = False
    
    If GetCA = True Then Declairs.GetCAFromReg
    CustomMenu.MNUONE.Caption = ONEname
    CustomMenu.MNUTWO.Caption = TWOname
    CustomMenu.MNUTHREE.Caption = THREEname
    CustomMenu.MNUFOUR.Caption = FOURname
    CustomMenu.MNUFIVE.Caption = FIVEname
    CustomMenu.MNUSIX.Caption = SIXname
    CustomMenu.MNUSEVEN.Caption = SEVENname
    CustomMenu.MNUEIGHT.Caption = EIGHTname
    CustomMenu.MNUNINE.Caption = NINEname
    CustomMenu.MNUTEN.Caption = TENname
    
    If CustomMenu.MNUONE.Caption <> "" Then
        CustomMenu.MNUONE.Visible = True
        CustomMenu.MNUCusEmpty.Visible = False
    End If
    
    If CustomMenu.MNUTWO.Caption <> "" Then
        CustomMenu.MNUTWO.Visible = True
        CustomMenu.MNUCusEmpty.Visible = False
    End If
    
    If CustomMenu.MNUTHREE.Caption <> "" Then
        CustomMenu.MNUTHREE.Visible = True
        CustomMenu.MNUCusEmpty.Visible = False
    End If
    
    If CustomMenu.MNUFOUR.Caption <> "" Then
        CustomMenu.MNUFOUR.Visible = True
        CustomMenu.MNUCusEmpty.Visible = False
    End If
    
    If CustomMenu.MNUFIVE.Caption <> "" Then
        CustomMenu.MNUFIVE.Visible = True
        CustomMenu.MNUCusEmpty.Visible = False
    End If
    
    If CustomMenu.MNUSIX.Caption <> "" Then
        CustomMenu.MNUSIX.Visible = True
        CustomMenu.MNUCusEmpty.Visible = False
    End If
    
    If CustomMenu.MNUSEVEN.Caption <> "" Then
        CustomMenu.MNUSEVEN.Visible = True
        CustomMenu.MNUCusEmpty.Visible = False
    End If
    
    If CustomMenu.MNUEIGHT.Caption <> "" Then
        CustomMenu.MNUEIGHT.Visible = True
        CustomMenu.MNUCusEmpty.Visible = False
    End If
    
    If CustomMenu.MNUNINE.Caption <> "" Then
        CustomMenu.MNUNINE.Visible = True
        CustomMenu.MNUCusEmpty.Visible = False
    End If
    
    If CustomMenu.MNUTEN.Caption <> "" Then
        CustomMenu.MNUTEN.Visible = True
        CustomMenu.MNUCusEmpty.Visible = False
    End If
End Function

Function GoButtonAppearance(ButtonAppearanceType As CharmelionButtonTypes) ' Change all Charmelion buttons to the selected style
    actions.command1.ButtonType = ButtonAppearanceType
    
    ChangeAppearance vcommands, ButtonAppearanceType        ' Add the form name here
    ChangeAppearance ss, ButtonAppearanceType
    ChangeAppearance SetProgPaths, ButtonAppearanceType
    ChangeAppearance options, ButtonAppearanceType
    ChangeAppearance mainfrm, ButtonAppearanceType
    ChangeAppearance Customreminders, ButtonAppearanceType
    ChangeAppearance actions, ButtonAppearanceType
    ChangeAppearance aboutplugs, ButtonAppearanceType
    ChangeAppearance DisableCommands, ButtonAppearanceType
End Function

Private Function ChangeAppearance(UseForm As Form, NewAppearance As CharmelionButtonTypes) ' Scans each control on the selected form and changes any charmelion buttons to the selected style
    For Each Control In UseForm.Controls ' Look at each control on the form
        If TypeOf Control Is chameleonButton Then ' Check if the control is a Charmelion button
            Control.ButtonType = NewAppearance ' Change the button's style
            Control.BackColor = RGB(255, 255, 255) ' Fix the colors
            Control.BackOver = RGB(200, 200, 200)  '      |
            Control.ColorScheme = Custom           '______|
        End If
    Next Control
End Function
