; This script was created using Pulover's Macro Creator
; www.macrocreator.com

#NoEnv
SetWorkingDir %A_ScriptDir%
CoordMode, Mouse, Window
SendMode Input
#SingleInstance Force
SetTitleMatchMode 2
#WinActivateForce
SetControlDelay 1
SetWinDelay 0
SetKeyDelay -1
SetMouseDelay -1
SetBatchLines -1


^F3::
Macro1:
MsgBox, 0, , Welcome to the Sands Close Jobs macro. Please ensure that the "Close Jobs" window is open in Adagio Jobcost and export it to excel before continuing. 
Loop
{
    WinActivate, ahk_exe EXCEL.EXE
    Sleep, 333
    IfWinActive, ahk_exe EXCEL.EXE
    {
        Break
    }
}
Send, {Enter}
Sleep, 50
Loop
{
    Send, {F5}
    IfWinExist, Go To
    {
        Break
    }
}
SendRaw, E1
Sleep, 150
Send, {Enter}
Sleep, 100
Send, {LAlt Down}
Sleep, 50
Send, {Down}
Sleep, 50
Send, {LAlt Up}
Sleep, 50
Send, {f}
Sleep, 50
Send, {g}
Sleep, 50
Send, {0}
Sleep, 50
Send, {Enter}
Sleep, 50
Loop
{
    Send, {F5}
    IfWinExist, Go To
    {
        Break
    }
}
SendRaw, D1
Sleep, 150
Send, {Enter}
Sleep, 100
Send, {LAlt Down}
Sleep, 50
Send, {Down}
Sleep, 50
Send, {LAlt Up}
Sleep, 50
Send, {f}
Sleep, 50
Send, {e}{Space}
Send, {Enter}
Sleep, 50
Loop
{
    Send, {F5}
    IfWinExist, Go To
    {
        Break
    }
}
SendRaw, H2
Sleep, 400
Send, {Enter}
Sleep, 100
Send, {F2}
Sleep, 50
SendRaw, =IF([@[Last Cost Entry Date]]=" "`,[@[Last Billing Date]]`,IF([@[Last Billing Date]]=" "`,[@[Last Cost Entry Date]]`,IF([@[Last Cost Entry Date]]>[@[Last Billing Date]]`,[@[Last Cost Entry Date]]`,[@[Last Billing Date]])))
Sleep, 2000
Send, {Enter}
Sleep, 1000
Send, {Control Down}
Sleep, 50
Send, {Space}
Sleep, 50
Send, {1}
Sleep, 300
Send, {Control Up}
Sleep, 50
Send, {Tab}
Sleep, 50
SendRaw, date
Sleep, 200
Send, {Tab}
Sleep, 50
SendRaw, 14/03/2012
Sleep, 1000
Send, {Enter}
Sleep, 200
Loop
{
    Send, {F5}
    IfWinExist, Go To
    {
        Break
    }
}
SendRaw, I2
Sleep, 400
Send, {Enter}
Sleep, 100
Send, {F2}
Sleep, 50
SendRaw, =NUMBERVALUE([@Code])
Sleep, 1000
Send, {Enter}
Sleep, 1000
Loop
{
    Send, {F5}
    IfWinExist, Go To
    {
        Break
    }
}
SendRaw, I1
Sleep, 400
Send, {Enter}
Sleep, 100
Loop
{
    Loop
    {
        WinActivate, ahk_exe EXCEL.EXE
        Sleep, 333
        IfWinActive, ahk_exe EXCEL.EXE
        {
            Break
        }
    }
    Send, {Down}
    Sleep, 50
    Send, {Control Down}
    Sleep, 50
    Send, {c}
    Sleep, 50
    Send, {Control Up}
    Sleep, 50
    If Clipboard not contains 7
    {
        If Clipboard not contains 6
        {
            If Clipboard not contains 5
            {
                If Clipboard not contains 4
                {
                    If Clipboard not contains 3
                    {
                        If Clipboard not contains 2
                        {
                            If Clipboard not contains 9
                            {
                                If Clipboard not contains 1
                                {
                                    If Clipboard not contains 8
                                    {
                                        Break
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }
    }
    Loop
    {
        WinActivate, Adagio JobCost - SANDS COMMERCIAL FLOOR COVERINGS, Close Jobs
        Sleep, 333
        IfWinActive, JobCost
        {
            Break
        }
    }
    Send, {Alt Down}
    Sleep, 50
    Send, {f}
    Sleep, 50
    Send, {Control Down}
    Sleep, 50
    Send, {Alt Up}
    Sleep, 50
    Send, {v}
    Sleep, 50
    Send, {Control Up}
    Sleep, 50
    Send, {Alt Down}
    Sleep, 50
    Send, {e}
    Sleep, 50
    Send, {Alt Up}
    Sleep, 50
    Send, {Enter}
    Sleep, 500
    IfWinExist, Warning
    {
        Send, {Escape}
        Sleep, 150
    }
    Loop
    {
        WinActivate, ahk_exe EXCEL.EXE
        Sleep, 333
        IfWinActive, ahk_exe EXCEL.EXE
        {
            Break
        }
    }
    Send, {Left}
    Sleep, 50
    Send, {Control Down}
    Sleep, 50
    Send, {c}
    Sleep, 50
    Send, {Control Up}
    Sleep, 50
    Send, {Right}
    Sleep, 50
    Loop
    {
        WinActivate, ahk_exe JobCost.exe
        Sleep, 333
        IfWinActive, ahk_exe JobCost.exe
        {
            Break
        }
    }
    Send, {Control Down}
    Sleep, 50
    Send, {v}
    Sleep, 50
    Send, {Control Up}
    Sleep, 50
    Send, {Enter}
    Sleep, 100
}
MsgBox, 0, , Task Complete!
Return


F8::ExitApp
