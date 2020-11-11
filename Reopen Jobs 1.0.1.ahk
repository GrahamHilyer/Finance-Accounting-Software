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


Macro1:
MsgBox, 0, , 
(LTrim
Welcome to the Sands reopen job macro program. 

Please export the error cost batch to excel and ensure that the "Close Jobs" window is open in Adagio JobCost before continuing.
)
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
Sleep, 100
Loop
{
    Send, {F5}
    IfWinExist, Go To
    {
        Break
    }
}
SendRaw, L2
Sleep, 150
Send, {Enter}
Sleep, 100
Send, {F2}
Sleep, 50
SendRaw, =NUMBERVALUE(MID([@[Job-Ph-Cat]]`,1`,5))
Sleep, 1000
Send, {Enter}
Sleep, 1000
LoopNumber := 2
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
    Loop
    {
        Send, {F5}
        IfWinExist, Go To
        {
            Break
        }
    }
    SendRaw, L%LoopNumber%
    Sleep, 400
    Send, {Enter}
    Sleep, 100
    Send, {Control Down}
    Sleep, 50
    Send, {c}
    Sleep, 50
    Send, {Control Up}
    Sleep, 50
    If Clipboard not contains 9
    {
        If Clipboard not contains 8
        {
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
                                    If Clipboard not contains 1
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
        IfWinActive, Adagio JobCost - SANDS COMMERCIAL FLOOR COVERINGS, Close Jobs
        {
            Break
        }
    }
    Send, {Alt Down}
    Sleep, 50
    Send, {f}
    Sleep, 50
    Send, {Alt Up}
    Sleep, 50
    Send, {Control Down}
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
    Send, {Delete}
    Sleep, 50
    Send, {Enter}
    Sleep, 50
    LoopNumber += 1
}
MsgBox, 0, , Job Closing Completed. Please post the batch.
Return


F8::ExitApp
