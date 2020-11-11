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


OpeningJobs:
SetNumLockState, On
SetScrollLockState, Off
MsgBox, 0, , 
(LTrim
This macro will enter all jobs from the open work order report. Once initiated please do not use the mouse or keyboard until the "task complete" message appears. 

Please Ensure: 
1. Job Costs (with "estimates" and "jobs" tabs opened) is open.
2. Google Chrome is open & logged into the work order system.
3. The open work order report excel sheet is opened.

Hit OK to continue!
)
Loop
{
    WinActivate, Excel
    Sleep, 333
    IfWinActive, Excel
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
SendRaw, A1
Sleep, 150
Send, {Enter}
Sleep, 100
Send, {LShift Down}
Sleep, 50
Send, {Down 4}
Sleep, 30
Send, {LShift Up}
Sleep, 50
Send, {LControl Down}
Loop
{
    Send, {NumpadSub}
    IfWinExist, Delete
    {
        Break
    }
}
Send, {LControl Up}
Send, {r}
Sleep, 50
Send, {Enter}
Sleep, 50
Loop
{
    Loop
    {
        WinActivate, JobCost
        Sleep, 333
        IfWinActive, JobCost
        {
            Break
        }
    }
    Click, 653, 159 Left, 1
    Sleep, 10
    Send, {Alt Down}
    Sleep, 30
    Send, {n}
    Send, {Alt Up}
    Sleep, 30
    Loop
    {
        WinActivate, Excel
        Sleep, 333
        IfWinActive, Excel
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
    SendRaw, A1
    Sleep, 100
    Send, {Enter}
    Sleep, 150
    Send, {LControl Down}
    Sleep, 30
    Send, {c}
    Send, {LControl Up}
    Sleep, 30
    Loop
    {
        WinActivate, New Job
        Sleep, 333
        IfWinExist, New Job
        {
            Break
        }
    }
    WOValue := Clipboard
    Send, {LControl Down}
    Sleep, 50
    Send, {v}
    Send, {LControl Up}
    Sleep, 60
    Loop
    {
        WinActivate, Excel
        Sleep, 333
        IfWinActive, Excel
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
    SendRaw, D1
    Sleep, 100
    Send, {Enter}
    Sleep, 150
    Send, {LControl Down}
    Sleep, 30
    Send, {c}
    Send, {LControl Up}
    Sleep, 30
    Loop
    {
        WinActivate, New Job
        Sleep, 333
        IfWinActive, New Job
        {
            Break
        }
    }
    Send, {Tab}
    Sleep, 50
    Send, {LControl Down}
    Sleep, 30
    Send, {v}
    Send, {LControl Up}
    Sleep, 30
    Loop
    {
        WinActivate, Excel
        Sleep, 333
        IfWinActive, Excel
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
    SendRaw, H1
    Sleep, 100
    Send, {Enter}
    Sleep, 100
    Clipboard := 0
    Send, {F2}
    Sleep, 100
    SendRaw, =JobCostInitials(E1)
    Sleep, 500
    Send, {Enter}
    Sleep, 400
    Send, {Up}
    Sleep, 50
    Delay := 100
    Loop
    {
        Sleep, %Delay%
        Send, {LControl Down}
        Sleep, 30
        Send, {c}
        Send, {LControl Up}
        Sleep, 30
        If Clipboard !=
        {
            IfNotInString, Clipboard, 0
            {
                Break
            }
        }
        Delay += 100
    }
    Loop
    {
        WinActivate, New Job
        Sleep, 333
        IfWinActive, New Job
        {
            Break
        }
    }
    Send, {Tab 2}
    Sleep, 40
    Send, {LControl Down}
    Sleep, 30
    Send, {v}
    Send, {LControl Up}
    Sleep, 80
    Run, "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe" http://win2012:8088/WorkOrder/Customer?woNo=%WOValue%
    Clipboard := ""
    Loop
    {
        WinActivate, Chrome
        Sleep, 333
        IfWinExist, Chrome
        {
            Break
        }
    }
    Loop
    {
        Send, {LControl Down}
        Sleep, 30
        Send, {a}
        Sleep, 30
        Send, {c}
        Send, {LControl Up}
        Sleep, 180
        IfInString, Clipboard, Cust #
        {
            Break
        }
    }
    StringGetPos, CustPos, Clipboard, Cust #:, L1, 0
    CustPos += 9
    StringTrimLeft, String1, Clipboard, %CustPos%
    StringGetPos, CustPos, String1, Name:, R1, 0
    CustPos -= 2
    StringLeft, Clipboard, String1, %CustPos%
    Send, {LControl Down}
    Sleep, 30
    Send, {w}
    Sleep, 20
    Send, {LControl Up}
    Sleep, 180
    WinActivate, ahk_class TfrmJobEdit
    Sleep, 333
    WinWaitActive, ahk_class TfrmJobEdit
    Sleep, 333
    Send, {Tab}
    Sleep, 100
    Send, {LControl Down}
    Sleep, 30
    Send, {v}
    Send, {LControl Up}
    Sleep, 80
    Send, {Tab}
    Sleep, 1000
    IfWinExist, Customer Alert
    {
        WinActivate, Customer Alert
        Sleep, 333
        WinWaitActive, Customer Alert
        Sleep, 333
        Send, {Enter}
    }
    Loop
    {
        WinActivate, Excel
        Sleep, 333
        IfWinActive, Excel
        {
            Break
        }
    }
    Send, {Delete}
    Sleep, 50
    Loop
    {
        Send, {F5}
        IfWinExist, Go To
        {
            Break
        }
    }
    SendRaw, F1
    Sleep, 150
    Send, {Enter}
    Sleep, 100
    Send, {LControl Down}
    Sleep, 30
    Send, {c}
    Send, {LControl Up}
    Sleep, 80
    Loop
    {
        WinActivate, New Job
        Sleep, 333
        IfWinActive, New Job
        {
            Break
        }
    }
    Send, {Tab 10}
    Sleep, 150
    Sleep, 100
    Send, {LControl Down}
    Sleep, 30
    Send, {v}
    Send, {LControl Up}
    Sleep, 80
    Loop
    {
        WinActivate, Excel
        Sleep, 333
        IfWinActive, Excel
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
    SendRaw, L1
    Sleep, 150
    Send, {Enter}
    Sleep, 150
    Send, {F2}
    Sleep, 50
    Clipboard := ""
    SendRaw, =SalesAccount(E1)
    Sleep, 250
    Send, {Enter}
    Sleep, 20
    Send, {Up}
    Sleep, 20
    Delay := 100
    Loop
    {
        Sleep, %Delay%
        Send, {LControl Down}
        Sleep, 30
        Send, {c}
        Send, {LControl Up}
        Sleep, 30
        If Clipboard !=
        {
            If Clipboard != 0
            {
                Break
            }
        }
        Delay += 100
    }
    Loop
    {
        WinActivate, New Job
        Sleep, 333
        IfWinActive, New Job
        {
            Break
        }
    }
    Send, {Tab 10}
    Sleep, 150
    Sleep, 200
    Send, {LControl Down}
    Sleep, 30
    Send, {v}
    Send, {LControl Up}
    Sleep, 30
    Send, {LAlt Down}
    Sleep, 30
    Send, {o}
    Sleep, 30
    Send, {LAlt Up}
    Sleep, 30
    Sleep, 1500
    CoordMode, Pixel, Window
    PixelSearch, FoundX, FoundY, -414, -366, 1370, 504, 0xA60A0A, 0, Fast RGB
    IfWinExist, Error
    {
        Send, {Enter}
        Sleep, 100
        Send, {Escape}
        Sleep, 500
        Send, {n}
        Sleep, 50
    }
    Else
    {
        Sleep, 500
        Loop
        {
            WinActivate, JobCost
            Sleep, 333
            IfWinActive, JobCost
            {
                Break
            }
        }
        Sleep, 300
        Send, {LControl Down}
        Sleep, 30
        Send, {Tab}
        Sleep, 30
        Send, {LControl Up}
        Sleep, 30
        Sleep, 200
        Send, {LAlt Down}
        Sleep, 30
        Send, {n}
        Sleep, 30
        Send, {LAlt Up}
        Sleep, 30
        Loop
        {
            WinActivate, Excel
            Sleep, 333
            IfWinActive, Excel
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
        Sleep, 100
        SendRaw, A1
        Sleep, 150
        Send, {Enter}
        Sleep, 30
        Send, {LControl Down}
        Sleep, 30
        Send, {c}
        Sleep, 30
        Send, {LControl Up}
        Sleep, 30
        Loop
        {
            WinActivate, New Job Estimate
            Sleep, 333
            IfWinActive, New Job Estimate
            {
                Break
            }
        }
        Send, {LControl Down}
        Sleep, 30
        Send, {v}
        Sleep, 30
        Send, {LControl Up}
        Sleep, 30
        Send, {Tab}
        Sleep, 100
        WinActivate, Warning
        Sleep, 333
        CoordMode, Pixel, Window
        ImageSearch, FoundX, FoundY, 0, 0, 1920, 1080, C:\Users\frontdesk\Desktop\RECEPTION DESK\MacroCreatorPortable\x86\MacroCreator\Screenshots\Screen_20200714113754.png
        If ErrorLevel = 0
        {
            MsgBox, 49, Continue?, Image / Pixel Found at %FoundX%x%FoundY%.`nPress OK to continue.
            IfMsgBox, Cancel
            	Return
        }
        Send, {1}
        Sleep, 100
        Send, {Tab}
        Sleep, 100
        Send, {1}
        Sleep, 50
        Loop
        {
            WinActivate, Excel
            Sleep, 333
            IfWinActive, Excel
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
        Sleep, 300
        SendRaw, G1
        Sleep, 100
        Sleep, 30
        Send, {Enter}
        Sleep, 200
        Send, {LControl Down}
        Sleep, 30
        Send, {c}
        Sleep, 30
        Send, {LControl Up}
        Sleep, 30
        Loop
        {
            WinActivate, New Job Estimate
            Sleep, 333
            IfWinActive, New Job Estimate
            {
                Break
            }
        }
        Send, {Tab 5}
        Sleep, 150
        Sleep, 150
        Send, {LControl Down}
        Sleep, 30
        Send, {v}
        Sleep, 30
        Send, {LControl Up}
        Sleep, 30
        Sleep, 150
        Send, {Enter}
        Sleep, 100
    }
    Loop
    {
        WinActivate, Excel
        Sleep, 333
        IfWinActive, Excel
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
    Sleep, 300
    SendRaw, A1
    Send, {Enter}
    Sleep, 300
    Send, {LControl Down}{NumpadSub}{LControl Up}
    Sleep, 50
    Send, {r}{Enter}
    Sleep, 150
    Loop
    {
        Send, {F5}
        IfWinExist, Go To
        {
            Break
        }
    }
    Sleep, 100
    SendRaw, A1
    Send, {Enter}
    Sleep, 150
    Send, {LControl Down}{c}{LControl Up}
    Sleep, 150
    IfNotInString, Clipboard, 1
    {
        IfNotInString, Clipboard, 9
        {
            IfNotInString, Clipboard, 8
            {
                Break
            }
        }
    }
}
MsgBox, 0, , Task Complete!
Return


F8::ExitApp
