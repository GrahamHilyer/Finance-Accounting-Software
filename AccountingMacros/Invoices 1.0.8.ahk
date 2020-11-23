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


WIPINVOICES:
SendMode, Input
SetNumLockState, On
SetScrollLockState, Off
MsgBox, 4, , 
(LTrim
Welcome to the Sands invoicing macro. This program has been written to quickly and accurately populate a Sands invoice in Adagio.

Before proceeding please ensure that Adagio is opened with an invoice created and that Google chrome is opened and logged into the work order system. 

For best results have as few programs open as possible and do not touch the keyboard or mouse until prompted. 

Would you like to include the excel project line in the ship to?
)
IfWinNotExist, Invoices
{
    MsgBox, 0, Error, Adagio Invoice does not seem to be opened. Program will now terminate.
    ExitApp
}
IfWinNotExist, WorkOrderReport
{
    MsgBox, 0, , Work order report does not seem to be open in excel. Program will now terminate. 
    ExitApp
}
IfWinNotExist, Invoices, Invoice
{
    MsgBox, 0, , Please ensure invoice is open. Program will now terminate. 
    ExitApp
}
BlockInput, ON
IfMsgBox, Yes
{
    AddressFormat := 1
}
Else
{
    AddressFormat := 0
}
Sleep, 600
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
Sleep, 400
Loop
{
    Send, {F5}
    IfWinExist, Go To
    {
        Break
    }
}
SendRaw, C8
Send, {Enter}
While Clipboard!=""
{
    Clipboard := ""
}
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 20
SendEvent, {F2}^a^c{Escape}
SetKeyDelay, %CurrentKeyDelay%
ClipWait, 0.2
Description := Clipboard
If ErrorLevel = 1
{
    Description := ""
    ErrorLevel := 0
}
Loop
{
    Send, {F5}
    IfWinExist, Go To
    {
        Break
    }
}
SendRaw, L6
Sleep, 150
Send, {Enter}
Sleep, 100
While Clipboard!=""
{
    Clipboard := ""
}
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 20
SendEvent, {F2}^a^c{Escape}
SetKeyDelay, %CurrentKeyDelay%
ClipWait, 0.2
SalesConv()
Loop
{
    Send, {F5}
    IfWinExist, Go To
    {
        Break
    }
}
SendRaw, E18
Sleep, 100
Send, {Enter}
Sleep, 100
While Clipboard!=""
{
    Clipboard := ""
}
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 20
SendEvent, {F2}^a^c{Escape}
SetKeyDelay, %CurrentKeyDelay%
ClipWait, 0.2
CustPO := Clipboard
If ErrorLevel = 1
{
    CustPO := ""
    ErrorLevel := 0
}
Loop
{
    Send, {F5}
    IfWinExist, Go To
    {
        Break
    }
}
SendRaw, L2
Sleep, 100
Send, {Enter}
Sleep, 100
While Clipboard!=""
{
    Clipboard := ""
}
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 20
SendEvent, {F2}^a^c{Escape}
SetKeyDelay, %CurrentKeyDelay%
ClipWait, 0.2
WOValue := Clipboard
If ErrorLevel = 1
{
    WOValue := ""
    ErrorLevel := 0
}
Loop
{
    Send, {F5}
    IfWinExist, Go To
    {
        Break
    }
}
SendRaw, A18
Sleep, 100
Send, {Enter}
Sleep, 100
While Clipboard!=""
{
    Clipboard := ""
}
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 20
SendEvent, {F2}^a^c{Escape}
SetKeyDelay, %CurrentKeyDelay%
ClipWait, 0.2
Contact := Clipboard
If ErrorLevel = 1
{
    Contact := ""
    ErrorLevel := 0
}
Loop
{
    Send, {F5}
    IfWinExist, Go To
    {
        Break
    }
}
SendRaw, I12
Sleep, 100
Send, {Enter}
Sleep, 100
While Clipboard!=""
{
    Clipboard := ""
}
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 20
SendEvent, {F2}^a^c{Escape}
SetKeyDelay, %CurrentKeyDelay%
ClipWait, 0.2
Shipto1 := Clipboard
If ErrorLevel = 1
{
    Shipto1 := ""
    ErrorLevel := 0
}
Loop
{
    Send, {F5}
    IfWinExist, Go To
    {
        Break
    }
}
SendRaw, I13
Sleep, 100
Send, {Enter}
Sleep, 100
While Clipboard!=""
{
    Clipboard := ""
}
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 20
SendEvent, {F2}^a^c{Escape}
SetKeyDelay, %CurrentKeyDelay%
ClipWait, 0.2
Shipto2 := Clipboard
If ErrorLevel = 1
{
    Shipto2 := ""
    ErrorLevel := 0
}
Loop
{
    Send, {F5}
    IfWinExist, Go To
    {
        Break
    }
}
SendRaw, I14
Sleep, 100
Send, {Enter}
Sleep, 100
While Clipboard!=""
{
    Clipboard := ""
}
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 20
SendEvent, {F2}^a^c{Escape}
SetKeyDelay, %CurrentKeyDelay%
ClipWait, 0.2
Shipto3 := Clipboard
If ErrorLevel = 1
{
    Shipto3 := ""
    ErrorLevel := 0
}
Loop
{
    Send, {F5}
    IfWinExist, Go To
    {
        Break
    }
}
SendRaw, I15
Sleep, 100
Send, {Enter}
Sleep, 100
While Clipboard!=""
{
    Clipboard := ""
}
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 20
SendEvent, {F2}^a^c{Escape}
SetKeyDelay, %CurrentKeyDelay%
ClipWait, 0.2
Shipto4 := Clipboard
If ErrorLevel = 1
{
    Shipto4 := ""
    ErrorLevel := 0
}
Loop
{
    Send, {F5}
    IfWinExist, Go To
    {
        Break
    }
}
SendRaw, I16
Sleep, 100
Send, {Enter}
Sleep, 100
While Clipboard!=""
{
    Clipboard := ""
}
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 20
SendEvent, {F2}^a^c{Escape}
SetKeyDelay, %CurrentKeyDelay%
ClipWait, 0.2
Shipto5 := Clipboard
If ErrorLevel = 1
{
    Shipto5 := ""
    ErrorLevel := 0
}
Loop
{
    WinActivate, Invoices
    Sleep, 333
    IfWinActive, Invoices
    {
        Break
    }
}
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 40
SendEvent, !h!v
SetKeyDelay, %CurrentKeyDelay%
SendRaw, %Description%
Sleep, 100
Send, {Tab 2}
Sleep, 100
SendRaw, %Reference%
Sleep, 100
Send, {Tab}
Sleep, 100
SendRaw, %CustPO%
Sleep, 100
Send, {Tab}
Sleep, 50
SendRaw, %WOValue%
Sleep, 100
Send, {Tab 2}
Sleep, 100
SendRaw, %Header%
Sleep, 100
Send, {Tab 6}
Sleep, 50
SendRaw, %Contact%
Sleep, 100
BlockInput, OFF
MsgBox, 0, , Please review the tax information and return the cursor to the "ship via" input box after making any necessary changes.
BlockInput, ON
Sleep, 600
Loop
{
    WinActivate, Invoices
    Sleep, 333
    IfWinActive, Invoices
    {
        Break
    }
}
Send, {Tab 6}
Sleep, 100
SendRaw, %WOValue%
Sleep, 200
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 50
Loop, 5
{
    SendEvent, +{Tab}
}
SetKeyDelay, %CurrentKeyDelay%
IfWinActive, Warning
{
    Sleep, 100
    Send, {Escape}
}
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 30
SendEvent, !v!b
SetKeyDelay, %CurrentKeyDelay%
BlockInput, Off
MsgBox, 0, , Please review the bill to information and email recipients before hitting OK
BlockInput, ON
Sleep, 600
Loop
{
    WinActivate, Invoices
    Sleep, 333
    IfWinActive, Invoices
    {
        Break
    }
}
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 30
SendEvent, {Tab}!h!v!s
SetKeyDelay, %CurrentKeyDelay%
If AddressFormat = 1
{
    Send, {Tab}
    Sleep, 100
    SendRaw, %Description%
    Sleep, 200
}
Send, {Tab}
Sleep, 100
SendRaw, %Shipto1%
Sleep, 100
If AddressFormat = 1
{
    Send, {Tab 2}
    Sleep, 50
}
Else
{
    Send, {Tab}
    Sleep, 50
}
SendRaw, %Shipto2%
Sleep, 100
If AddressFormat = 0
{
    Send, {Tab 2}
    Sleep, 50
}
Else
{
    Send, {Tab}
    Sleep, 50
}
SendRaw, %Shipto3%
Sleep, 100
Send, {Tab}
Sleep, 100
SendRaw, %Shipto4%
Sleep, 100
Send, {Tab}
Sleep, 100
SendRaw, %Shipto5%
Sleep, 100
Send, {Tab}
Sleep, 100
Send, {Delete}
Sleep, 100
Send, {Tab}
Sleep, 100
Send, {Delete}
Sleep, 100
Send, {Tab}
Sleep, 100
Send, {Delete}
Sleep, 100
BlockInput, OFF
MsgBox, 0, , Please review the ship to information before hitting OK
BlockInput, ON
Sleep, 200
Run, "C:\Program Files \Google\Chrome\Application\chrome.exe" http://win2012:8088/WorkOrder/Invoices?woNo=%WOValue%, , UseErrorLevel
Run, "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe" http://win2012:8088/WorkOrder/Invoices?woNo=%WOValue%, , UseErrorLevel
Loop
{
    WinActivate, Invoices
    Sleep, 333
    IfWinActive, Invoices
    {
        Break
    }
}
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 30
SendEvent, !h!v!t
SetKeyDelay, %CurrentKeyDelay%
Send, {Tab}
Sleep, 50
Send, ^c
Sleep, 50
InvoiceNum := Clipboard
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
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 30
    SendEvent, ^a^c
    SetKeyDelay, %CurrentKeyDelay%
    IfInString, Clipboard, Invoice #
    {
        Break
    }
}
SendRaw, ^+{Home}
Sleep, 30
Send, {Tab 18}
Sleep, 100
Send, {Enter}
Sleep, 1000
Send, {Tab 3}
Sleep, 50
SendRaw, %InvoiceNum%
Sleep, 200
Send, {Tab 2}
Sleep, 100
Send, {Enter}
Sleep, 1000
Run, "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe" http://win2012:8088/WorkOrder/Edit?woNo=%WOValue%
Loop
{
    WinActivate, Invoices
    Sleep, 333
    IfWinActive, Invoices
    {
        Break
    }
}
Send, !i
Sleep, 50
Send, {n}
SendRaw, TOT--
Sleep, 150
Send, {Enter}
Sleep, 500
Send, {e}
Sleep, 50
Send, {Tab}
Sleep, 50
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
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 30
    SendEvent, ^a^c
    SetKeyDelay, %CurrentKeyDelay%
    IfInString, Clipboard, Comments
    {
        Break
    }
}
SendRaw, ^+{Home}
Sleep, 50
Send, {Tab 31}
Sleep, 50
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 30
SendEvent, ^a^c
SetKeyDelay, %CurrentKeyDelay%
Loop
{
    WinActivate, Chrome
    Sleep, 333
    IfWinActive, Chrome
    {
        Break
    }
}
Loop
{
    WinActivate, Invoices
    Sleep, 333
    IfWinActive, Invoices
    {
        Break
    }
}
Send, ^v
Sleep, 30
Send, {Tab}
Sleep, 50
StringLen, Len1, Clipboard
StringGetPos, Pos1, Clipboard, $, R1
StringGetPos, Pos2, Clipboard, HST, R1
Len1 -= Pos2
StringTrimRight, Clipboard, Clipboard, %Len1%
StringTrimLeft, Clipboard, Clipboard, %Pos1%
Send, ^v
Sleep, 30
Send, {Tab}
Sleep, 30
BlockInput, OFF
MsgBox, 0, , 
(LTrim
Task complete. Please finish filling out the invoicing lines. 

Remember to create a holdback invoice if required. 
)
Return

SalesConv()
{
    global
    If Clipboard contains BREEN
    {
        Account := 3104
        Reference := "BREEN"
        Header := "BREEN"
        Initials := "DA"
    }
    If Clipboard contains PETER
    {
        Account := 3101
        Header := "BREEN"
        Reference := "SANDS"
        Initials := "PE"
    }
    If Clipboard contains WEATHERALL
    {
        Account := 3107
        Header := "WEATHE"
        Reference := "WEATHERALL"
        Initials := "WE"
    }
    If Clipboard contains JEFFEREY
    {
        Account := 3108
        Header := "JEFFER"
        Reference := "JEFFEREY"
        Initials := "DJ"
    }
    If Clipboard contains DUNK
    {
        Account := 3110
        Header := "DUNK"
        Reference := "DUNK"
        Initials := "AD"
    }
    If Clipboard contains RUFFIEUX
    {
        Account := 3106
        Header := "RUFFIE"
        Reference := "RUFFIEUX"
        Initials := "ER"
    }
    If Clipboard contains CAHILL
    {
        Header := "CAHILL"
        Reference := "CAHILL"
        Account := 3102
        Initials := "KC"
    }
    If Clipboard contains BRYAN
    {
        Reference := "SANDS B"
        Header := "BRYAN"
        Account := 3114
        Initials := "BS"
    }
    If Clipboard contains THOMPSON
    {
        Reference := "THOMPSON"
        Header := "THOMPS"
        Account := 3121
        Initials := "MT"
    }
    If Clipboard contains WADDEN
    {
        Account := 3119
        Header := "WADDEN"
        Reference := "WADDEN"
        Initials := "DW"
    }
    If Clipboard contains PARIC
    {
        Account := 3122
        Header := "PARIC"
        Reference := "PARIC"
        Initials := "DP"
    }
    If Clipboard contains COWARD
    {
        Account := 3113
        Header := "COWARD"
        Reference := "COWARD"
        Initials := "WC"
    }
    If Clipboard contains Barclay
    {
        Account := 3115
        Header := "BARCLAY"
        Reference := "BARCLAY"
        Initials := "PB"
    }
    If Clipboard contains NELSON
    {
        Reference := "BREEN/NELSON"
        Header := "NELSON"
        Account := 3103
        Initials := "JN"
    }
    If Clipboard contains Owen
    {
        Reference := "OWEN"
        Header := "RJOWEN"
        Account := 3111
        Initials := "RJ"
    }
}


F8::ExitApp
