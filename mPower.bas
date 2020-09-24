Attribute VB_Name = "mPower"
Option Explicit

'Power Monitor
'Inhibits StandBy and Sleep Modes by hooking the appropriate notification or query messages

Private Enum ApiConsts
    WM_POWER = &H48
    PWR_SUSPENDREQUEST = 1

    WM_POWERBROADCAST = &H218
    PBT_APMQUERYSUSPEND = 0
    'PBT_APMQUERYSTANDBY = 1
    'PBT_APMQUERYSUSPENDFAILED = 2
    'PBT_APMQUERYSTANDBYFAILED = 3
    'PBT_APMSUSPEND = 4
    'PBT_APMSTANDBY = 5
    'PBT_APMRESUMECRITICAL = 6
    'PBT_APMRESUMESUSPEND = 7
    'PBT_APMRESUMESTANDBY = 8

    PWR_OK = 1
    PWR_FAIL = -1
    DENY_QUERY = &H424D5144 'DQMB

    IDX_WNDPROC = -4
End Enum

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private prevProcAddress As Long
Private hWndActive      As Long

Private Declare Function Beeper Lib "kernel32" Alias "Beep" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long

Private Declare Function GetSystemPowerStatus Lib "kernel32" (lpSystemPowerStatus As SYSTEM_POWER_STATUS) As Long
Private Type SYSTEM_POWER_STATUS
    ACLineStatus        As Byte
    BatteryFlag         As Byte
    BatteryLifePercent  As Byte
    Reserved1           As Byte
    BatteryLifeTime     As Long
    BatteryFullLifeTime As Long
End Type
Private SysPwrStat      As SYSTEM_POWER_STATUS

Public Sub ActivatePowerMonitor(hWnd As Long)

    If prevProcAddress = 0 Then
        prevProcAddress = SetWindowLong(hWnd, IDX_WNDPROC, AddressOf MessageHook)
        hWndActive = hWnd
    End If

End Sub

Public Sub DeactivatePowerMonitor()

    If prevProcAddress Then
        SetWindowLong hWndActive, IDX_WNDPROC, prevProcAddress
        prevProcAddress = 0
    End If

End Sub

Private Function IsOnMainsSupply() As Boolean

    GetSystemPowerStatus SysPwrStat
    IsOnMainsSupply = (SysPwrStat.ACLineStatus = 1)

End Function

Private Function MessageHook(ByVal hWnd As Long, ByVal nMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

    MessageHook = CallWindowProc(prevProcAddress, hWnd, nMsg, wParam, lParam)
    If IsOnMainsSupply Then
        Select Case True
          Case nMsg = WM_POWER And wParam = PWR_SUSPENDREQUEST
            MessageHook = PWR_FAIL
            Beeper 1000, 50
          Case nMsg = WM_POWERBROADCAST And wParam = PBT_APMQUERYSUSPEND
            MessageHook = DENY_QUERY
            Beeper 1000, 50
        End Select
    End If

End Function

':) Ulli's VB Code Formatter V2.17.3 (2004-Jul-25 13:32) 44 + 43 = 87 Lines
