VERSION 5.00
Begin VB.Form fTestPowerMonitor 
   BackColor       =   &H0080C0FF&
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Power Monitor"
   ClientHeight    =   4425
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3645
   ForeColor       =   &H0080C0FF&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4425
   ScaleWidth      =   3645
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.Line Line1 
      BorderColor     =   &H000040C0&
      BorderWidth     =   2
      X1              =   855
      X2              =   1230
      Y1              =   1125
      Y2              =   1125
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   $"fTestPowerMonitor.frx":0000
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   1440
      Left            =   300
      TabIndex        =   2
      Top             =   1935
      Width           =   3015
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Close this form to enable those modes again."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   615
      Left            =   300
      TabIndex        =   1
      Top             =   3585
      Width           =   3015
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "If everyting works as it is supposed to, you should now not be able to put Windows into to StandBy or Sleep Mode..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   1530
      Left            =   300
      TabIndex        =   0
      Top             =   240
      Width           =   3015
   End
End
Attribute VB_Name = "fTestPowerMonitor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

    ActivatePowerMonitor hWnd

End Sub

Private Sub Form_Unload(Cancel As Integer)

    DeactivatePowerMonitor

End Sub

':) Ulli's VB Code Formatter V2.17.3 (2004-Jul-25 13:32) 1 + 14 = 15 Lines
