VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Form1"
   ClientHeight    =   3720
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4065
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   4065
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command29 
      Caption         =   "Cascade Wins"
      Height          =   255
      Left            =   2760
      TabIndex        =   28
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton Command28 
      Caption         =   "Tile Wins"
      Height          =   255
      Left            =   2760
      TabIndex        =   27
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton Command27 
      Caption         =   "UnMap Net drv"
      Height          =   255
      Left            =   2760
      TabIndex        =   26
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton Command26 
      Caption         =   "System Error"
      Height          =   255
      Left            =   2760
      TabIndex        =   25
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton Command25 
      Caption         =   "Kill Windows"
      Height          =   255
      Left            =   2760
      TabIndex        =   24
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton Command24 
      Caption         =   "Map Net drv"
      Height          =   255
      Left            =   2760
      TabIndex        =   23
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton Command23 
      Caption         =   "Swap Buttons"
      Height          =   255
      Left            =   2760
      TabIndex        =   22
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton Command22 
      Caption         =   "Load Excels"
      Height          =   255
      Left            =   2760
      TabIndex        =   21
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton Command21 
      Caption         =   "Repaint Scrn"
      Height          =   255
      Left            =   2760
      TabIndex        =   20
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Command20 
      Caption         =   "Cascade Wins"
      Height          =   255
      Left            =   1440
      TabIndex        =   19
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton Command19 
      Caption         =   "Tile Windows"
      Height          =   255
      Left            =   1440
      TabIndex        =   18
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton Command18 
      Caption         =   "Play  .wav"
      Height          =   255
      Left            =   1440
      TabIndex        =   17
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton Command17 
      Caption         =   "Restart Win"
      Height          =   255
      Left            =   1440
      TabIndex        =   16
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton Command16 
      Caption         =   "Reboot Sys"
      Height          =   255
      Left            =   1440
      TabIndex        =   15
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton Command15 
      Caption         =   "Exit Windows"
      Height          =   255
      Left            =   1440
      TabIndex        =   14
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton Command14 
      Caption         =   "Start Dialup"
      Height          =   255
      Left            =   1440
      TabIndex        =   13
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Quick Restart"
      Height          =   255
      Left            =   1440
      TabIndex        =   12
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Clear Uninstall"
      Height          =   255
      Left            =   1440
      TabIndex        =   11
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Open Inbox"
      Height          =   255
      Left            =   1440
      TabIndex        =   10
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Print To"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Print HTML"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Browse Folder"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Control Panel"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Open With..."
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Screen Saver"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Telnet"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "News"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Mail"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "URL"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "AntiSociaL"
      Height          =   195
      Left            =   3000
      TabIndex        =   29
      Top             =   3360
      Width           =   765
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Shell "rundll32.exe shdocvw.dll,OpenURL favorite.url", vbMaximizedFocus 'Opens a favoreite internet shortcut .url file
End Sub

Private Sub Command10_Click()
Shell "rundll32.exe MSHTML.DLL,PrintHTML filename.html Output.txt", vbNormalFocus 'Dump an HTML to txt file (i havent tested this one)
End Sub

Private Sub Command11_Click()
Shell "rundll32.exe C:\PROGRA~1\INTERN~1\HMMAPI.DLL,OpenInboxHandler", vbNormalFocus ' opens hotmail inbox on IE
End Sub

Private Sub Command12_Click()
Shell "rundll32.exe setupapi.dll,InstallHinfSection DiskCleanup.Uninstall 0 setupc.inf" 'delets the windows 98 uninstall info, might need to be modified to work on other machines
End Sub

Private Sub Command13_Click()
Shell "RUNDLL.EXE user.exe,exitwindowsexec", vbNormalFocus
End Sub

Private Sub Command14_Click()
Shell "rundll32.exe rnaui.dll,RnaDial Connection Name", vbNormalFocus 'Start a dialup connection dialog
'use the following line to have VB send Enter (hits default button which is "Connect")
'SendKeys "{enter}"
End Sub

Private Sub Command15_Click()
Shell "RUNDLL.EXE user.exe,exitwindows", vbNormalFocus
End Sub

Private Sub Command16_Click()
Shell "RUNDLL.EXE user.exe,rebootsystem", vbNormalFocus
End Sub

Private Sub Command17_Click()
Shell "RUNDLL.EXE user.exe,restartwindows, vbNormalFocus"
End Sub

Private Sub Command18_Click()
Shell "rundll32.exe amovie.ocx,RunDll /play or /close Wav Filename", vbNormalFocus
End Sub

Private Sub Command19_Click()
Shell "rundll32.exe user.exe,TILECHILDWINDOWS", vbNormalFocus
End Sub

Private Sub Command2_Click()
Shell "rundll32.exe url.dll,MailToProtocolHandler email@address.com", vbNormalFocus
End Sub

Private Sub Command20_Click()
Shell "rundll32.exe user.exe,CASCADEWINDOWS", vbNormalFocus
End Sub

Private Sub Command21_Click()
Shell "rundll32.exe user.exe,REPAINTSCREEN", vbNormalFocus
End Sub

Private Sub Command22_Click()
Shell "rundll32.exe user.exe,LOADACCELERATORS", vbNormalFocus
End Sub

Private Sub Command23_Click()
Shell "rundll32.exe user.exe,SWAPMOUSEBUTTON", vbNormalFocus
MsgBox "To get your buttons back to normal goto the Mouse menu on the Control Panel", vbOKOnly, "Sorry I don't got the switchback"
End Sub

Private Sub Command24_Click()
Shell "rundll32.exe user.exe,WNETCONNECTIONDIALOG", vbNormalFocus
End Sub

Private Sub Command25_Click()
Shell "rundll32.exe user.exe,DISABLEOEMLAYER", vbNormalFocus 'ENABLEOEMLAYER also exists but I'm not sure how you can invoke it after DISABLE has been invoked
End Sub

Private Sub Command26_Click()
Shell "rundll32.exe user.exe,SYSERRORBOX", vbNormalFocus
End Sub

Private Sub Command27_Click()
Shell "rundll32.exe user.exe,WNETDISCONNECTDIALOG", vbNormalFocus
End Sub

Private Sub Command28_Click()
Shell "rundll32.exe user.exe,TILECHILDWINDOWS", vbNormalFocus
End Sub

Private Sub Command29_Click()
Shell "rundll32.exe user.exe,CASCADECHILDWINDOWS", vbNormalFocus
End Sub

Private Sub Command3_Click()
Shell "rundll32.exe url.dll,NewsProtocolHandler News.Arress.net", vbNormalFocus
End Sub

Private Sub Command4_Click()
Shell "rundll32.exe url.dll,TelnetProtocolHandler 127.0.0.1:80", vbNormalFocus 'Format IP : Port
End Sub

Private Sub Command5_Click()
Shell "rundll32.exe desk.cpl,InstallScreenSaver", vbNormalFocus 'You can use "rundll32.exe desk.cpl,InstallScreenSaver ScreenSaverName.scr" to install a specific screen saver
End Sub

Private Sub Command6_Click()
Shell "rundll32.exe shell32.dll,OpenAs_RunDLL", vbNormalFocus 'Show 'Open with' dialog
End Sub

Private Sub Command7_Click()
Shell "rundll32.exe shell32.dll,Control_RunDLL", vbNormalFocus
End Sub

Private Sub Command8_Click()
Shell "rundll32.exe url.dll,FileProtocolHandler", vbNormalFocus 'opens a webbrwoser to C:\
End Sub

Private Sub Command9_Click()
Shell "rundll32.exe MSHTML.DLL,PrintHTML filename.html", vbNormalFocus 'prints a HTML file, may work with other stuff too
End Sub
