VERSION 5.00
Begin VB.Form Form1 
   ClientHeight    =   1215
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   3675
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   1215
   ScaleWidth      =   3675
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      Height          =   375
      Left            =   960
      TabIndex        =   2
      Top             =   840
      Width           =   1575
   End
   Begin VB.Timer Timer1 
      Interval        =   30
      Left            =   1560
      Top             =   120
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   480
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   480
      TabIndex        =   0
      Top             =   120
      Width           =   2775
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This is a simple example ilustrating an API Call.
'It gives you information about the memory on the
'PC.
'***************************************************
'The program is freeware.You can make money selling
'it -- YEAH, right! :) --, you can modify it, you
'can take credit for it! (see if i care!) You can do
'anything to it! Also, i wouldn't mind if you
'mentioned my name in a program made by you,
'but that's up to you.
'
'Made by Cristian Oftez (But you can call me Armageddon)
'E-Mail: armageddon_tvel@yahoo.com
'Don't send any messages unless you have a serious
'question. AND !!!PLEASE!!! DON'T SPAM!!!!
'***************************************************

'Declaration of variables
Public TotalPhysicalMemory, AvailablePhysicalMemory, TotalPageFile, AvailablePageFile, TotalVirtualMemory, AvailableVirtualMemory As Long


'Optional function used to make
'it easier
Public Function GetMemoryStats()
'declare a variable similar to
'the TYPE MEMORYSTATUS
Dim ms As Module1.MEMORYSTATUS
'Call the function in the module
'on the specified variable
Module1.GlobalMemoryStatus ms
'Assign the return values to the variables
TotalPhysicalMemory = ms.dwTotalPhys \ 1024
AvailablePhysicalMemory = ms.dwAvailPhys \ 1024
TotalPageFile = ms.dwTotalPageFile \ 1024
AvailablePageFile = ms.dwAvailPageFile \ 1024
TotalVirtualMemory = ms.dwTotalVirtual \ 1024
AvailableVirtualMemory = ms.dwAvailVirtual \ 1024
End Function


Private Sub Timer1_Timer()
'Call the function
GetMemoryStats
'Type the returned values
Label1.Caption = "You have " & TotalPhysicalMemory & " KB of RAM"
Label2.Caption = "of which " & AvailablePhysicalMemory & " KB are not used"
End Sub
Private Sub Command1_Click()
'Exit
Unload Me
End Sub

