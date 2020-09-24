VERSION 5.00
Begin VB.Form frmInfoMouse 
   Caption         =   "Info Mouse"
   ClientHeight    =   1590
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3870
   LinkTopic       =   "Form1"
   ScaleHeight     =   1590
   ScaleWidth      =   3870
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   420
      Left            =   2520
      TabIndex        =   5
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton cmdAction 
      Caption         =   "Start"
      Height          =   420
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Coordinates"
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1095
      Begin VB.Label lblY 
         Alignment       =   2  'Center
         Caption         =   "Y = "
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   855
      End
      Begin VB.Label lblX 
         Alignment       =   2  'Center
         Caption         =   "X = "
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Label lblInfo 
      Height          =   735
      Left            =   1320
      TabIndex        =   3
      Top             =   240
      Width           =   2415
   End
End
Attribute VB_Name = "frmInfoMouse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'the 'POINTAPI' type is declared
Dim COORDINATES As POINTAPI

'this will be used by the main 'Do Loop' to know whether to stop or go on
Dim STATUS As Integer

'variable that will hold the current window value
Dim currentWindow As Long
'I think you can figure this one out
Dim previousWindow As Long

'will hold a title of a window or part of window
Dim className As String
'not used for anything significant
Dim retVal As Long

Private Sub cmdAction_Click()

'if the button 'Start' has not been clicked then
If cmdAction.Caption = "Start" Then
      
   previousWindow = 0

   'set the 'Start' button caption to "Stop"
   cmdAction.Caption = "Stop"
   
   'setting STATUS to 1 here, so that the following 'Do Loop' could go on
   STATUS = 1
   
   'the famouse 'Do Loop'
   Do
     
     'if we haven't set STATUS to 1 then it'd be equal to 0, and this loop wouldn't
     'work
     If STATUS = 0 Then Exit Do
     
     'this API function returns the position of the cursor with Y and X
     Call GetCursorPos(COORDINATES)
     
     'setting up the labels...
     lblX.Caption = "X = " & COORDINATES.X
     lblY.Caption = "Y = " & COORDINATES.Y
     
     'this statement finds out the hWnd value of the window
     '(or part of window) where the cursor's at
     currentWindow = WindowFromPoint(COORDINATES.X, COORDINATES.Y)
     
     If currentWindow <> previousWindow Then
     
        'sets up the buffer
        className = Space(256)
        
        previousWindow = currentWindow
        
        'gets the name of a window(or part of window) where the cursor is currently at
        'this statement puts the title into the buffer
        retVal = GetClassName(currentWindow, className, 255)
        
        'if the name returned is "SysListView32", actually if only a piece of returned
        'string is "SysLstVew32".
        '"Why?" you ask
        'the answer is very simple:
        'the Win32 API DLL files are written with C, and as I hope you know that C
        'adds the so-called NullChar in the end of every string, so I, instead
        'of making a function to remove that null char, just simply read the first 13
        'letters(26 bytes) of the string
        If Left(className, 13) = "SysListView32" Then
           'if the conditins above are met then the string is substituted with 'Desktop'
           className = "Desktop"
           
        End If
            'sets the label(lblInfo)'s caption to the following...
            lblInfo.Caption = "Position Info: " & className
                      
      End If
    'this is what makes you able to clck 'Stop'
    'this simply makes a break in the loop for windows to finish up other tasks **
    DoEvents
    
    Loop
    
Else
   'if the 'Start' button has been clck it becomes the 'Stop' button, so now this
   'is what the 'Stop' button does, sets everything to like it was when the
   'form was loaded...well, almost everything
   
   cmdAction.Caption = "Start"
   
   STATUS = 0
   
   lblInfo = "Program: Stopped"
   
End If
'** - I think ***
'*** - almost sure ****
'**** - for sure
End Sub

Private Sub cmdExit_Click()
 
'unloads the form
Unload frmInfoMouse

'end the run
End

End Sub
