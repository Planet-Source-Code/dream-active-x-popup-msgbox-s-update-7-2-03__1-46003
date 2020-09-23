VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Popup Control Box"
   ClientHeight    =   4800
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9975
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   4800
   ScaleWidth      =   9975
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "PopUp"
      Height          =   855
      Left            =   0
      TabIndex        =   32
      Top             =   120
      Width           =   9975
      Begin VB.CommandButton Command4 
         Caption         =   "Show Popup Message !"
         Height          =   375
         Left            =   8040
         TabIndex        =   36
         Top             =   360
         Width           =   1815
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         ItemData        =   "Form1.frx":0000
         Left            =   6720
         List            =   "Form1.frx":000D
         TabIndex        =   35
         Text            =   "Combo1"
         Top             =   360
         Width           =   735
      End
      Begin VB.TextBox Text22 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   5280
         TabIndex        =   34
         Text            =   "3000"
         Top             =   360
         Width           =   495
      End
      Begin VB.TextBox Text11 
         Height          =   285
         Left            =   840
         TabIndex        =   33
         Text            =   "You Have No New Emails!"
         Top             =   360
         Width           =   3495
      End
      Begin VB.Label Label3 
         Caption         =   "Sound:"
         Height          =   255
         Left            =   6000
         TabIndex        =   40
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "Duration:"
         Height          =   255
         Left            =   4440
         TabIndex        =   39
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Message:"
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label8 
         Caption         =   "Set Duration To 0 For Ok Button"
         Height          =   255
         Left            =   4440
         TabIndex        =   37
         Top             =   120
         Width           =   2535
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Balloon"
      Height          =   3135
      Left            =   0
      TabIndex        =   1
      Top             =   1200
      Width           =   9975
      Begin VB.ComboBox Combo5 
         Height          =   315
         ItemData        =   "Form1.frx":0020
         Left            =   3600
         List            =   "Form1.frx":002A
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Top             =   2400
         Width           =   615
      End
      Begin VB.ComboBox Combo4 
         Height          =   315
         ItemData        =   "Form1.frx":0034
         Left            =   6600
         List            =   "Form1.frx":003E
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   2400
         Width           =   1455
      End
      Begin VB.Frame Frame3 
         Caption         =   "Buttons"
         Height          =   1815
         Left            =   5280
         TabIndex        =   18
         Top             =   240
         Width           =   4575
         Begin VB.OptionButton Option1 
            Caption         =   "Show Three"
            Height          =   255
            Index           =   3
            Left            =   240
            TabIndex        =   25
            Top             =   1440
            Width           =   1215
         End
         Begin VB.TextBox txt3 
            Height          =   285
            Left            =   1560
            TabIndex        =   24
            Text            =   "Ignore"
            Top             =   1440
            Width           =   1215
         End
         Begin VB.TextBox txt2 
            Height          =   285
            Left            =   1560
            TabIndex        =   23
            Text            =   "Retry"
            Top             =   1080
            Width           =   1215
         End
         Begin VB.TextBox txt1 
            Height          =   285
            Left            =   1560
            TabIndex        =   22
            Text            =   "Abort"
            Top             =   720
            Width           =   1215
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Show Two"
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   21
            Top             =   1080
            Width           =   1695
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Show One"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   20
            Top             =   720
            Width           =   1695
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Show None"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   19
            Top             =   360
            Width           =   1695
         End
         Begin VB.Label Label7 
            Caption         =   "<--2 Button Show"
            Height          =   255
            Left            =   2880
            TabIndex        =   28
            Top             =   1440
            Width           =   1575
         End
         Begin VB.Label Label6 
            Caption         =   "<--2 Button Show"
            Height          =   255
            Left            =   2880
            TabIndex        =   27
            Top             =   1080
            Width           =   1575
         End
         Begin VB.Label Label5 
            Caption         =   "<--1 Button Show"
            Height          =   255
            Left            =   2880
            TabIndex        =   26
            Top             =   720
            Width           =   1575
         End
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Custom Balloon"
         Height          =   375
         Left            =   8400
         TabIndex        =   17
         Top             =   2400
         Width           =   1455
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         Height          =   285
         Left            =   4080
         MaxLength       =   2
         TabIndex        =   16
         Text            =   "3"
         Top             =   1560
         Width           =   375
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   960
         TabIndex        =   10
         Text            =   "Arial"
         Top             =   1200
         Width           =   1455
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "Form1.frx":0052
         Left            =   960
         List            =   "Form1.frx":006E
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   720
         Width           =   1455
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "Form1.frx":008F
         Left            =   960
         List            =   "Form1.frx":009F
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   240
         Width           =   1455
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Auto close after          seconds"
         Height          =   255
         Left            =   2640
         TabIndex        =   7
         Top             =   1560
         Value           =   1  'Checked
         Width           =   2655
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Put at current mouse position"
         Height          =   255
         Index           =   0
         Left            =   2640
         TabIndex        =   6
         Top             =   240
         Value           =   -1  'True
         Width           =   2415
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Show close button"
         Height          =   255
         Left            =   2640
         TabIndex        =   5
         Top             =   1920
         Value           =   1  'Checked
         Width           =   1695
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Set position"
         Height          =   255
         Index           =   1
         Left            =   2640
         TabIndex        =   4
         Top             =   600
         Width           =   2055
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   0
         Left            =   3000
         TabIndex        =   3
         Text            =   "34"
         Top             =   960
         Width           =   375
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   1
         Left            =   4080
         TabIndex        =   2
         Text            =   "34"
         Top             =   960
         Width           =   375
      End
      Begin VB.Label Label10 
         Caption         =   "Popup Style"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5400
         TabIndex        =   41
         Top             =   2400
         Width           =   1095
      End
      Begin VB.Label Label9 
         Caption         =   "Number of text boxes on InputBox?"
         Height          =   255
         Left            =   720
         TabIndex        =   29
         Top             =   2460
         Width           =   2655
      End
      Begin VB.Label Label1 
         Caption         =   "Font Face"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   15
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Font Size"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   14
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Icon"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Y"
         Height          =   255
         Index           =   4
         Left            =   3840
         TabIndex        =   12
         Top             =   960
         Width           =   135
      End
      Begin VB.Label Label1 
         Caption         =   "X"
         Height          =   255
         Index           =   5
         Left            =   2760
         TabIndex        =   11
         Top             =   960
         Width           =   135
      End
   End
   Begin VB.Label Label1 
      Caption         =   "* Remember you must register any active-x before running as plugin"
      Height          =   255
      Index           =   0
      Left            =   2160
      TabIndex        =   0
      Top             =   4440
      Width           =   5175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'***********************************************************************
' Popup Message Box ActiveX
'------------------------------------------------------
' Written by Dream
' Date:  2nd July 2003
' Email:  baddest_attitude@hotmail.com
'*************************************
' Feel free to use and modify this code as
' you see fit.  If you make any improvements
' it would be nice if you would send me a copy.
'*************************************
' Displays Three Types Of Customizable Popup Message Box's
' Also Plays Sounds From Resource File
'*************************************
' Date:  2nd July 2003
' Added: Removed frmInput reducing size to 1 Popup form Optimizing code.
' Date:  23rd June 2003
' Added Dual Text InputBox, fixed minor resize bug in InputBox
' Date:  18th June 2003
' Added  Input Box Balloon to control
' Date:  6th June 2003
' Added customizable buttons to the balloon popup
' Automatically Resizes / Repositions the buttons according
' to how many there are on the ballon and the size of the balloon,
' which is determined by how much you write on the balloon.
 
' Credit to the authors of the original msn and balloon codes where due !
'
' Please Vote And Leave Comments!!
'***********************************************************************
Option Explicit
Private PopResponse As String
Private InputResponseA As String
Private InputResponseB As String
Dim PluginName

Public Enum IconType
    vbExclamation = 48
    vbCritical = 16
    vbInformation = 64
    vbNone = 0
End Enum

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
'for loops doevents function
Private Declare Function GetInputState Lib "user32" () As Long

Private Sub Form_Load()
    Combo1.ListIndex = 3: Combo2.ListIndex = 1: Combo3.ListIndex = 0
    Combo4.ListIndex = 0: Combo5.ListIndex = 0
    If App.PrevInstance Then 'Verifies the state of the app
    MsgBox App.EXEName & " is already running.", vbInformation
       If App.TaskVisible = True Then
            End 'Free the current app from memory
        End If
    End If
       Form1.Show 'If the app is Not load, Then load it
End Sub
     
Private Sub Command3_Click()
On Error GoTo 1
'use a variable to define the plugin
 Dim objPlugIn As Object
'Variable contains plugin's response
 Dim strResponse As String
'The format for CreateObject is [Project name].[Class module name]
 Set objPlugIn = CreateObject("Popup.DisplayMsg")
'Clear our variables
 PopResponse = vbNullString
 InputResponseA = vbNullString
 InputResponseB = vbNullString
 
'======================================================================================
'======================================================================================
'       Message     - Required - The message you want to show
'       Title       - Optional - The title of the balloon; Default is the Project Title
'       Icon        - Optional - The icon you want to show;
'                                Critical, Exclamation or Information; Default is no Icon
'       ShowClose   - Optional - Want to show the close button; Default is True
'       autoCloseTime Optional - Specify nonzero time in seconds if you want to close
'                                the popUp automatically
'       fontSize    - Optional - Specifies the font size; Default is 8
'       fontFace    - Optional - Specifies a custom font; Default is MS Sans Serif
'       PutAtCurrentMousePos   - Optional; Default true; will place the balloon at the
'                                current cursor position
'       XinPixel    - Optional - Specifies the X-coordinate in pixel where ballon is to
'                                be placed;
'                                works only if PutAtCurrentMousePos=False
'       YinPixel    - Optional - Specifies the Y-coordinate in pixel where ballon is to
'                                be placed;
'                                works only if PutAtCurrentMousePos=False
'       Button1     - Optional - Button1 visible state true/false
'       Caption1    - Optional - Button1 Caption
'       Button2     - Optional - Button2 visible state true/false
'       Caption2    - Optional - Button2 Caption
'       Button3     - Optional - Button3 visible state true/false
'       Caption3    - Optional - Button3 Caption
'       FormX       - Optional - Name of the control (Having hwnd property otherwise current
'                                Cursor position will be used) where you want the balloon to
'                                be placed
'                                works only if PutAtCurrentMousePos=False
'       vbMod       - Optional - Sets vbModal to .hWnd true/false
'       intBoxes    - Optional - Number of Textboxes on the Inputbox Popup
'       Style       - Optional - Popup Balloon Style "Balloon"/"Input"
'=======================================================================================
'=======================================================================================

'Call the dll
 strResponse = objPlugIn.popUpBalloon("See my Icon..." & vbCr & _
                            "If Im a balloon, I can self close after " & Check2 * Val(Text2) & " seconds," & vbCr & _
                            "I'm" & IIf(Check3.Value, "", " not ") & " having the close button, see my top right corner.." & vbCr & _
                            "My Font size is " & Combo2.List(Combo2.ListIndex) & vbCr & _
                            "My Font face is " & Text1 & vbCr & _
                            "I'm located at " & IIf(Option2(0).Value, "Current Cursor Position", IIf(Option2(1).Value, "(" & Text3(0) & "," & Text3(1) & ")", """" & "Click On Me Button" & """")), _
                            "I'm a customized Popup", _
                            getIcon, Check3.Value, _
                            Check2 * Val(Text2), _
                            Combo2.List(Combo2.ListIndex), _
                            Text1, _
                            Option2(0).Value, _
                            Text3(0).Text, _
                            Text3(1).Text, _
                            Option1(1).Value, _
                            txt1, _
                            Option1(2).Value, _
                            txt2, _
                            Option1(3).Value, _
                            txt3, _
                            Me, _
                            False, _
                            Combo5.Text, _
                            Combo4.Text)

  'if the plug-in returns an error, let us know ' (Screen.Width / 2) / 15  (Screen.Height / 2) / 15
   If strResponse <> vbNullString Then
        MsgBox strResponse
   End If

    Select Case Combo4.Text
        Case "Balloon"
            Do Until PopResponse <> vbNullString
               Sleep 10
               DoEvents
            Loop
            If PopResponse <> "MiscEv" Then MsgBox PopResponse
            
        Case "Input"
            Do Until InputResponseA <> vbNullString
               Sleep 10
               DoEvents  ' If GetInputState() <> 0 Then
            Loop
            
            Select Case Combo5.Text
                Case 1: MsgBox InputResponseA
                Case 2: MsgBox InputResponseA & " & " & InputResponseB
            End Select
    End Select
  
Exit Sub
1:
    Select Case Err.Number
        Case 429 'can't create object
            'The ProgID can't be found. Either it is misspelled or the component hasn't been registered!
            MsgBox "You have selected an invalid plug-in ID. Please check that the name is correct and the component is registered."
            Exit Sub
        Case 5 'Invalid proceedure call or argument
            'The 'popUpBalloon' function cannot be found in the class module
            MsgBox "The plug-in you have selected does not have a valid entry point. Please verify the object module with specified guidelines."
            Exit Sub
        Case Else
              MsgBox Err.Number & "  " & Err.Description
    End Select
End Sub

Private Sub Command4_Click()
    On Error GoTo errhandler
    'use a variable to define the plugin
    Dim objPlugIn As Object
    'Variable contains plugin's response
    Dim strResponse As String
    Dim Indentity As String
    'The format for CreateObject is [Project name].[Class module name]
    Set objPlugIn = CreateObject("Popup.DisplayMsg")

    'Call the entry function
    strResponse = objPlugIn.DisplayAlert(Me, Text11, Text22, Combo3.Text)
    'if the plugin contains an error, show us in a message box
    If strResponse <> vbNullString Then
        MsgBox strResponse
    End If
    Me.SetFocus
    Exit Sub

errhandler:
    Select Case Err.Number
        Case 429 'can't create object
            'The ProgID can't be found. Either it is misspelled or the component hasn't been registered!
            MsgBox "You have selected an invalid plug-in ID. Please check that the name is correct and the component is registered."
            Exit Sub
        Case 5 'Invalid proceedure call or argument
            'The 'DisplayAlert' function cannot be found in the class module
            MsgBox "The plug-in you have selected does not have a valid entry point. Please verify the object module with specified guidelines."
            Exit Sub
        Case Else
              MsgBox Err.Number & "  " & Err.Description
    End Select
End Sub

Private Function getIcon() As IconType
    Select Case Combo1.ListIndex
        Case 0: getIcon = vbNone
        Case 1: getIcon = vbCritical
        Case 2: getIcon = vbExclamation
        Case 3: getIcon = vbInformation
    End Select
End Function

Public Sub MsgBack(Message As String)
    PopResponse = Message
End Sub

Public Sub InputBack(Message As String)
    Dim arr() As String
    arr() = Split(Message, ",")
    InputResponseA = arr(LBound(arr()))
    InputResponseB = arr(UBound(arr()))
End Sub
