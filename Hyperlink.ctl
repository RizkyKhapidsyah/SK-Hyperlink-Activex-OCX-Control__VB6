VERSION 5.00
Begin VB.UserControl vbcHyperlink 
   ClientHeight    =   315
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   780
   ScaleHeight     =   315
   ScaleWidth      =   780
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   360
      Top             =   0
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hyperlink"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   660
   End
End
Attribute VB_Name = "vbcHyperlink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long


Private Type POINTAPI
        x As Long
        y As Long
End Type

'******************************************************************************
'Internal variables
'******************************************************************************

Private mvarTextColor As Variant
Private mvarHotColor As Variant
Private mvarURL As Variant
Private mvarText As Variant
Private mvarTop As Integer
Private mvarLeft As Integer

'Control properties

Public Property Let Text(ByVal vData As Variant)
    mvarText = vData
    Label1.Caption = mvarText
    ResizeControl
End Property

Public Property Get Text() As Variant
    If IsObject(mvarText) Then
        Set Text = mvarText
    Else
        Text = mvarText
    End If
End Property

Public Property Let URL(ByVal vData As Variant)
    mvarURL = vData
    Label1.ToolTipText = mvarURL
End Property

Public Property Get URL() As Variant
    If IsObject(mvarURL) Then
        Set URL = mvarURL
    Else
        URL = mvarURL
    End If
End Property

Public Property Let TextColor(ByVal vData As Variant)
    mvarTextColor = vData
    Label1.ForeColor = mvarTextColor
End Property

Public Property Get TextColor() As Variant
    If IsObject(mvarTextColor) Then
        Set TextColor = mvarTextColor
    Else
        TextColor = mvarTextColor
    End If
End Property

Public Property Let HotColor(ByVal vData As Variant)
    mvarHotColor = vData
End Property

Public Property Get HotColor() As Variant
    If IsObject(mvarHotColor) Then
        Set HotColor = mvarHotColor
    Else
        HotColor = mvarHotColor
    End If
End Property

'******************************************************************************
'Events
'******************************************************************************

Private Sub Label1_Click()
    GoURL (mvarURL)
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    'set the text color to the alternate color
    Label1.ForeColor = mvarHotColor
    'reset the timer
    With Timer1
        .Interval = 5000
        .Enabled = True
    End With
End Sub

Private Sub Timer1_Timer()
    'reset the color of the link if there has been no movement
    Label1.ForeColor = mvarTextColor
    Timer1.Enabled = False
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    'reset text color as cursor moves away from text
    Label1.ForeColor = mvarTextColor
    Timer1.Enabled = False
End Sub

Private Sub ResizeControl()
    'resizes the control to the length of the text
    UserControl.Width = Label1.Width + 120
    UserControl.Height = Label1.Height + 120
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    'Read properties
    mvarTextColor = PropBag.ReadProperty("TextColor", &HFF0000)
    mvarHotColor = PropBag.ReadProperty("HotColor", &HFFFF00)
    mvarURL = PropBag.ReadProperty("URL", "http://www.freevbcode.com/")
    mvarText = PropBag.ReadProperty("Text", "My Link")
    'Set Properties internally
    With Label1
        .Caption = mvarText
        .ForeColor = mvarTextColor
    End With
    ResizeControl
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    'save user defined properties
    With PropBag
        .WriteProperty "TextColor", mvarTextColor
        .WriteProperty "HotColor", mvarHotColor
        .WriteProperty "URL", mvarURL
        .WriteProperty "Text", mvarText
    End With
End Sub

'******************************************************************************
'Methods
'******************************************************************************

Public Sub GoURL(Destination As Variant)
    On Error GoTo ErrHandler
    'check and see if there is a valid link
    If Destination = "" Then
        'The programmer did not enter a link
        Err.Raise 100
    End If
        'execute the link
        ShellExecute hWnd, "open", Destination, vbNullString, vbNullString, conSwNormal
    Exit Sub
ErrHandler:
    'need i say more?
    MsgBox "Could not open URL " & Chr(10) & Destination, vbCritical, "Hyperlink"
End Sub
