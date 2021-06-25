VERSION 5.00
Object = "{4DA5FE99-C05C-45FC-9C6C-27263198FAE6}#1.0#0"; "Hyperlink.ocx"
Begin VB.Form frmHyperlink 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Hyperlink Control Example"
   ClientHeight    =   915
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4065
   ForeColor       =   &H00404040&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   915
   ScaleWidth      =   4065
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Hyperlink.vbcHyperlink vbcHyperlink2 
      Height          =   315
      Left            =   1560
      TabIndex        =   1
      Top             =   480
      Width           =   825
      _ExtentX        =   1455
      _ExtentY        =   556
      TextColor       =   16711680
      HotColor        =   16776960
      URL             =   "mailto:contoh@contoh.com"
      Text            =   "EMail Me!"
   End
   Begin Hyperlink.vbcHyperlink vbcHyperlink1 
      Height          =   315
      Left            =   1320
      TabIndex        =   0
      Top             =   120
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   556
      TextColor       =   16711680
      HotColor        =   16776960
      URL             =   "https://google.com"
      Text            =   "Pergi ke Google!"
   End
End
Attribute VB_Name = "frmHyperlink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
