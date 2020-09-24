VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2685
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   179
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Transparent"
      Height          =   375
      Left            =   3420
      TabIndex        =   2
      Top             =   480
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Done"
      Height          =   420
      Left            =   1755
      TabIndex        =   1
      Top             =   2160
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   240
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   105
      Width           =   4260
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const GWL_EXSTYLE = -20
Private Const LWA_COLORKEY = 1
Private Const LWA_ALPHA = 2
Private Const WS_EX_LAYERED = &H80000

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal cKey As Long, ByVal bAlpha As Long, ByVal dwFlags As Long) As Long


Private Sub Command1_Click()
   Call Unload(Me)
End Sub

Private Sub Command2_Click()
   Call SetWindowLong(Form1.hwnd, GWL_EXSTYLE, GetWindowLong(Form1.hwnd, GWL_EXSTYLE) Or WS_EX_LAYERED)
   '/* change the '50' to reflect a transparent percentage from 0-100 */
   Call SetLayeredWindowAttributes(Form1.hwnd, 0, (255 * 50) / 100, LWA_ALPHA)
End Sub
