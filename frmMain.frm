VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Begin VB.Form frmMain 
   Caption         =   "Dictionary.com"
   ClientHeight    =   4485
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6000
   FillColor       =   &H8000000F&
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4485
   ScaleWidth      =   6000
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtWord 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   0
      TabIndex        =   1
      Text            =   "Enter word"
      Top             =   0
      Width           =   8775
   End
   Begin SHDocVwCtl.WebBrowser WbDictionary 
      Height          =   2895
      Left            =   0
      TabIndex        =   0
      Top             =   315
      Width           =   10335
      ExtentX         =   18230
      ExtentY         =   5106
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   1
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Done"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   8760
      TabIndex        =   2
      Top             =   0
      Width           =   1575
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Resize()
      
    txtWord.Width = Me.Width - Label1.Width - 60
    Label1.Left = txtWord.Width - 60
    WbDictionary.Top = Label1.Height
    WbDictionary.Height = Me.Height - 670
    WbDictionary.Width = Me.Width - 100
    
End Sub

Private Sub txtWord_GotFocus()
    txtWord.SelStart = 0: txtWord.SelLength = Len(txtWord)
End Sub

Private Sub txtWord_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then WbDictionary.Navigate ("http://www.dictionary.com/cgi-bin/dict.pl?term=" & txtWord.Text): KeyAscii = 0
End Sub

Private Sub WbDictionary_ProgressChange(ByVal Progress As Long, ByVal ProgressMax As Long)
    Dim Percent As Integer, RedBlue As Integer
            
    If ProgressMax > 0 Then
        Percent = Progress / (ProgressMax / 100)
        
        Label1.BackColor = RGB(Percent + 145, Percent + 145, Percent + 145)
        Label1.Caption = Percent & "%"
    Else
        Label1.Caption = "Done"
    End If
End Sub

