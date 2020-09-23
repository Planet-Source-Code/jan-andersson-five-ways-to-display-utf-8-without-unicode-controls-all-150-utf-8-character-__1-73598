VERSION 5.00
Begin VB.Form frmP 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   4035
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5400
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00800000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   269
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   360
   StartUpPosition =   3  'Windows Default
   Begin VB.VScrollBar vsc 
      CausesValidation=   0   'False
      Height          =   1575
      LargeChange     =   100
      Left            =   4260
      SmallChange     =   10
      TabIndex        =   1
      Top             =   660
      Width           =   255
   End
   Begin VB.PictureBox pic 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1755
      Left            =   540
      ScaleHeight     =   117
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   181
      TabIndex        =   0
      Top             =   480
      Width           =   2715
   End
   Begin VB.Menu mnuPop 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuP 
         Caption         =   "Copy"
         Index           =   0
      End
      Begin VB.Menu mnuP 
         Caption         =   "Paste"
         Index           =   1
      End
   End
End
Attribute VB_Name = "frmP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*************************************************************************************************
'**************************   Display UTF-8 & modW.bas November 17 2010   ************************
'*************************************************************************************************
'*  Copyright© Pappsegull Sweden, http://freetranslator.webs.com <pappsegull@yahoo.se>
'*
'*  This form is used to demonstrate the capabilities of the module modW.bas.
'*
'*  If you see lots of squares instead of characters, install Far East language support.
'*  Have a look in Windows help if you do not know how to.
'*
'* FEATURES
'* --------
'*  Five ways of displaying UTF-8 characters without unicode controls.
'*  All 150 UTF-8 character code tables with enum, the name and character code range.
'*  Open/Save files as UTF-8 or ANSI.
'*  Copy and paste unicode text.
'*  Storing UTF-8 without resource file/property page, 50+ different languages stored in the form.
'*  Functions to easily store your own UTF-8 texts.
'*  I'm into an ActiveX control that can quickly translate text, so this feature will-
'*  hopefully be redundant soon;-)
'*
'* This software is provided "as-is," without any express or implied warranty.
'* In no event shall the author be held liable for any damages arising from the use of this software.
'* If you do not agree with these terms, do not use this code.
'*
'* CREDITS
'* -------
'*  Planet Source & Vesa Piittinen aka Merri for his work on the unicode controls.
'*  I have also use his code CaptionWLet() & CaptionWGet() in my modW.bas
'*  They did inspired me to my other project, "Free SRT-File Translator"
'*  (http://freetranslator.webs.com) and the translation ActiveX (GAT) I'm working on.
'*  Where I use Merri's UniLabel and UniText controls. Paljon kiitoksia Merri :-)
'*  Long live VB6, I'm to old and tired to learn some new programming language :-P
'*
'*************************************************************************************************

Option Explicit
Dim sTextW$, Margin%, RTL As Boolean, vscFactor!: Const Max = 32767

Sub PrintTextW(TextW$, Optional PercentMargin%, Optional RightToLeft As Boolean)
    Icon = frm.Icon: sTextW$ = TextW$: RTL = RightToLeft: Caption = App.Title
    Margin% = PercentMargin%: Show: vscFactor! = 1: Form_Resize
End Sub

Private Sub Form_Resize() 'Print the text sent to the picture box in thisform when resize.
Dim W&, h&, W2&: Const C = 5: W& = ScaleWidth: h& = ScaleHeight: W2& = W& * 0.05
    Enabled = False: DoEvents
    Select Case W2&
        Case Is > 17: W2& = 17
        Case Is < 12: W2& = 12
    End Select
'Arrange controls
    If RTL Then
        vsc.Move ScaleLeft, ScaleTop, W2&, h&
        pic.Move W2& + C, C, W& - W2& - (C * 2), h&
    Else
        vsc.Move W& - W2&, ScaleTop, W2&, h&
        pic.Move C, C, W& - W2& - (C * 2), h&
    End If
    'Redraw the text and get the hight to W&
    W& = TextWPrint(pic, sTextW$, Margin%, RTL, True) - h&
    If W& > Max Then
        vscFactor! = W& / Max: W& = Max
    Else: vscFactor! = 1: End If
    vsc.Max = IIf(W& < 0, 0, W&)
    pic.Refresh: Enabled = True
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then PopupMenu mnuPop 'Show Copy & Paste popup menu
End Sub

Private Sub pic_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then PopupMenu mnuPop 'Show Copy & Paste popup menu
End Sub

Sub mnuP_Click(Index As Integer) 'Copy & Paste popup menu click
Dim s$
    If Index = 1 Then                               'Paste
        s$ = ClipboardGetTextW$
        If LenB(s$) = 0 Then
            MsgBoxW MwEmptyCB, 48
        Else: sTextW$ = s$: Form_Resize: End If
    Else: ClipboardLetTextW$ sTextW$: End If         'Copy
End Sub

Private Sub vsc_Change()
Dim l& 'Scroll look of picturebox
    l& = vsc.Value: l& = (l& * vscFactor!) * -1: pic.Top = l&
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frm.WindowState = 0: frm.SetFocus
End Sub
