VERSION 5.00
Begin VB.Form frm 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5880
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   7410
   Icon            =   "frm.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   392
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   494
   StartUpPosition =   2  'CenterScreen
   Tag             =   $"frm.frx":058A
   Begin VB.TextBox txtASCII 
      Height          =   1875
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   13
      Text            =   "frm.frx":070E
      Top             =   5940
      Width           =   7215
      Visible         =   0   'False
   End
   Begin VB.PictureBox pic 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   3615
      Index           =   0
      Left            =   120
      ScaleHeight     =   237
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   473
      TabIndex        =   6
      Top             =   2100
      Width           =   7155
      Begin VB.PictureBox pic 
         BackColor       =   &H0080FFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   1140
         Index           =   2
         Left            =   840
         ScaleHeight     =   76
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   101
         TabIndex        =   7
         Top             =   1020
         Width           =   1515
      End
      Begin VB.VScrollBar vsc 
         CausesValidation=   0   'False
         Height          =   1275
         LargeChange     =   100
         Left            =   3180
         SmallChange     =   10
         TabIndex        =   8
         Top             =   660
         Width           =   255
      End
   End
   Begin VB.Frame fra 
      Caption         =   "Please select the way to display the text and  language "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7155
      Begin VB.PictureBox pic 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   1380
         Index           =   1
         Left            =   3180
         MousePointer    =   14  'Arrow and Question
         Picture         =   "frm.frx":3820
         ScaleHeight     =   1380
         ScaleWidth      =   1425
         TabIndex        =   15
         Tag             =   "http://freetranslator.webs.com"
         ToolTipText     =   "Click here to visit my homepage"
         Top             =   360
         Width           =   1425
      End
      Begin VB.ComboBox cmb 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   150
         TabIndex        =   14
         Tag             =   $"frm.frx":6ACF
         Top             =   1530
         Width           =   2895
      End
      Begin VB.Frame fra 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Index           =   1
         Left            =   4740
         TabIndex        =   9
         Tag             =   "Standard Controls "
         Top             =   240
         Width           =   2235
         Begin VB.CommandButton cmd 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   120
            TabIndex        =   10
            Tag             =   "http://www.planet-source-code.com/vb/scripts/BrowseCategoryOrSearchResults.asp?txtCriteria=Jan+andersson&lngWId=1"
            Top             =   960
            Width           =   1995
         End
         Begin VB.OptionButton opt 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   5
            Left            =   110
            TabIndex        =   12
            Top             =   180
            Width           =   2055
         End
         Begin VB.CheckBox chk 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   120
            TabIndex        =   11
            Top             =   480
            Width           =   2055
         End
      End
      Begin VB.OptionButton opt 
         Caption         =   "Show in a message box"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   5
         Top             =   1260
         Width           =   2055
      End
      Begin VB.OptionButton opt 
         Caption         =   "Open in Notepad"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   4
         Top             =   1020
         Width           =   1755
      End
      Begin VB.OptionButton opt 
         Caption         =   "As HTML in your browser"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   3
         Top             =   780
         Width           =   2475
      End
      Begin VB.OptionButton opt 
         Caption         =   "In a new window"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   2
         Top             =   540
         Width           =   1755
      End
      Begin VB.OptionButton opt 
         Caption         =   "On the picture box"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   300
         Value           =   -1  'True
         Width           =   2055
      End
   End
   Begin VB.Menu mnu 
      Caption         =   "More..."
      Begin VB.Menu mnuUT 
         Caption         =   "UTF-8 characters table"
         Begin VB.Menu mnuU 
            Caption         =   "Create all UTF-8 characters tables in one file"
            Index           =   0
         End
         Begin VB.Menu mnuU 
            Caption         =   "Create all UTF-8 characters tables in 150 separate files"
            Index           =   1
         End
         Begin VB.Menu mnuU 
            Caption         =   "-"
            Index           =   2
         End
      End
      Begin VB.Menu mnuM 
         Caption         =   ""
         Index           =   0
      End
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
Attribute VB_Name = "frm"
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
Dim bMsg As Boolean, sWide$, vscFactor!, CW$(): Const c_GB = 10, c_Max = 32767, c_Paste = 100, c_S = ":-)"


Private Sub Form_Load()
Dim X&, Y&, Z&, l&, v$(), v2$(), s$, a$

    Call InitModuleW                        'Init the module (modW.bas)
    v$() = Split(Tag, ","): Y& = UBound(v$)
    For X& = 0 To Y&                        'Build the "More.." submenu
        If X& > 0 Then Load mnuM(X&)
        mnuM(X&).Caption = v$(X&)
    Next
    For X& = 3 To TablesU% + 3              'Build "UTF-8 characters tables" submenu
        With modW.UTF8(X& - 3)
            Load mnuU(X&): mnuU(X&).Caption = "Index = " & Format(X& - 3, "000") & _
              String(5, " ") & Format(.From, modW.Mw5x0) & " - " & _
              Format(.To, modW.Mw5x0) & String(5, " ") & .Name
        End With
    Next
    v$() = Split(cmb.Tag, ";"): Y& = UBound(v$): Show: DoEvents
    s$ = LCase(modW.GetLocalLanguageCountry) 'Get system language
    For X& = 0 To Y& 'Add language names to combo, stored in cmb.Tag
        cmb.AddItem v$(X&): If InStr(1, LCase(v$(X&)), s$) Then _
          l& = X&                            'Found default language
    Next
    ReDim CW$(Y&, 3)
'Get my stored charcodes from txtASCII and convert it to wide
    v$() = Split(txtASCII, "|")     'Split txtASCII to separate languages
    For X& = 0 To Y&                'Loop all languages
        v2$() = Split(v$(X&), ":")  'Split to get separate sentences
        For Z& = 0 To 3             'Convert the ascii codes to characters
            CW$(X&, Z&) = modW.StringA2W$(v2$(Z&))
        Next                        'Store the wide chars in CW$()
    Next
    l& = IIf(l& = 0, c_GB, l&)      'Use english as default language if not found.
    Caption = App.Title: Erase v$(): Erase v2$(): cmb.ListIndex = l&
End Sub

Private Sub cmb_Click() '//Change language in combobox
    opt_Click 0
End Sub

Private Sub opt_Click(Index As Integer) '//Click on a option button
On Error GoTo opt_ClickErr
Dim s$, T$, W$, X&, Y&, R2L As Boolean
    If Index = 5 Then Run cmd.Tag: Exit Sub 'Vote:-)
    MousePointer = 11
    Select Case cmb.Text 'Right to Left languages
        Case "Arabic", "Hebrew", "Persian": R2L = True
    End Select
    X& = cmb.ListIndex: ArrangeVsc R2L 'Arrange vscroll and pic(2) in pic(0)
    s$ = CW$(X&, 1)                               ' "Please vote" in 50+ languages
    s$ = IIf(s$ = vbNullString, CW$(c_GB, 1), s$) ' Use English if missing
    'Let unicode caption to the standard controls
    s$ = s$ & ";-)": CaptionWLet opt(5), s$:  CaptionWLet cmd, s$
    CaptionWLet chk, s$: CaptionWLet fra(1), fra(1).Tag & s$
    'You can use the CaptionWGet$() function to read the unicode caption i.e:
    'MsgBoxW CaptionWGet(fra(1)) 'Uncomment this row and try
    If Index <> c_Paste Then
        For Y& = 0 To 3 'Get all text from selected language.
            If Y& <> 1 Then
                W$ = IIf(CW$(X&, Y&) = vbNullString, CW$(c_GB, Y&), CW$(X&, Y&))
                If Y& = 3 Then W$ = W$ & c_S         'Use English if missing
            Else: W$ = s$: End If
            T$ = T$ & W$ & " "
        Next
        s$ = T$: T$ = ""
        Y& = GetRndNo&(3, 20)                'Get random number of times to display the text
        For X& = 0 To Y&
            If GetRndNo&(0, 2) < 2 Then      'Add linefeed
                T$ = T$ & s$ & vbNewLine
            Else: T$ = T$ & s$ & " ": End If 'Add space
        Next
    Else: T$ = sWide$: End If 'Use clipboard text
PasteText:
    For X& = 0 To 4 'Find the seleceted option button for display
        If opt(X&) Then Exit For
    Next
    Select Case X&
    'Display the text the way you choose
        Case 0  '//On the picture box
            X& = TextWPrint(pic(2), T$, 2, R2L, True)
            'Returns height in pixels of the printed text to X&
            X& = X& - pic(0).Height
            If X& > c_Max Then
                vscFactor! = X& / c_Max: X& = c_Max
            Else: vscFactor! = 1: End If
            vsc.Max = IIf(X& < 0, 0, X&)
        Case 1  '//On a picture box in a new window
            Dim f As New frmP: f.PrintTextW T$, 2, R2L
        Case 2  '//As HTML in your browser
            s$ = AppPath$ & "Demo of TextW2HTML.htm"
            If Dir(s$) <> "" Then Kill s$ 'Remove if exists so no error occur.
            TextW2HTML T$, s$, True, "Demo of the TextW2HTML() function", "Verdana", 4, R2L
        Case 3  '//Open in Notepad
            s$ = AppPath$ & "Demo of Save as UTF-8.txt"
            If Dir(s$) <> "" Then Kill s$ 'Remove if exists so no error occur.
            TextWOpenInNotepad T$, s$
            'Right click in Notepad to get the option Right to Left.
            'You can use the TextWSave() function to just save the text-
            'or TextWLoad() to load the text from a file, (UTF-8 or ANSI) i.e:
            'MsgBoxW TextWLoad$(s$), 'Uncomment this row to try.
        Case 4  '//In a unicode message box
            If R2L Then
                s$ = "Right to Left": X& = vbInformation + vbMsgBoxRtlReading
            Else: s$ = App.Title: X& = vbInformation: End If
            MsgBoxW T$, X&, s$
    End Select
    sWide$ = T$
    MousePointer = 0
    Exit Sub
opt_ClickErr:
    MsgBox Err.Description, vbCritical: Err.Clear: Exit Sub
    Resume
End Sub

Sub mnuM_Click(Index As Integer) '//Click on the More... menu
Dim s$, X&, Y&, v: Const WF = "Wide.txt", AF = "Charcode string.txt", MC = "You can right click in the form/picture box to show a popup menu.": MousePointer = 11
    
    Select Case Index
        Case 0 'Copy the text to clipboard
            If Not bMsg Then MsgBoxW MC: bMsg = True
            ClipboardLetTextW sWide$
        Case 1 'Paste the text from clipboard
            s$ = ClipboardGetTextW$
            If Not bMsg Then MsgBoxW MC: bMsg = True
            If LenB(s$) Then
                sWide$ = s$: opt_Click c_Paste
            Else: MsgBoxW modW.MwEmptyCB, 48: End If
        Case 3 'Convert "Wide.txt" to charcode string
            ' *** Tip! use http://translate.google.com to create ur file ***'
            If Dir(AppPath & WF) = vbNullString Then 'File is missing
                For X& = 0 To 50 'Create from my built in wide chars.
                    For Y& = 0 To 3
                        s$ = s$ & CW$(X&, Y&) & vbNewLine
                    Next
                Next
                TextWSave s$, AppPath & WF 'Save the wide to disk
            Else: s$ = TextWLoad(AppPath & WF): End If 'Load file
            'Convert the text to a string with charcodes and save as US ASCII
            TextWOpenInNotepad StringW2A$(s$), AppPath & AF, True 'Open in Notepad.
        Case 4 'Convert the clipboard text to charcode string and show in Notepad
            sWide$ = ClipboardGetTextW$
            If LenB(sWide$) Then 'Open in a temp file in Notepad.
                TextWOpenInNotepad StringW2A$(sWide$), "", True
            Else: MsgBoxW modW.MwEmptyCB, 48: End If
        Case 5 'Load the file "Charcode string.txt" with the charcode string
            'Convert the codes to wide characters and display the result in Notepad.
            s$ = TextWLoad(AppPath & AF): TextWOpenInNotepad StringA2W$(s$)
        Case 7 'Find UTF-8 characters table by name
            s$ = InputBox("Type a string to search with InStr().")
            If s$ <> vbNullString Then FindTable s$, True
        Case 8 'Show info of a character
            s$ = InputBox("Type one charcther to get info about it.")
            If s$ <> vbNullString Then s$ = Left$(s$, 1): ShowCharCodeInfo CLng(UniW(s$))
        Case 9 'Show info of a charcode
            On Error Resume Next
            s$ = InputBox("Select charcode between 0 - " & MwMaxUTF8): X& = Val(s$)
            If Err Then Err.Clear: X& = MwMaxUTF8 + 1
            If s$ <> vbNullString Then
                If Not ShowCharCodeInfo(X&) Then mnuM_Click 6 'Retry
            End If
        Case 10 'Get some random number of characters from random char table.
            MsgBoxW GetRndChrs$(GetRndNo&(0, TablesU%), GetRndNo&(10, 50), True)
        Case 12 'Go to Google Translate's homepage
            Run "http://translate.google.com"
        Case 13 'Go to my homepage
            Run pic(1).Tag
        Case 14 'Contact me by e-mail
            Contact
    End Select
    MousePointer = 0
End Sub

Private Sub mnuU_Click(Index As Integer) '//Click on UTF-8 characters table sub menu.
Dim X&, s$(), f$: MousePointer = 11
    Select Case Index
        Case 0 'Creates all chartables in a single file.
            ReDim s$(TablesU%)
            For X& = 0 To TablesU%: s$(X&) = CreateCharTable(X&): Next
            s$(0) = Join(s$, vbNewLine & vbNewLine): s$(1) = AppPath & "UTF-8 charcode 0 - " & _
              c_Max & MwDotTXT: TextWSave s$(0), s$(1)
            If MsgBoxW("Done! The file was saved as:" & vbLf & s(1) & vbLf & vbLf & _
              "Do you want to open the file now?", 36) = vbYes Then TextWOpenInNotepad "", s(1)
        Case 1 'Creates all chartables in separate files.
            ReDim s$(1)
            For X& = 0 To TablesU%
                s$(0) = CreateCharTable(X&, , , s$(1)): TextWSave s$(0), s$(1)
            Next
            MsgBoxW "Done! " & TablesU% + 1 & " files was created in " & AppPath
        Case Is > 2 'Creates and show the whole selected chartable in Notepad.
            CreateCharTable Index - 3, True
    End Select
    Erase s$(): MousePointer = 0
End Sub

Sub mnuP_Click(Index As Integer)
    bMsg = True: mnuM_Click Index  'Copy & Paste popup menu
End Sub

Private Sub vsc_Change()
Dim l& 'Scroll look of picturebox
    l& = vsc.Value: l& = (l& * vscFactor!) * -1: pic(2).Top = l&
End Sub

Private Sub ArrangeVsc(RTL) 'Arrange vscroll and pic(2) in pic(0)
    If vscFactor! <> 0 Then
        If vsc.Tag = "" And Not RTL Then Exit Sub
        If vsc.Tag <> "" And RTL Then Exit Sub
    Else: pic(2).BackColor = vbWhite: End If
Dim W&, h&, W2&: Const C = 3, WS = 17
    W& = pic(0).ScaleWidth: h& = pic(0).ScaleHeight
    If RTL Then
        vsc.Move ScaleLeft, ScaleTop, WS, h&: vsc.Tag = vbVerticalTab
        pic(2).Move WS + C, C, W& - WS - (C * 2), h&
    Else
        vsc.Move W& - WS, ScaleTop, WS, h&: vsc.Tag = vbNullString
        pic(2).Move C, C, W& - WS - (C * 2), h&
    End If
    DoEvents
End Sub

Private Sub pic_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Index = 1 Then Run pic(1).Tag                              'Go to my homepage
    If Index = 2 And Button = vbRightButton Then PopupMenu mnuPop 'Show Copy & Paste popup menu
End Sub

Private Sub cmd_Click()
    Run cmd.Tag 'Vote:-)
End Sub

Private Sub chk_Click()
    Run cmd.Tag 'Vote:-)
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim X&, s$, W$: X& = cmb.ListIndex: s$ = AppPath & CW$(c_GB, 3) & modW.MwDotTXT
    If Dir(s$) = vbNullString Then
        W$ = IIf(CW$(X&, 2) = vbNullString, CW$(c_GB, 2), CW$(X&, 2))
        If MsgBoxW(W$, 292) = vbYes Then 'Ask if you have vote:P
            W$ = IIf(CW$(X&, 3) = vbNullString, CW$(c_GB, 3), CW$(X&, 3)) & c_S
            modW.TextWSave W$, s$: MsgBoxW W$
        Else: Run cmd.Tag: End If 'Vote:-)
    End If
    Erase CW$(): Erase UTF8(): End
End Sub

