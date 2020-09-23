Attribute VB_Name = "modW"
'*************************************************************************************************
'**************************   Display UTF-8 & modW.bas November 17 2010   ************************
'*************************************************************************************************
'*  Copyright© Pappsegull Sweden, http://freetranslator.webs.com <pappsegull@yahoo.se>
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

Private Declare Function DefWindowProcW Lib "user32" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Sub PutMem4 Lib "msvbvm60" (Destination As Any, Value As Any)
Private Declare Function SysAllocStringLen Lib "oleaut32" (ByVal OleStr As Long, ByVal bLen As Long) As Long
Private Declare Function MessageBoxW Lib "user32.dll" (ByVal hWnd As Long, ByVal lpText As Long, ByVal lpCaption As Long, ByVal uType As Long) As Long
Private Declare Function TextOut Lib "gdi32" Alias "TextOutW" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As Long, ByVal nCount As Long) As Long
Private Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32W" (ByVal hDC As Long, ByVal lpsz As Long, ByVal cbString As Long, lpSize As POINTAPI) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function OpenClipboard Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SetClipboardData Lib "user32" (ByVal Format As Long, ByVal hMem As Long) As Long
Private Declare Function CloseClipboard Lib "user32" () As Long
Private Declare Function EmptyClipboard Lib "user32" () As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal Flags As Long, ByVal l As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (ByVal pDest As Long, ByVal pSource As Long, ByVal l As Long)
Private Declare Function GetClipboardData Lib "user32" (ByVal wFormat As Long) As Long
Private Declare Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
Private Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long

Private Const c_UTF8 = "0;Basic Latin|128;Latin-1 Supplement|256;Latin Extended-A|384;Latin Extended-B|592;IPA Extensions|688;Spacing Modifier Letters|768;Combining Diacritical Marks|880;Greek and Coptic|1024;Cyrillic|1280;Cyrillic Supplement|1328;Armenian|1424;Hebrew|1536;Arabic|1792;Syriac|1872;Arabic Supplement|1920;Thaana|1984;NKo|2048;Samaritan|2304;Devanagari|2432;Bengali|2560;Gurmukhi|2688;Gujarati|2816;Oriya|2944;Tamil|3072;Telugu|3200;Kannada|3328;Malayalam|3456;Sinhala|3584;Thai|3712;Lao|3840;Tibetan|4096;Myanmar|4256;Georgian|4352;Hangul Jamo|4608;Ethiopic|4992;Ethiopic Supplement|5024;Cherokee|5120;Unified Canadian Aboriginal Syllabics|5760;Ogham|5792;Runic|5888;Tagalog|5920;Hanunoo|5952;Buhid|5984;Tagbanwa|6016;Khmer|6144;Mongolian|6320;Unified Canadian Aboriginal Syllabics Extended|6400;Limbu|6480;Tai Le|6528;New Tai Lue|6624;Khmer Symbols|6656;Buginese|6688;Tai Tham|6912;Balinese|7040;Sundanese|7168;Lepcha|7248;Ol Chiki|7376;Vedic Extensions|7424;Phonetic Extensions|7552;Phonetic " & _
  "Extensions Supplement|7616;Combining Diacritical Marks Supplement|7680;Latin Extended Additional|7936;Greek Extended|8192;General Punctuation|8304;Superscripts and Subscripts|8352;Currency Symbols|8400;Combining Diacritical Marks for Symbols|8448;Letterlike Symbols|8528;Number Forms|8592;Arrows|8704;Mathematical Operators|8960;Miscellaneous Technical|9216;Control Pictures|9280;Optical Character Recognition|9312;Enclosed Alphanumerics|9472;Box Drawing|9600;Block Elements|9632;Geometric Shapes|9728;Miscellaneous Symbols|9984;Dingbats|10176;Miscellaneous Mathematical Symbols-A|10224;Supplemental Arrows-A|10240;Braille Patterns|10496;Supplemental Arrows-B|10624;Miscellaneous Mathematical Symbols-B|10752;Supplemental Mathematical Operators|11008;Miscellaneous Symbols and Arrows|11264;Glagolitic|11360;Latin Extended-C|11392;Coptic|11520;Georgian Supplement|11568;Tifinagh|11648;Ethiopic Extended|11744;Cyrillic Extended-A|11776;Supplemental Punctuation|11904;CJK Radicals Supplement|12032;Kangxi Radicals|" & _
  "12272;Ideographic Description Characters|12288;CJK Symbols and Punctuation|12352;Hiragana|12448;Katakana|12544;Bopomofo|12592;Hangul Compatibility Jamo|12688;Kanbun|12704;Bopomofo Extended|12736;CJK Strokes|12784;Katakana Phonetic Extensions|12800;Enclosed CJK Letters and Months|13056;CJK Compatibility|13312;CJK Unified Ideographs Extension A|19904;Yijing Hexagram Symbols|19968;CJK Unified Ideographs|40960;Yi Syllables|42128;Yi Radicals|42192;Lisu|42240;Vai|42560;Cyrillic Extended-B|42656;Bamum|42752;Modifier Tone Letters|42784;Latin Extended-D|43008;Syloti Nagri|43056;Common Indic Number Forms|43072;Phags-pa|43136;Saurashtra|43232;Devanagari Extended|43264;Kayah Li|43312;Rejang|43360;Hangul Jamo Extended-A|43392;Javanese|43520;Cham|43616;Myanmar Extended-A|43648;Tai Viet|43968;Meetei Mayek|44032;Hangul Syllables|55216;Hangul Jamo Extended-B|55296;High Surrogates|56192;High Private Use Surrogates|56320;Low Surrogates|57344;Private Use Area|63744;CJK Compatibility Ideographs|" & _
  "64256;Alphabetic Presentation Forms|64336;Arabic Presentation Forms-A|65024;Variation Selectors|65040;Vertical Forms|65056;Combining Half Marks|65072;CJK Compatibility Forms|65104;Small Form Variants|65136;Arabic Presentation Forms-B|65280;Halfwidth and Fullwidth Forms|65520;Specials"

Private Type POINTAPI
   X As Long
   Y As Long
End Type

Public Enum TablesUTF8
    [Basic Latin]
    [Latin-1 Supplement]
    [Latin Extended-A]
    [Latin Extended-B]
    [IPA Extensions]
    [Spacing Modifier Letters]
    [Combining Diacritical Marks]
    [Greek and Coptic]
    [Cyrillic]
    [Cyrillic Supplement]
    [Armenian]
    [Hebrew]
    [Arabic]
    [Syriac]
    [Arabic Supplement]
    [Thaana]
    [NKo]
    [Samaritan]
    [Devanagari]
    [Bengali]
    [Gurmukhi]
    [Gujarati]
    [Oriya]
    [Tamil]
    [Telugu]
    [Kannada]
    [Malayalam]
    [Sinhala]
    [Thai]
    [Lao]
    [Tibetan]
    [Myanmar]
    [Georgian]
    [Hangul Jamo]
    [Ethiopic]
    [Ethiopic Supplement]
    [Cherokee]
    [Unified Canadian Aboriginal Syllabics]
    [Ogham]
    [Runic]
    [Tagalog]
    [Hanunoo]
    [Buhid]
    [Tagbanwa]
    [Khmer]
    [Mongolian]
    [Unified Canadian Aboriginal Syllabics Extended]
    [Limbu]
    [Tai Le]
    [New Tai Lue]
    [Khmer Symbols]
    [Buginese]
    [Tai Tham]
    [Balinese]
    [Sundanese]
    [Lepcha]
    [Ol Chiki]
    [Vedic Extensions]
    [Phonetic Extensions]
    [Phonetic Extensions Supplement]
    [Combining Diacritical Marks Supplement]
    [Latin Extended Additional]
    [Greek Extended]
    [General Punctuation]
    [Superscripts and Subscripts]
    [Currency Symbols]
    [Combining Diacritical Marks for Symbols]
    [Letterlike Symbols]
    [Number Forms]
    [Arrows]
    [Mathematical Operators]
    [Miscellaneous Technical]
    [Control Pictures]
    [Optical Character Recognition]
    [Enclosed Alphanumerics]
    [Box Drawing]
    [Block Elements]
    [Geometric Shapes]
    [Miscellaneous Symbols]
    [Dingbats]
    [Miscellaneous Mathematical Symbols-A]
    [Supplemental Arrows-A]
    [Braille Patterns]
    [Supplemental Arrows-B]
    [Miscellaneous Mathematical Symbols-B]
    [Supplemental Mathematical Operators]
    [Miscellaneous Symbols and Arrows]
    [Glagolitic]
    [Latin Extended-C]
    [Coptic]
    [Georgian Supplement]
    [Tifinagh]
    [Ethiopic Extended]
    [Cyrillic Extended-A]
    [Supplemental Punctuation]
    [CJK Radicals Supplement]
    [Kangxi Radicals]
    [Ideographic Description Characters]
    [CJK Symbols and Punctuation]
    [Hiragana]
    [Katakana]
    [Bopomofo]
    [Hangul Compatibility Jamo]
    [Kanbun]
    [Bopomofo Extended]
    [CJK Strokes]
    [Katakana Phonetic Extensions]
    [Enclosed CJK Letters and Months]
    [CJK Compatibility]
    [CJK Unified Ideographs Extension A]
    [Yijing Hexagram Symbols]
    [CJK Unified Ideographs]
    [Yi Syllables]
    [Yi Radicals]
    [Lisu]
    [Vai]
    [Cyrillic Extended-B]
    [Bamum]
    [Modifier Tone Letters]
    [Latin Extended-D]
    [Syloti Nagri]
    [Common Indic Number Forms]
    [Phags-pa]
    [Saurashtra]
    [Devanagari Extended]
    [Kayah Li]
    [Rejang]
    [Hangul Jamo Extended-A]
    [Javanese]
    [Cham]
    [Myanmar Extended-A]
    [Tai Viet]
    [Meetei Mayek]
    [Hangul Syllables]
    [Hangul Jamo Extended-B]
    [High Surrogates]
    [High Private Use Surrogates]
    [Low Surrogates]
    [Private Use Area]
    [CJK Compatibility Ideographs]
    [Alphabetic Presentation Forms]
    [Arabic Presentation Forms-A]
    [Variation Selectors]
    [Vertical Forms]
    [Combining Half Marks]
    [CJK Compatibility Forms]
    [Small Form Variants]
    [Arabic Presentation Forms-B]
    [Halfwidth and Fullwidth Forms]
    [Specials]
    [Linear B Syllabary]
End Enum

Public Type UTF8Table
    Name As String
    From As Long
    To As Long
End Type
Global Const Mw5x0 = "00000", MwMaxUTF8 = 65535, MwTablesU = 149, MwDotTXT = ".txt", MwErrTable = "Invalid charcode table, a valid value is 0 - " & MwTablesU & ".", MwEmptyCB = "There are no unicode text on the clipboard."
Public UTF8(MwTablesU) As UTF8Table, TablesU%, AppPath$

' !!!! ALLWAYS RUN THIS SUB FIRST IF YOU LIKE TO USE ALL FUNCTIONS !!!!

Public Sub InitModuleW()
Dim X&, s$(), v$() 'Init the UTF-8 characters table
    AppPath$ = App.Path: If Right(AppPath$, 1) <> "\" Then AppPath$ = AppPath$ & "\"
    s$ = Split(c_UTF8, "|")
    For X& = 0 To MwTablesU
        v$() = Split(s$(X&), ";")
        With UTF8(X&)
            .Name = v$(1): .From = v$(0)
        End With
        If X& > 0 Then UTF8(X& - 1).To = Val(v$(0)) - 1
    Next
    UTF8(MwTablesU).To = MwMaxUTF8: Erase s$(): Erase v$(): TablesU% = MwTablesU
End Sub

Public Function TextWPrint&(ToObject As Object, TextW$, Optional PercentMargin% = 2, Optional RightToLeft, Optional AutoHeight As Boolean = True)
Dim s$, T$(), l&, n&, p&, X&, XP&, Y&, h&, W&, RH&, PT As POINTAPI, b As Boolean
On Error GoTo ErrTextWPrint 'Print text to a picturebox or form
    If Not TypeOf ToObject Is PictureBox And Not TypeOf ToObject Is Form Then Exit Function
ResizeFont:
    With ToObject
        W& = .ScaleX(.ScaleWidth, .ScaleMode, vbPixels): .Cls: .AutoRedraw = True
        h& = .ScaleY(.ScaleHeight, .ScaleMode, vbPixels)
        GetTextExtentPoint32 .hDC, StrPtr(TextW$), Len(TextW$), PT: h& = PT.Y
        XP& = W& * (PercentMargin% / 100): T$ = Split(TextW$, vbNewLine)
        s$ = Join(T$, "|"): T$ = Split(s$, " "): s$ = "": n& = UBound(T): Y& = XP&
        For l& = 0 To n&
            GetTextExtentPoint32 .hDC, StrPtr(s$ & T(l&)), Len(s$ & T(l&)), PT
            If PT.X + (XP& * 2) > W& Then 'Text to long to fit the width.
                If RightToLeft Then
                    X& = W& - XP& - PT.X
                Else
                    X& = XP&: If InStr(1, T(l&), "|") = 0 Then s$ = s$ & "-"
                End If
                If AutoHeight Then .Height = Y& + (XP& * 2)
                TextOut .hDC, X&, Y, (StrPtr(s$)), _
                  (Len(s$)): s$ = "": Y& = Y& + h&
            End If
LineFeed:
            p& = InStr(1, T(l&), "|")
            If p& Then 'Is a linefeed
                s$ = s$ & Mid$(T(l&), 1, p& - 1)
                GetTextExtentPoint32 .hDC, StrPtr(s$), Len(s$), PT
                If RightToLeft Then
                    X& = W& - XP& - PT.X
                Else: X& = XP&: End If
                If AutoHeight Then .Height = Y& + (XP& * 2)
                TextOut .hDC, X&, Y, (StrPtr(s$)), (Len(s$)) 'Print the text
                s$ = Mid$(T(l&), p& + 1, Len(T(l&)) - p&): Y& = Y& + h&
                If InStr(1, s$, "|") Then T(l&) = s$: s$ = "": GoTo LineFeed
            Else: s$ = s$ & T(l&): End If
            s$ = s$ & " "
        Next
        If LenB(s$) Then
            GetTextExtentPoint32 .hDC, StrPtr(s$), Len(s$), PT
            If RightToLeft Then
                X& = W& - XP& - PT.X
            Else: X& = XP&: End If
            If AutoHeight Then .Height = Y& + (XP& * 2)
            TextOut .hDC, X&, Y&, (StrPtr(s$)), (Len(s$)) 'Print the text
        End If
        Erase T(): TextWPrint& = Y& + (XP& * 2) + h&      'Return the height
        If AutoHeight Then .Height = TextWPrint&
    End With
    Exit Function
ErrTextWPrint:
    MsgBox "Error number " & Err & " in function TextWPrint():" & _
      vbLf & vbLf & Err.Description, vbCritical: Err.Clear
End Function

Public Function TextW2HTML(TextW$, Optional ToFile$, Optional OpenFile As Boolean, Optional ByVal Title$, Optional ByVal FontName$ = "Tahoma", Optional ByVal FontSize% = 3, Optional Right2Left As Boolean) As String
On Local Error GoTo ErrTextW2HTML 'Create a HTML page of unicode or ansi text

    If Right2Left Then TextW2HTML = " dir=""rtl""" 'Right to left
    TextW2HTML = "<html" & TextW2HTML & "><head><meta http-equiv=""Content-Type"" " & _
      "content=""text/html; charset=UTF-8""><title>" & Title$ & _
      "</title></head><body>" & vbNewLine & _
      "<p><font face=""" & FontName$ & """ size=""" & FontSize% & """>" & vbNewLine & _
      Join(Split(TextW$, vbNewLine), "<br>" & vbNewLine) & vbNewLine & _
      "</font></p></body></html>"
    If LenB(ToFile$) = 0 Then Exit Function
    If Not TextWSave(TextW2HTML, ToFile$) Then Exit Function
    If OpenFile Then Run ToFile$
    Exit Function
ErrTextW2HTML:
    MsgBox "Error number " & Err & " in function TextW2HTML():" & _
      vbLf & vbLf & Err.Description, vbCritical: Err.Clear
End Function

Public Function MsgBoxW(ByVal Prompt$, Optional ByVal Buttons As VbMsgBoxStyle = vbInformation, Optional ByVal Title$, Optional ByVal hWndOwner&) As VbMsgBoxResult
On Local Error Resume Next 'Show a unicode message box...
    If hWndOwner = 0 Then _
        If Not Screen.ActiveForm Is Nothing Then _
          hWndOwner& = Screen.ActiveForm.hWnd
    If LenB(Title$) = 0 Then Title$ = App.Title
    MsgBoxW = MessageBoxW(hWndOwner&, StrPtr(Prompt$), StrPtr(Title$), Buttons)
End Function

Public Function CaptionWGet$(ByRef Control As Object)
Dim l&, lPtr& 'Get the wide caption from a control
    If Not Control Is Nothing Then 'Validate supported control
        If (TypeOf Control Is CheckBox) Or (TypeOf Control Is CommandButton) Or (TypeOf Control Is Form) Or (TypeOf Control Is Frame) Or (TypeOf Control Is MDIForm) Or (TypeOf Control Is OptionButton) Then
            l& = DefWindowProcW(Control.hWnd, &HE, 0, ByVal 0) 'Get length of text
            If l& Then
                lPtr& = SysAllocStringLen(0, l&): PutMem4 ByVal VarPtr(CaptionWGet), ByVal lPtr&
                DefWindowProcW Control.hWnd, &HD, l& + 1, ByVal lPtr&
            End If
        Else 'Try the default property
            On Local Error Resume Next
            CaptionWGet = Control.Caption
            If LenB(CaptionWGet) = 0 Then CaptionWGet = Control.Text
            If Err Then Err.Clear
        End If
    End If
End Function

Public Sub CaptionWLet(ByRef Control As Object, ByRef NewValue$)
'Let the wide caption from a control
    If Not Control Is Nothing Then 'Validate supported control
        If (TypeOf Control Is Menu) Or (TypeOf Control Is CheckBox) Or (TypeOf Control Is CommandButton) Or (TypeOf Control Is Form) Or (TypeOf Control Is Frame) Or (TypeOf Control Is MDIForm) Or (TypeOf Control Is OptionButton) Then
            DefWindowProcW Control.hWnd, &HC, 0, ByVal StrPtr(NewValue)
        Else 'Try the default property
            On Local Error Resume Next
            Control.Text = NewValue: Control.Caption = NewValue
            If Err Then Err.Clear
        End If
        Control.Refresh
    End If
End Sub

Public Sub ClipboardLetTextW(ByVal sText$, Optional hWnd&)
Dim h&, p&: On Local Error Resume Next 'Let Unicode text to the clipboard
    OpenClipboard hWnd&: EmptyClipboard
    h& = GlobalAlloc(&O2& Or &O40&, LenB(sText$))
    p& = GlobalLock(h&): RtlMoveMemory p&, StrPtr(sText$), LenB(sText$)
    GlobalUnlock h&: SetClipboardData &HD&, h&: CloseClipboard
    If Err <> 0 Then Err.Clear
End Sub

Public Function ClipboardGetTextW$(Optional hWnd&)
Dim hM&, SzM&, lPtr&, b() As Byte 'Get Unicode text from the clipboard
On Error GoTo ErrClipboardGetUniText
    'Get Unicode text in a global, movable memory block
    OpenClipboard hWnd&: hM& = GetClipboardData(&HD&): SzM& = GlobalSize(hM&)
    If SzM& <= 0 Then GoTo ErrClipboardGetUniText
    'Get a non-movable global pointer to the Unicode text
    lPtr& = GlobalLock(hM&): If (lPtr& = 0) Then GoTo ErrClipboardGetUniText
    ReDim b(0 To SzM& - 1): CopyMemory b(0), ByVal lPtr&, SzM&
    GlobalUnlock hM&: ClipboardGetTextW$ = StripNullChars$(b)
ErrClipboardGetUniText:
    CloseClipboard
End Function

Public Function UniW(ByVal sChr$, Optional ToHTML As Boolean)
'This function returns what AscW() should return, input one character only
Dim l&: l& = Len(sChr$) 'I have notice that AscW() fail sometimes...
    If l& = 0 Then Exit Function
    If l& > 1 Then sChr$ = Left$(sChr$, 1)
    l& = AscB(MidB(sChr$, 2, 1)): UniW = AscB(MidB(sChr$, 1, 1)) + (l& * 256)
    If ToHTML Then UniW = "&#" & UniW & ";" 'Can also return as HTML code for the char.
End Function

Public Function StringW2A$(TextW$, Optional JoinString = ";")
'Creates a string with ascii codes that you can store like my sample
'of the vote text I have store in the opt().tag, just open a file and try;)
If LenB(TextW$) = 0 Then Exit Function
Dim l&, n&, s$(): n& = Len(TextW$): ReDim s$(n& - 1)
    For l& = 1 To n&: s$(l& - 1) = UniW(Mid$(TextW$, l&, 1)): Next
    StringW2A$ = Join(s$, JoinString): Erase s$
End Function

Public Function StringA2W$(StringUniW$, Optional SplitChar = ";")
'Convert the asciicode string from the function above to wide characters.
If LenB(StringUniW$) = 0 Then Exit Function
Dim X&, n&, s$(), v$()
    v$ = Split(StringUniW$, SplitChar): n& = UBound(v$): ReDim s$(n&)
    For X& = 0 To n&: s$(X&) = ChrW(CLng(Val(v(X&)))): Next
    StringA2W$ = StripNullChars$(Join(s$, "")): Erase v$: Erase s$
End Function

Public Function GetRndNo&(ByVal Min&, ByVal Max&)
    Randomize: GetRndNo& = CLng((Max& - Min& + 1) * Rnd) + Min& 'Get random numbers
End Function

Public Function GetRndChrs$(ByVal Table As TablesUTF8, Optional NoOffChars% = 1, Optional IncludeInfo As Boolean)
If NoOffChars% < 1 Then Exit Function
    If Not TablesU% Then InitModuleW
    If Table > TablesU% Or Table < 0 Then _
      MsgBoxW MwErrTable, 48: Exit Function
'Get some random characters from selected UTF-8 char table
Dim X&, f&, T&, s$(): ReDim s$(NoOffChars% - 1)
    f& = UTF8(Table).From: T& = UTF8(Table).To
    For X& = 0 To NoOffChars% - 1: s$(X&) = ChrW(GetRndNo&(f&, T&)): Next
    GetRndChrs = Join(s$, ""): Erase s$()
    If IncludeInfo Then
        GetRndChrs = f& & " - " & T& & ", " & UTF8(Table).Name & vbLf & vbLf & GetRndChrs
    End If
End Function

Function ShowCharCodeInfo(CharCode&) As Boolean 'Show info of the charcode in a messagebox.
Dim X&, s$: Const C = "The intervall for UTF-8 is "
    If CharCode& > MwMaxUTF8 Or CharCode& < 0 Then _
      MsgBoxW "Invalid charcode, " & LCase(C) & "0 - " & _
      MwMaxUTF8 & ".", 48: Exit Function
    If Not TablesU% Then InitModuleW
    For X& = 0 To TablesU%
        With UTF8(X&)
            If CharCode& >= .From And CharCode& <= .To Then
                ShowCharCodeInfo = True: If CharCode& = 0 Then _
                  s$ = "Null" Else s$ = ChrW(CharCode&)
                MsgBoxW "The charcode " & CharCode& & " '" & _
                  s$ & "' belongs to " & .Name & "." & vbLf & _
                  vbLf & C & .From & " - " & .To & ".": Exit Function
            End If
        End With
    Next
End Function

Public Function CreateCharTable$(ByVal Table As TablesUTF8, Optional SaveAndShowInNotepad As Boolean, Optional ReplaceFileIfExist As Boolean, Optional RetFileName$)
    If Not TablesU% Then InitModuleW
    If Table > TablesU% Or Table < 0 Then _
      MsgBoxW MwErrTable, 48: Exit Function
'Show info of the selected UTF-8 char table, creats a UTF-8 file and open it in Notepad.
Dim X&, f&, T&, n&, s$()
    f& = UTF8(Table).From: T& = UTF8(Table).To: ReDim s$(T& - f&)
    For X& = f& To T&
        s$(n&) = Format(X&, Mw5x0) & " = " & ChrW(X&): n& = n& + 1
    Next
    RetFileName$ = Format(f&, Mw5x0) & " - " & Format(T&, Mw5x0) & " - " & UTF8(Table).Name
    CreateCharTable$ = vbNewLine & String(50, "*") & vbNewLine & RetFileName$ & _
      " (" & T& - f& + 1 & " characters)" & vbNewLine & String(50, "*") & _
      vbNewLine & vbNewLine & Join(s$, vbNewLine)
    RetFileName$ = AppPath & RetFileName$ & MwDotTXT: Erase s$()
    If SaveAndShowInNotepad Then
        If Dir(RetFileName$) <> vbNullString And ReplaceFileIfExist Then Kill RetFileName$ 'Remove if file exist.
        TextWOpenInNotepad CreateCharTable$, RetFileName$
    End If
End Function

Function FindTable(ByVal SerchString$, Optional ShowInfo As Boolean) As UTF8Table()
Dim X&, n&, s$, m%(), T() As UTF8Table 'Find table by name
    SerchString$ = LCase(SerchString$): ReDim T(0)
    For X& = 0 To TablesU%
        With UTF8(X&)
            If InStr(1, LCase(.Name), SerchString$) Then
                ReDim Preserve T(n&): ReDim Preserve m%(n&): T(n&) = UTF8(n&): m%(n&) = X&
                s$ = s$ & "Index = " & X& & vbTab & Format(.From, Mw5x0) & " - " & _
                  Format(.To, Mw5x0) & ", " & .Name & vbNewLine: n& = n& + 1
            End If
        End With
    Next
    If ShowInfo Then 'Show info if selected
        If n& Then
            s$ = "Did find " & n& & " table" & IIf(n& > 1, "s", "") & _
              ":" & vbLf & vbLf & s$ & vbLf & vbLf & _
              "Do you want to create " & IIf(n& > 1, "them?", "it?"): X& = 292
        Else: s$ = "Did not find any matching table.": X& = 48: End If
        If MsgBoxW(s$, X&) = vbYes Then  'Create or open the selected char table files.
            For X& = 0 To n& - 1
                CreateCharTable m%(X&), True
            Next
        End If
    End If
    FindTable = T: Erase T: Erase m%()
End Function

'//Mixed

Public Function GetLocalLanguageCountry$(Optional bGetCountry As Boolean, Optional bNativeName As Boolean)
Dim l& 'Return a string like Swedish or Svenska or Sweden or Sverige
    If bGetCountry Then
        l& = IIf(bNativeName, &H8, &H1002)
    Else
        l& = IIf(bNativeName, &H4, &H1001)
    End If
    GetLocalLanguageCountry$ = GetLocalInfo$(l&)
End Function

Public Function StripNullChars$(ByVal sInput$)
Dim l&: l& = InStr(1, sInput$, vbNullChar)
    If l& Then 'Remove ending nullchars from a string
        StripNullChars = Left$(sInput$, l& - 1)
    Else: StripNullChars = sInput$: End If
End Function

Public Sub Run(CommandString, Optional hWnd&) 'Execute a command string
    ShellExecute hWnd&, vbNullString, CommandString, vbNullString, vbNullString, 0&
End Sub

Public Sub Contact(Optional eMail$, Optional Name$, Optional Subject$, Optional BodyText$, Optional AttachedFile$, Optional CC$, Optional BCC$)
Dim s$ 'Send e-mail to me or someone else, AttachedFile$ do not work with i.e Outlook Express!!
    If eMail$ = vbNullString Then eMail$ = "pappsegull@yahoo.se"
    If Name$ = vbNullString Then Name$ = "Pappsegull Sweden"
    If Subject$ = vbNullString Then _
      Subject$ = App.Title & " " & App.Major & "." & _
      App.Minor & "." & App.Revision
    s$ = "mailto:" & Name$ & " <" & eMail$ & ">"
    s$ = s$ & "?subject=" & Subject$
    If LenB(BodyText$) Then s$ = s$ & "&Body=" & _
      Replace(BodyText$, vbNewLine, "%0D%0A")
    If LenB(CC$) Then s$ = s$ & "&cc=" & CC$
    If LenB(BCC$) Then s$ = s$ & "&bcc=" & BCC$
    If LenB(AttachedFile$) Then
        If Dir(AttachedFile$) = vbNullString Then 'Can not find attached file...
            If MsgBoxW("Cant find the attached file '" & AttachedFile$ & "'." & _
              vbLf & vbLf & "Do you want to continue?", 308) = vbNo Then Exit Sub
        Else: s$ = s$ & "&attach=" & ChrW$(34) & AttachedFile$ & ChrW$(34): End If
    End If
    Run s$
End Sub

'//File stuff

Public Function TextWOpenInNotepad(TextW$, Optional ByVal SaveToFile$, Optional UsASCII As Boolean) As Boolean
'Open text in notepad as UTF8 if not UsASCII = True then it is as US ASCII
    TextWOpenInNotepad = SaveOpenFile(TextW$, SaveToFile$, True, , , UsASCII)
End Function

Public Function TextWSave(TextW$, SaveToFile$, Optional UsASCII As Boolean) As Boolean
'Save text as UTF8 if not UsASCII = True then it is as US ASCII
    TextWSave = SaveOpenFile(TextW$, SaveToFile$, , , , UsASCII)
End Function

Public Function TextWLoad$(ReadFromFile$, Optional UsASCII As Boolean, Optional OpenInNotepad As Boolean)
'Load text as UTF8 if not UsASCII = True then it is as US ASCII
Dim s$: SaveOpenFile "", ReadFromFile$, OpenInNotepad, True, s$, UsASCII: TextWLoad$ = s$
End Function

'//Private

Private Function SaveOpenFile(TextW$, Optional ByVal sFile$, Optional OpenInNotepad As Boolean, Optional ReadFromFile As Boolean, Optional RetText$, Optional AsASCII As Boolean) As Boolean
On Local Error GoTo ErrSaveOpenFile 'Open/Save as UTF-8 or ANSI and open file in Notepad
Dim OBJ As Object, s$, b As Boolean
    If Not ReadFromFile Then
        If LenB(sFile$) Then
            If Dir(sFile$) <> "" Then GoTo ShowInNotepad
        Else
            sFile$ = AppPath$ & IIf(AsASCII, "~ansi", "~utf8") & ".tmp"
            If Dir(sFile$) <> "" Then Kill sFile$
        End If
    End If
    Set OBJ = CreateObject("ADODB.Stream")
    With OBJ
        .Open: .Type = 2: .Charset = IIf(AsASCII, "us-ascii", "utf-8")
        If ReadFromFile Then                                          'Open file
            .LoadFromFile sFile$: RetText$ = .ReadText: b = True
        Else: .WriteText TextW$: .SaveToFile sFile$: b = True: End If 'Save file
ExitSaveOpenFile:
        If b Then 'Close ADODB object if no error and open the file if selected.
            .Close: Set OBJ = Nothing
            If OpenInNotepad Then
ShowInNotepad:
                Set OBJ = CreateObject("Wscript.Shell")
                OBJ.Run "%windir%\notepad " & sFile$: b = True
            End If
        End If
        Set OBJ = Nothing
    End With
    SaveOpenFile = b
    Exit Function
ErrSaveOpenFile:
    MsgBox "Error number " & Err & " in function SaveOpenFile():" & _
      vbLf & sFile$ & vbLf & vbLf & Err.Description, vbCritical: Resume ExitSaveOpenFile
End Function

Private Function GetLocalInfo$(ByVal lInfo&)
Dim Buffer$, Ret$ 'Get local information
    Buffer = String$(256, 0)
    Ret$ = GetLocaleInfo(&H400, lInfo&, Buffer$, Len(Buffer$))
    If Ret > 0 Then
        GetLocalInfo$ = Left$(Buffer$, Ret$ - 1)
    End If
End Function

