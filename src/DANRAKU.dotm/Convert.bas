Attribute VB_Name = "Convert"
Option Explicit

Public Sub ’i—‹L†•ÏŠ·EN2JP()
    Call ConvertParagraph("\[[0-9]{4}\]", "y‚O‚O‚O‚Oz")
    Call ’i—”Ô†‚È‚ÇU‚è’¼‚µJP
End Sub

Public Sub ’i—‹L†•ÏŠ·JP2EN()
    Call ConvertParagraph("y[‚O‚P‚Q‚R‚S‚T‚U‚V‚W‚X]{4}z", "[0000]")
    Call ’i—”Ô†U‚è’¼‚µEN
End Sub

Private Sub ConvertParagraph(’uŠ·‘O, ’uŠ·Œã)
    Dim AddStr As String
    Dim ParagraphNum As Integer: ParagraphNum = 1
    Dim myRange

    Set myRange = ActiveDocument.Range()
    With myRange.Find
        .ClearFormatting
        .Text = ’uŠ·‘O
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchCase = False
        .MatchWholeWord = True
        .MatchByte = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
        .MatchFuzzy = False
        Do While .Execute = True
            With .parent
                .Delete
                AddStr = ’uŠ·Œã
                .Font.Reset
                .InsertAfter AddStr
                .Move
            End With
            ParagraphNum = ParagraphNum + 1
        Loop
    End With
End Sub
