Attribute VB_Name = "ReNumEN"
'各種番号振りなおしマクロ
'   Created By D*suke YAMAKWA
'   last modified : 14/4/7
'
'段落番号の振りなおし部分(RenumberingParagraph)に関して
'   original code : 段落番号置換マクロ 03/08/23 By 岡田
'   modified By D*suke YAMAKWA
'
'【マクロの概要】
' 1. 文書中の＠を段落番号に置き換えた後、
'    文書中の段落番号を、連番になるように書き直します。
'
'【※注意点】
'   "＠"が文書の途中にある場合でも置換されます。

Sub 段落番号振り直しEN()
    Call ReplaceAll("@", "[0000]")
    Call RenumberingParagraphEn
End Sub

Private Sub ReplaceAll(before As String, after As String)
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = before
        .Replacement.Text = after
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchKashida = False
        .MatchDiacritics = False
        .MatchAlefHamza = False
        .MatchControl = False
        .MatchByte = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = False
        .MatchFuzzy = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
End Sub

Private Sub RenumberingParagraphEn()
    Dim AddStr As String
    Dim ParagraphNum As Integer: ParagraphNum = 1

    Set myRange = ActiveDocument.Range()
    With myRange.Find
        .ClearFormatting
        .Text = "\[[0123456789]{4}\]"
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
                AddStr = "[" + Format(ParagraphNum, "0000") + "]"
                .Font.Reset
                .InsertAfter (AddStr)
                .Move
            End With
            ParagraphNum = ParagraphNum + 1
        Loop
    End With
End Sub

