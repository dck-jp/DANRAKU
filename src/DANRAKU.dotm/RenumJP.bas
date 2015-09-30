Attribute VB_Name = "RenumJP"
'Šeí”Ô†U‚è‚È‚¨‚µƒ}ƒNƒ
'   Created By D*suke YAMAKWA
'   last modified : 12/10/16
'
'’i—”Ô†‚ÌU‚è‚È‚¨‚µ•”•ª(RenumberingParagraph)‚ÉŠÖ‚µ‚Ä
'   original code : ’i—”Ô†’uŠ·ƒ}ƒNƒ 03/08/23 By ‰ª“c
'   modified By D*suke YAMAKWA
'
'yƒ}ƒNƒ‚ÌŠT—vz
' 1. •¶‘’†‚Ì—‚ğ’i—”Ô†‚É’u‚«Š·‚¦‚½ŒãA
'    •¶‘’†‚Ì’i—”Ô†‚ğA˜A”Ô‚É‚È‚é‚æ‚¤‚É‘‚«’¼‚µ‚Ü‚·B
'
' 2. •¶‘’†‚Ì–‚ğ¿‹€”Ô†‚É’u‚«Š·‚¦‚½ŒãA
'    •¶‘’†‚Ì¿‹€”Ô†‚ğA˜A”Ô‚É‚È‚é‚æ‚¤‚É‘‚«’¼‚µ‚Ü‚·B
'
' 3. }A‰»Šw®A”®A•\‚Ì”Ô†‚ğ˜A”Ô‚É‚È‚é‚æ‚¤‚É‘‚«’¼‚µ‚Ü‚·B
'
'y¦’ˆÓ“_z
'   "—"A"–"‚ª•¶‘‚Ì“r’†‚É‚ ‚éê‡‚Å‚à’uŠ·‚³‚ê‚Ü‚·B
'   •¶‚Ì“r’†‚Å‚Í"—"A"–"‚ğg‚í‚È‚¢A‚à‚µ‚­‚ÍA‰º‹Ls‚ğíœ‚µ‚Ä‰º‚³‚¢B

Public Sub ’i—”Ô†‚È‚ÇU‚è’¼‚µJP()
    Call ReplaceAll("—", "y‚O‚O‚O‚Oz") '—‚Ì’uŠ·‹@”\‚ª•s—v‚Èê‡‚Í‚±‚Ìs‚ğíœ
    Call RenumberingParagraph

    'Call ReplaceAll("–", "y¿‹€‚Oz") '–‚Ì’uŠ·‹@”\‚ª•s—v‚Èê‡‚Í‚±‚Ìs‚ğíœ
    'Call Renumbering("¿‹€")
    
    'Call Renumbering("}")
    'Call Renumbering("‰»")
    'Call Renumbering("”")
    'Call Renumbering("•\")
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

Private Sub RenumberingParagraph()
    Dim AddStr As String
    Dim ParagraphNum As Integer: ParagraphNum = 1

    Set myRange = ActiveDocument.Range()
    With myRange.Find
        .ClearFormatting
        .Text = "y[‚O‚P‚Q‚R‚S‚T‚U‚V‚W‚X]*z"
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
                AddStr = "y" + StrConv(Format(ParagraphNum, "0000"), vbWide) + "z"
                .Font.Reset
                .InsertAfter (AddStr)
                .Move
            End With
            ParagraphNum = ParagraphNum + 1
        Loop
    End With
End Sub


Private Sub Renumbering(moji As String)
    Dim num As Integer: num = 1

    Set myRange = ActiveDocument.Range()
    With myRange.Find
        .ClearFormatting
        .Text = "y" & moji & "[‚O‚P‚Q‚R‚S‚T‚U‚V‚W‚X]*z"
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
                AddStr = "y" & moji & StrConv(num, vbWide) & "z"
                .Font.Reset
                .InsertAfter (AddStr)
                .Move
            End With
            num = num + 1
        Loop
    End With
End Sub


