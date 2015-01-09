Attribute VB_Name = "DeleteNumJP"
Option Explicit

Public Sub íióéî‘çÜçÌèúJP()
    Dim num: num = "Åy[ÇOÇPÇQÇRÇSÇTÇUÇVÇWÇX]{4}Åz"
    Dim space: space = "[ Å@^t]{1,10}"
    Dim r: r = "[^11^13]"
    
    Call ReplaceToNull(space & num & space & r)
    Call ReplaceToNull(space & num & r)
    Call ReplaceToNull(num & space & r)
    Call ReplaceToNull(num & r)
End Sub

Private Sub ReplaceToNull(before)
    With ActiveDocument.Range(0, 0).Find
         .ClearFormatting
         .Text = before
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
         .Execute Replace:=wdReplaceAll
    End With
End Sub
