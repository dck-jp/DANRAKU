Attribute VB_Name = "DeleteNumEN"
Option Explicit

Public Sub íióéî‘çÜçÌèúEN()
    Dim num: num = "\[[0-9]{4}\]"
    Dim space: space = "[ ^t]{1,10}"
    Dim r: r = "[^11^13]"
    Dim anyString: anyString = "[!^11^13]{1,}"
    
    Call ReplaceToNull(space & num & space & r)
    Call ReplaceToNull(space & num & r)
    Call ReplaceToNull(num & space & r)
    Call ReplaceToNull(num & r)
    
    Call ReplaceTo(space & num & "(" & anyString & ")", "\1")
    Call ReplaceTo(num & "(" & anyString & ")", "\1")
End Sub

Private Sub ReplaceToNull(before)
    Call ReplaceTo(before, "")
End Sub

Private Sub ReplaceTo(before, after)
    With ActiveDocument.Range(0, 0).Find
         .ClearFormatting
         .Text = before
         .Replacement.Text = after
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
