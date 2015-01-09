Attribute VB_Name = "DeleteNumJP"
Option Explicit

Public Sub �i���ԍ��폜JP()
    Dim num: num = "�y[�O�P�Q�R�S�T�U�V�W�X]{4}�z"
    Dim space: space = "[ �@^t]{1,10}"
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
