Attribute VB_Name = "DeleteNumJP"
Option Explicit

Public Sub �i���ԍ��폜JP()
    With ActiveDocument.Range().Find
        .ClearFormatting
        .Text = "�y[�O�P�Q�R�S�T�U�V�W�X]*�z[^11^13]"
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
            End With
        Loop
    End With
End Sub
