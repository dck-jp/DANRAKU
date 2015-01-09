Attribute VB_Name = "DeleteNumEN"
Option Explicit

Public Sub íióéî‘çÜçÌèúEN()
    With ActiveDocument.Range().Find
        .ClearFormatting
        .Text = "\[[0-9]*\][^11^13]"
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

