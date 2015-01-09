Attribute VB_Name = "DeleteNumJP"
Option Explicit

Public Sub íióéî‘çÜçÌèúJP()
    With ActiveDocument.Range().Find
        .ClearFormatting
        .Text = "Åy[ÇOÇPÇQÇRÇSÇTÇUÇVÇWÇX]*Åz[^11^13]"
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
