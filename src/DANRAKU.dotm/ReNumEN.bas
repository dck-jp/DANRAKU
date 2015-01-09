Attribute VB_Name = "ReNumEN"
'�e��ԍ��U��Ȃ����}�N��
'   Created By D*suke YAMAKWA
'   last modified : 14/4/7
'
'�i���ԍ��̐U��Ȃ�������(RenumberingParagraph)�Ɋւ���
'   original code : �i���ԍ��u���}�N�� 03/08/23 By ���c
'   modified By D*suke YAMAKWA
'
'�y�}�N���̊T�v�z
' 1. �������́���i���ԍ��ɒu����������A
'    �������̒i���ԍ����A�A�ԂɂȂ�悤�ɏ��������܂��B
'
'�y�����ӓ_�z
'   "��"�������̓r���ɂ���ꍇ�ł��u������܂��B

Sub �i���ԍ��U�蒼��EN()
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

