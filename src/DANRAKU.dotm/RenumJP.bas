Attribute VB_Name = "RenumJP"
'�e��ԍ��U��Ȃ����}�N��
'   Created By D*suke YAMAKWA
'   last modified : 12/10/16
'
'�i���ԍ��̐U��Ȃ�������(RenumberingParagraph)�Ɋւ���
'   original code : �i���ԍ��u���}�N�� 03/08/23 By ���c
'   modified By D*suke YAMAKWA
'
'�y�}�N���̊T�v�z
' 1. �������́���i���ԍ��ɒu����������A
'    �������̒i���ԍ����A�A�ԂɂȂ�悤�ɏ��������܂��B
'
' 2. �������́��𐿋����ԍ��ɒu����������A
'    �������̐������ԍ����A�A�ԂɂȂ�悤�ɏ��������܂��B
'
' 3. �}�A���w���A�����A�\�̔ԍ���A�ԂɂȂ�悤�ɏ��������܂��B
'
'�y�����ӓ_�z
'   "��"�A"��"�������̓r���ɂ���ꍇ�ł��u������܂��B
'   ���̓r���ł�"��"�A"��"���g��Ȃ��A�������́A���L�s���폜���ĉ������B

Public Sub �i���ԍ��ȂǐU�蒼��JP()
    Call ReplaceAll("��", "�y�O�O�O�O�z") '���̒u���@�\���s�v�ȏꍇ�͂��̍s���폜
    Call RenumberingParagraph

    'Call ReplaceAll("��", "�y�������O�z") '���̒u���@�\���s�v�ȏꍇ�͂��̍s���폜
    'Call Renumbering("������")
    
    'Call Renumbering("�}")
    'Call Renumbering("��")
    'Call Renumbering("��")
    'Call Renumbering("�\")
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
        .Text = "�y[�O�P�Q�R�S�T�U�V�W�X]*�z"
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
                AddStr = "�y" + StrConv(Format(ParagraphNum, "0000"), vbWide) + "�z"
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
        .Text = "�y" & moji & "[�O�P�Q�R�S�T�U�V�W�X]*�z"
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
                AddStr = "�y" & moji & StrConv(num, vbWide) & "�z"
                .Font.Reset
                .InsertAfter (AddStr)
                .Move
            End With
            num = num + 1
        Loop
    End With
End Sub


