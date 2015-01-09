Attribute VB_Name = "AutoNumJP"
Option Explicit

Public Sub �i���ԍ������t�^JP()
    �i���ԍ��폜JP
    AddParagraphMarker
    �i���ԍ��ȂǐU�蒼��JP
End Sub

Private Sub AddParagraphMarker()
    Dim subTitles: subTitles = Array("�y�Z�p����z", "�y�w�i�Z�p�z", "�y���������z", "�y����������z" _
                                        , "�y�������������悤�Ƃ���ۑ�z", "�y�ۑ���������邽�߂̎�i�z" _
                                        , "�y�����̌��ʁz", "�y�}�ʂ̊ȒP�Ȑ����z", "�y���������{���邽�߂̌`�ԁz" _
                                        , "�y���{��z", "�y�Y�Ə�̗��p�\���z", "�y�����̐����z")

    ' �e���ڂ̌��ɕϊ��p�̋L����t�^
    Dim i: For Each i In subTitles
    Call ReplaceX(i & "([^11^13])", i & "\1��\1")
    Next
    ' �y��,��,�\ �������s�̏�Ɂ��ǉ�
    Dim formula: formula = "�y[�����\]"
    Call ReplaceX("([^11^13])([�@ ]{1,10})" & formula & "([�O-�X]{1,2})(�z)", "\1��\1\2" & formula & "\3\4")
    Call ReplaceX("([^11^13])" & formula & "([�O-�X]{1,2})(�z)", "\1��\1" & formula & "\2\3")
    ' �y���{�ၖ�z�̒���Ɂ��ǉ��B
    Call ReplaceX("(�y���{��[�O-�X]{1,2}�z)([^11^13])", "\1\2��\2")
    ' �w�@�B\r\n�@�y�ȊO�@�x�̏ꍇ�A�B�̎��̍s���Ɂ���ǉ�
    Call ReplaceX("�B([^11^13])�@([!�y])", "�B\1��\1�@\2")
End Sub

Private Sub ReplaceX(before, after)
    With ActiveDocument.Range(0, 0).Find
         .Text = before
         .Replacement.Text = after
         .MatchFuzzy = False
         .MatchWildcards = True
         .Execute Replace:=wdReplaceAll
    End With
End Sub
