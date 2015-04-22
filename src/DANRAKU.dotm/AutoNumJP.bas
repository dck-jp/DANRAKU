Attribute VB_Name = "AutoNumJP"
Option Explicit

Public Sub �i���ԍ������t�^JP()
    �i���ԍ��폜JP
    AddParagraphMarker
    �i���ԍ��ȂǐU�蒼��JP
End Sub

Private Sub AddParagraphMarker()
    Dim paragraphMarker: paragraphMarker = "�@��"
    Dim subTitles: subTitles = Array("�y�Z�p����z", "�y�w�i�Z�p�z", "�y���������z", "�y����������z" _
                                        , "�y�������������悤�Ƃ���ۑ�z", "�y�ۑ���������邽�߂̎�i�z" _
                                        , "�y�����̌��ʁz", "�y�}�ʂ̊ȒP�Ȑ����z", "�y���������{���邽�߂̌`�ԁz" _
                                        , "�y���{��z", "�y�Y�Ə�̗��p�\���z", "�y�����̐����z")

    ' �e���ڂ̌��ɕϊ��p�̋L����t�^
    Dim i: For Each i In subTitles
    Call ReplaceX(i & "([^11^13])", i & "\1" & paragraphMarker & "\1")
    Next
    ' �y���{�ၖ�z�̒���Ɂ��ǉ��B
    Call ReplaceX("(�y���{��[�O-�X]{1,2}�z)([^11^13])", "\1\2" & paragraphMarker & "\2")
    
    ' �w�@�B\r\n�@�y�ȊO�@�x�̏ꍇ�A�B�̎��̍s���Ɂ���ǉ�
    Call ReplaceX("�B([^11^13])�@([!�y])", "�B\1" & paragraphMarker & "\1�@\2")
    
    ' �y��,��,�\,�i,�m �������s�̏�Ɂ��ǉ�
    Dim userMarkers: userMarkers = Array("�i", "�m", "�y��", "�y��", "�y�\")
    Dim userMarker: For Each userMarker In userMarkers
        Call ReplaceX("([^11^13])" & _
                      "([�@ ]{1,10}" & userMarker & ")", _
 _
                      "\1" & _
                      userMarker & "\1" _
                      & "\2")
    
        Call ReplaceX("([^11^13])" & _
                      "(" & userMarker & ")", _
 _
                      "\1" & _
                      paragraphMarker & "\1" _
                      & "\2")
    
    Next
    
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
