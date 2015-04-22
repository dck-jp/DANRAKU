Attribute VB_Name = "AutoNumJP"
Option Explicit

Public Sub ’i—”Ô†©“®•t—^JP()
    ’i—”Ô†íœJP
    AddParagraphMarker
    ’i—”Ô†‚È‚ÇU‚è’¼‚µJP
End Sub

Private Sub AddParagraphMarker()
    Dim paragraphMarker: paragraphMarker = "@—"
    Dim subTitles: subTitles = Array("y‹Zp•ª–ìz", "y”wŒi‹Zpz", "y“Á‹–•¶Œ£z", "y”ñ“Á‹–•¶Œ£z" _
                                        , "y”­–¾‚ª‰ğŒˆ‚µ‚æ‚¤‚Æ‚·‚é‰Û‘èz", "y‰Û‘è‚ğ‰ğŒˆ‚·‚é‚½‚ß‚Ìè’iz" _
                                        , "y”­–¾‚ÌŒø‰Êz", "y}–Ê‚ÌŠÈ’P‚Èà–¾z", "y”­–¾‚ğÀ{‚·‚é‚½‚ß‚ÌŒ`‘Ôz" _
                                        , "yÀ{—áz", "yY‹Æã‚Ì—˜—p‰Â”\«z", "y•„†‚Ìà–¾z")

    ' Še€–Ú‚ÌŒã‚ë‚É•ÏŠ·—p‚Ì‹L†‚ğ•t—^
    Dim i: For Each i In subTitles
    Call ReplaceX(i & "([^11^13])", i & "\1" & paragraphMarker & "\1")
    Next
    ' yÀ{—á–z‚Ì’¼Œã‚É—’Ç‰ÁB
    Call ReplaceX("(yÀ{—á[‚O-‚X]{1,2}z)([^11^13])", "\1\2" & paragraphMarker & "\2")
    
    ' w@B\r\n@yˆÈŠO@x‚Ìê‡AB‚ÌŸ‚Ìs“ª‚É—‚ğ’Ç‰Á
    Call ReplaceX("B([^11^13])@([!y])", "B\1" & paragraphMarker & "\1@\2")
    
    ' y‰»,”,•\,i,m ‚ªŒ»‚ê‚és‚Ìã‚É—’Ç‰Á
    Dim userMarkers: userMarkers = Array("i", "m", "y‰»", "y”", "y•\")
    Dim userMarker: For Each userMarker In userMarkers
        Call ReplaceX("([^11^13])" & _
                      "([@ ]{1,10}" & userMarker & ")", _
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
