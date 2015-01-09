Attribute VB_Name = "Callbacks"
Option Explicit

'Callback for Auto_JP onAction
Sub CallbackAuto_JP(control As IRibbonControl)
    段落番号自動付与JP
End Sub

'Callback for Renum_JP onAction
Sub CallbackRenum_JP(control As IRibbonControl)
    段落番号など振り直しJP
End Sub

'Callback for Conv_JP onAction
Sub CallbackConv_JP(control As IRibbonControl)
    段落記号変換JP2EN
End Sub

'Callback for Delete_JP onAction
Sub CallbackDelete_JP(control As IRibbonControl)
    段落番号削除JP
End Sub

'Callback for Renum_EN onAction
Sub CallbackRenum_EN(control As IRibbonControl)
    段落番号振り直しEN
End Sub

'Callback for Conv_EN onAction
Sub CallbackConv_EN(control As IRibbonControl)
    段落記号変換EN2JP
End Sub

'Callback for Delete_EN onAction
Sub CallbackDelete_EN(control As IRibbonControl)
    段落番号削除EN
End Sub


'Callback for Help onAction
Sub CallbackHelp(control As IRibbonControl)
    段落番号ヘルプ
End Sub


