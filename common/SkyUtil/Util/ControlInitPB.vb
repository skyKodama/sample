#Region "宣言"
Option Explicit On 
Option Strict On

Imports System.Text
Imports IG = Infragistics.Win.UltraWinEditors
Imports IGM = Infragistics.Win.Misc
Imports IGMASK = Infragistics.Win.UltraWinMaskedEdit

#End Region

'*************************************************************************
'*　機能　　：共通初期化・状態Sharedクラス
'*            <Edit・Combo・Date・RadioButton・Numberをクリア>
'*            <Edit・Combo・Dateの状態変更>
'*　作成日　：2006.04.25    黄
'*
'*  変更日  ：
'*
'*************************************************************************
''' <summary>
''' 共通初期化・状態Sharedクラス
''' </summary>
''' <remarks>画面コントロール一括使用可不可・初期化をおこないます。</remarks>
Public Class ControlInitPB

#Region "★※Sharedメッソド※★"

#Region "Control初期化：クリア・TextVAlign=Middleｾｯﾄ"
    ' --------------------------------------------------------------------------------
    ' <summary>
    '     指定したコントロール内に含まれる 
    '      ：Edit・Combo・Date・RadioButton・Number・Checkboxをクリア。</summary>
    ' <param name="ctrlParent">
    '     検索対象となる親コントロール。</param>
    ' --------------------------------------------------------------------------------
    ''' <summary>
    ''' 指定したコントロール内に含まれるコントロールオブジェクトを初期化する
    ''' </summary>
    ''' <param name="ctrlParent">呼び出し元コントロール(主にフォーム)</param>
    ''' <param name="WithDate">日付を処理するかどうか 0：処理する(初期値)</param>
    ''' <remarks></remarks>
    Public Shared Sub doInitItems(ByVal ctrlParent As Control, Optional ByVal WithDate As Integer = 0)

        ' ctrlParent 内のすべてのコントロールを列挙する
        For Each frmControl As Control In ctrlParent.Controls

            ' 列挙したコントロールにコントロールが含まれている場合は再帰呼び出しする
            If frmControl.HasChildren = True Then
                doInitItems(frmControl, WithDate)
            End If

            '①コントロールの型(Edit)
            If TypeOf frmControl Is IG.UltraTextEditor Then
                If Not CType(frmControl, IG.UltraTextEditor).Tag Is "NotInit" Then
                    CType(frmControl, IG.UltraTextEditor).Clear()
                End If

            End If

            '①コントロールの型(Mask)
            If TypeOf frmControl Is Infragistics.Win.UltraWinMaskedEdit.UltraMaskedEdit Then
                If Not CType(frmControl, Infragistics.Win.UltraWinMaskedEdit.UltraMaskedEdit).Tag Is "NotInit" Then
                    CType(frmControl, Infragistics.Win.UltraWinMaskedEdit.UltraMaskedEdit).ResetText()
                End If

            End If

            '②コントロールの型(ComboBox)
            If TypeOf frmControl Is IG.UltraComboEditor Then
                If Not CType(frmControl, IG.UltraComboEditor).Tag Is "NotInit" Then
                    CType(frmControl, IG.UltraComboEditor).Clear()
                End If
            End If

            '③コントロールの型(Date)
            If WithDate = 0 Then
                If TypeOf frmControl Is IG.UltraDateTimeEditor Then
                    If Not CType(frmControl, IG.UltraDateTimeEditor).Tag Is "NotInit" Then
                        CType(frmControl, IG.UltraDateTimeEditor).Value = Nothing
                        'CType(frmControl, IG.UltraDateTimeEditor).TextVAlign = IM.AlignVertical.Middle
                        'CType(frmControl, IG.UltraDateTimeEditor).TextHAlign = IM.AlignHorizontal.Center
                        CType(frmControl, IG.UltraDateTimeEditor).ImeMode = ImeMode.Disable  '← ADD 2006.07.19
                    End If
                End If
            End If


            '③コントロールの型(Date)
            If WithDate = 0 Then
                If TypeOf frmControl Is IG.UltraDateTimeEditor Then
                    If Not CType(frmControl, IG.UltraDateTimeEditor).Tag Is "NotInit" Then
                        CType(frmControl, IG.UltraDateTimeEditor).Value = Nothing
                        'CType(frmControl, IG.UltraDateTimeEditor).TextVAlign = IM.AlignVertical.Middle
                        'CType(frmControl, IG.UltraDateTimeEditor).TextHAlign = IM.AlignHorizontal.Center
                        CType(frmControl, IG.UltraDateTimeEditor).ImeMode = ImeMode.Disable  '← ADD 2006.07.19
                    End If

                End If
            End If

            '④コントロールの型(Number)
            If TypeOf frmControl Is IG.UltraNumericEditor Then
                CType(frmControl, IG.UltraNumericEditor).Value = Nothing
                CType(frmControl, IG.UltraNumericEditor).Appearance.TextVAlign = Infragistics.Win.VAlign.Middle
                CType(frmControl, IG.UltraNumericEditor).ImeMode = ImeMode.Disable  '← ADD 2006.07.19
            End If

            '⑤コントロールの型(RadioButton)
            If TypeOf frmControl Is RadioButton Then
                If Not CType(frmControl, RadioButton).Tag Is "NotInit" Then
                    CType(frmControl, RadioButton).Checked = False
                End If

            End If

            '⑥チェックボックス(CheckBox)
            If TypeOf frmControl Is CheckBox Then
                If Not CType(frmControl, CheckBox).Tag Is "NotInit" Then
                    CType(frmControl, CheckBox).Checked = False
                End If
            End If
            If TypeOf frmControl Is IG.UltraCheckEditor Then
                If Not CType(frmControl, IG.UltraCheckEditor).Tag Is "NotInit" Then
                    CType(frmControl, IG.UltraCheckEditor).Checked = False
                End If
            End If

            '⑦リンクラベル
            If TypeOf frmControl Is LinkLabel Then
                If Not CType(frmControl, LinkLabel).Tag Is "NotInit" Then
                    CType(frmControl, LinkLabel).Text = ""
                End If

            End If

            '⑧テキストボックス
            If TypeOf frmControl Is TextBox Then
                CType(frmControl, TextBox).Text = ""
            End If


            ''ステータスバー
            If TypeOf frmControl Is Infragistics.Win.UltraWinStatusBar.UltraStatusBar Then
                If CType(frmControl, Infragistics.Win.UltraWinStatusBar.UltraStatusBar).Panels.Count > 0 Then
                    CType(frmControl, Infragistics.Win.UltraWinStatusBar.UltraStatusBar).Panels(1).Text = ""
                End If
            End If

        Next frmControl
    End Sub
#End Region

#Region "Control初期化：状態"
    ' -----------------------------------------------------------------------------------------
    ' <summary>
    '     指定したコントロール内に含まれる 
    '     Edit・Combo・Date・Number・Checkbox・Button・LinkLabelの状態変更。</summary>
    ' <param name="ctrlParent">
    '     検索対象となる親コントロール。</param>
    ' -----------------------------------------------------------------------------------------
    ''' <summary>
    ''' 指定したコントロール内に含まれるコントロールオブジェクトの状態を変更する
    ''' </summary>
    ''' <param name="ctrlParent">呼び出し元コントロール(主にフォーム)</param>
    ''' <param name="blFlg">Enabled =blFlg/Readonly=blFlg</param>
    ''' <remarks></remarks>
    Public Shared Sub doInitEnabled(ByVal ctrlParent As Control, ByVal blFlg As Boolean)
        '使用不可
        Dim Frc As Color = Color.Blue
        Dim Bkc As Color = Color.WhiteSmoke
        '使用可
        Dim FrcN As Color = System.Drawing.SystemColors.WindowText
        Dim BkcN As Color = System.Drawing.SystemColors.Window
        '必須
        Dim BkcNeed As Color = Color.FromArgb(CType(255, Byte), CType(255, Byte), CType(192, Byte))

        For Each frmControl As Control In ctrlParent.Controls

            ' 列挙したコントロールにコントロールが含まれている場合は再帰呼び出しする
            If frmControl.HasChildren = True Then
                doInitEnabled(frmControl, blFlg)
            End If

            '①コントロールの型(UltraTextEditor)
            If TypeOf frmControl Is IG.UltraTextEditor Then
                CType(frmControl, IG.UltraTextEditor).Appearance.BackColorDisabled = Bkc
                CType(frmControl, IG.UltraTextEditor).Appearance.ForeColorDisabled = Frc

                If CType(frmControl, IG.UltraTextEditor).Tag Is Nothing Then
                    CType(frmControl, IG.UltraTextEditor).Enabled = blFlg

                ElseIf CType(frmControl, IG.UltraTextEditor).Tag Is "" Then
                    CType(frmControl, IG.UltraTextEditor).Enabled = blFlg

                ElseIf CType(frmControl, IG.UltraTextEditor).Tag Is "ROALWAYS" Then
                    ''常にReadonly
                    CType(frmControl, IG.UltraTextEditor).Enabled = True
                    CType(frmControl, IG.UltraTextEditor).ReadOnly = True
                    CType(frmControl, IG.UltraTextEditor).BackColor = Bkc
                    CType(frmControl, IG.UltraTextEditor).ForeColor = Frc


                ElseIf CType(frmControl, IG.UltraTextEditor).Tag Is "RO" Then

                    'ReadLonyに切り替える 20080711_1
                    CType(frmControl, IG.UltraTextEditor).Enabled = True
                    CType(frmControl, IG.UltraTextEditor).ReadOnly = Not blFlg
                    'CType(frmControl, IG.UltraTextEditor).ReadOnly = Not blFlg
                    If Not blFlg Then
                        CType(frmControl, IG.UltraTextEditor).BackColor = Bkc
                        CType(frmControl, IG.UltraTextEditor).ForeColor = Frc
                    Else
                        '必須と通常を切り分ける
                        If CType(frmControl, IG.UltraTextEditor).AccessibleDescription = "NEED" Then
                            CType(frmControl, IG.UltraTextEditor).BackColor = BkcNeed
                        Else
                            CType(frmControl, IG.UltraTextEditor).BackColor = BkcN
                        End If
                        CType(frmControl, IG.UltraTextEditor).ForeColor = FrcN
                    End If
                End If
            End If


            '①コントロールの型(UltraTextEditor)
            If TypeOf frmControl Is IGMASK.UltraMaskedEdit Then
                CType(frmControl, IGMASK.UltraMaskedEdit).Appearance.BackColorDisabled = Bkc
                CType(frmControl, IGMASK.UltraMaskedEdit).Appearance.ForeColorDisabled = Frc

                If CType(frmControl, IGMASK.UltraMaskedEdit).Tag Is Nothing Then
                    CType(frmControl, IGMASK.UltraMaskedEdit).Enabled = blFlg

                ElseIf CType(frmControl, IGMASK.UltraMaskedEdit).Tag Is "" Then
                    CType(frmControl, IGMASK.UltraMaskedEdit).Enabled = blFlg

                ElseIf CType(frmControl, IG.UltraTextEditor).Tag Is "ROALWAYS" Then
                    ''常にReadonly
                    CType(frmControl, IGMASK.UltraMaskedEdit).Enabled = True
                    CType(frmControl, IGMASK.UltraMaskedEdit).ReadOnly = True

                ElseIf CType(frmControl, IGMASK.UltraMaskedEdit).Tag Is "RO" Then

                    'ReadLonyに切り替える 20080711_1
                    CType(frmControl, IGMASK.UltraMaskedEdit).Enabled = True
                    CType(frmControl, IGMASK.UltraMaskedEdit).ReadOnly = True
                    'CType(frmControl, IGMASK.UltraMaskedEdit).ReadOnly = Not blFlg
                    If Not blFlg Then
                        CType(frmControl, IGMASK.UltraMaskedEdit).BackColor = Bkc
                        CType(frmControl, IGMASK.UltraMaskedEdit).ForeColor = Frc
                    Else
                        '必須と通常を切り分ける
                        If CType(frmControl, IGMASK.UltraMaskedEdit).AccessibleDescription = "NEED" Then
                            CType(frmControl, IGMASK.UltraMaskedEdit).BackColor = BkcNeed
                        Else
                            CType(frmControl, IGMASK.UltraMaskedEdit).BackColor = BkcN
                        End If
                        CType(frmControl, IGMASK.UltraMaskedEdit).ForeColor = FrcN
                    End If
                End If
            End If


            '②コントロールの型(ComboBox)
            If TypeOf frmControl Is IG.UltraComboEditor Then
                CType(frmControl, IG.UltraComboEditor).Appearance.BackColorDisabled = Bkc
                CType(frmControl, IG.UltraComboEditor).Appearance.ForeColorDisabled = Frc

                If Not CType(frmControl, IG.UltraComboEditor).Tag Is "EF" Then
                    CType(frmControl, IG.UltraComboEditor).Enabled = blFlg
                End If
            End If

            '③コントロールの型(Date)
            If TypeOf frmControl Is IG.UltraDateTimeEditor Then
                CType(frmControl, IG.UltraDateTimeEditor).Appearance.BackColorDisabled = Bkc
                CType(frmControl, IG.UltraDateTimeEditor).Appearance.ForeColorDisabled = Frc

                If Not CType(frmControl, IG.UltraDateTimeEditor).Tag Is "EF" Then
                    CType(frmControl, IG.UltraDateTimeEditor).Enabled = blFlg
                End If

            End If

            '④コントロールの型(Number)
            If TypeOf frmControl Is IG.UltraNumericEditor Then
                CType(frmControl, IG.UltraNumericEditor).Appearance.BackColorDisabled = Bkc
                CType(frmControl, IG.UltraNumericEditor).Appearance.ForeColorDisabled = Frc

                If Not CType(frmControl, IG.UltraNumericEditor).Tag Is "EF" Then
                    CType(frmControl, IG.UltraNumericEditor).Enabled = blFlg
                End If
            End If

            '⑤コントロールの型(Checkbox)
            If TypeOf frmControl Is CheckBox Then
                CType(frmControl, CheckBox).Enabled = blFlg
            End If

            If TypeOf frmControl Is IG.UltraCheckEditor Then

                    CType(frmControl, IG.UltraCheckEditor).Enabled = blFlg

            End If


            '⑥コントロールの型(Button)
            If TypeOf frmControl Is Button Then
                CType(frmControl, Button).Enabled = blFlg
            End If

            '⑥コントロールの型(Button)
            If TypeOf frmControl Is IGM.UltraButton Then
                CType(frmControl, IGM.UltraButton).Enabled = blFlg
            End If

            '↓ADD 2006.07.11
            '⑦コントロールの型(LinkLabel)
            If TypeOf frmControl Is LinkLabel Then
                If CType(frmControl, LinkLabel).Tag Is "EF" Then
                    CType(frmControl, LinkLabel).Enabled = False

                ElseIf CType(frmControl, LinkLabel).Tag Is "ET" Then
                    CType(frmControl, LinkLabel).Enabled = True

                Else
                    CType(frmControl, LinkLabel).Enabled = blFlg
                End If
            End If

            '↓ADD 2006.07.21
            '⑧コントロールの型(RadioButton)
            If TypeOf frmControl Is RadioButton Then
                If Not CType(frmControl, RadioButton).Tag Is "EF" Then
                    CType(frmControl, RadioButton).Enabled = blFlg
                End If
            End If
        Next frmControl
    End Sub
#End Region

#End Region

End Class

