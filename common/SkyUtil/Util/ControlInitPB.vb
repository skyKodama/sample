#Region "�錾"
Option Explicit On 
Option Strict On

Imports System.Text
Imports IG = Infragistics.Win.UltraWinEditors
Imports IGM = Infragistics.Win.Misc
Imports IGMASK = Infragistics.Win.UltraWinMaskedEdit

#End Region

'*************************************************************************
'*�@�@�\�@�@�F���ʏ������E���Shared�N���X
'*            <Edit�ECombo�EDate�ERadioButton�ENumber���N���A>
'*            <Edit�ECombo�EDate�̏�ԕύX>
'*�@�쐬���@�F2006.04.25    ��
'*
'*  �ύX��  �F
'*
'*************************************************************************
''' <summary>
''' ���ʏ������E���Shared�N���X
''' </summary>
''' <remarks>��ʃR���g���[���ꊇ�g�p�s�E�������������Ȃ��܂��B</remarks>
Public Class ControlInitPB

#Region "����Shared���b�\�h����"

#Region "Control�������F�N���A�ETextVAlign=Middle���"
    ' --------------------------------------------------------------------------------
    ' <summary>
    '     �w�肵���R���g���[�����Ɋ܂܂�� 
    '      �FEdit�ECombo�EDate�ERadioButton�ENumber�ECheckbox���N���A�B</summary>
    ' <param name="ctrlParent">
    '     �����ΏۂƂȂ�e�R���g���[���B</param>
    ' --------------------------------------------------------------------------------
    ''' <summary>
    ''' �w�肵���R���g���[�����Ɋ܂܂��R���g���[���I�u�W�F�N�g������������
    ''' </summary>
    ''' <param name="ctrlParent">�Ăяo�����R���g���[��(��Ƀt�H�[��)</param>
    ''' <param name="WithDate">���t���������邩�ǂ��� 0�F��������(�����l)</param>
    ''' <remarks></remarks>
    Public Shared Sub doInitItems(ByVal ctrlParent As Control, Optional ByVal WithDate As Integer = 0)

        ' ctrlParent ���̂��ׂẴR���g���[����񋓂���
        For Each frmControl As Control In ctrlParent.Controls

            ' �񋓂����R���g���[���ɃR���g���[�����܂܂�Ă���ꍇ�͍ċA�Ăяo������
            If frmControl.HasChildren = True Then
                doInitItems(frmControl, WithDate)
            End If

            '�@�R���g���[���̌^(Edit)
            If TypeOf frmControl Is IG.UltraTextEditor Then
                If Not CType(frmControl, IG.UltraTextEditor).Tag Is "NotInit" Then
                    CType(frmControl, IG.UltraTextEditor).Clear()
                End If

            End If

            '�@�R���g���[���̌^(Mask)
            If TypeOf frmControl Is Infragistics.Win.UltraWinMaskedEdit.UltraMaskedEdit Then
                If Not CType(frmControl, Infragistics.Win.UltraWinMaskedEdit.UltraMaskedEdit).Tag Is "NotInit" Then
                    CType(frmControl, Infragistics.Win.UltraWinMaskedEdit.UltraMaskedEdit).ResetText()
                End If

            End If

            '�A�R���g���[���̌^(ComboBox)
            If TypeOf frmControl Is IG.UltraComboEditor Then
                If Not CType(frmControl, IG.UltraComboEditor).Tag Is "NotInit" Then
                    CType(frmControl, IG.UltraComboEditor).Clear()
                End If
            End If

            '�B�R���g���[���̌^(Date)
            If WithDate = 0 Then
                If TypeOf frmControl Is IG.UltraDateTimeEditor Then
                    If Not CType(frmControl, IG.UltraDateTimeEditor).Tag Is "NotInit" Then
                        CType(frmControl, IG.UltraDateTimeEditor).Value = Nothing
                        'CType(frmControl, IG.UltraDateTimeEditor).TextVAlign = IM.AlignVertical.Middle
                        'CType(frmControl, IG.UltraDateTimeEditor).TextHAlign = IM.AlignHorizontal.Center
                        CType(frmControl, IG.UltraDateTimeEditor).ImeMode = ImeMode.Disable  '�� ADD 2006.07.19
                    End If
                End If
            End If


            '�B�R���g���[���̌^(Date)
            If WithDate = 0 Then
                If TypeOf frmControl Is IG.UltraDateTimeEditor Then
                    If Not CType(frmControl, IG.UltraDateTimeEditor).Tag Is "NotInit" Then
                        CType(frmControl, IG.UltraDateTimeEditor).Value = Nothing
                        'CType(frmControl, IG.UltraDateTimeEditor).TextVAlign = IM.AlignVertical.Middle
                        'CType(frmControl, IG.UltraDateTimeEditor).TextHAlign = IM.AlignHorizontal.Center
                        CType(frmControl, IG.UltraDateTimeEditor).ImeMode = ImeMode.Disable  '�� ADD 2006.07.19
                    End If

                End If
            End If

            '�C�R���g���[���̌^(Number)
            If TypeOf frmControl Is IG.UltraNumericEditor Then
                CType(frmControl, IG.UltraNumericEditor).Value = Nothing
                CType(frmControl, IG.UltraNumericEditor).Appearance.TextVAlign = Infragistics.Win.VAlign.Middle
                CType(frmControl, IG.UltraNumericEditor).ImeMode = ImeMode.Disable  '�� ADD 2006.07.19
            End If

            '�D�R���g���[���̌^(RadioButton)
            If TypeOf frmControl Is RadioButton Then
                If Not CType(frmControl, RadioButton).Tag Is "NotInit" Then
                    CType(frmControl, RadioButton).Checked = False
                End If

            End If

            '�E�`�F�b�N�{�b�N�X(CheckBox)
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

            '�F�����N���x��
            If TypeOf frmControl Is LinkLabel Then
                If Not CType(frmControl, LinkLabel).Tag Is "NotInit" Then
                    CType(frmControl, LinkLabel).Text = ""
                End If

            End If

            '�G�e�L�X�g�{�b�N�X
            If TypeOf frmControl Is TextBox Then
                CType(frmControl, TextBox).Text = ""
            End If


            ''�X�e�[�^�X�o�[
            If TypeOf frmControl Is Infragistics.Win.UltraWinStatusBar.UltraStatusBar Then
                If CType(frmControl, Infragistics.Win.UltraWinStatusBar.UltraStatusBar).Panels.Count > 0 Then
                    CType(frmControl, Infragistics.Win.UltraWinStatusBar.UltraStatusBar).Panels(1).Text = ""
                End If
            End If

        Next frmControl
    End Sub
#End Region

#Region "Control�������F���"
    ' -----------------------------------------------------------------------------------------
    ' <summary>
    '     �w�肵���R���g���[�����Ɋ܂܂�� 
    '     Edit�ECombo�EDate�ENumber�ECheckbox�EButton�ELinkLabel�̏�ԕύX�B</summary>
    ' <param name="ctrlParent">
    '     �����ΏۂƂȂ�e�R���g���[���B</param>
    ' -----------------------------------------------------------------------------------------
    ''' <summary>
    ''' �w�肵���R���g���[�����Ɋ܂܂��R���g���[���I�u�W�F�N�g�̏�Ԃ�ύX����
    ''' </summary>
    ''' <param name="ctrlParent">�Ăяo�����R���g���[��(��Ƀt�H�[��)</param>
    ''' <param name="blFlg">Enabled =blFlg/Readonly=blFlg</param>
    ''' <remarks></remarks>
    Public Shared Sub doInitEnabled(ByVal ctrlParent As Control, ByVal blFlg As Boolean)
        '�g�p�s��
        Dim Frc As Color = Color.Blue
        Dim Bkc As Color = Color.WhiteSmoke
        '�g�p��
        Dim FrcN As Color = System.Drawing.SystemColors.WindowText
        Dim BkcN As Color = System.Drawing.SystemColors.Window
        '�K�{
        Dim BkcNeed As Color = Color.FromArgb(CType(255, Byte), CType(255, Byte), CType(192, Byte))

        For Each frmControl As Control In ctrlParent.Controls

            ' �񋓂����R���g���[���ɃR���g���[�����܂܂�Ă���ꍇ�͍ċA�Ăяo������
            If frmControl.HasChildren = True Then
                doInitEnabled(frmControl, blFlg)
            End If

            '�@�R���g���[���̌^(UltraTextEditor)
            If TypeOf frmControl Is IG.UltraTextEditor Then
                CType(frmControl, IG.UltraTextEditor).Appearance.BackColorDisabled = Bkc
                CType(frmControl, IG.UltraTextEditor).Appearance.ForeColorDisabled = Frc

                If CType(frmControl, IG.UltraTextEditor).Tag Is Nothing Then
                    CType(frmControl, IG.UltraTextEditor).Enabled = blFlg

                ElseIf CType(frmControl, IG.UltraTextEditor).Tag Is "" Then
                    CType(frmControl, IG.UltraTextEditor).Enabled = blFlg

                ElseIf CType(frmControl, IG.UltraTextEditor).Tag Is "ROALWAYS" Then
                    ''���Readonly
                    CType(frmControl, IG.UltraTextEditor).Enabled = True
                    CType(frmControl, IG.UltraTextEditor).ReadOnly = True
                    CType(frmControl, IG.UltraTextEditor).BackColor = Bkc
                    CType(frmControl, IG.UltraTextEditor).ForeColor = Frc


                ElseIf CType(frmControl, IG.UltraTextEditor).Tag Is "RO" Then

                    'ReadLony�ɐ؂�ւ��� 20080711_1
                    CType(frmControl, IG.UltraTextEditor).Enabled = True
                    CType(frmControl, IG.UltraTextEditor).ReadOnly = Not blFlg
                    'CType(frmControl, IG.UltraTextEditor).ReadOnly = Not blFlg
                    If Not blFlg Then
                        CType(frmControl, IG.UltraTextEditor).BackColor = Bkc
                        CType(frmControl, IG.UltraTextEditor).ForeColor = Frc
                    Else
                        '�K�{�ƒʏ��؂蕪����
                        If CType(frmControl, IG.UltraTextEditor).AccessibleDescription = "NEED" Then
                            CType(frmControl, IG.UltraTextEditor).BackColor = BkcNeed
                        Else
                            CType(frmControl, IG.UltraTextEditor).BackColor = BkcN
                        End If
                        CType(frmControl, IG.UltraTextEditor).ForeColor = FrcN
                    End If
                End If
            End If


            '�@�R���g���[���̌^(UltraTextEditor)
            If TypeOf frmControl Is IGMASK.UltraMaskedEdit Then
                CType(frmControl, IGMASK.UltraMaskedEdit).Appearance.BackColorDisabled = Bkc
                CType(frmControl, IGMASK.UltraMaskedEdit).Appearance.ForeColorDisabled = Frc

                If CType(frmControl, IGMASK.UltraMaskedEdit).Tag Is Nothing Then
                    CType(frmControl, IGMASK.UltraMaskedEdit).Enabled = blFlg

                ElseIf CType(frmControl, IGMASK.UltraMaskedEdit).Tag Is "" Then
                    CType(frmControl, IGMASK.UltraMaskedEdit).Enabled = blFlg

                ElseIf CType(frmControl, IG.UltraTextEditor).Tag Is "ROALWAYS" Then
                    ''���Readonly
                    CType(frmControl, IGMASK.UltraMaskedEdit).Enabled = True
                    CType(frmControl, IGMASK.UltraMaskedEdit).ReadOnly = True

                ElseIf CType(frmControl, IGMASK.UltraMaskedEdit).Tag Is "RO" Then

                    'ReadLony�ɐ؂�ւ��� 20080711_1
                    CType(frmControl, IGMASK.UltraMaskedEdit).Enabled = True
                    CType(frmControl, IGMASK.UltraMaskedEdit).ReadOnly = True
                    'CType(frmControl, IGMASK.UltraMaskedEdit).ReadOnly = Not blFlg
                    If Not blFlg Then
                        CType(frmControl, IGMASK.UltraMaskedEdit).BackColor = Bkc
                        CType(frmControl, IGMASK.UltraMaskedEdit).ForeColor = Frc
                    Else
                        '�K�{�ƒʏ��؂蕪����
                        If CType(frmControl, IGMASK.UltraMaskedEdit).AccessibleDescription = "NEED" Then
                            CType(frmControl, IGMASK.UltraMaskedEdit).BackColor = BkcNeed
                        Else
                            CType(frmControl, IGMASK.UltraMaskedEdit).BackColor = BkcN
                        End If
                        CType(frmControl, IGMASK.UltraMaskedEdit).ForeColor = FrcN
                    End If
                End If
            End If


            '�A�R���g���[���̌^(ComboBox)
            If TypeOf frmControl Is IG.UltraComboEditor Then
                CType(frmControl, IG.UltraComboEditor).Appearance.BackColorDisabled = Bkc
                CType(frmControl, IG.UltraComboEditor).Appearance.ForeColorDisabled = Frc

                If Not CType(frmControl, IG.UltraComboEditor).Tag Is "EF" Then
                    CType(frmControl, IG.UltraComboEditor).Enabled = blFlg
                End If
            End If

            '�B�R���g���[���̌^(Date)
            If TypeOf frmControl Is IG.UltraDateTimeEditor Then
                CType(frmControl, IG.UltraDateTimeEditor).Appearance.BackColorDisabled = Bkc
                CType(frmControl, IG.UltraDateTimeEditor).Appearance.ForeColorDisabled = Frc

                If Not CType(frmControl, IG.UltraDateTimeEditor).Tag Is "EF" Then
                    CType(frmControl, IG.UltraDateTimeEditor).Enabled = blFlg
                End If

            End If

            '�C�R���g���[���̌^(Number)
            If TypeOf frmControl Is IG.UltraNumericEditor Then
                CType(frmControl, IG.UltraNumericEditor).Appearance.BackColorDisabled = Bkc
                CType(frmControl, IG.UltraNumericEditor).Appearance.ForeColorDisabled = Frc

                If Not CType(frmControl, IG.UltraNumericEditor).Tag Is "EF" Then
                    CType(frmControl, IG.UltraNumericEditor).Enabled = blFlg
                End If
            End If

            '�D�R���g���[���̌^(Checkbox)
            If TypeOf frmControl Is CheckBox Then
                CType(frmControl, CheckBox).Enabled = blFlg
            End If

            If TypeOf frmControl Is IG.UltraCheckEditor Then

                    CType(frmControl, IG.UltraCheckEditor).Enabled = blFlg

            End If


            '�E�R���g���[���̌^(Button)
            If TypeOf frmControl Is Button Then
                CType(frmControl, Button).Enabled = blFlg
            End If

            '�E�R���g���[���̌^(Button)
            If TypeOf frmControl Is IGM.UltraButton Then
                CType(frmControl, IGM.UltraButton).Enabled = blFlg
            End If

            '��ADD 2006.07.11
            '�F�R���g���[���̌^(LinkLabel)
            If TypeOf frmControl Is LinkLabel Then
                If CType(frmControl, LinkLabel).Tag Is "EF" Then
                    CType(frmControl, LinkLabel).Enabled = False

                ElseIf CType(frmControl, LinkLabel).Tag Is "ET" Then
                    CType(frmControl, LinkLabel).Enabled = True

                Else
                    CType(frmControl, LinkLabel).Enabled = blFlg
                End If
            End If

            '��ADD 2006.07.21
            '�G�R���g���[���̌^(RadioButton)
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

