Option Explicit On
Option Strict On

Imports System.Text
Imports System.Data.SqlClient
Imports GrapeCity.Win.Input.DateTimeEx
Imports IM = GrapeCity.Win.Input
Imports skysystem.common.SystemUtil
Imports skysystem.common.SystemConst
Imports skysystem.common
'*********************************************************************************************
'*�@�@�\�@�@�F���ʃn���h���N���X
'*            <Button�ECombo�EEnter�R���g���[���C�x���g����N���X>
'*�@�쐬���@�F2006.04.25    ��
'*
'*�@���ύX���e��
'*  2006.05.24    ��    �R���{�{�b�N�X��s�}��
'*  2006.06.27    ��    �R���{�{�b�N�XContents ������skylog(targetCombo_ValueChanged)
'*  2006.06.29    ��    InputMan Number���۰ׂ�NegativeValueskylog(targetIMNumber_ValueChanged)
'*  20060630�@    ��`  �����ޯ��������(PBS_SetComboDS)
'*  2006.07.05    ��    �����ޯ���ر(PBS_SetComboDS) 
'*                      StackOverflowException�����h�~(targetCombo_ValueChanged)
'*  2006.07.10    ��    DropDownStyle(�����l�FDropDown), Field��p����(strAS)
'*                      ChangedComboValue(Combo�l�ύX���AEvent����������) 
'*  2006.07.11    ��    LinkLabel �g�p�¥�s�ɂ��ALabel��ԕύX
'*  2006.07.13    ��    Content�̏�����skylog(targetCombo_ValueChanged)���ď���
'*  2006.07.19    ��    ImeMode�ǉ�(PBS_SetComboDS), LINK_LABEL��TabStop�ݒ�(Default:False)
'*  2006.07.26    ��    PBS_SetComboDS�F���޾�Ă��AMaxDropDownItems�̕s��C��
'*  2006.07.28    ��    targetCombo_ValueChanged����ē�񔭐��̕s��C��
'*  20060807_1    ��`  PBS_SetComboDS DEL
'*  2006.09.06    ��    �����ޯ��Width����(PBS_SetComboDS, intMinus�ǉ�)
'*  2006.09.11    ��    �����ޯ��DropDownList�ݒ�
'*  20060921_1    ��    PBS_SetComboDS(Tran�ǉ�)
'*  20061122_1    ��`  OrderBy��̒ǉ�
'*********************************************************************************************
''' <summary>
''' ���ʃn���h���N���X
''' </summary>
''' <remarks></remarks>
Public Class ControlHandlesPB

#Region "Private�ϐ�"
    Private WithEvents targetModeButton As Button   '���[�h�{�^��
    Private targetModeLabel As Label                '���[�h�̃��x��

    Private WithEvents targetCtrl As Control        'Enter�L�[���󂯓����R���g���[��
    Private startCtrl As Control                    '�t�H�[�J�X�̈ړ���ƂȂ�R���g���[��
    Private baseForm As Form                        '�R���g���[����ێ����Ă���t�H�[��

    'Private WithEvents targetCombo As IM.Combo      '�^�[�Q�b�g�̃R���{�{�b�N�X
    'Private strPattern As String                    '�R���{�{�b�N�XPattern�Z�b�g

    'Private WithEvents targetIMNumber As IM.Number  'HighlightText��ĺ��۰�(
    'Private WithEvents targetIMEdit As IM.Edit      '   (Default=True�ł����A���܂���Ăł��Ȃ������̂ŁA�����I��Ă̂��߁B)

    Private WithEvents targetStbar As StatusBar      ''�X�e�C�^�X�o�[ 20061225_1 

    Private oda As SqlDataAdapter
    Private ocd As SqlCommand
    Private dts As DataSet
    Private blnFlgInit As Boolean = False           '�Z�b�g�R���{�������t���O
    Private strSQL As String

    Private bytHighNega As Byte                     'IM.Number�^�̲���Đ���(0�FHighLightText�s�skylog, 1�FNegativeColorskylog)

    '��ADD 2006.07.10 (�����ޯ���l���ύX���ꂽ���A̫�ё��ŏ����������ꍇ)
    Public Event ChangedComboValue(ByVal sender As Object, ByVal strKey As String)

    '��ADD 2006.07.11
    Private WithEvents lbl_LinkLabel As LinkLabel
    Private lbl_TitleLabel As Label
    Private fontStyleLeave As FontStyle     'MouseLeave����FontStyle(Title���������F���͌n��}�X�^�n)

#End Region


#Region "�R���X�g���N�^"
    ' ------------------------------------------------------------------ 
    '�@�@�\�F�X�e�[�^�X�o�[ �쐬�ҁF�X�V�҂̒ǉ�
    '�@�����FfrmStbar
    ' ------------------------------------------------------------------ 
    Public Sub New(ByVal frmStbar As StatusBar)
        Me.targetStbar = frmStbar

    End Sub '20061225_1

    ' ------------------------------------------------------------------ 
    '�@�@�\�FEdit�ENumber��ImputMan���۰ׂ̕s��΍�(HighLightText)
    '�@�����FIMNumber�EIMEdit
    ' ------------------------------------------------------------------ 
    Public Sub New(ByVal frmIMNumber As IM.Number, Optional ByVal highNega As Byte = 0)
        Me.targetIMNumber = frmIMNumber
        Me.bytHighNega = highNega       'HighLight(0)�ENegative�敪�׸�(1)
    End Sub
    Public Sub New(ByVal frmIMEdit As IM.Edit)
        Me.targetIMEdit = frmIMEdit
    End Sub

    ' ------------------------------------------------------------------ 
    '�@�@�\�F�{�^���C�x���g�����p(Enabled=True�EFalse�ɂ��A���x���̏�ԕύX)
    '  �����FfrmButton (��ԕύX�ɂ��A���x���ύX)
    '        frmLabel (�{�^���ɂ���āA��ԕύX)
    ' ------------------------------------------------------------------ 
    Public Sub New(ByVal frmButton As Button, _
                   ByVal frmLabel As Label)
        Me.targetModeButton = frmButton
        Me.targetModeLabel = frmLabel
    End Sub

    ' ------------------------------------------------------------------ 
    '�@�@�\�F�L�[Enter�n���h���p
    '  �����FtargetCtrl (Enter�L�[���󂯓����Control)
    '        startCtrl (�t�H�[�J�X�ړ����Control)
    '        baseForm (Control��ێ�����t�H�[��)
    ' ------------------------------------------------------------------ 
    Public Sub New(ByVal targetCtrl As Control, _
                   ByVal startCtrl As Control, _
                   ByVal baseForm As Form)
        Me.targetCtrl = targetCtrl
        Me.startCtrl = startCtrl
        Me.baseForm = baseForm
    End Sub

    ' ------------------------------------------------------------------ 
    '�@�@�\�F�R���{�`�F���W�n���h���p
    '  �����FtargetCombo (�R���{ValueChanged����Control)
    ' ------------------------------------------------------------------ 
    Public Sub New(ByVal targetCombo As IM.Combo)
        Me.targetCombo = targetCombo
        ''Me.blnFlgSpace = blnSpace
    End Sub


    '��ADD 2006.07.11
    ' ------------------------------------------------------------------------- 
    '�@�@�\�FLinkLabel�n���h���p
    '  �����Flbl_LinkLabel (MouseEnter�MouseLeave�EnabledChanged����Control)
    '        lbl_TitleLabel(LinkLabel���۰ق�Enabled=False�̏ꍇ�A��p���)
    ' ------------------------------------------------------------------------- 
    Public Sub New(ByVal targetLinkLabel As LinkLabel, _
                   ByVal titleLabel As Label, _
                   Optional ByVal style As FontStyle = FontStyle.Regular, _
                   Optional ByVal blnTabStop As Boolean = False)
        Me.lbl_LinkLabel = targetLinkLabel
        Me.lbl_TitleLabel = titleLabel
        Me.lbl_LinkLabel.LinkBehavior = LinkBehavior.HoverUnderline
        Me.lbl_LinkLabel.ActiveLinkColor = Color.Red
        Me.lbl_LinkLabel.LinkColor = Color.FromArgb(CByte(0), CByte(0), CByte(255))
        Me.lbl_LinkLabel.TabStop = blnTabStop   '�� ADD 2006.07.19
        Me.fontStyleLeave = style
    End Sub
#End Region


#Region "�������b�\�h����"

#Region "�L���ȃR���g���[���Ƀt�H�[�J�X���ړ�����"
    ' ------------------------------------------------------------------ 
    '�@�@�\�F���̃R���g���[���ֈڍs
    '
    ' ------------------------------------------------------------------ 
    Private Sub PRS_NextControl()
        Dim nextCtrl As Control = startCtrl
        Do
            If (TypeOf nextCtrl Is RadioButton) Then
                '���W�I�{�^���̎���,�`�F�b�N����Ă�����̂ɑ΂���,�t�H�[�J�X���Z�b�g����
                If CType(nextCtrl, RadioButton).Checked Then
                    '���x���ȊO�̃t�H�[�J�X���󂯓����R���g���[���̏ꍇ�t�H�[�J�X�ړ�
                    nextCtrl.Focus()
                    Exit Do
                End If
            ElseIf Not (TypeOf nextCtrl Is Label) And nextCtrl.Visible And nextCtrl.Enabled Then
                '���x���ȊO�̃t�H�[�J�X���󂯓����R���g���[���̏ꍇ�t�H�[�J�X�ړ�
                nextCtrl.Focus()
                Exit Do
            End If
            nextCtrl = baseForm.GetNextControl(nextCtrl, True)
        Loop Until nextCtrl Is startCtrl
    End Sub
#End Region

#End Region

#Region "�R���{�{�b�N�X�l�Z�b�g�F�o�C���h"
    '********************************************************************
    '* �@�\�@�@�@: �R���{�{�b�N�X�o�C���h����
    '* �Ԃ�l�@�@: �Ȃ�
    '* �������@�@: target           -in GrapeCity.Win.Input.Combo   �ΏۃR���{�{�b�N�X
    '*             dtTbl            -in DataTable                   �f�[�^�e�[�u��
    '*             valueMember      -in String                      value�i�[�J������
    '*             displayMember    -in String                      ���̊i�[�J������
    '*             addSpaceItem     -in Boolean                     �󔒍s�ǉ��t���O(True..�ǉ�����AFalse..�ǉ����Ȃ�)
    '* �@�\�����@:
    '* ���l�@    :
    '* �쐬  �@  : 2007/04/11 Hoshiya
    '* �X�V����  : 20071025_1 �h���b�v�_�E�����X�g ��������
    '********************************************************************
    ''' <summary>
    '''  �R���{�{�b�N�X�o�C���h����
    ''' </summary>
    ''' <param name="con"></param>
    ''' <param name="strTable"></param>
    ''' <param name="valueMember"></param>
    ''' <param name="displayMember"></param>
    ''' <param name="strWhere"></param>
    ''' <param name="szOrder"></param>
    ''' <param name="blnSpace"></param>
    ''' <param name="tran"></param>
    ''' <remarks>SkyBaseCombo���g�p���Ă��邽�߁A���ݖ��g�p</remarks>
    Public Sub BindCmbData(ByVal con As SqlConnection, ByVal strTable As String, _
                         ByVal valueMember As String, ByVal displayMember As String, _
                         Optional ByVal strWhere As String = "", Optional ByVal szOrder As String = Nothing, _
                         Optional ByVal blnSpace As Boolean = True, Optional ByVal tran As SqlTransaction = Nothing)

        blnFlgInit = False   '�Z�b�g�R���{�������t���O(False)
        Dim strAS As String = "NM" 'As��


        'SQL���쐬
        strSQL = ""
        strSQL = strSQL & " SELECT " & valueMember & "," & displayMember & " AS " & strAS
        strSQL = strSQL & " FROM " & strTable & " "
        If strWhere <> "" Then
            strSQL = strSQL & " WHERE " & strWhere
        End If

        If szOrder = "" Then
            strSQL = strSQL & " ORDER BY " & valueMember
        Else
            strSQL = strSQL & " ORDER BY " & szOrder
        End If

        '�R�}���h�쐬
        oda = New SqlDataAdapter
        dts = New DataSet

        'Modify 20060921_1
        'ocd = New OracleCommand(strSQL, con)
        ocd = New SqlCommand(strSQL, con, tran)

        oda.SelectCommand = ocd
        oda.Fill(dts, strTable)

        '** ADD 2006.07.05
        targetCombo.Items.Clear()  '��Ă���O������


        If dts.Tables(0).Rows.Count > 0 Then

            '������
            targetCombo.DataSource = Nothing
            targetCombo.Items.Clear()


            Dim dtTbl_Clone As DataTable = dts.Tables(0).Copy()

            'ValuMenb�̍ő啶�������擾
            Dim iMaxLength As Integer = 0
            For i As Integer = 0 To dtTbl_Clone.Rows.Count - 1
                If PBCStr(dtTbl_Clone.Rows.Item(i)(valueMember)).Length > iMaxLength Then
                    iMaxLength = PBCStr(dtTbl_Clone.Rows.Item(i)(valueMember)).Length
                End If
            Next

            '�z��N���X�Ɋi�[
            Dim cmbAry As New ArrayList

            '�󍀖ڂ�ǉ����邩�H
            If blnSpace Then
                cmbAry.Add(New cConf("", "", 0))
            End If

            For i As Integer = 0 To dtTbl_Clone.Rows.Count - 1
                cmbAry.Add(New cConf(PBCStr(dtTbl_Clone.Rows.Item(i)(valueMember)), _
                                                        PBCStr(dtTbl_Clone.Rows.Item(i)(strAS)), iMaxLength))
            Next

            '�f�[�^�\�[�X�ݒ�
            targetCombo.DataSource = cmbAry
            targetCombo.ValueMember = "ValueData"
            targetCombo.DisplayMember = "DisplayDataEdting"

            '�������� 20071025_1
            targetCombo.DropDownAutoSize = True
            '���擾
            targetCombo.AutoSelect = True

        End If
    End Sub
#End Region

#Region "SetStBar �X�e�C�^�X�o�[ �쐬�ҁA�X�V�҂̃Z�b�g"
    ''' <summary>
    ''' �X�e�C�^�X�o�[ �쐬�ҁA�X�V�҂̃Z�b�g
    ''' </summary>
    ''' <param name="con"></param>
    ''' <param name="szTblName"></param>
    ''' <param name="szKey"></param>
    ''' <param name="tran"></param>
    ''' <remarks></remarks>
    Friend Sub SetStBar(ByVal con As SqlConnection, _
                            ByVal szTblName As String, _
                            ByVal szKey As String, Optional ByVal tran As SqlTransaction = Nothing)
        Dim szSQL As String
        Dim stBarText As New StringBuilder

        Try

            szSQL = ""
            szSQL = szSQL & " SELECT A.IN_CODE,TO_CHAR(A.IN_DATE,'YY/MM/DD hh24:MI:SS') AS IN_DATE ,"
            szSQL = szSQL & " A.UP_CODE,TO_CHAR(A.UP_DATE,'YY/MM/DD hh24:MI:SS') AS UP_DATE"
            szSQL = szSQL & " ,B.EMP_NMEMP AS IN_NAME ,C.EMP_NMEMP AS UP_NAME "
            szSQL = szSQL & " FROM " & szTblName & " A , M_EMP B , M_EMP C "
            szSQL = szSQL & " WHERE  A.IN_CODE = B.EMP_CDEMP(+) "
            szSQL = szSQL & " AND  A.UP_CODE = C.EMP_CDEMP(+) "
            szSQL = szSQL & " AND " & szKey

            'con = skysystem.common(con)

            '�R�}���h�쐬
            oda = New SqlDataAdapter
            dts = New DataSet
            ocd = New SqlCommand(szSQL, con, tran)

            oda.SelectCommand = ocd
            oda.Fill(dts)

            If dts.Tables(0).Rows.Count > 0 Then

                With dts.Tables(0)

                    stBarText.Append("�y�V�K�쐬�z")
                    stBarText.Append(PBCStr(.Rows(0)("IN_NAME")))
                    stBarText.Append("(" & PBCStr(.Rows(0)("IN_DATE")) & ")")
                    stBarText.Append("�@�y�ŏI�X�V�z")
                    stBarText.Append(PBCStr(.Rows(0)("UP_NAME")))
                    stBarText.Append("(" & PBCStr(.Rows(0)("UP_DATE")) & ")")


                    ''�X�e�[�^�X�o�[�ɃZ�b�g
                    Me.targetStbar.Panels(0).Text = stBarText.ToString

                End With
            End If


        Catch ex As Exception
            Throw ex
        End Try

    End Sub
#End Region '20061225_1

#Region "�����C�x���g����"

#Region "EnabledChanged�C�x���g�F�{�^����ԃC�x���g"
    ' ------------------------------------------------------------------ 
    '
    '�@�@�\�F���[�h�ɂ��{�^���̏�ԕύX�̏ꍇ�A���x����ԕύX
    '
    '
    ' ------------------------------------------------------------------ 
    Private Sub EnabledChangedEvent(ByVal sender As Object, _
                                    ByVal e As System.EventArgs) Handles targetModeButton.EnabledChanged
        If targetModeButton.Enabled Then
            targetModeLabel.Enabled = True
        Else
            targetModeLabel.Enabled = False
        End If
    End Sub
#End Region

#Region "KeyDown�C�x���g�FEnterKeyHandles"
    ' ------------------------------------------------------------------ 
    '�@�@�\�F�t�H�[�J�X�ړ�
    '
    ' ------------------------------------------------------------------ 
    Private Sub targetCtrl_KeyDown(ByVal sender As Object, _
                                   ByVal e As System.Windows.Forms.KeyEventArgs) Handles targetCtrl.KeyDown
        Select Case e.KeyCode
            Case Keys.Enter, Keys.Return
                PRS_NextControl()
        End Select
    End Sub
#End Region

#Region "ValueChanged�C�x���g�F�R���{�`�F���W�C�x���g"
    ' DEL 20080331_1 Leave�C�x���g�ֈڍs
    '------------------------------------------------------------------ 
    '�@�@�\�F�R���{�e�L�X�g�Z�b�g�C�x���g(Content & �b & Desciption)
    '
    '�@���l�F
    ' ------------------------------------------------------------------ 
    Private Sub targetCombo_ValueChanged(ByVal sender As Object, _
                                         ByVal e As System.EventArgs) Handles targetCombo.ValueChanged

        'Dim intLength As Integer

        '�R���{�{�b�N�X�l�Z�b�g���A�C�x���g�������Ȃ��悤�ɁB
        If blnFlgInit = False Then Exit Sub

        Dim strContent, strDesciption As String

        Try
            If targetCombo.Value <> "" Then

                strContent = GetCmbContent(targetCombo)
                If strContent = "" Then Exit Sub

                For i As Integer = 0 To targetCombo.Items.Count - 1

                    'Comment ADD 2006.06.27
                    'If Not IsNumeric(strContent) Then Exit Sub
                    'Comment END

                    'Comment ADD 2006.07.13
                    'If strContent = Trim(CStr(targetCombo.Items.Item(i).Content)) Then
                    '    strDesciption = PBCstr(targetCombo.Items.Item(i).Description)
                    '    targetCombo.Value = CStr(targetCombo.Items.Item(i).Content) + PBCSTR_VERTICAL + strDesciption
                    '    RaiseEvent ChagedComboValue(sender, strContent) '�� ADD 2006.07.10 
                    '    Exit For    '�� ADD 2006.07.05 StackOverflowException �����h�~
                    'End If

                    If strContent.ToUpper = CStr(targetCombo.Items.Item(i).Content) OrElse _
                        strContent.ToLower = CStr(targetCombo.Items.Item(i).Content) Then

                        blnFlgInit = False  '�� ADD 2006.07.28(����ē�񔭐��h�~)

                        strDesciption = PBCStr(targetCombo.Items.Item(i).Description)
                        targetCombo.Value = CStr(targetCombo.Items.Item(i).Content) & PBCSTR_VERTICAL & strDesciption

                        blnFlgInit = True  '�� ADD 2006.07.28

                        RaiseEvent ChangedComboValue(targetCombo, strContent) '�� ADD 2006.07.10 
                        Exit For    '�� ADD 2006.07.05 StackOverflowException �����h�~
                    End If
                Next

                'For i As Integer = 0 To targetCombo.Items.Count - 1
                '    If Not IsNumeric(strContent) Then Exit Sub
                '    'If Rtn_Int(Left(targetCombo.Value, inDegit)) = _
                '    '   Rtn_Int(targetCombo.Items.Item(i).Content) Then
                '    'If Left(targetCombo.Value, inDegit) = CStr(targetCombo.Items.Item(i).Content) Then
                '    If strContent = Trim(CStr(targetCombo.Items.Item(i).Content)) Then
                '        'targetCombo.Format.Pattern = String.Empty
                '        strDesciption = PBCstr(targetCombo.Items.Item(i).Description)
                '        targetCombo.Value = CStr(targetCombo.Items.Item(i).Content) + PBCSTR_VERTICAL + strDesciption
                '    End If
                'Next
            Else
                RaiseEvent ChangedComboValue(targetCombo, "") '�� ADD 2006.07.14 
                Exit Sub
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
#End Region

#Region "ValueChanged�C�x���g�FNumber���۰ق�NegativeValue��skylog"
    Private Sub targetIMNumber_ValueChanged(ByVal sender As Object, _
                                            ByVal e As System.EventArgs) Handles targetIMNumber.ValueChanged
        '*** NegativeValueskylog�̏ꍇ�̂ݏ������s�킹��
        If bytHighNega = 1 Then
            If TypeOf targetIMNumber.Value Is Decimal Then
                If CDec(targetIMNumber.Value) < 0 Then
                    targetIMNumber.DisabledForeColor = targetIMNumber.NegativeColor
                Else
                    targetIMNumber.DisabledForeColor = System.Drawing.SystemColors.WindowText
                End If
            End If
        End If
    End Sub
#End Region

#Region "GotFocus�C�x���g�F�I�����ꂽ���(HighlightText���)"
    Private Sub targetIMNumber_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) _
                                        Handles targetIMNumber.GotFocus

        '*** HighLight�s��̏ꍇ�̂ݏ������s�킹��
        If bytHighNega = 0 Then
            '*** �������I�����ꂽ���
            Dim iLength As Integer = targetIMNumber.Text.IndexOf(targetIMNumber.Text)
            If iLength > -1 Then
                targetIMNumber.SelectionStart = iLength
                targetIMNumber.SelectionLength = targetIMNumber.Text.Length
            End If
        End If
    End Sub
    Private Sub targetIMEdit_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) _
                                  Handles targetIMEdit.GotFocus

        '*** �������I�����ꂽ���
        Dim iLength As Integer = targetIMEdit.Text.IndexOf(targetIMEdit.Text)
        If iLength > -1 Then
            targetIMEdit.SelectionStart = iLength
            targetIMEdit.SelectionLength = targetIMEdit.Text.Length
        End If
    End Sub
#End Region

#Region "LinkLabel�C�x���g"
#Region "MouseEnter(LinkLabel)�C�x���g�FLinkLabel�̃t�H�g�ύX"
    '�}�E�X���e�L�X�g��ɂ���ꍇ�̏���
    Private Sub LinkLabel_MouseEnter(ByVal sender As Object, _
                                     ByVal e As System.EventArgs) Handles lbl_LinkLabel.MouseEnter

        Dim link As LinkLabel = CType(sender, LinkLabel)
        link.Font = New Font(link.Font, FontStyle.Bold)     '�t�H���g�𑾎��ɂ��� 
    End Sub
#End Region
#Region "MouseLeave(LinkLabel)�C�x���g�FLinkLabel�̃t�H�g�ύX"
    '�}�E�X���e�L�X�g�ォ�痣�ꂽ�ꍇ�̏��� 
    Private Sub LinkLabel_MouseLeave(ByVal sender As Object, _
                                     ByVal e As System.EventArgs) Handles lbl_LinkLabel.MouseLeave

        Dim link As LinkLabel = CType(sender, LinkLabel)
        link.Font = New Font(link.Font, fontStyleLeave)     '�t�H���g���(���͌n�E�}�X�^�n)
    End Sub
#End Region
#Region "EnabledChanged(LinkLabel)�C�x���g�F��pLabel�\�����\��"
    Private Sub LinkLabel_EnabledChanged(ByVal sender As Object, _
                                         ByVal e As System.EventArgs) Handles lbl_LinkLabel.EnabledChanged
        If lbl_LinkLabel.Enabled Then
            lbl_TitleLabel.Visible = False
        Else
            lbl_TitleLabel.Visible = True
        End If
    End Sub
#End Region
#End Region

#Region "MouseHover,Leave�����"
#Region "MouseHover�C�x���g�F���ʏ���"
    'Friend Sub MouseHoverEvent(ByVal sender As System.Object, ByVal e As System.EventArgs) _
    '                        Handles targetStbar.MouseHover
    '    If sender Is targetStbar Then
    '        Dim asm As System.Reflection.Assembly = System.Reflection.Assembly.GetExecutingAssembly()
    '        Dim ver As System.Version = asm.GetName().Version  '�o�[�W�����̎擾
    '        Me.targetStbar.Panels(3).Text = ver.ToString
    '    End If

    'End Sub
#End Region

#Region "MouseLeave�C�x���g�F���ʏ���"
    'Friend Sub MouseLeaveEvent(ByVal sender As System.Object, ByVal e As System.EventArgs) _
    '                        Handles targetStbar.MouseLeave
    '    'If sender Is targetStbar Then
    '    '    Me.targetStbar.Panels(3).Text = stu_EMP.PGMID
    '    'End If

    'End Sub
#End Region

#End Region '20070328_1

#End Region

#Region "�ۗ�"

    ''DEL 20060807_1
    '#Region "�R���{�{�b�N�X�l�Z�b�g"
    '    ' ------------------------------------------------------------------------------ 
    '    '�@�@�\�F�R���{��DataSource�l�Z�b�g
    '    '�@      (�X�y�[�X�v���O�ɂ��A�ŏ��s�ɃX�y�[�X�s�}���E�����ɂ���)
    '    '  �����F�R�l�N�V����(con)
    '    '        Table��(strTable), �L�[(strKey)
    '    '        �ڍ�(strDescription)
    '    '        Optional�FWhere����(strWhere)�EDropDownStyle(bytDropDownStyle)
    '    '                  ALias(strAS, ��ȏ��Field����������ꍇ)
    '    '                  ImeMode(ime, �������͉\�̏ꍇ)
    '    '                  MaxDropDownItems(�ő�List�𒴂����ꍇ�A�װ����)
    '    ' ------------------------------------------------------------------------------  
    '    Public Overloads Sub PBS_SetComboDS(ByVal con As SqlConnection, _
    '                              ByVal strTable As String, ByVal strKey As String, _
    '                              ByVal strDescription As String, _
    '                              Optional ByVal strWhere As String = "", _
    '                              Optional ByVal bytDropDownStyle As ComboBoxStyle = ComboBoxStyle.DropDown, _
    '                              Optional ByVal strAS As String = "", _
    '                              Optional ByVal ime As ImeMode = ImeMode.Disable, _
    '                              Optional ByVal iMaxDownItems As Integer = 10)

    '        Try
    '            blnFlgInit = False   '�Z�b�g�R���{�������t���O(False)

    '            'SQL���쐬
    '            strSQL = ""
    '            strSQL = strSQL & " SELECT " & strKey & "," & strDescription
    '            strSQL = strSQL & " FROM " & strTable & " "
    '            If strWhere <> "" Then
    '                strSQL = strSQL & " WHERE " & strWhere
    '            End If
    '            strSQL = strSQL & " ORDER BY " & strKey

    '            con = PBFCON_ChkConnection(con)

    '            '�R�}���h�쐬
    '            oda = New sqlDataReader
    '            dts = New DataSet
    '            ocd = New sqlCommand(strSQL, con)
    '            oda.SelectCommand = ocd
    '            oda.Fill(dts, strTable)

    '            '** 20060630
    '            targetCombo.Items.Clear()  '��Ă���O������

    '            '** ADD 2006.07.10
    '            If strAS <> "" Then strDescription = strAS

    '            If dts.Tables(0).Rows.Count > 0 Then
    '                With targetCombo

    '                    If blnFlgSpace Then
    '                        '��DataSource�ɃX�y�[�X�s�ǉ��̏ꍇ

    '                        .BeginUpdate()
    '                        For i As Integer = 0 To dts.Tables(0).Rows.Count
    '                            If i = 0 Then
    '                                .Items.AddRange(New IM.ComboItem() _
    '                                               {New IM.ComboItem(0, Nothing, "", "", "")})
    '                            Else
    '                                .Items.AddRange(New IM.ComboItem() _
    '                                               {New IM.ComboItem(0, Nothing, _
    '                                                    PBCstr(dts.Tables(0).Rows(i - 1)(strKey)), _
    '                                                    PBCstr(dts.Tables(0).Rows(i - 1)(strDescription)), _
    '                                                    PBCstr(dts.Tables(0).Rows(i - 1)(strKey)))})
    '                            End If
    '                        Next
    '                        .EndUpdate()
    '                        '.DisplayMember = strKey
    '                        '.ValueMember = strKey
    '                        '.DescriptionMember = strDescription
    '                    Else
    '                        '��DataSource�ɃX�y�[�X�s�ǉ����Ȃ��̏ꍇ

    '                        '�p�����[�^�[�ݒ�
    '                        .DataSource = dts.Tables(0)

    '                        '�������Ƃ��ĕ\������f�[�^�\�[�X�̃v���p�e�B�������������ݒ肵�܂��B
    '                        .DisplayMember = strKey

    '                        '�l�Ƃ��Ĉ����f�[�^�\�[�X�̃v���p�e�B�������������ݒ肵�܂��B
    '                        .ValueMember = strKey

    '                        '�������Ƃ��ĕ\������f�[�^�\�[�X�̃v���p�e�B�������������ݒ�
    '                        .DescriptionMember = strDescription
    '                    End If


    '                    '##<< ���ʐݒ� >>##
    '                    .AutoSelect = True
    '                    .HighlightText = IM.HighlightText.All
    '                    '.ImeMode = ImeMode.Disable
    '                    .ImeMode = ime   '�� Modify 2006.07.19
    '                    .ListBoxStyle = IM.ListBoxStyle.TextWithDescription
    '                    .TextBoxStyle = IM.TextBoxStyle.TextOnly
    '                    .TextHAlign = IM.AlignHorizontal.Left
    '                    .TextVAlign = IM.AlignVertical.Middle
    '                    .DropDownWidth = .Width
    '                    .ImageWidth = 0
    '                    .DropDownStyle = bytDropDownStyle               '�� ADD 2006.07.10
    '                    .ShowScrollBar = True                           '�� ADD 2006.07.28
    '                    .ScrollBarMode = IM.ScrollBarMode.Automatic     '�� ADD 2006.07.28

    '                    Dim maxDigit As Integer 'Content�̍ő包���擾
    '                    For i As Integer = 0 To .Items.Count - 1
    '                        If maxDigit < CStr(.Items.Item(i).Content).Length Then
    '                            maxDigit = CStr(.Items.Item(i).Content).Length
    '                        End If
    '                    Next

    '                    '������ADD 2006.07.26
    '                    '�h���b�v�_�E�������ɕ\������鍀�ڂ̍ő吔���擾�܂��͐ݒ�
    '                    '.MaxDropDownItems = .Items.Count + 1
    '                    If .Items.Count > iMaxDownItems Then
    '                        .MaxDropDownItems = iMaxDownItems + 1

    '                        '����Modify 2006.07.28 
    '                        'Content�̕������ŁA���������Ȃ��Ȃ�s�����skylog
    '                        If .DropDownWidth > (maxDigit * 10) Then

    '                            'Content((�ő包�� + 1) *12)  �b Description�T�C�Y���
    '                            .DescriptionWidth = .DropDownWidth - ((maxDigit + 1) * 12)
    '                        Else
    '                            .DescriptionWidth = .DropDownWidth - CInt(.DropDownWidth / 4)
    '                        End If
    '                    Else
    '                        .MaxDropDownItems = .Items.Count + 1

    '                        If .DropDownWidth > (maxDigit * 10) Then

    '                            'Content((�ő包�� + 1) *10)  �b Description�T�C�Y���
    '                            .DescriptionWidth = .DropDownWidth - ((maxDigit + 1) * 10)
    '                        Else
    '                            .DescriptionWidth = .DropDownWidth - CInt(.DropDownWidth / 4)
    '                        End If
    '                    End If

    '                    'Comment 2006.07.28 (���ړ�)
    '                    ''Content((�ő包�� + 1) *10)  �b Description�T�C�Y���
    '                    'If .DropDownWidth > (maxDigit * 10) Then
    '                    '    .DescriptionWidth = .DropDownWidth - ((maxDigit + 1) * 10)
    '                    'Else
    '                    '    .DescriptionWidth = .DropDownWidth - CInt(.DropDownWidth / 4)
    '                    'End If

    '                End With
    '            End If

    '            '����������������(�Z�b�g�R���{�������t���O=True)
    '            blnFlgInit = True

    '        Catch ex As Exception
    '            SkyLog.Error(ex.Message, ex)
    '        End Try
    '    End Sub
    ''#End Region
    '#Region "KeyPress�C�x���g�F�����ABackSpace�ȊO�̂��͓̂��͕s��"
    '    Private Sub targetCombo_KeyPress(ByVal sender As Object, _
    '                                     ByVal e As System.Windows.Forms.KeyPressEventArgs) _
    '                                     Handles targetCombo.KeyPress
    '        If ControlFlg <> 2 Then Exit Sub
    '        '�����ABackSpace�ȊO�̂��͓̂��͕s�ɂ���
    '        If (e.KeyChar < "0"c Or e.KeyChar > "9"c) And e.KeyChar <> vbBack Then
    '            e.Handled = True
    '            Exit Sub
    '        End If
    '        targetCombo.Format.Pattern = strPattern
    '    End Sub
    '#End Region
#End Region
End Class

Class cConf
    '*********************************************************************************************
    '*�@�@�\�@�@�F�z��i�[�N���X
    '*            
    '*�@
    '*********************************************************************************************
    Private DisMeb As String
    Private DescMeb As String
    Private ValMeb As String
    Private pMaxLength As Integer

#Region "�R���X�g���N�^"
    Sub New(ByVal prmDisMeb As String, ByVal prmDescMeb As String, ByVal prmValMeb As String)
        MyBase.New()
        Me.DisMeb = prmDisMeb
        Me.DescMeb = prmDescMeb
        Me.ValMeb = prmValMeb
    End Sub
    Sub New(ByVal prmValMeb As String, ByVal prmDisMeb As String, ByVal prmMaxLengeth As Integer)
        MyBase.New()
        Me.DisMeb = prmDisMeb
        Me.ValMeb = prmValMeb
        Me.pMaxLength = prmMaxLengeth
    End Sub
#End Region

#Region "�v���p�e�B"
#Region "DisplayMemberEditing "
    ReadOnly Property DisplayDataEdting() As String
        Get
            If ValMeb.Equals("") Then
                Return ""
            Else
                Return ValMeb.PadRight(pMaxLength) & "�b" & DisMeb
            End If
        End Get
    End Property
#End Region
#Region "DisplayMember "
    ReadOnly Property DisplayData() As String
        Get
            Return DisMeb
        End Get
    End Property
#End Region
#Region "DescriptionMember "
    ReadOnly Property DescriptionData() As String
        Get
            Return DescMeb
        End Get
    End Property
#End Region
#Region "ValueMember"
    ReadOnly Property ValueData() As String
        Get
            Return ValMeb
        End Get
    End Property
#End Region
#Region "ImageMenber"
    'Public ReadOnly Property ImageData() As Object
    '    Get
    '        Return aImageData
    '    End Get
    'End Property
#End Region
#End Region

End Class
