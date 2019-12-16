
Option Explicit On
Option Strict On

Imports skysystem.common.SystemUtil
Imports skysystem.common.MessageUtil
Imports skysystem.common.SystemConst
Imports System.Data.OleDb
Imports Devart.Data.Universal


'****************************************************************************************
'*�@�@�@�\�@�F����MSG�N���X(DB��A��)
'*�@�쐬���@�F2007/07/08
'*
'****************************************************************************************
''' <summary>
'''  ���ʃ��b�Z�[�W���C�u����
''' </summary>
''' <remarks></remarks>
Public Class MessageUtil

#Region "Private�萔(�Œ�MSG)"
    Private Const PRCSTR_MISSING_MSG As String = "���b�Z�[�W�擾���s"

    '*** �₢���킹MSG
    '* �V�K�F�o�^���܂��B��낵���ł����H
    Private Const PRCSTR_MSGCTG_INSERT As Integer = 2
    Private Const PRCSTR_MSGID_INSERT As Integer = 0

    '* �X�V�F�X�V���܂��B��낵���ł����H
    Private Const PRCSTR_MSGCTG_UPDATE As Integer = 2
    Private Const PRCSTR_MSGID_UPDATE As Integer = 1

    '* �폜�F�\������Ă���f�[�^���폜���܂��B��낵���ł����H
    Private Const PRCSTR_MSGCTG_DELETE As Integer = 2
    Private Const PRCSTR_MSGID_DELETE As Integer = 2

    '* �I���F�I�����Ă�낵���ł����H
    Private Const PRCSTR_MSGCTG_EXIT As Integer = 1
    Private Const PRCSTR_MSGID_EXIT As Integer = 9
#End Region


#Region "Public�萔(�Œ�MSG)"

    '*** ����MSG(�o�^�E�C��)
    '* �V�K�E����
    '��) �˗���CD G0000 & '�œo�^����܂����B'
    Public Const PBC_MSGCTG_REGISTED As Integer = 10
    Public Const PBC_MSGID_REGISTED As Integer = 0

    '* �C��
    '��) �˗���CD G0001 & '���C������܂����B'
    Public Const PBCSTR_MSGCTG_UPDATED As Integer = 10
    Public Const PBCSTR_MSGID_UPDATED As Integer = 1

    '**************************
    '*** �Ώۃf�[�^�`�F�b�N ***
    '**************************
    '�ΏۂƂȂ�f�[�^������܂���B
    Public Const PBCSTR_MSGCTG_NODATA As Integer = 0
    Public Const PBCSTR_MSGID_NODATA As Integer = 0

    '�w�肳�ꂽ�R�[�h�͓o�^����Ă��܂���B & vbCrLf & �ʂ̃R�[�h�Ō������ĉ������B
    Public Const PBCSTR_MSGCTG_NOT_REGISTED As Integer = 0
    Public Const PBCSTR_MSGID_NOT_REGISTED As Integer = 1

    '���̍��ڂ̓��X�g���ɂ��鍀�ڂ���I�����ĉ������B
    Public Const PBCSTR_MSGCTG_NO_LIST As Integer = 0
    Public Const PBCSTR_MSGID_NO_LIST As Integer = 2

    '*�o�^�ς݁F���ɓo�^����Ă��܂��B
    Public Const PBCSTR_MSGCTG_ERR_REGISTED As Integer = 0
    Public Const PBCSTR_MSGID_ERR_REGISTED As Integer = 4

    '(�d���f�[�^)����ԍ������ɓo�^����Ă��܂��B
    Public Const PBCSTR_MSGCTG_KEY_CONFLICT As Integer = 0
    Public Const PBCSTR_MSGID_KEY_CONFLICT As Integer = 10

    '(�K�{�`�F�b�N)���̍��ڂ͕K�{���͂ł��B & vbCrLf & �������l����͂��ĉ������B
    Public Const PBC_MSGCTG_MUST_INPUT As Integer = 0
    Public Const PBC_MSGID_MUST_INPUT As Integer = 11

    '(�͈͊O)�w��͈͊O�ł��B������x���͂��Ă��������B
    Public Const PBCSTR_MSGCTG_OVERFLOW As Integer = 0
    Public Const PBCSTR_MSGID_OVERFLOW As Integer = 12

    'ADD 2006.08.30
    '�Y���̺��ނ����݂��܂���B
    Public Const PBCSTR_MSGCTG_NOCODE As Integer = 0
    Public Const PBCSTR_MSGID_NOCODE As Integer = 13

    'ADD 2006.07.06
    '����t�̑召���قȂ�܂��B�
    Public Const PBCSTR_MSGCTG_DIFFER_DAY_SIZE As Integer = 0
    Public Const PBCSTR_MSGID_DIFFER_DAY_SIZE As Integer = 9

    'ADD 2006.08.08
    '������������Ă��܂��B�
    Public Const PBCINT_MSGCTG_OVER_DIGIT As Integer = 0
    Public Const PBCINT_MSGID_OVER_DIGIT As Integer = 14

    'ADD 20070525_1
    '������ꂩ�̃f�[�^�Ƀ`�F�b�N���s���Ă��������B�
    Public Const PBCINT_MSGCTG_NOCHECK As Integer = 0
    Public Const PBCINT_MSGID_NOCHECK As Integer = 16

    'ADD 20080714_1
    '��召���قȂ�܂��B�
    Public Const PBCINT_MSGCTG_DIFFER_SIZE As Integer = 0
    Public Const PBCSTR_MSGID_DIFFER_SIZE As Integer = 15

    'ADD 20080714_1
    '��o�͂��܂��B��낵���ł����H�
    Public Const PBCINT_MSGCTG_OUT_ACTION As Integer = 2
    Public Const PBCSTR_MSGID_OUT_ACTION As Integer = 3

    '**************************
    '*** ���گ�޼�Ċ֘A MSG ***
    '**************************

    '(���ד�������)����������͂��Ă��������B
    Public Const PBCSTR_MSGCTG_CELLNULL As Integer = 10
    Public Const PBCSTR_MSGID_CELLNULL As Integer = 11

    '(���ד�������)0�ȏ�̐��l����͂��Ă��������B
    Public Const PBCSTR_MSGCTG_CELLZERO As Integer = 0
    Public Const PBCSTR_MSGID_CELLZERO As Integer = 5

    '(���ז�����)���ׂ����͂���Ă��܂���B
    Public Const PBCSTR_MSGCTG_SPREADNULL As Integer = 6
    Public Const PBCSTR_MSGID_SPREADNULL As Integer = 0

    '(���ז�����)���ׂ��`�F�b�N����Ă��܂���B
    Public Const PBCSTR_MSGCTG_SPREADNOTCHECK As Integer = 6
    Public Const PBCSTR_MSGID_SPREADNOTCHECK As Integer = 1

    'ADD 2006/06/26
    '(���ד���)�u�������v���d�����Ă��܂��B
    Public Const PBCSTR_MSGCTG_SPREAD_CONFLICT As Integer = 6
    Public Const PBCSTR_MSGID_SPREAD_CONFLICT As Integer = 2


    '*** ������o�X����̋��ʃ��W���[���Ƃ��ĕ���(\common\Mei3\Mei3PB.vb�Q��)
    'Public Const PBCSTR_TITLE_MSG As String = "�Ɩ��x���V�X�e��"

    Public Const PBC_MSG_INIT_ERR As String = "�����ݒ�Ɏ��s���܂����B"
    Public Const PBC_MSG_NO_CONNECTION As String = "�R�l�N�V�����ݒ肪�擾�ł��܂���B"
    Public Const PBC_MSG_ERROR_DB As String = "�G���[���������܂����B������������܂��B"
    Public Const PBC_MSG_ALREADY_STARTED As String = "�v���O�����͊��ɋN������Ă��܂��B"
    'ADD 2006.08.01
    Public Const PBC_MSG_ERROR_STOP As String = "�G���[���������܂����B�����𒆎~���܂��B"
    Public Const PBC_MSG_STOP As String = "�����𒆎~���܂��B"


    Public Const PBC_MSG_START As String = " �X�^�[�g"
    Public Const PBC_MSG_END As String = " �I��"

    'ADD 2006.07.05
    '�****���t����͂��Ă��������B�
    Public Const PBCSTR_MSGCTG_INPUT_DAY As Integer = 10
    Public Const PBCSTR_MSGID_INPUT_DAY As Integer = 13


    'ADD 2006.07.06
    '**************************************
    '*** ��PGM������ ����MSG ��1 �`   ***
    '**************************************
    Public Const PBCINT_NO0 As Integer = 0              'ADD 2006.07.31 
    Public Const PBCINT_NO1 As Integer = 1
    Public Const PBCINT_NO2 As Integer = 2
    Public Const PBCINT_NO3 As Integer = 3
    Public Const PBCINT_NO4 As Integer = 4
    Public Const PBCINT_NO5 As Integer = 5
    Public Const PBCINT_NO6 As Integer = 6
    Public Const PBCINT_NO7 As Integer = 7
    Public Const PBCINT_NO8 As Integer = 8
    Public Const PBCINT_NO9 As Integer = 9
    Public Const PBCINT_NO10 As Integer = 10        'ADD 2006.09.08 

    'ADD 2006.07.25 
    Public Const PBCSTR_MSGCTG_XLS As String = "XLS"
    Public Const PBCINT_MSGID_XLS1 As Integer = 1      'EXCEL�o�͂Ɏ��s���܂����B
    Public Const PBCINT_MSGID_XLS2 As Integer = 2      '���m�F�ł��܂���B�V���ɍ쐬���Ă���낵���ł����H
    Public Const PBCINT_MSGID_XLS3 As Integer = 3      'EXCEL���`�̎擾�Ɏ��s���܂����B
    Public Const PBCINT_MSGID_XLS4 As Integer = 4      'EXCEL�����Ɏ��s���܂����B
    Public Const PBCINT_MSGID_XLS5 As Integer = 5      'EXCEL�ۑ��Ɏ��s���܂����B
    Public Const PBCINT_MSGID_XLS6 As Integer = 6      '�o�͐�̃p�X���m�F�ł��܂���B
    Public Const PBCINT_MSGID_XLS7 As Integer = 7      '���Ƀt�@�C�������݂��܂��B�㏑�����܂����H
    Public Const PBCINT_MSGID_XLS8 As Integer = 8      '�ǂݎ���p�t�@�C���ł��B�������݂ł��܂���B

    'ADD 2006.07.26
    Public Const PBCSTR_MSGCTG_FLD As String = "FLD"
    Public Const PBCINT_MSGID_FLD0 As Integer = 0       '�Y������t�H���_�����݂��܂���B       
    Public Const PBCINT_MSGID_FLD1 As Integer = 1       '�Y������t�H���_�����݂��܂���B�쐬���܂����H
    Public Const PBCINT_MSGID_FLD2 As Integer = 2       '�t�H���_���폜���܂��B��낵���ł����H

    'ADD 2006.08.30
    Public Const PBCSTR_MSGCTG_NUM As String = "NUM"
    'ID_NO0     '�̔ԃ}�X�^�̎w��͈͊O�ł��B
    'ID_NO1     '
    'ID_NO2     '
    'ID_NO3     '

    'ADD 2006.07.31 
    Public Const PBCSTR_MSG_NOWAIT As String = "�����[�U�[���ɂ��f�[�^���g�p����Ă��܂��B"

    'ADD 20061011_1
    Public Const PBCSTR_MSG_ERROR_DB_UNIQUE As String = "���ɓo�^����Ă��܂��B������x�A���s���Ă��������B"

    'ADD 2006.08.01
    '***************************
    '*** ORACLE ERROR CODE   ***
    '***************************
    Public Const PBCINT_ORAERR_CODE54 As Integer = 54   '�uORA-0054�F���\�[�X �r�W�[NoWait�v

    'ADD 20061011_1
    Public Const PBCINT_ORAERR_CODE1 As Integer = 1     '�uORA-00001: ��Ӑ���(MEI3.M_TOK_IDX1)�ɔ����Ă��܂��v

#End Region

#Region "MSG Eunm�萔"

    '*** �{�^���A�C�R��
    Public Enum MIcon
        Info = 64 : [Error] = 16 : Warning = 48 : Question = 32
    End Enum

    '*** �I���{�^����
    Public Enum MButton
        OK = 0 : YesNo = 4
        AbortRetryIgnore = 2    '�\��
        OKCancel = 1            '�\��    
        RetryCancel = 5         '�\��
        YesNoCancel = 3         '�\��
    End Enum

    '*** �{�^����SF�ʒu
    Public Enum MPosition
        Button1 = 0 : Button2 = 256
        Button3 = 512   '�\��
    End Enum

    '*** �I������
    Public Enum MResult
        OK = 1 : Yes = 6 : No = 7
        Abort = 3   '�\��
        Cancel = 2  '�\��
        Ignore = 5  '�\��
        Retry = 4   '�\��
        None = 0    '�\��
    End Enum
#End Region

#Region "�R�l�N�V��������̃��b�Z�[�W�\�����\�b�h"

#Region "MSG�\����"
    Public Structure STU_MSG

        Private objMsgCTG As Object       '(�e)�J�e�S���[(���ʕ���, PG_ID)
        Private intMsgID As Integer       '(�q)MSG_ID
        Private strMsgTitle As String     'MSG�^�C�g��
        Private strMsgText As String      'MSG�e�L�X�g
        Private intMsgIcon As Integer     '�A�C�R�����
        Private intMsgPtn As Integer      '�{�^���̃p�^��(�{�^���̐�)
        Private intMsgDef As Integer      '�{�^���̈ʒu
        Private intMsgTan As Integer      '�\���F�X�V�S����
        Private strMsgDate As String      '�\���F�X�V���t

        Public Property MSG_CTG() As Object
            Get
                Return objMsgCTG
            End Get
            Set(ByVal Value As Object)
                objMsgCTG = Value
            End Set
        End Property

        Public Property MSG_ID() As Integer
            Get
                Return intMsgID
            End Get
            Set(ByVal Value As Integer)
                intMsgID = Value
            End Set
        End Property

        Public Property MSG_TITLE() As String
            Get
                Return strMsgTitle
            End Get
            Set(ByVal Value As String)
                If PB_ChkNUll(Value) Then
                    strMsgTitle = PBCSTR_TITLE_MSG
                Else
                    strMsgTitle = Value
                End If
            End Set
        End Property

        Public Property MSG_TEXT() As String
            Get
                Return strMsgText
            End Get
            Set(ByVal Value As String)
                strMsgText = Value
            End Set
        End Property

        Public Property MSG_ICON() As Integer
            Get
                Return intMsgIcon
            End Get
            Set(ByVal Value As Integer)
                Select Case Value
                    Case 1  'Info
                        intMsgIcon = MIcon.Info

                    Case 2  'Error
                        intMsgIcon = MIcon.Error

                    Case 3  'Warning
                        intMsgIcon = MIcon.Warning

                    Case 4  'Question
                        intMsgIcon = MIcon.Question

                    Case Else
                        intMsgIcon = MIcon.Error
                End Select
            End Set
        End Property

        Public Property MSG_PTN() As Integer
            Get
                Return intMsgPtn
            End Get
            Set(ByVal Value As Integer)
                Select Case Value

                    Case 1  'OK(�P��{�^��)
                        intMsgPtn = MButton.OK

                    Case 2  'YesNo(�I���{�^��)
                        intMsgPtn = MButton.YesNo

                    Case Else
                        intMsgPtn = MButton.OK

                End Select
            End Set
        End Property

        Public Property MSG_DEF() As Integer
            Get
                Return intMsgDef
            End Get
            Set(ByVal Value As Integer)

                Select Case Value

                    Case 1  '��Ԗ�
                        intMsgDef = MPosition.Button1

                    Case 2  '��Ԗ�
                        intMsgDef = MPosition.Button2

                    Case Else
                        intMsgDef = MPosition.Button1

                End Select
            End Set
        End Property

        Public Property MSG_TAN() As Integer
            Get
                Return intMsgTan
            End Get
            Set(ByVal Value As Integer)
                intMsgTan = Value
            End Set
        End Property

        Public Property MSG_DATE() As String
            Get
                Return strMsgDate
            End Get
            Set(ByVal Value As String)
                strMsgDate = Value
            End Set
        End Property
    End Structure
#End Region

#Region "���E�x��MSG(�P��I��)"
    ''' <summary>
    ''' �P�ꃁ�b�Z�[�W�\��(�R�l�N�V��������)
    ''' </summary>
    ''' <param name="con">�R�l�N�V����</param>
    ''' <param name="objCTG">���b�Z�[�W�J�e�S��</param>
    ''' <param name="intID">���b�Z�[�WID</param>
    ''' <param name="strText">���b�Z�[�W���e(�ړ�)</param>
    ''' <param name="strTitle">���b�Z�[�W�^�C�g��</param>
    ''' <remarks></remarks>
    Private Shared Sub DB_ShowMsg(ByVal con As UniConnection, _
                            ByVal objCTG As Object, ByVal intID As Integer, _
                            Optional ByVal strText As String = "", _
                            Optional ByVal strTitle As String = "")

        Dim stuMSG As New STU_MSG
        stuMSG.MSG_CTG = objCTG         '�J�e�S��
        stuMSG.MSG_ID = intID           'ID
        stuMSG.MSG_TITLE = strTitle     'ү��������

        If DB_GetMSG(con, stuMSG, strText) Then
            MsgBoxPB.Show(stuMSG.MSG_TEXT, stuMSG.MSG_TITLE, stuMSG.MSG_ICON, stuMSG.MSG_PTN)
        Else
            'Modify 2006.08.30
            'PRS_MissingMSG()    'MSG�擾���s��
            PRS_MissingMSG(objCTG, intID)    'MSG�擾���s��
        End If
    End Sub
    ''' <summary>
    ''' ���b�Z�[�W�\�����\�b�h(�R�l�N�V��������)
    ''' </summary>
    ''' <param name="con">�R�l�N�V����</param>
    ''' <param name="objCTG">���b�Z�[�W�J�e�S��</param>
    ''' <param name="intID">���b�Z�[�WID</param>
    ''' <remarks></remarks>
    Public Overloads Shared Sub ShowMsgWithCon(ByVal con As UniConnection, _
                                     ByVal objCTG As Object, ByVal intID As Integer)
        Call DB_ShowMsg(con, objCTG, intID)
    End Sub
    ''' <summary>
    ''' ���b�Z�[�W�\�����\�b�h(�R�l�N�V��������)
    ''' </summary>
    ''' <param name="con">�R�l�N�V����</param>
    ''' <param name="objCTG">���b�Z�[�W�J�e�S��</param>
    ''' <param name="intID">���b�Z�[�WID</param>
    ''' <param name="strText">���b�Z�[�W���e(�O����)</param>
    ''' <remarks>���b�Z�[�W�^�C�g���̓V�X�e�����𗘗p</remarks>
    Public Overloads Shared Sub ShowMsgWithCon(ByVal con As UniConnection, _
                                     ByVal objCTG As Object, ByVal intID As Integer, _
                                     ByVal strText As String)
        Call DB_ShowMsg(con, objCTG, intID, strText)
    End Sub
    ''' <summary>
    ''' ���b�Z�[�W�\�����\�b�h(�R�l�N�V��������)
    ''' </summary>
    ''' <param name="con">�R�l�N�V����</param>
    ''' <param name="objCTG">���b�Z�[�W�J�e�S��</param>
    ''' <param name="intID">���b�Z�[�WID</param>
    ''' <param name="strText">���b�Z�[�W���e(�O����)</param>
    ''' <param name="strTitle">���b�Z�[�W�^�C�g��</param>
    ''' <remarks></remarks>
    Public Overloads Shared Sub ShowMsgWithCon(ByVal con As UniConnection, _
                                     ByVal objCTG As Object, ByVal intID As Integer, _
                                     ByVal strText As String, ByVal strTitle As String)
        Call DB_ShowMsg(con, objCTG, intID, strText, strTitle)
    End Sub
#End Region

#Region "�⍇��MSG(�I��)"
    '----------------------------------------------------------------------
    '�@�\�@�F�⍇��MSG(�I���FYes�ENo)
    '�����@�F
    '�߂�l�FTrue�^False
    '----------------------------------------------------------------------
    Public Shared Function PRFBLN_QUser(ByVal con As UniConnection, _
                                ByVal objCTG As Object, ByVal intID As Integer, _
                                Optional ByVal strText As String = "", _
                                Optional ByVal strTitle As String = "") As Boolean
        Dim stuMsg As New STU_MSG
        Dim dialRslt As DialogResult = DialogResult.No
        stuMsg.MSG_CTG = objCTG         '�J�e�S��
        stuMsg.MSG_ID = intID           'ID
        stuMsg.MSG_TITLE = strTitle     'ү��������

        If DB_GetMSG(con, stuMsg, strText) Then
            dialRslt = MsgBoxPB.Show(stuMsg.MSG_TEXT, stuMsg.MSG_TITLE, stuMsg.MSG_ICON, stuMsg.MSG_PTN, stuMsg.MSG_DEF)
        Else
            'Modify 2006.08.30
            'PRS_MissingMSG()    'MSG�擾���s��
            PRS_MissingMSG(objCTG, intID)    'MSG�擾���s��
        End If
        Return dialRslt = DialogResult.Yes
    End Function
    ''' <summary>
    ''' �⍇�����b�Z�[�W�\��(�R�l�N�V��������)
    ''' </summary>
    ''' <param name="con">�R�l�N�V����</param>
    ''' <param name="objCTG">���b�Z�[�W�J�e�S��</param>
    ''' <param name="intID">���b�Z�[�WID</param>
    ''' <Retuen>True�F�͂������@False�F����������</Retuen>
    ''' <remarks></remarks>
    Public Overloads Shared Function PBFBLN_QUser(ByVal con As UniConnection, _
                                         ByVal objCTG As Object, ByVal intID As Integer) As Boolean
        Return PRFBLN_QUser(con, objCTG, intID)
    End Function
    ''' <summary>
    ''' �⍇�����b�Z�[�W�\��(�R�l�N�V��������)
    ''' </summary>
    ''' <param name="con">�R�l�N�V����</param>
    ''' <param name="objCTG">���b�Z�[�W�J�e�S��</param>
    ''' <param name="intID">���b�Z�[�WID</param>
    ''' <param name="strText">���b�Z�[�W���e(�O����)</param>
    ''' <Retuen>True�F�͂������@False�F����������</Retuen>
    ''' <remarks></remarks>
    Public Overloads Function PBFBLN_QUser(ByVal con As UniConnection, _
                                           ByVal objCTG As Object, ByVal intID As Integer, _
                                           ByVal strText As String) As Boolean
        Return PRFBLN_QUser(con, objCTG, intID, strText)
    End Function
    ''' <summary>
    ''' �⍇�����b�Z�[�W�\��(�R�l�N�V��������)
    ''' </summary>
    ''' <param name="con">�R�l�N�V����</param>
    ''' <param name="objCTG">���b�Z�[�W�J�e�S��</param>
    ''' <param name="intID">���b�Z�[�WID</param>
    ''' <param name="strText">���b�Z�[�W���e(�O����)</param>
    ''' <Retuen>True�F�͂������@False�F����������</Retuen>
    ''' <remarks></remarks>
    Public Shared Function ShowUserMsgWithCon(ByVal con As UniConnection, _
                                           ByVal objCTG As Object, ByVal intID As Integer, _
                                           Optional ByVal strText As String = "") As Boolean
        Return PRFBLN_QUser(con, objCTG, intID, strText)
    End Function
    ''' <summary>
    ''' �⍇�����b�Z�[�W�\��(�R�l�N�V��������)
    ''' </summary>
    ''' <param name="con">�R�l�N�V����</param>
    ''' <param name="objCTG">���b�Z�[�W�J�e�S��</param>
    ''' <param name="intID">���b�Z�[�WID</param>
    ''' <param name="strText">���b�Z�[�W���e(�O����)</param>
    ''' <param name="strTitle">���b�Z�[�W�^�C�g��</param>
    ''' <Retuen>True�F�͂������@False�F����������</Retuen>
    ''' <remarks></remarks>
    Public Overloads Function DB_QUser(ByVal con As UniConnection, _
                                           ByVal objCTG As Object, ByVal intID As Integer, _
                                           ByVal strText As String, ByVal strTitle As String) As Boolean
        Return PRFBLN_QUser(con, objCTG, intID, strText, strTitle)
    End Function

#End Region

#Region "�₢���킹MSG(�V�K�E�X�V�E�폜�E�I��)"
    ''' <summary>
    ''' �o�^�p�⍇�����b�Z�[�W(�V�K��)
    ''' </summary>
    ''' <param name="con">�R�l�N�V����</param>
    ''' <param name="strText">���b�Z�[�W���e(�ړ���)</param>
    ''' <param name="strTitle">���b�Z�[�W�^�C�g��</param>
    ''' <Retuen>True�F�͂������@False�F����������</Retuen>
    ''' <remarks></remarks>
    Public Shared Function ShowMsgQUserInsert(ByVal con As UniConnection, _
                                     Optional ByVal strText As String = "", _
                                     Optional ByVal strTitle As String = "") As Boolean
        Return PRFBLN_QUser(con, PRCSTR_MSGCTG_INSERT, PRCSTR_MSGID_INSERT, strText, strTitle)
    End Function

    ''' <summary>
    ''' �o�^�p�⍇�����b�Z�[�W(�X�V��)
    ''' </summary>
    ''' <param name="con">�R�l�N�V����</param>
    ''' <param name="strText">���b�Z�[�W���e(�ړ���)</param>
    ''' <param name="strTitle">���b�Z�[�W�^�C�g��</param>
    ''' <Retuen>True�F�͂������@False�F����������</Retuen>
    ''' <remarks></remarks>
    Public Shared Function ShowMsgQUserUpdate(ByVal con As UniConnection, _
                                     Optional ByVal strText As String = "", _
                                     Optional ByVal strTitle As String = "") As Boolean
        Return PRFBLN_QUser(con, PRCSTR_MSGCTG_UPDATE, PRCSTR_MSGID_UPDATE, strText, strTitle)
    End Function

    ''' <summary>
    ''' �o�^�p�⍇�����b�Z�[�W(�폜��)
    ''' </summary>
    ''' <param name="con">�R�l�N�V����</param>
    ''' <param name="strText">���b�Z�[�W���e(�ړ���)</param>
    ''' <param name="strTitle">���b�Z�[�W�^�C�g��</param>
    ''' <Retuen>True�F�͂������@False�F����������</Retuen>
    ''' <remarks></remarks>
    Public Shared Function ShowMsgQUserDelete(ByVal con As UniConnection, _
                                     Optional ByVal strText As String = "", _
                                     Optional ByVal strTitle As String = "") As Boolean
        Return PRFBLN_QUser(con, PRCSTR_MSGCTG_DELETE, PRCSTR_MSGID_DELETE, strText, strTitle)
    End Function

    ''' <summary>
    ''' �I�����m�F���b�Z�[�W
    ''' </summary>
    ''' <param name="con">�R�l�N�V����</param>
    ''' <param name="strText">���b�Z�[�W���e(�ړ���)</param>
    ''' <param name="strTitle">���b�Z�[�W�^�C�g��</param>
    ''' <Retuen>True�F�͂������@False�F����������</Retuen>
    ''' <remarks></remarks>
    Public Function ShowMsgQUserExit(ByVal con As UniConnection, _
                                       Optional ByVal strText As String = "", _
                                       Optional ByVal strTitle As String = "") As Boolean
        Return PRFBLN_QUser(con, PRCSTR_MSGCTG_EXIT, PRCSTR_MSGID_EXIT, strText, strTitle)
    End Function
#End Region

#Region "Private���\�b�h"
#Region "MSG�}�X�^�擾"
    '----------------------------------------------------------------------
    '�@�\�@�FMSG�}�X�^�擾
    '�����@�FConnection�^STU_MSG�\���́^Optional(�e�L�X�g�ǉ���)
    '�߂�l�FTrue�^False, �\����(ByRef)
    '���l�@�F
    '----------------------------------------------------------------------
    Private Shared Function DB_GetMSG(ByVal con As UniConnection, _
                                   ByRef stuMSG As STU_MSG, _
                                   Optional ByVal strText As String = "") As Boolean
        Dim strSQL As String : Dim arlMSG As New ArrayList

        strSQL = ""
        strSQL = strSQL & " SELECT MSG_CTG, MSG_ID, MSG_TEXT "
        strSQL = strSQL & "      , MSG_ICON, MSG_PTN, MSG_DEF "
        'strSQL = strSQL & "      , MSG_TAN,  CONVERT(DateTime,MSG_DATE) "
        'strSQL = strSQL & "      , MSG_TAN, TO_CHAR(MSG_DATE, 'YYYY/MM/DD HH24:MI:SS')" 'ORACLE�Ή�
        strSQL = strSQL & " FROM M_SYS_MSG "
        strSQL = strSQL & " WHERE MSG_CTG = " & PBFSTR_SetQTT(stuMSG.MSG_CTG)
        strSQL = strSQL & "   AND MSG_ID = " & PBFSTR_SetQTT(stuMSG.MSG_ID)

        arlMSG = getAryDataDB(con, strSQL)

        If arlMSG.Count > 0 Then

            stuMSG.MSG_CTG = PBCStr(arlMSG(0))
            stuMSG.MSG_ID = PBCint(arlMSG(1))
            If PB_ChkNUll(strText) Then
                stuMSG.MSG_TEXT = PBCStr(arlMSG(2))
            Else
                stuMSG.MSG_TEXT = strText & PBCStr(arlMSG(2))
            End If

            stuMSG.MSG_ICON = PBCint(arlMSG(3))
            stuMSG.MSG_PTN = PBCint(arlMSG(4))
            stuMSG.MSG_DEF = PBCint(arlMSG(5))
            'stuMSG.MSG_TAN = PBCint(arlMSG(6))
            'stuMSG.MSG_DATE = PBCStr(arlMSG(7))
            Return True
        Else
            Return False
        End If
    End Function
#End Region
#Region "�f�[�^�擾(ONE RECORD)"
    '---------------------------------------------------------
    '�@�@�\�F�f�[�^�Q�b�g(Return ArrayList)
    '
    '�@�����@�FConnection, SQL��, Optional(Transaction)
    '�@�߂�l�FArrayList(�Q�b�g��������)
    '---------------------------------------------------------
    Private Shared Function getAryDataDB(ByVal ocon As UniConnection, ByVal SQL As String, _
                                        Optional ByVal tran As UniTransaction = Nothing) As ArrayList
        Dim ocd As New UniCommand
        Dim odr As UniDataReader
        Dim arlData As New ArrayList

        Try

            If IsNothing(tran) Then
                ocon = ChkConnection(ocon)
            End If

            ocd.Connection = ocon
            ocd.CommandText = SQL

            If Not tran Is Nothing Then
                ''ocd.Transaction = tran
            End If

            odr = ocd.ExecuteReader

            While (odr.Read)
                For i As Integer = 0 To odr.FieldCount - 1
                    With arlData
                        .Add(PBCStr(odr.Item(i)))
                    End With
                Next
            End While

            odr.Close()
            Return arlData
        Catch ex As Exception
            Throw ex
        End Try
    End Function
#End Region

#Region "�R�l�N�V�����`�F�b�N"
    '-------------------------------------------------------------------------------------
    '  �@�\    �FOracle Connection�`�F�b�N
    '            (��Ԋm�F���A�Đڑ����ĕԂ�)
    '  ����    �F1�DSqlConnection
    '  �߂�l  �FConnection 
    '
    '  �쐬��  �F2006.04.25  ��
    '-------------------------------------------------------------------------------------
    Private Shared Function ChkConnection(ByVal ocon As UniConnection) As UniConnection

        '�R�l�N�V�����������Ă���ꍇ�̂�
        If (ocon.State = ConnectionState.Closed) Then
            Dim Str As String = XMLReadConnection()
            ocon.ConnectionString = Str
            ocon.Open()
        End If

        Return ocon
    End Function
#End Region
#Region "XMLReadConnection�FXML�ڑ�������擾(CONNECTION)"
    '--------------------------------------------------------
    '  �@�\    �F�ڑ���������擾����B
    '  ����    �F�P�D�l
    '  �߂�l  �F�ڑ�������
    '  �쐬��  �F
    '--------------------------------------------------------
    Private Shared Function XMLReadConnection() As String
        Return PB_ReadXML("/SKY/SKY_DB/CONNECTION", "", SystemConst.C_SYSTEMPRM)
    End Function
#End Region

#Region "MSG�}�X�^�擾���sү����"
    Private Shared Sub PRS_MissingMSG(ByVal objCTG As Object, ByVal intID As Integer)
        Dim ERROR_MSG As String = PRCSTR_MISSING_MSG & vbCrLf & _
                        "(" & Convert.ToString(objCTG) & "," & intID & ")"
        '"�J�e�S���F" & Convert.ToString(objCTG) & vbCrLf &  "ID�F" & intID
        MsgBoxPB.Show(ERROR_MSG, PBCSTR_TITLE_MSG, MIcon.Error, MButton.OK)
    End Sub
#End Region
#End Region

#Region "����̃��b�Z�[�W"
#Region "�ΏۂƂȂ�f�[�^�����݂��܂���B"
    ''' <summary>
    ''' �ΏۂƂȂ�f�[�^�����݂��܂���B(�R�l�N�V�����L)
    ''' </summary>
    ''' <param name="con"></param>
    ''' <remarks></remarks>
    Public Shared Sub ShowMSG_NODATA(ByVal con As UniConnection)
        DB_ShowMsg(con, 0, 0)
    End Sub
#End Region

#End Region


#End Region

#Region "�R�l�N�V�����Ȃ��̃��b�Z�[�W�\�����\�b�h"
#Region "�װMSG"
    ''' <summary>
    ''' �G���[���b�Z�[�W�\��(�R�l�N�V�����Ȃ�)
    ''' </summary>
    ''' <param name="strError">���b�Z�[�W���e</param>
    ''' <param name="strTitle">���b�Z�[�W�e�L�X�g</param>
    ''' <remarks></remarks>
    Public Shared Sub ShowErrorMsg(ByVal strError As String, Optional ByVal strTitle As String = "")
        MsgBoxPB.Show(strError, getTitle(strTitle), MIcon.Error, MButton.OK)
    End Sub
#End Region
#Region "�x��MSG"
    ''' <summary>
    '''  �x�����b�Z�[�W�\��(�R�l�N�V�����Ȃ�)
    ''' </summary>
    ''' <param name="strWarning">���b�Z�[�W���e</param>
    ''' <param name="strTitle">���b�Z�[�W�e�L�X�g</param>
    ''' <remarks></remarks>
    Public Shared Sub ShowWarningMsg(ByVal strWarning As String, Optional ByVal strTitle As String = "")
        MsgBoxPB.Show(strWarning, getTitle(strTitle), MIcon.Warning, MButton.OK)
    End Sub
#End Region
#Region "�x��MSG(Yes/No)"
    ''' <summary>
    '''  �x�����b�Z�[�W�⍇���\��(�R�l�N�V�����Ȃ�)
    ''' </summary>
    ''' <param name="strWarning">���b�Z�[�W���e</param>
    ''' <param name="strTitle">���b�Z�[�W�e�L�X�g</param>
    ''' <Retuen>True�F�͂������@False�F����������</Retuen>
    ''' <remarks></remarks>
    Public Shared Function ShowWarningUserMsg(ByVal strWarning As String, Optional ByVal strTitle As String = "") As Boolean
        Dim dialRslt As DialogResult = DialogResult.No
        dialRslt = MsgBoxPB.Show(strWarning, getTitle(strTitle), MIcon.Warning, MButton.YesNo, MPosition.Button2)
        Return dialRslt = DialogResult.Yes
    End Function
#End Region
#Region "�m�FMSG"
    ''' <summary>
    '''  ��񃁃b�Z�[�W�\��(�R�l�N�V�����Ȃ�)
    ''' </summary>
    ''' <param name="strInfo">���b�Z�[�W���e</param>
    ''' <param name="strTitle">���b�Z�[�W�e�L�X�g</param>
    ''' <remarks></remarks>
    Public Shared Sub ShowInfoMsg(ByVal strInfo As String, Optional ByVal strTitle As String = "")
        MsgBoxPB.Show(strInfo, getTitle(strTitle), MIcon.Info, MButton.OK)
    End Sub
#End Region
#Region "�⍇��MSG�F20060626�ǉ�"
    ''' <summary>
    '''  ��񃁃b�Z�[�W�⍇���\��(�R�l�N�V�����Ȃ�)
    ''' </summary>
    ''' <param name="strInfo">���b�Z�[�W���e</param>
    ''' <param name="strTitle">���b�Z�[�W�e�L�X�g</param>
    ''' <Retuen>True�F�͂������@False�F����������</Retuen>
    ''' <remarks></remarks>
    Public Overloads Shared Function ShowUserMsg(ByVal strInfo As String, Optional ByVal strTitle As String = "") As Boolean
        Dim dialRslt As DialogResult = DialogResult.No
        dialRslt = MsgBoxPB.Show(strInfo, getTitle(strTitle), MIcon.Question, MButton.YesNo, MPosition.Button2)
        Return dialRslt = DialogResult.Yes
    End Function
    'END 2006/06/26
#End Region

#Region "Private����āF���ٕԂ�"
    Private Shared Function getTitle(ByVal strTitle As String) As String
        Dim strCaption As String
        If PB_ChkNUll(strTitle) Then
            strCaption = PBCSTR_TITLE_MSG
        Else
            strCaption = strTitle
        End If
        Return strCaption
    End Function
#End Region
#End Region
    ''' <summary>
    ''' XXXX���o�^����܂����B
    ''' </summary>
    ''' <param name="prmText"></param>
    ''' <param name="con"></param>
    ''' <remarks></remarks>
    Public Shared Sub ShowInfoMsgRegistIns(ByVal prmText As String, Optional ByVal con As UniConnection = Nothing)

        If Not con Is Nothing Then
            MessageUtil.ShowMsgWithCon(con, PBC_MSGCTG_REGISTED, PBC_MSGID_REGISTED, prmText & vbCrLf)
        Else
            ShowErrorMsg("�f�[�^���o�^����܂����B")
        End If
    End Sub
    ''' <summary>
    ''' XXXX���C������܂���
    ''' </summary>
    ''' <param name="prmText"></param>
    ''' <param name="con"></param>
    ''' <remarks></remarks>
    Public Shared Sub ShowInfoMsgRegistUpd(ByVal prmText As String, Optional ByVal con As UniConnection = Nothing)

        If Not con Is Nothing Then
            MessageUtil.ShowMsgWithCon(con, PBCSTR_MSGCTG_UPDATED, PBCSTR_MSGID_UPDATED, prmText & vbCrLf)
        Else
            ShowErrorMsg("�f�[�^���C������܂����B")
        End If
    End Sub
    ''' <summary>
    ''' �f�[�^���o�͂��܂���
    ''' </summary>
    ''' <param name="con"></param>
    ''' <remarks></remarks>
    Public Shared Sub ShowInfoMsgOutPutData(Optional ByVal con As UniConnection = Nothing)

        If Not con Is Nothing Then
            ShowMsgWithCon(con, 0, 31)
        Else
            ShowErrorMsg("�f�[�^���o�͂��܂����B")
        End If
    End Sub
    ''' <summary>
    ''' XXXXX��I�����Ă��������B
    ''' </summary>
    ''' <param name="prmText"></param>
    ''' <param name="con"></param>
    ''' <remarks></remarks>
    Public Shared Sub ShowErrorMsgMustISelect(ByVal prmText As String, Optional ByVal con As UniConnection = Nothing)

        If Not con Is Nothing Then
            MessageUtil.ShowMsgWithCon(con, 3, 1, prmText & vbCrLf)
        Else
            ShowErrorMsg(prmText & " ��I�����Ă��������B")
        End If
    End Sub
    ''' <summary>
    ''' ���̍��ڂ͕K�{���͂ł��B�������l����͂��ĉ������B
    ''' </summary>
    ''' <param name="prmText"></param>
    ''' <param name="con"></param>
    ''' <remarks></remarks>
    Public Shared Sub ShowErrorMsgMustInput(ByVal prmText As String, Optional ByVal con As UniConnection = Nothing)

        If Not con Is Nothing Then
            MessageUtil.ShowMsgWithCon(con, PBC_MSGCTG_MUST_INPUT, PBC_MSGID_MUST_INPUT, prmText & vbCrLf)
        Else
            ShowErrorMsg(prmText & vbCrLf & "���̍��ڂ͕K�{���͂ł��B�������l����͂��ĉ������B")
        End If
    End Sub
    ''' <summary>
    ''' �e���v���[�g�t�@�C����������܂���B
    ''' </summary>
    ''' <param name="prmText"></param>
    ''' <param name="con"></param>
    ''' <remarks></remarks>
    Public Shared Sub ShowErrorMsgNotExistTemplate(ByVal prmText As String, Optional ByVal con As UniConnection = Nothing)

        If Not con Is Nothing Then
            MessageUtil.ShowMsgWithCon(con, 0, 15, prmText & vbCrLf)
        Else
            ShowErrorMsg(prmText & vbCrLf & " �e���v���[�g�t�@�C����������܂���B")
        End If
    End Sub
    ''' <summary>
    ''' �L���Ȗ��ׂ����݂��܂���B
    ''' </summary>
    ''' <param name="con"></param>
    ''' <remarks></remarks>
    Public Shared Sub ShowErrorMsgNoItemData(Optional ByVal con As UniConnection = Nothing)

        If Not con Is Nothing Then
            MessageUtil.ShowMsgWithCon(con, 0, 19)
        Else
            ShowErrorMsg("�L���Ȗ��ׂ����݂��܂���B")
        End If
    End Sub
    ''' <summary>
    ''' �Y���f�[�^�����݂��܂���B
    ''' </summary>
    ''' <param name="con"></param>
    ''' <remarks></remarks>
    Public Shared Sub ShowErrorMsgNoData(Optional ByVal con As UniConnection = Nothing)

        If Not con Is Nothing Then
            ShowMsgWithCon(con, PBCSTR_MSGCTG_NODATA, PBCSTR_MSGID_NODATA)
        Else
            ShowErrorMsg("�Y���f�[�^�����݂��܂���B")
        End If
    End Sub
    ''' <summary>
    ''' 1���ȏ㌋�ʂ��Ɖ�Ă��������B
    ''' </summary>
    ''' <param name="con"></param>
    ''' <remarks></remarks>
    Public Shared Sub ShowErrorMsgNoResult(Optional ByVal con As UniConnection = Nothing)

        If Not con Is Nothing Then
            ShowMsgWithCon(con, PBCSTR_MSGCTG_NODATA, 21)
        Else
            ShowErrorMsg("1���ȏ㌋�ʂ��Ɖ�Ă�������")
        End If
    End Sub
    ''' <summary>
    ''' �o�^�ΏۂƂȂ�f�[�^�����݂��܂���B
    ''' </summary>
    ''' <param name="con"></param>
    ''' <remarks></remarks>
    Public Shared Sub ShowErrorMsgNoDataRegist(Optional ByVal con As UniConnection = Nothing)

        If Not con Is Nothing Then
            ShowMsgWithCon(con, PBCSTR_MSGCTG_NODATA, 20)
        Else
            ShowErrorMsg("�o�^�ΏۂƂȂ�f�[�^�����݂��܂���B")
        End If
    End Sub
    ''' <summary>
    ''' MSG�F���t�̑召���قȂ�܂��B"
    ''' </summary>
    ''' <param name="prmText">�O����</param>
    ''' <param name="con"></param>
    ''' <remarks></remarks>
    Public Shared Sub ShowErrorMsgDaySize(ByVal prmText As String, Optional ByVal con As UniConnection = Nothing)

        If Not con Is Nothing Then
            ShowMsgWithCon(con, PBCSTR_MSGCTG_DIFFER_DAY_SIZE, PBCSTR_MSGID_DIFFER_SIZE, prmText)
        Else
            ShowErrorMsg("���t�̑召���قȂ�܂��B")
        End If
    End Sub
    ''' <summary>
    ''' MSG�F���łɓ���ԍ������݂��܂�"
    ''' </summary>
    ''' <param name="prmText">�O����</param>
    ''' <param name="con"></param>
    ''' <remarks></remarks>
    Public Shared Sub ShowErrorMsgOverLap(ByVal prmText As String, Optional ByVal con As UniConnection = Nothing)

        If Not con Is Nothing Then
            ShowMsgWithCon(con, PBCSTR_MSGCTG_KEY_CONFLICT, PBCSTR_MSGID_KEY_CONFLICT, prmText & vbCrLf)
        Else
            ShowErrorMsg("���łɓ���ԍ������݂��܂�")
        End If
    End Sub
    ''' <summary>
    ''' MSG�F�P�ȏ�`�F�b�N�������Ȃ��Ă�������"
    ''' </summary>
    ''' <param name="prmText">�O����</param>
    ''' <param name="con"></param>
    ''' <remarks></remarks>
    Public Shared Sub ShowErrorItemCheckIsNothing(ByVal prmText As String, Optional ByVal con As UniConnection = Nothing)

        If Not con Is Nothing Then
            ShowMsgWithCon(con, 6, 1, prmText & vbCrLf)
        Else
            ShowErrorMsg("�P�ȏ�`�F�b�N��ݒ肵�Ă��������B")
        End If
    End Sub
    ''' <summary>
    ''' MSG�F�[���ȏ���w�肵�Ă�������"
    ''' </summary>
    ''' <param name="prmText">�O����</param>
    ''' <remarks></remarks>
    Public Shared Sub ShowErrorMsgMustZeroOver(ByVal prmText As String)

        ShowErrorMsg(prmText & vbCrLf & "�[���ȏ���w�肵�Ă��������B")
    End Sub
    ''' <summary>
    ''' MSG�F���׍s��I�����Ă��������B"
    ''' </summary>
    ''' <param name="con"></param>
    ''' <remarks></remarks>
    Public Shared Sub ShowErrorItemNoSelect(Optional ByVal con As UniConnection = Nothing)

        If Not con Is Nothing Then
            ShowMsgWithCon(con, "SELECT", 1)
        Else
            ShowErrorMsg("���׍s��I�����Ă��������B")
        End If
    End Sub
    ''' <summary>
    ''' MSG�F�G���[���������܂����B�����𒆎~���܂��B
    ''' </summary>
    ''' <param name="con"></param>
    ''' <remarks></remarks>
    Public Shared Sub ShowErrorActionStop(Optional ByVal con As UniConnection = Nothing)

        ShowErrorMsg("�G���[���������܂����B�����𒆎~���܂��B")
    End Sub
    ''' <summary>
    ''' MSG�FXXXX�����s���܂��B��낵���ł����H"
    ''' </summary>
    ''' <param name="prmText">�O����</param>
    ''' <param name="con"></param>
    ''' <remarks></remarks>
    Public Shared Function ShowUserMsgAction(ByVal prmText As String, Optional ByVal con As UniConnection = Nothing) As Boolean

        If Not con Is Nothing Then
            Return PRFBLN_QUser(con, 1, 5, prmText & vbCrLf)
        Else
            ShowErrorMsg(prmText & vbCrLf & " �����s���܂��B��낵���ł����H")
        End If
    End Function
    ''' <summary>
    ''' MSG�FXXXX���o�͂��܂��B��낵���ł����H"
    ''' </summary>
    ''' <param name="prmText">�O����</param>
    ''' <param name="con"></param>
    ''' <remarks></remarks>
    Public Shared Function ShowUserMsgOutPut(ByVal prmText As String, Optional ByVal con As UniConnection = Nothing) As Boolean

        If Not con Is Nothing Then
            Return PRFBLN_QUser(con, 1, 8, prmText & vbCrLf)
        Else
            ShowErrorMsg(prmText & vbCrLf & " ���o�͂��܂��B��낵���ł����H")
        End If
    End Function
    ''' <summary>
    ''' MSG�FXXXX���������܂����B"
    ''' </summary>
    ''' <param name="prmText">�O����</param>
    ''' <param name="con"></param>
    ''' <remarks></remarks>
    Public Shared Function ShowInfoMsgComplete(ByVal prmText As String, Optional ByVal con As UniConnection = Nothing) As Boolean

        If Not con Is Nothing Then
            Return PRFBLN_QUser(con, 0, 32, prmText & vbCrLf)
        Else
            ShowErrorMsg(prmText & vbCrLf & " ���������܂����B")
        End If
    End Function
End Class



#Region "MsgBox Shared �׽"
'*************************************************************************
'*�@�@�@�\�@�F����MSGShared �N���X(MessageBox��Show���b�\�hOverloads)
'*�@�쐬���@�F2006.05.18    ��
'*
'*  �ύX��  �F
'*  �ύX���e�F
'*  ���@�l�@�F
'*        
'*************************************************************************
Class MsgBoxPB

    '------------------------------------------------------------------------------
    '�@�\�@�FMSG�\���y �I�� Yes(�͂�)�ENo(������) �z
    '�����@�FMSG���e(strText)�^MSG�^�C�g��(strTitle)�^
    '       �@�A�C�R�����(intIcon)�^�{�^�����(intButton)�^�{�^���ʒu(intDfBtn)
    '�߂�l�F�����ꂽ�{�^���̈ʒu(Yes(6)�^No(7))
    '���l�@�F
    '------------------------------------------------------------------------------
    Overloads Shared Function Show(ByVal strText As String, ByVal strTitle As String, _
                                          ByVal intIcon As Integer, _
                                          ByVal intButton As Integer, _
                                          ByVal intDfBtn As Integer) As DialogResult
        Dim msgBtn As MessageBoxButtons
        Dim msgIcon As MessageBoxIcon
        Dim msgDftBtn As MessageBoxDefaultButton

        msgBtn = CType(intButton, MessageBoxButtons)
        msgIcon = CType(intIcon, MessageBoxIcon)
        msgDftBtn = CType(intDfBtn, MessageBoxDefaultButton)

        Return MessageBox.Show(strText, strTitle, msgBtn, msgIcon, msgDftBtn)
    End Function


    '------------------------------------------------------------------------------
    '�@�\�@�FMSG�\�� �y �P��I�� OK �z
    '�����@�FMSG���e(strText)�^MSG�^�C�g��(strTitle)�^
    '       �@�A�C�R�����(intIcon)�^�{�^�����(intButton)
    '�߂�l�F�m�F�{�^��OK(1)
    '���l�@�F
    '------------------------------------------------------------------------------
    Overloads Shared Function Show(ByVal strText As String, ByVal strTitle As String, _
                                          ByVal intIcon As Integer, ByVal intButton As Integer) As DialogResult
        Dim msgBtn As MessageBoxButtons
        Dim msgIcon As MessageBoxIcon

        msgBtn = CType(intButton, MessageBoxButtons)
        msgIcon = CType(intIcon, MessageBoxIcon)

        Return (MessageBox.Show(strText, strTitle, msgBtn, msgIcon))
    End Function

End Class
#End Region





