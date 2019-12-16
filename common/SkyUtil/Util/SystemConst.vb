
'********************************************************************
'* �\�[�X�t�@�C���� : SystemConst.vb
'* �N���X���@�@	    : SystemConst
'* �N���X�����@	    : �V�X�e���萔�ꗗ
'* ���l�@           :
'* �쐬  �@         : 
'* �X�V����         :
'********************************************************************
''' <summary>
''' �V�X�e���萔�ꗗ
''' </summary>
''' <remarks></remarks>
Public Class SystemConst

#Region "Public �萔"
    ''' <summary>
    ''' �V�X�e���萔�F�c�_
    ''' </summary>
    ''' <remarks></remarks>
    Public Const PBCSTR_VERTICAL As String = "�b"
    ''' <summary>
    ''' �V�X�e���萔�F�X�y�[�X(���p)
    ''' </summary>
    ''' <remarks></remarks>
    Public Const PBCSTR_SPACE_H As String = " "
    ''' <summary>
    ''' �V�X�e���萔�F�X�y�[�X(�S�p)
    ''' </summary>
    ''' <remarks></remarks>
    Public Const PBCSTR_SPACE As String = "�@"
    ''' <summary>
    ''' �V�X�e���萔�F�V�X�e����
    ''' </summary>
    ''' <remarks></remarks>
    Public Shared PBCSTR_TITLE_MSG As String = ""
    ''' <summary>
    ''' �V�X�e���萔�F�V�X�e����(����)
    ''' </summary>
    ''' <remarks></remarks>
    Public Shared PBCSTR_TITLE_SHORT As String = " "
    ''' <summary>
    ''' �V�X�e���萔�F�V�X�e���p�����[�^�t�@�C����
    ''' </summary>
    ''' <remarks></remarks>
    Public Const C_SYSTEMPRM As String = "Skysystem.xml"
    ''' <summary>
    ''' �V�X�e���萔�F�V�X�e��FTP�\���t�@�C����
    ''' </summary>
    ''' <remarks></remarks>
    Public Const C_FTPCONFIG As String = "ftpConfig.xml"
    ''' <summary>
    ''' �V�X�e���萔�Fհ�ް�\���t�@�C��
    ''' </summary>
    ''' <remarks></remarks>
    Public Const C_USRCONF_XML As String = "userConfig.xml"
    ''' <summary>
    ''' �V�X�e���萔�F���[�J���I�v�V����
    ''' </summary>
    ''' <remarks></remarks>
    Public Const C_TABLEHIST_XML As String = "datatableHisotry.xml"

#End Region

#Region "Public�萔(�Œ�MSG)"

    '*** ����MSG(�o�^�E�C��)
    '* �V�K�E����
    '��) �˗���CD G0000 & '�œo�^����܂����B'
    Public Const PBCSTR_MSGCTG_REGISTED As Integer = 10
    Public Const PBCSTR_MSGID_REGISTED As Integer = 0

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
    Public Const PBCSTR_MSGCTG_MUST_INPUT As Integer = 0
    Public Const PBCSTR_MSGID_MUST_INPUT As Integer = 11

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

    Public Const PBCSTR_MSG_INIT_ERR As String = "�����ݒ�Ɏ��s���܂����B"
    Public Const PBCSTR_MSG_NO_CONNECTION As String = "�R�l�N�V�����ݒ肪�擾�ł��܂���B"
    Public Const PBCSTR_MSG_ERROR_DB As String = "�װ���������܂����B������������܂��B"
    Public Const PBCSTR_MSG_ALREADY_STARTED As String = "�v���O�����͊��ɋN������Ă��܂��B"
    'ADD 2006.08.01
    Public Const PBCSTR_MSG_ERROR_STOP As String = "�װ���������܂����B�����𒆎~���܂��B"
    Public Const PBCSTR_MSG_STOP As String = "�����𒆎~���܂��B"
    Public Const PBCSTR_MSG_WARN_1 As String = "�����͂���Ă��܂���B�������p�����܂����H"
    Public Const PBCSTR_MSG_WARN_2 As String = "�s�ڂ��傫�����t���w�肳��Ă��܂��B" & vbCrLf & "�������p�����܂����H"
    Public Const PBCSTR_MSG_WARN_3 As String = "�s�ڂ�藂�N�ȏ�̔N�x���w�肳��Ă��܂��B" & vbCrLf & "�������p�����܂����H"
    Public Const PBCSTR_MSG_WARN_4 As String = "�s�ڂ��傫���w�N���w�肳��Ă��܂��B" & vbCrLf & "�������p�����܂����H"
    Public Const PBCSTR_MSG_WARN_5 As String = "�͍X�V����܂���B�����𑱂��܂����H"

    Public Const PBCSTR_MSG_START As String = " �X�^�[�g"
    Public Const PBCSTR_MSG_END As String = " �I��"

    Public Const PBCSTR_MSG_ERROR_1 As String = "���̍��ڂ͕K�{���͂ł��B�������l����͂��Ă��������B"
    Public Const PBCSTR_MSG_ERROR_2 As String = "���̍��ڂ̓��X�g���ɂ��鍀�ڂ���I�����Ă��������B"

    'ADD 2006.07.05
    '�****���t����͂��Ă��������B�
    Public Const PBCSTR_MSGCTG_INPUT_DAY As Integer = 10
    Public Const PBCSTR_MSGID_INPUT_DAY As Integer = 13

    '�m�F���b�Z�[�W
    Public Const PBCSTR_RPT_OUT As String = "���o�͂��܂����B"


    '20080103_1 DataValidating�Œ胁�b�Z�[�W
    Public Const PBC_NULL As String = "�F��l�̂��ߍX�V�ł��܂���B"
    Public Const PBC_NotDate As String = "�F���t�Ƃ��ĔF�߂��܂���B"
    Public Const PBC_NotHalf As String = "�F���p�ȊO�̕������͂��F�߂��܂��B"
    Public Const PBC_Camma As String = "�F�J���}���܂܂�Ă��邽�ߘA�g�ł��܂���B"





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

#End Region '���b�Z�[�W�pConst

#Region "Public �񋓑�"
    ''' <summary>
    ''' �ېŋ敪
    ''' </summary>
    ''' <remarks></remarks>
    Enum KAZEI
        KAZEI = 0
        HIKAZEI = 1
    End Enum
    ''' <summary>
    ''' �ō��敪
    ''' </summary>
    ''' <remarks></remarks>
    Enum ZEIK
        ZEIIN = 0
        ZEIOUT = 1
    End Enum
    ''' <summary>
    ''' �l�ϊ�
    ''' </summary>
    ''' <remarks></remarks>
    Enum CVT_KIND
        CHR = 0
        MAIL = 1
        NUM = 2
    End Enum
    ''' <summary>
    ''' �󕥎��
    ''' </summary>
    ''' <remarks></remarks>
    Enum MOVE_TP
        NKA = 10
        SKA = 20
        IDO = 30
        TANKA = 50
        DEL = 90
    End Enum
    ''' <summary>
    ''' �󕥋敪
    ''' </summary>
    ''' <remarks></remarks>
    Enum MOVE_KB
        SKA = 0
        NKA = 1
    End Enum
    ''' <summary>
    ''' �@�\���
    ''' </summary>
    ''' <remarks></remarks>
    Enum FUNC_TP
        BTN = 0
        LABEL = 1
    End Enum
    ''' <summary>
    ''' OK/NG
    ''' </summary>
    ''' <remarks></remarks>
    Enum ONFLG
        NG = 0
        OK = 1
    End Enum
    ''' <summary>
    ''' ON.OFF
    ''' </summary>
    ''' <remarks></remarks>
    Enum OFFLG
        OFF = 0
        [ON] = 1
    End Enum
#Region "�l�̌ܓ�"
    ''' <summary>
    ''' �l�̌ܓ�
    ''' </summary>
    ''' <remarks></remarks>
    Enum Round
        UP = 0
        Down = 1
        Half = 2
    End Enum
    ''' <summary>
    ''' �ύX���������敪
    ''' </summary>
    Enum COM_Kbn
        CUSTOMER = 1
        PURCHASE = 2
    End Enum

#End Region

#End Region



End Class
