#Region "�錾"
Option Explicit On
Option Strict On
Imports System.Data.SqlClient
Imports skysystem.common.SystemUtil
#End Region

'*************************************************************************
'*�@�@�\�@�@�FDBconnection
'*            <DB�ڑ�>
'*�@�쐬���@�F
'*
'*�@���ύX���e��
'*
'*************************************************************************



''' <summary>
''' �f�[�^�x�[�X�R�l�N�V�����N���X
''' </summary>
''' <remarks>�R�l�N�V�����̐ڑ��E�ؒf�E�߂�ҿ���</remarks>
Public Class DBconnectionSS

    Private Const DecryptPass As String = "#2030" '20101208_1

#Region "Private�ϐ�"
    Private OraCon As New SqlConnection             '�I���N���R�l�N�V����
#End Region

#Region "�R���X�g���N�^"
    Public Sub New()

    End Sub
#End Region

#Region "ChkConnection�FOracleOpen�ڑ��F�`�F�b�N"
    ''' <summary>
    ''' �ڑ��I�[�v������
    ''' </summary>
    ''' <remarks>�R�l�N�V�����������Ă���ꍇ�́A�Đڑ����s���B</remarks>
    Public Sub Open()

        '�R�l�N�V�����������Ă���ꍇ�̂�
        If OraCon Is Nothing OrElse (OraCon.State = ConnectionState.Closed) Then

            Dim strDataSorce As String = PB_ReadXML("/SKY/SKY_DB/DataSorce", "", SystemConst.C_SYSTEMPRM)
            Dim strCatalog As String = PB_ReadXML("/SKY/SKY_DB/Catalog", "", SystemConst.C_SYSTEMPRM)
            Dim strUserId As String = PB_ReadXML("/SKY/SKY_DB/UserId", "", SystemConst.C_SYSTEMPRM)
            Dim strPassword As String = PB_ReadXML("/SKY/SKY_DB/Password", "", SystemConst.C_SYSTEMPRM)
            Dim strTimeOut As String = PB_ReadXML("/SKY/SKY_DB/TimeOut", "", SystemConst.C_SYSTEMPRM)

            Dim Str As String = CreateConnectionString(strDataSorce, strCatalog, strUserId, strPassword, strTimeOut)
            OraCon.ConnectionString = Str

            OraCon.Open()
        End If


    End Sub
#End Region

#Region "RtnCon�F�ڑ���Ԃ�"
    ''' <summary>
    ''' �ڑ�����߂�
    ''' </summary>
    ''' <returns>Sql�R�l�N�V����</returns>
    ''' <remarks></remarks>
    Public Function rtncon() As SqlConnection

        '�R�l�N�V�����������Ă���ꍇ�̂�
        Open()

        Return OraCon
    End Function
#End Region

#Region "Close�F�ڑ������"
    ''' <summary>
    ''' �ڑ������
    ''' </summary>
    Public Sub Close()
        OraCon.Close()
    End Sub
#End Region

#Region "CreateConnectionString:ConectionString�̐���"
    ''20101213_1
    Private Function CreateConnectionString(ByVal strDataSorce As String, ByVal strCatalog As String, ByVal strUserId As String, ByVal strPassword As String, ByVal strTimeOut As String) As String
        ''�Í�������Ă��镔���͕���������B
        Dim rtn As String = ""
        rtn += "Data Source=" + strDataSorce & ";"
        rtn += "Initial Catalog=" + SystemUtil.doDecrypt(strCatalog, DecryptPass) & ";"
        rtn += "User ID=" + SystemUtil.doDecrypt(strUserId, DecryptPass) & ";"
        rtn += "Password=" + SystemUtil.doDecrypt(strPassword, DecryptPass) & ";"
        rtn += "Connection Lifetime=" + strTimeOut & ";"
        Return rtn
    End Function
#End Region


End Class
