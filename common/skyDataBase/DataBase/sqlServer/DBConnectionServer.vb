#Region "�錾"
Option Explicit On
Option Strict On
Imports System.Data.SqlClient
Imports skysystem.common.SystemUtil
#End Region

'*************************************************************************
'*�@�@�\�@�@�FDatabaseConnectionServer�i���C���T�[�o�[�ւ̐ڑ��j
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
Public Class DBconnectionServer

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
            Dim Str As String = XMLReadConnectionS()
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

#Region "XMLReadConnection�FXML�ڑ�������擾(CONNECTION)"
    ''' <summary>
    ''' XML���ڑ���������擾����
    ''' </summary>
    ''' <returns>�ڑ�������</returns>
    ''' <remarks></remarks>
    Private Function XMLReadConnectionS() As String
        Return PB_ReadXML("/SKY/SKY_DB/CONNECTIONS", "", SystemConst.C_SYSTEMPRM)
    End Function
#End Region



End Class
