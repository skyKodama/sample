#Region "�錾"
Option Explicit On
Option Strict On

Imports System.IO
Imports System.Xml
Imports System.Text
Imports skysystem.common
Imports skysystem.common.SystemConst
Imports skysystem.common.MessageUtil

#End Region

'********************************************************************
'* �\�[�X�t�@�C���� : ValueChange.vb
'* �N���X���@�@	    : ValueChange
'* �N���X�����@	    : �l�ύX�N���X
'* ���l�@           :
'* �쐬  �@         : 2007/02/22 ���
'* �X�V����         :
'********************************************************************
Public Class ValueConvert
    Private Shared hRandom As New System.Random()

#Region "�l�ύX�����F�e�X�g�ȂǂɎg�p"
    Public Shared Function doValVCng(ByVal prmValue As String, ByVal prmCvtKind As CVT_KIND) As String

        Dim CvtChar() As String = {"�q", "��", "��", "�e", "�C", "��", "��", "��", "�\", "��", "��", "��"}
        Dim CvtNum() As String = {"0", "1", "2", "3", "4", "5", "6", "7", "8", "9"}
        Dim CvtHalf() As String = {"a", "b", "c", "d", "e", "f", "g", "h", "j", "k", "m", "n", "p", "q", "r", "s", "t", "u", "v", "x", "y", "z", "!", "#", "$", "&", "-", "_"}
        Dim iCnt As Integer = 0
        Dim Val As String
        Dim rtnVal As String = ""

        '�������擾
        iCnt = prmValue.Length
        If iCnt = 0 Then
            Return ""
        End If

        Select Case prmCvtKind


            '�Z���▼�̓�
            Case CVT_KIND.CHR
                '������������_���ɕϊ�����
                For i As Integer = 0 To iCnt - 1
                    Val = prmValue.Substring(i, 1)
                    If Val.Equals(" ") Or Val.Equals("�@") Then
                        rtnVal += Val
                    Else
                        rtnVal += CvtChar(hRandom.Next(0, CvtChar.Length))
                    End If
                Next

                '���[���A�h���X
            Case CVT_KIND.MAIL

                '������������_���ɕϊ�����
                For i As Integer = 0 To iCnt - 1
                    Val = prmValue.Substring(i, 1)
                    If Val.Equals(".") Or Val.Equals("@") Then
                        rtnVal += Val
                    Else
                        rtnVal += CvtHalf(hRandom.Next(0, CvtHalf.Length))
                    End If
                Next

                '�d�b��FAX��
            Case CVT_KIND.NUM

                '������������_���ɕϊ�����
                For i As Integer = 0 To iCnt - 1
                    Val = prmValue.Substring(i, 1)
                    If Val.Equals("-") Or Val.Equals("0") Then
                        rtnVal += Val
                    Else
                        rtnVal += CvtNum(hRandom.Next(0, CvtNum.Length))
                    End If
                Next


        End Select


        Return rtnVal

    End Function
#End Region


End Class
