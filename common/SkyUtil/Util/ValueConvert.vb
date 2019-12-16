#Region "宣言"
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
'* ソースファイル名 : ValueChange.vb
'* クラス名　　	    : ValueChange
'* クラス説明　	    : 値変更クラス
'* 備考　           :
'* 作成  　         : 2007/02/22 駒方
'* 更新履歴         :
'********************************************************************
Public Class ValueConvert
    Private Shared hRandom As New System.Random()

#Region "値変更処理：テストなどに使用"
    Public Shared Function doValVCng(ByVal prmValue As String, ByVal prmCvtKind As CVT_KIND) As String

        Dim CvtChar() As String = {"子", "牛", "虎", "兎", "辰", "巳", "午", "未", "申", "鳥", "戌", "亥"}
        Dim CvtNum() As String = {"0", "1", "2", "3", "4", "5", "6", "7", "8", "9"}
        Dim CvtHalf() As String = {"a", "b", "c", "d", "e", "f", "g", "h", "j", "k", "m", "n", "p", "q", "r", "s", "t", "u", "v", "x", "y", "z", "!", "#", "$", "&", "-", "_"}
        Dim iCnt As Integer = 0
        Dim Val As String
        Dim rtnVal As String = ""

        '文字数取得
        iCnt = prmValue.Length
        If iCnt = 0 Then
            Return ""
        End If

        Select Case prmCvtKind


            '住所や名称等
            Case CVT_KIND.CHR
                '文字列をランダムに変換する
                For i As Integer = 0 To iCnt - 1
                    Val = prmValue.Substring(i, 1)
                    If Val.Equals(" ") Or Val.Equals("　") Then
                        rtnVal += Val
                    Else
                        rtnVal += CvtChar(hRandom.Next(0, CvtChar.Length))
                    End If
                Next

                'メールアドレス
            Case CVT_KIND.MAIL

                '文字列をランダムに変換する
                For i As Integer = 0 To iCnt - 1
                    Val = prmValue.Substring(i, 1)
                    If Val.Equals(".") Or Val.Equals("@") Then
                        rtnVal += Val
                    Else
                        rtnVal += CvtHalf(hRandom.Next(0, CvtHalf.Length))
                    End If
                Next

                '電話やFAX等
            Case CVT_KIND.NUM

                '文字列をランダムに変換する
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
