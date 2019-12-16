
Option Explicit On
Option Strict On

Imports skysystem.common.SystemUtil
Imports skysystem.common.MessageUtil
Imports skysystem.common.SystemConst
Imports System.Data.OleDb
Imports Devart.Data.Universal


'****************************************************************************************
'*　機　能　：共通MSGクラス(DB非連結)
'*　作成日　：2007/07/08
'*
'****************************************************************************************
''' <summary>
'''  共通メッセージライブラリ
''' </summary>
''' <remarks></remarks>
Public Class MessageUtil

#Region "Private定数(固定MSG)"
    Private Const PRCSTR_MISSING_MSG As String = "メッセージ取得失敗"

    '*** 問い合わせMSG
    '* 新規：登録します。よろしいですか？
    Private Const PRCSTR_MSGCTG_INSERT As Integer = 2
    Private Const PRCSTR_MSGID_INSERT As Integer = 0

    '* 更新：更新します。よろしいですか？
    Private Const PRCSTR_MSGCTG_UPDATE As Integer = 2
    Private Const PRCSTR_MSGID_UPDATE As Integer = 1

    '* 削除：表示されているデータを削除します。よろしいですか？
    Private Const PRCSTR_MSGCTG_DELETE As Integer = 2
    Private Const PRCSTR_MSGID_DELETE As Integer = 2

    '* 終了：終了してよろしいですか？
    Private Const PRCSTR_MSGCTG_EXIT As Integer = 1
    Private Const PRCSTR_MSGID_EXIT As Integer = 9
#End Region


#Region "Public定数(固定MSG)"

    '*** 完了MSG(登録・修正)
    '* 新規・複写
    '例) 依頼者CD G0000 & 'で登録されました。'
    Public Const PBC_MSGCTG_REGISTED As Integer = 10
    Public Const PBC_MSGID_REGISTED As Integer = 0

    '* 修正
    '例) 依頼者CD G0001 & 'が修正されました。'
    Public Const PBCSTR_MSGCTG_UPDATED As Integer = 10
    Public Const PBCSTR_MSGID_UPDATED As Integer = 1

    '**************************
    '*** 対象データチェック ***
    '**************************
    '対象となるデータがありません。
    Public Const PBCSTR_MSGCTG_NODATA As Integer = 0
    Public Const PBCSTR_MSGID_NODATA As Integer = 0

    '指定されたコードは登録されていません。 & vbCrLf & 別のコードで検索して下さい。
    Public Const PBCSTR_MSGCTG_NOT_REGISTED As Integer = 0
    Public Const PBCSTR_MSGID_NOT_REGISTED As Integer = 1

    'この項目はリスト内にある項目から選択して下さい。
    Public Const PBCSTR_MSGCTG_NO_LIST As Integer = 0
    Public Const PBCSTR_MSGID_NO_LIST As Integer = 2

    '*登録済み：既に登録されています。
    Public Const PBCSTR_MSGCTG_ERR_REGISTED As Integer = 0
    Public Const PBCSTR_MSGID_ERR_REGISTED As Integer = 4

    '(重複データ)同一番号が既に登録されています。
    Public Const PBCSTR_MSGCTG_KEY_CONFLICT As Integer = 0
    Public Const PBCSTR_MSGID_KEY_CONFLICT As Integer = 10

    '(必須チェック)この項目は必須入力です。 & vbCrLf & 正しい値を入力して下さい。
    Public Const PBC_MSGCTG_MUST_INPUT As Integer = 0
    Public Const PBC_MSGID_MUST_INPUT As Integer = 11

    '(範囲外)指定範囲外です。もう一度入力してください。
    Public Const PBCSTR_MSGCTG_OVERFLOW As Integer = 0
    Public Const PBCSTR_MSGID_OVERFLOW As Integer = 12

    'ADD 2006.08.30
    '該当のｺｰﾄﾞが存在しません。
    Public Const PBCSTR_MSGCTG_NOCODE As Integer = 0
    Public Const PBCSTR_MSGID_NOCODE As Integer = 13

    'ADD 2006.07.06
    '｢日付の大小が異なります。｣
    Public Const PBCSTR_MSGCTG_DIFFER_DAY_SIZE As Integer = 0
    Public Const PBCSTR_MSGID_DIFFER_DAY_SIZE As Integer = 9

    'ADD 2006.08.08
    '｢桁数が超えています。｣
    Public Const PBCINT_MSGCTG_OVER_DIGIT As Integer = 0
    Public Const PBCINT_MSGID_OVER_DIGIT As Integer = 14

    'ADD 20070525_1
    '｢いずれかのデータにチェックを行ってください。｣
    Public Const PBCINT_MSGCTG_NOCHECK As Integer = 0
    Public Const PBCINT_MSGID_NOCHECK As Integer = 16

    'ADD 20080714_1
    '｢大小が異なります。｣
    Public Const PBCINT_MSGCTG_DIFFER_SIZE As Integer = 0
    Public Const PBCSTR_MSGID_DIFFER_SIZE As Integer = 15

    'ADD 20080714_1
    '｢出力します。よろしいですか？｣
    Public Const PBCINT_MSGCTG_OUT_ACTION As Integer = 2
    Public Const PBCSTR_MSGID_OUT_ACTION As Integer = 3

    '**************************
    '*** ｽﾌﾟﾚｯﾄﾞｼｰﾄ関連 MSG ***
    '**************************

    '(明細入力ﾁｪｯｸ)｢○○｣を入力してください。
    Public Const PBCSTR_MSGCTG_CELLNULL As Integer = 10
    Public Const PBCSTR_MSGID_CELLNULL As Integer = 11

    '(明細入力ﾁｪｯｸ)0以上の数値を入力してください。
    Public Const PBCSTR_MSGCTG_CELLZERO As Integer = 0
    Public Const PBCSTR_MSGID_CELLZERO As Integer = 5

    '(明細未入力)明細が入力されていません。
    Public Const PBCSTR_MSGCTG_SPREADNULL As Integer = 6
    Public Const PBCSTR_MSGID_SPREADNULL As Integer = 0

    '(明細未入力)明細がチェックされていません。
    Public Const PBCSTR_MSGCTG_SPREADNOTCHECK As Integer = 6
    Public Const PBCSTR_MSGID_SPREADNOTCHECK As Integer = 1

    'ADD 2006/06/26
    '(明細入力)「○○○」が重複しています。
    Public Const PBCSTR_MSGCTG_SPREAD_CONFLICT As Integer = 6
    Public Const PBCSTR_MSGID_SPREAD_CONFLICT As Integer = 2


    '*** ↓名阪バスさんの共通モジュールとして分離(\common\Mei3\Mei3PB.vb参照)
    'Public Const PBCSTR_TITLE_MSG As String = "業務支援システム"

    Public Const PBC_MSG_INIT_ERR As String = "初期設定に失敗しました。"
    Public Const PBC_MSG_NO_CONNECTION As String = "コネクション設定が取得できません。"
    Public Const PBC_MSG_ERROR_DB As String = "エラーが発生しました。処理を取消します。"
    Public Const PBC_MSG_ALREADY_STARTED As String = "プログラムは既に起動されています。"
    'ADD 2006.08.01
    Public Const PBC_MSG_ERROR_STOP As String = "エラーが発生しました。処理を中止します。"
    Public Const PBC_MSG_STOP As String = "処理を中止します。"


    Public Const PBC_MSG_START As String = " スタート"
    Public Const PBC_MSG_END As String = " 終了"

    'ADD 2006.07.05
    '｢****日付を入力してください。｣
    Public Const PBCSTR_MSGCTG_INPUT_DAY As Integer = 10
    Public Const PBCSTR_MSGID_INPUT_DAY As Integer = 13


    'ADD 2006.07.06
    '**************************************
    '*** 個別PGMが持つ 共通MSG №1 ～   ***
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
    Public Const PBCINT_MSGID_XLS1 As Integer = 1      'EXCEL出力に失敗しました。
    Public Const PBCINT_MSGID_XLS2 As Integer = 2      'が確認できません。新たに作成してもよろしいですか？
    Public Const PBCINT_MSGID_XLS3 As Integer = 3      'EXCEL雛形の取得に失敗しました。
    Public Const PBCINT_MSGID_XLS4 As Integer = 4      'EXCEL書込に失敗しました。
    Public Const PBCINT_MSGID_XLS5 As Integer = 5      'EXCEL保存に失敗しました。
    Public Const PBCINT_MSGID_XLS6 As Integer = 6      '出力先のパスが確認できません。
    Public Const PBCINT_MSGID_XLS7 As Integer = 7      '既にファイルが存在します。上書きしますか？
    Public Const PBCINT_MSGID_XLS8 As Integer = 8      '読み取り専用ファイルです。書き込みできません。

    'ADD 2006.07.26
    Public Const PBCSTR_MSGCTG_FLD As String = "FLD"
    Public Const PBCINT_MSGID_FLD0 As Integer = 0       '該当するフォルダが存在しません。       
    Public Const PBCINT_MSGID_FLD1 As Integer = 1       '該当するフォルダが存在しません。作成しますか？
    Public Const PBCINT_MSGID_FLD2 As Integer = 2       'フォルダを削除します。よろしいですか？

    'ADD 2006.08.30
    Public Const PBCSTR_MSGCTG_NUM As String = "NUM"
    'ID_NO0     '採番マスタの指定範囲外です。
    'ID_NO1     '
    'ID_NO2     '
    'ID_NO3     '

    'ADD 2006.07.31 
    Public Const PBCSTR_MSG_NOWAIT As String = "他ユーザーがによりデータが使用されています。"

    'ADD 20061011_1
    Public Const PBCSTR_MSG_ERROR_DB_UNIQUE As String = "既に登録されています。もう一度、実行してください。"

    'ADD 2006.08.01
    '***************************
    '*** ORACLE ERROR CODE   ***
    '***************************
    Public Const PBCINT_ORAERR_CODE54 As Integer = 54   '「ORA-0054：リソース ビジーNoWait」

    'ADD 20061011_1
    Public Const PBCINT_ORAERR_CODE1 As Integer = 1     '「ORA-00001: 一意制約(MEI3.M_TOK_IDX1)に反しています」

#End Region

#Region "MSG Eunm定数"

    '*** ボタンアイコン
    Public Enum MIcon
        Info = 64 : [Error] = 16 : Warning = 48 : Question = 32
    End Enum

    '*** 選択ボタン数
    Public Enum MButton
        OK = 0 : YesNo = 4
        AbortRetryIgnore = 2    '予備
        OKCancel = 1            '予備    
        RetryCancel = 5         '予備
        YesNoCancel = 3         '予備
    End Enum

    '*** ボタンへSF位置
    Public Enum MPosition
        Button1 = 0 : Button2 = 256
        Button3 = 512   '予備
    End Enum

    '*** 選択処理
    Public Enum MResult
        OK = 1 : Yes = 6 : No = 7
        Abort = 3   '予備
        Cancel = 2  '予備
        Ignore = 5  '予備
        Retry = 4   '予備
        None = 0    '予備
    End Enum
#End Region

#Region "コネクションありのメッセージ表示メソッド"

#Region "MSG構造体"
    Public Structure STU_MSG

        Private objMsgCTG As Object       '(親)カテゴリー(共通分類, PG_ID)
        Private intMsgID As Integer       '(子)MSG_ID
        Private strMsgTitle As String     'MSGタイトル
        Private strMsgText As String      'MSGテキスト
        Private intMsgIcon As Integer     'アイコン種類
        Private intMsgPtn As Integer      'ボタンのパタン(ボタンの数)
        Private intMsgDef As Integer      'ボタンの位置
        Private intMsgTan As Integer      '予備：更新担当者
        Private strMsgDate As String      '予備：更新日付

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

                    Case 1  'OK(単一ボタン)
                        intMsgPtn = MButton.OK

                    Case 2  'YesNo(選択ボタン)
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

                    Case 1  '一番目
                        intMsgDef = MPosition.Button1

                    Case 2  '二番目
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

#Region "情報・警告MSG(単一選択)"
    ''' <summary>
    ''' 単一メッセージ表示(コネクションあり)
    ''' </summary>
    ''' <param name="con">コネクション</param>
    ''' <param name="objCTG">メッセージカテゴリ</param>
    ''' <param name="intID">メッセージID</param>
    ''' <param name="strText">メッセージ内容(接頭)</param>
    ''' <param name="strTitle">メッセージタイトル</param>
    ''' <remarks></remarks>
    Private Shared Sub DB_ShowMsg(ByVal con As UniConnection, _
                            ByVal objCTG As Object, ByVal intID As Integer, _
                            Optional ByVal strText As String = "", _
                            Optional ByVal strTitle As String = "")

        Dim stuMSG As New STU_MSG
        stuMSG.MSG_CTG = objCTG         'カテゴリ
        stuMSG.MSG_ID = intID           'ID
        stuMSG.MSG_TITLE = strTitle     'ﾒｯｾｰｼﾞﾀｲﾄﾙ

        If DB_GetMSG(con, stuMSG, strText) Then
            MsgBoxPB.Show(stuMSG.MSG_TEXT, stuMSG.MSG_TITLE, stuMSG.MSG_ICON, stuMSG.MSG_PTN)
        Else
            'Modify 2006.08.30
            'PRS_MissingMSG()    'MSG取得失敗時
            PRS_MissingMSG(objCTG, intID)    'MSG取得失敗時
        End If
    End Sub
    ''' <summary>
    ''' メッセージ表示メソッド(コネクションあり)
    ''' </summary>
    ''' <param name="con">コネクション</param>
    ''' <param name="objCTG">メッセージカテゴリ</param>
    ''' <param name="intID">メッセージID</param>
    ''' <remarks></remarks>
    Public Overloads Shared Sub ShowMsgWithCon(ByVal con As UniConnection, _
                                     ByVal objCTG As Object, ByVal intID As Integer)
        Call DB_ShowMsg(con, objCTG, intID)
    End Sub
    ''' <summary>
    ''' メッセージ表示メソッド(コネクションあり)
    ''' </summary>
    ''' <param name="con">コネクション</param>
    ''' <param name="objCTG">メッセージカテゴリ</param>
    ''' <param name="intID">メッセージID</param>
    ''' <param name="strText">メッセージ内容(前頭句)</param>
    ''' <remarks>メッセージタイトルはシステム名を利用</remarks>
    Public Overloads Shared Sub ShowMsgWithCon(ByVal con As UniConnection, _
                                     ByVal objCTG As Object, ByVal intID As Integer, _
                                     ByVal strText As String)
        Call DB_ShowMsg(con, objCTG, intID, strText)
    End Sub
    ''' <summary>
    ''' メッセージ表示メソッド(コネクションあり)
    ''' </summary>
    ''' <param name="con">コネクション</param>
    ''' <param name="objCTG">メッセージカテゴリ</param>
    ''' <param name="intID">メッセージID</param>
    ''' <param name="strText">メッセージ内容(前頭句)</param>
    ''' <param name="strTitle">メッセージタイトル</param>
    ''' <remarks></remarks>
    Public Overloads Shared Sub ShowMsgWithCon(ByVal con As UniConnection, _
                                     ByVal objCTG As Object, ByVal intID As Integer, _
                                     ByVal strText As String, ByVal strTitle As String)
        Call DB_ShowMsg(con, objCTG, intID, strText, strTitle)
    End Sub
#End Region

#Region "問合せMSG(選択)"
    '----------------------------------------------------------------------
    '機能　：問合せMSG(選択：Yes・No)
    '引数　：
    '戻り値：True／False
    '----------------------------------------------------------------------
    Public Shared Function PRFBLN_QUser(ByVal con As UniConnection, _
                                ByVal objCTG As Object, ByVal intID As Integer, _
                                Optional ByVal strText As String = "", _
                                Optional ByVal strTitle As String = "") As Boolean
        Dim stuMsg As New STU_MSG
        Dim dialRslt As DialogResult = DialogResult.No
        stuMsg.MSG_CTG = objCTG         'カテゴリ
        stuMsg.MSG_ID = intID           'ID
        stuMsg.MSG_TITLE = strTitle     'ﾒｯｾｰｼﾞﾀｲﾄﾙ

        If DB_GetMSG(con, stuMsg, strText) Then
            dialRslt = MsgBoxPB.Show(stuMsg.MSG_TEXT, stuMsg.MSG_TITLE, stuMsg.MSG_ICON, stuMsg.MSG_PTN, stuMsg.MSG_DEF)
        Else
            'Modify 2006.08.30
            'PRS_MissingMSG()    'MSG取得失敗時
            PRS_MissingMSG(objCTG, intID)    'MSG取得失敗時
        End If
        Return dialRslt = DialogResult.Yes
    End Function
    ''' <summary>
    ''' 問合せメッセージ表示(コネクションあり)
    ''' </summary>
    ''' <param name="con">コネクション</param>
    ''' <param name="objCTG">メッセージカテゴリ</param>
    ''' <param name="intID">メッセージID</param>
    ''' <Retuen>True：はい押下　False：いいえ押下</Retuen>
    ''' <remarks></remarks>
    Public Overloads Shared Function PBFBLN_QUser(ByVal con As UniConnection, _
                                         ByVal objCTG As Object, ByVal intID As Integer) As Boolean
        Return PRFBLN_QUser(con, objCTG, intID)
    End Function
    ''' <summary>
    ''' 問合せメッセージ表示(コネクションあり)
    ''' </summary>
    ''' <param name="con">コネクション</param>
    ''' <param name="objCTG">メッセージカテゴリ</param>
    ''' <param name="intID">メッセージID</param>
    ''' <param name="strText">メッセージ内容(前頭句)</param>
    ''' <Retuen>True：はい押下　False：いいえ押下</Retuen>
    ''' <remarks></remarks>
    Public Overloads Function PBFBLN_QUser(ByVal con As UniConnection, _
                                           ByVal objCTG As Object, ByVal intID As Integer, _
                                           ByVal strText As String) As Boolean
        Return PRFBLN_QUser(con, objCTG, intID, strText)
    End Function
    ''' <summary>
    ''' 問合せメッセージ表示(コネクションあり)
    ''' </summary>
    ''' <param name="con">コネクション</param>
    ''' <param name="objCTG">メッセージカテゴリ</param>
    ''' <param name="intID">メッセージID</param>
    ''' <param name="strText">メッセージ内容(前頭句)</param>
    ''' <Retuen>True：はい押下　False：いいえ押下</Retuen>
    ''' <remarks></remarks>
    Public Shared Function ShowUserMsgWithCon(ByVal con As UniConnection, _
                                           ByVal objCTG As Object, ByVal intID As Integer, _
                                           Optional ByVal strText As String = "") As Boolean
        Return PRFBLN_QUser(con, objCTG, intID, strText)
    End Function
    ''' <summary>
    ''' 問合せメッセージ表示(コネクションあり)
    ''' </summary>
    ''' <param name="con">コネクション</param>
    ''' <param name="objCTG">メッセージカテゴリ</param>
    ''' <param name="intID">メッセージID</param>
    ''' <param name="strText">メッセージ内容(前頭句)</param>
    ''' <param name="strTitle">メッセージタイトル</param>
    ''' <Retuen>True：はい押下　False：いいえ押下</Retuen>
    ''' <remarks></remarks>
    Public Overloads Function DB_QUser(ByVal con As UniConnection, _
                                           ByVal objCTG As Object, ByVal intID As Integer, _
                                           ByVal strText As String, ByVal strTitle As String) As Boolean
        Return PRFBLN_QUser(con, objCTG, intID, strText, strTitle)
    End Function

#End Region

#Region "問い合わせMSG(新規・更新・削除・終了)"
    ''' <summary>
    ''' 登録用問合せメッセージ(新規時)
    ''' </summary>
    ''' <param name="con">コネクション</param>
    ''' <param name="strText">メッセージ内容(接頭句)</param>
    ''' <param name="strTitle">メッセージタイトル</param>
    ''' <Retuen>True：はい押下　False：いいえ押下</Retuen>
    ''' <remarks></remarks>
    Public Shared Function ShowMsgQUserInsert(ByVal con As UniConnection, _
                                     Optional ByVal strText As String = "", _
                                     Optional ByVal strTitle As String = "") As Boolean
        Return PRFBLN_QUser(con, PRCSTR_MSGCTG_INSERT, PRCSTR_MSGID_INSERT, strText, strTitle)
    End Function

    ''' <summary>
    ''' 登録用問合せメッセージ(更新時)
    ''' </summary>
    ''' <param name="con">コネクション</param>
    ''' <param name="strText">メッセージ内容(接頭句)</param>
    ''' <param name="strTitle">メッセージタイトル</param>
    ''' <Retuen>True：はい押下　False：いいえ押下</Retuen>
    ''' <remarks></remarks>
    Public Shared Function ShowMsgQUserUpdate(ByVal con As UniConnection, _
                                     Optional ByVal strText As String = "", _
                                     Optional ByVal strTitle As String = "") As Boolean
        Return PRFBLN_QUser(con, PRCSTR_MSGCTG_UPDATE, PRCSTR_MSGID_UPDATE, strText, strTitle)
    End Function

    ''' <summary>
    ''' 登録用問合せメッセージ(削除時)
    ''' </summary>
    ''' <param name="con">コネクション</param>
    ''' <param name="strText">メッセージ内容(接頭句)</param>
    ''' <param name="strTitle">メッセージタイトル</param>
    ''' <Retuen>True：はい押下　False：いいえ押下</Retuen>
    ''' <remarks></remarks>
    Public Shared Function ShowMsgQUserDelete(ByVal con As UniConnection, _
                                     Optional ByVal strText As String = "", _
                                     Optional ByVal strTitle As String = "") As Boolean
        Return PRFBLN_QUser(con, PRCSTR_MSGCTG_DELETE, PRCSTR_MSGID_DELETE, strText, strTitle)
    End Function

    ''' <summary>
    ''' 終了時確認メッセージ
    ''' </summary>
    ''' <param name="con">コネクション</param>
    ''' <param name="strText">メッセージ内容(接頭句)</param>
    ''' <param name="strTitle">メッセージタイトル</param>
    ''' <Retuen>True：はい押下　False：いいえ押下</Retuen>
    ''' <remarks></remarks>
    Public Function ShowMsgQUserExit(ByVal con As UniConnection, _
                                       Optional ByVal strText As String = "", _
                                       Optional ByVal strTitle As String = "") As Boolean
        Return PRFBLN_QUser(con, PRCSTR_MSGCTG_EXIT, PRCSTR_MSGID_EXIT, strText, strTitle)
    End Function
#End Region

#Region "Privateメソッド"
#Region "MSGマスタ取得"
    '----------------------------------------------------------------------
    '機能　：MSGマスタ取得
    '引数　：Connection／STU_MSG構造体／Optional(テキスト追加分)
    '戻り値：True／False, 構造体(ByRef)
    '備考　：
    '----------------------------------------------------------------------
    Private Shared Function DB_GetMSG(ByVal con As UniConnection, _
                                   ByRef stuMSG As STU_MSG, _
                                   Optional ByVal strText As String = "") As Boolean
        Dim strSQL As String : Dim arlMSG As New ArrayList

        strSQL = ""
        strSQL = strSQL & " SELECT MSG_CTG, MSG_ID, MSG_TEXT "
        strSQL = strSQL & "      , MSG_ICON, MSG_PTN, MSG_DEF "
        'strSQL = strSQL & "      , MSG_TAN,  CONVERT(DateTime,MSG_DATE) "
        'strSQL = strSQL & "      , MSG_TAN, TO_CHAR(MSG_DATE, 'YYYY/MM/DD HH24:MI:SS')" 'ORACLE対応
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
#Region "データ取得(ONE RECORD)"
    '---------------------------------------------------------
    '　機能：データゲット(Return ArrayList)
    '
    '　引数　：Connection, SQL文, Optional(Transaction)
    '　戻り値：ArrayList(ゲットしたもの)
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

#Region "コネクションチェック"
    '-------------------------------------------------------------------------------------
    '  機能    ：Oracle Connectionチェック
    '            (状態確認し、再接続して返す)
    '  引数    ：1．SqlConnection
    '  戻り値  ：Connection 
    '
    '  作成日  ：2006.04.25  黄
    '-------------------------------------------------------------------------------------
    Private Shared Function ChkConnection(ByVal ocon As UniConnection) As UniConnection

        'コネクションが閉じられている場合のみ
        If (ocon.State = ConnectionState.Closed) Then
            Dim Str As String = XMLReadConnection()
            ocon.ConnectionString = Str
            ocon.Open()
        End If

        Return ocon
    End Function
#End Region
#Region "XMLReadConnection：XML接続文字列取得(CONNECTION)"
    '--------------------------------------------------------
    '  機能    ：接続文字列を取得する。
    '  引数    ：１．値
    '  戻り値  ：接続文字列
    '  作成日  ：
    '--------------------------------------------------------
    Private Shared Function XMLReadConnection() As String
        Return PB_ReadXML("/SKY/SKY_DB/CONNECTION", "", SystemConst.C_SYSTEMPRM)
    End Function
#End Region

#Region "MSGマスタ取得失敗ﾒｯｾｰｼﾞ"
    Private Shared Sub PRS_MissingMSG(ByVal objCTG As Object, ByVal intID As Integer)
        Dim ERROR_MSG As String = PRCSTR_MISSING_MSG & vbCrLf & _
                        "(" & Convert.ToString(objCTG) & "," & intID & ")"
        '"カテゴリ：" & Convert.ToString(objCTG) & vbCrLf &  "ID：" & intID
        MsgBoxPB.Show(ERROR_MSG, PBCSTR_TITLE_MSG, MIcon.Error, MButton.OK)
    End Sub
#End Region
#End Region

#Region "特定のメッセージ"
#Region "対象となるデータが存在しません。"
    ''' <summary>
    ''' 対象となるデータが存在しません。(コネクション有)
    ''' </summary>
    ''' <param name="con"></param>
    ''' <remarks></remarks>
    Public Shared Sub ShowMSG_NODATA(ByVal con As UniConnection)
        DB_ShowMsg(con, 0, 0)
    End Sub
#End Region

#End Region


#End Region

#Region "コネクションなしのメッセージ表示メソッド"
#Region "ｴﾗｰMSG"
    ''' <summary>
    ''' エラーメッセージ表示(コネクションなし)
    ''' </summary>
    ''' <param name="strError">メッセージ内容</param>
    ''' <param name="strTitle">メッセージテキスト</param>
    ''' <remarks></remarks>
    Public Shared Sub ShowErrorMsg(ByVal strError As String, Optional ByVal strTitle As String = "")
        MsgBoxPB.Show(strError, getTitle(strTitle), MIcon.Error, MButton.OK)
    End Sub
#End Region
#Region "警告MSG"
    ''' <summary>
    '''  警告メッセージ表示(コネクションなし)
    ''' </summary>
    ''' <param name="strWarning">メッセージ内容</param>
    ''' <param name="strTitle">メッセージテキスト</param>
    ''' <remarks></remarks>
    Public Shared Sub ShowWarningMsg(ByVal strWarning As String, Optional ByVal strTitle As String = "")
        MsgBoxPB.Show(strWarning, getTitle(strTitle), MIcon.Warning, MButton.OK)
    End Sub
#End Region
#Region "警告MSG(Yes/No)"
    ''' <summary>
    '''  警告メッセージ問合せ表示(コネクションなし)
    ''' </summary>
    ''' <param name="strWarning">メッセージ内容</param>
    ''' <param name="strTitle">メッセージテキスト</param>
    ''' <Retuen>True：はい押下　False：いいえ押下</Retuen>
    ''' <remarks></remarks>
    Public Shared Function ShowWarningUserMsg(ByVal strWarning As String, Optional ByVal strTitle As String = "") As Boolean
        Dim dialRslt As DialogResult = DialogResult.No
        dialRslt = MsgBoxPB.Show(strWarning, getTitle(strTitle), MIcon.Warning, MButton.YesNo, MPosition.Button2)
        Return dialRslt = DialogResult.Yes
    End Function
#End Region
#Region "確認MSG"
    ''' <summary>
    '''  情報メッセージ表示(コネクションなし)
    ''' </summary>
    ''' <param name="strInfo">メッセージ内容</param>
    ''' <param name="strTitle">メッセージテキスト</param>
    ''' <remarks></remarks>
    Public Shared Sub ShowInfoMsg(ByVal strInfo As String, Optional ByVal strTitle As String = "")
        MsgBoxPB.Show(strInfo, getTitle(strTitle), MIcon.Info, MButton.OK)
    End Sub
#End Region
#Region "問合せMSG：20060626追加"
    ''' <summary>
    '''  情報メッセージ問合せ表示(コネクションなし)
    ''' </summary>
    ''' <param name="strInfo">メッセージ内容</param>
    ''' <param name="strTitle">メッセージテキスト</param>
    ''' <Retuen>True：はい押下　False：いいえ押下</Retuen>
    ''' <remarks></remarks>
    Public Overloads Shared Function ShowUserMsg(ByVal strInfo As String, Optional ByVal strTitle As String = "") As Boolean
        Dim dialRslt As DialogResult = DialogResult.No
        dialRslt = MsgBoxPB.Show(strInfo, getTitle(strTitle), MIcon.Question, MButton.YesNo, MPosition.Button2)
        Return dialRslt = DialogResult.Yes
    End Function
    'END 2006/06/26
#End Region

#Region "Privateｲﾍﾞﾝﾄ：ﾀｲﾄﾙ返す"
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
    ''' XXXXが登録されました。
    ''' </summary>
    ''' <param name="prmText"></param>
    ''' <param name="con"></param>
    ''' <remarks></remarks>
    Public Shared Sub ShowInfoMsgRegistIns(ByVal prmText As String, Optional ByVal con As UniConnection = Nothing)

        If Not con Is Nothing Then
            MessageUtil.ShowMsgWithCon(con, PBC_MSGCTG_REGISTED, PBC_MSGID_REGISTED, prmText & vbCrLf)
        Else
            ShowErrorMsg("データが登録されました。")
        End If
    End Sub
    ''' <summary>
    ''' XXXXを修正されました
    ''' </summary>
    ''' <param name="prmText"></param>
    ''' <param name="con"></param>
    ''' <remarks></remarks>
    Public Shared Sub ShowInfoMsgRegistUpd(ByVal prmText As String, Optional ByVal con As UniConnection = Nothing)

        If Not con Is Nothing Then
            MessageUtil.ShowMsgWithCon(con, PBCSTR_MSGCTG_UPDATED, PBCSTR_MSGID_UPDATED, prmText & vbCrLf)
        Else
            ShowErrorMsg("データが修正されました。")
        End If
    End Sub
    ''' <summary>
    ''' データを出力しました
    ''' </summary>
    ''' <param name="con"></param>
    ''' <remarks></remarks>
    Public Shared Sub ShowInfoMsgOutPutData(Optional ByVal con As UniConnection = Nothing)

        If Not con Is Nothing Then
            ShowMsgWithCon(con, 0, 31)
        Else
            ShowErrorMsg("データを出力しました。")
        End If
    End Sub
    ''' <summary>
    ''' XXXXXを選択してください。
    ''' </summary>
    ''' <param name="prmText"></param>
    ''' <param name="con"></param>
    ''' <remarks></remarks>
    Public Shared Sub ShowErrorMsgMustISelect(ByVal prmText As String, Optional ByVal con As UniConnection = Nothing)

        If Not con Is Nothing Then
            MessageUtil.ShowMsgWithCon(con, 3, 1, prmText & vbCrLf)
        Else
            ShowErrorMsg(prmText & " を選択してください。")
        End If
    End Sub
    ''' <summary>
    ''' この項目は必須入力です。正しい値を入力して下さい。
    ''' </summary>
    ''' <param name="prmText"></param>
    ''' <param name="con"></param>
    ''' <remarks></remarks>
    Public Shared Sub ShowErrorMsgMustInput(ByVal prmText As String, Optional ByVal con As UniConnection = Nothing)

        If Not con Is Nothing Then
            MessageUtil.ShowMsgWithCon(con, PBC_MSGCTG_MUST_INPUT, PBC_MSGID_MUST_INPUT, prmText & vbCrLf)
        Else
            ShowErrorMsg(prmText & vbCrLf & "この項目は必須入力です。正しい値を入力して下さい。")
        End If
    End Sub
    ''' <summary>
    ''' テンプレートファイルが見つかりません。
    ''' </summary>
    ''' <param name="prmText"></param>
    ''' <param name="con"></param>
    ''' <remarks></remarks>
    Public Shared Sub ShowErrorMsgNotExistTemplate(ByVal prmText As String, Optional ByVal con As UniConnection = Nothing)

        If Not con Is Nothing Then
            MessageUtil.ShowMsgWithCon(con, 0, 15, prmText & vbCrLf)
        Else
            ShowErrorMsg(prmText & vbCrLf & " テンプレートファイルが見つかりません。")
        End If
    End Sub
    ''' <summary>
    ''' 有効な明細が存在しません。
    ''' </summary>
    ''' <param name="con"></param>
    ''' <remarks></remarks>
    Public Shared Sub ShowErrorMsgNoItemData(Optional ByVal con As UniConnection = Nothing)

        If Not con Is Nothing Then
            MessageUtil.ShowMsgWithCon(con, 0, 19)
        Else
            ShowErrorMsg("有効な明細が存在しません。")
        End If
    End Sub
    ''' <summary>
    ''' 該当データが存在しません。
    ''' </summary>
    ''' <param name="con"></param>
    ''' <remarks></remarks>
    Public Shared Sub ShowErrorMsgNoData(Optional ByVal con As UniConnection = Nothing)

        If Not con Is Nothing Then
            ShowMsgWithCon(con, PBCSTR_MSGCTG_NODATA, PBCSTR_MSGID_NODATA)
        Else
            ShowErrorMsg("該当データが存在しません。")
        End If
    End Sub
    ''' <summary>
    ''' 1件以上結果を照会してください。
    ''' </summary>
    ''' <param name="con"></param>
    ''' <remarks></remarks>
    Public Shared Sub ShowErrorMsgNoResult(Optional ByVal con As UniConnection = Nothing)

        If Not con Is Nothing Then
            ShowMsgWithCon(con, PBCSTR_MSGCTG_NODATA, 21)
        Else
            ShowErrorMsg("1件以上結果を照会してください")
        End If
    End Sub
    ''' <summary>
    ''' 登録対象となるデータが存在しません。
    ''' </summary>
    ''' <param name="con"></param>
    ''' <remarks></remarks>
    Public Shared Sub ShowErrorMsgNoDataRegist(Optional ByVal con As UniConnection = Nothing)

        If Not con Is Nothing Then
            ShowMsgWithCon(con, PBCSTR_MSGCTG_NODATA, 20)
        Else
            ShowErrorMsg("登録対象となるデータが存在しません。")
        End If
    End Sub
    ''' <summary>
    ''' MSG：日付の大小が異なります。"
    ''' </summary>
    ''' <param name="prmText">前頭句</param>
    ''' <param name="con"></param>
    ''' <remarks></remarks>
    Public Shared Sub ShowErrorMsgDaySize(ByVal prmText As String, Optional ByVal con As UniConnection = Nothing)

        If Not con Is Nothing Then
            ShowMsgWithCon(con, PBCSTR_MSGCTG_DIFFER_DAY_SIZE, PBCSTR_MSGID_DIFFER_SIZE, prmText)
        Else
            ShowErrorMsg("日付の大小が異なります。")
        End If
    End Sub
    ''' <summary>
    ''' MSG：すでに同一番号が存在します"
    ''' </summary>
    ''' <param name="prmText">前頭句</param>
    ''' <param name="con"></param>
    ''' <remarks></remarks>
    Public Shared Sub ShowErrorMsgOverLap(ByVal prmText As String, Optional ByVal con As UniConnection = Nothing)

        If Not con Is Nothing Then
            ShowMsgWithCon(con, PBCSTR_MSGCTG_KEY_CONFLICT, PBCSTR_MSGID_KEY_CONFLICT, prmText & vbCrLf)
        Else
            ShowErrorMsg("すでに同一番号が存在します")
        End If
    End Sub
    ''' <summary>
    ''' MSG：１つ以上チェックをおこなってください"
    ''' </summary>
    ''' <param name="prmText">前頭句</param>
    ''' <param name="con"></param>
    ''' <remarks></remarks>
    Public Shared Sub ShowErrorItemCheckIsNothing(ByVal prmText As String, Optional ByVal con As UniConnection = Nothing)

        If Not con Is Nothing Then
            ShowMsgWithCon(con, 6, 1, prmText & vbCrLf)
        Else
            ShowErrorMsg("１つ以上チェックを設定してください。")
        End If
    End Sub
    ''' <summary>
    ''' MSG：ゼロ以上を指定してください"
    ''' </summary>
    ''' <param name="prmText">前頭句</param>
    ''' <remarks></remarks>
    Public Shared Sub ShowErrorMsgMustZeroOver(ByVal prmText As String)

        ShowErrorMsg(prmText & vbCrLf & "ゼロ以上を指定してください。")
    End Sub
    ''' <summary>
    ''' MSG：明細行を選択してください。"
    ''' </summary>
    ''' <param name="con"></param>
    ''' <remarks></remarks>
    Public Shared Sub ShowErrorItemNoSelect(Optional ByVal con As UniConnection = Nothing)

        If Not con Is Nothing Then
            ShowMsgWithCon(con, "SELECT", 1)
        Else
            ShowErrorMsg("明細行を選択してください。")
        End If
    End Sub
    ''' <summary>
    ''' MSG：エラーが発生しました。処理を中止します。
    ''' </summary>
    ''' <param name="con"></param>
    ''' <remarks></remarks>
    Public Shared Sub ShowErrorActionStop(Optional ByVal con As UniConnection = Nothing)

        ShowErrorMsg("エラーが発生しました。処理を中止します。")
    End Sub
    ''' <summary>
    ''' MSG：XXXXを実行します。よろしいですか？"
    ''' </summary>
    ''' <param name="prmText">前頭句</param>
    ''' <param name="con"></param>
    ''' <remarks></remarks>
    Public Shared Function ShowUserMsgAction(ByVal prmText As String, Optional ByVal con As UniConnection = Nothing) As Boolean

        If Not con Is Nothing Then
            Return PRFBLN_QUser(con, 1, 5, prmText & vbCrLf)
        Else
            ShowErrorMsg(prmText & vbCrLf & " を実行します。よろしいですか？")
        End If
    End Function
    ''' <summary>
    ''' MSG：XXXXを出力します。よろしいですか？"
    ''' </summary>
    ''' <param name="prmText">前頭句</param>
    ''' <param name="con"></param>
    ''' <remarks></remarks>
    Public Shared Function ShowUserMsgOutPut(ByVal prmText As String, Optional ByVal con As UniConnection = Nothing) As Boolean

        If Not con Is Nothing Then
            Return PRFBLN_QUser(con, 1, 8, prmText & vbCrLf)
        Else
            ShowErrorMsg(prmText & vbCrLf & " を出力します。よろしいですか？")
        End If
    End Function
    ''' <summary>
    ''' MSG：XXXXを完了しました。"
    ''' </summary>
    ''' <param name="prmText">前頭句</param>
    ''' <param name="con"></param>
    ''' <remarks></remarks>
    Public Shared Function ShowInfoMsgComplete(ByVal prmText As String, Optional ByVal con As UniConnection = Nothing) As Boolean

        If Not con Is Nothing Then
            Return PRFBLN_QUser(con, 0, 32, prmText & vbCrLf)
        Else
            ShowErrorMsg(prmText & vbCrLf & " を完了しました。")
        End If
    End Function
End Class



#Region "MsgBox Shared ｸﾗｽ"
'*************************************************************************
'*　機　能　：共通MSGShared クラス(MessageBoxのShowメッソドOverloads)
'*　作成日　：2006.05.18    黄
'*
'*  変更日  ：
'*  変更内容：
'*  備　考　：
'*        
'*************************************************************************
Class MsgBoxPB

    '------------------------------------------------------------------------------
    '機能　：MSG表示【 選択 Yes(はい)・No(いいえ) 】
    '引数　：MSG内容(strText)／MSGタイトル(strTitle)／
    '       　アイコン種類(intIcon)／ボタン種類(intButton)／ボタン位置(intDfBtn)
    '戻り値：押されたボタンの位置(Yes(6)／No(7))
    '備考　：
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
    '機能　：MSG表示 【 単一選択 OK 】
    '引数　：MSG内容(strText)／MSGタイトル(strTitle)／
    '       　アイコン種類(intIcon)／ボタン種類(intButton)
    '戻り値：確認ボタンOK(1)
    '備考　：
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





