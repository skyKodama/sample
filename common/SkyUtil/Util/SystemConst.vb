
'********************************************************************
'* ソースファイル名 : SystemConst.vb
'* クラス名　　	    : SystemConst
'* クラス説明　	    : システム定数一覧
'* 備考　           :
'* 作成  　         : 
'* 更新履歴         :
'********************************************************************
''' <summary>
''' システム定数一覧
''' </summary>
''' <remarks></remarks>
Public Class SystemConst

#Region "Public 定数"
    ''' <summary>
    ''' システム定数：縦棒
    ''' </summary>
    ''' <remarks></remarks>
    Public Const PBCSTR_VERTICAL As String = "｜"
    ''' <summary>
    ''' システム定数：スペース(半角)
    ''' </summary>
    ''' <remarks></remarks>
    Public Const PBCSTR_SPACE_H As String = " "
    ''' <summary>
    ''' システム定数：スペース(全角)
    ''' </summary>
    ''' <remarks></remarks>
    Public Const PBCSTR_SPACE As String = "　"
    ''' <summary>
    ''' システム定数：システム名
    ''' </summary>
    ''' <remarks></remarks>
    Public Shared PBCSTR_TITLE_MSG As String = ""
    ''' <summary>
    ''' システム定数：システム名(略名)
    ''' </summary>
    ''' <remarks></remarks>
    Public Shared PBCSTR_TITLE_SHORT As String = " "
    ''' <summary>
    ''' システム定数：システムパラメータファイル名
    ''' </summary>
    ''' <remarks></remarks>
    Public Const C_SYSTEMPRM As String = "Skysystem.xml"
    ''' <summary>
    ''' システム定数：システムFTP構成ファイル名
    ''' </summary>
    ''' <remarks></remarks>
    Public Const C_FTPCONFIG As String = "ftpConfig.xml"
    ''' <summary>
    ''' システム定数：ﾕｰｻﾞｰ構成ファイル
    ''' </summary>
    ''' <remarks></remarks>
    Public Const C_USRCONF_XML As String = "userConfig.xml"
    ''' <summary>
    ''' システム定数：ローカルオプション
    ''' </summary>
    ''' <remarks></remarks>
    Public Const C_TABLEHIST_XML As String = "datatableHisotry.xml"

#End Region

#Region "Public定数(固定MSG)"

    '*** 完了MSG(登録・修正)
    '* 新規・複写
    '例) 依頼者CD G0000 & 'で登録されました。'
    Public Const PBCSTR_MSGCTG_REGISTED As Integer = 10
    Public Const PBCSTR_MSGID_REGISTED As Integer = 0

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
    Public Const PBCSTR_MSGCTG_MUST_INPUT As Integer = 0
    Public Const PBCSTR_MSGID_MUST_INPUT As Integer = 11

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

    Public Const PBCSTR_MSG_INIT_ERR As String = "初期設定に失敗しました。"
    Public Const PBCSTR_MSG_NO_CONNECTION As String = "コネクション設定が取得できません。"
    Public Const PBCSTR_MSG_ERROR_DB As String = "ｴﾗｰが発生しました。処理を取消します。"
    Public Const PBCSTR_MSG_ALREADY_STARTED As String = "プログラムは既に起動されています。"
    'ADD 2006.08.01
    Public Const PBCSTR_MSG_ERROR_STOP As String = "ｴﾗｰが発生しました。処理を中止します。"
    Public Const PBCSTR_MSG_STOP As String = "処理を中止します。"
    Public Const PBCSTR_MSG_WARN_1 As String = "が入力されていません。処理を継続しますか？"
    Public Const PBCSTR_MSG_WARN_2 As String = "行目より大きい日付が指定されています。" & vbCrLf & "処理を継続しますか？"
    Public Const PBCSTR_MSG_WARN_3 As String = "行目より翌年以上の年度が指定されています。" & vbCrLf & "処理を継続しますか？"
    Public Const PBCSTR_MSG_WARN_4 As String = "行目より大きい学年が指定されています。" & vbCrLf & "処理を継続しますか？"
    Public Const PBCSTR_MSG_WARN_5 As String = "は更新されません。処理を続けますか？"

    Public Const PBCSTR_MSG_START As String = " スタート"
    Public Const PBCSTR_MSG_END As String = " 終了"

    Public Const PBCSTR_MSG_ERROR_1 As String = "この項目は必須入力です。正しい値を入力してください。"
    Public Const PBCSTR_MSG_ERROR_2 As String = "この項目はリスト内にある項目から選択してください。"

    'ADD 2006.07.05
    '｢****日付を入力してください。｣
    Public Const PBCSTR_MSGCTG_INPUT_DAY As Integer = 10
    Public Const PBCSTR_MSGID_INPUT_DAY As Integer = 13

    '確認メッセージ
    Public Const PBCSTR_RPT_OUT As String = "を出力しました。"


    '20080103_1 DataValidating固定メッセージ
    Public Const PBC_NULL As String = "：空値のため更新できません。"
    Public Const PBC_NotDate As String = "：日付として認められません。"
    Public Const PBC_NotHalf As String = "：半角以外の文字入力が認められます。"
    Public Const PBC_Camma As String = "：カンマが含まれているため連携できません。"





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

#End Region 'メッセージ用Const

#Region "Public 列挙対"
    ''' <summary>
    ''' 課税区分
    ''' </summary>
    ''' <remarks></remarks>
    Enum KAZEI
        KAZEI = 0
        HIKAZEI = 1
    End Enum
    ''' <summary>
    ''' 税込区分
    ''' </summary>
    ''' <remarks></remarks>
    Enum ZEIK
        ZEIIN = 0
        ZEIOUT = 1
    End Enum
    ''' <summary>
    ''' 値変換
    ''' </summary>
    ''' <remarks></remarks>
    Enum CVT_KIND
        CHR = 0
        MAIL = 1
        NUM = 2
    End Enum
    ''' <summary>
    ''' 受払種別
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
    ''' 受払区分
    ''' </summary>
    ''' <remarks></remarks>
    Enum MOVE_KB
        SKA = 0
        NKA = 1
    End Enum
    ''' <summary>
    ''' 機能種別
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
#Region "四捨五入"
    ''' <summary>
    ''' 四捨五入
    ''' </summary>
    ''' <remarks></remarks>
    Enum Round
        UP = 0
        Down = 1
        Half = 2
    End Enum
    ''' <summary>
    ''' 変更履歴処理区分
    ''' </summary>
    Enum COM_Kbn
        CUSTOMER = 1
        PURCHASE = 2
    End Enum

#End Region

#End Region



End Class
