
Option Explicit On
Option Strict On
Imports System.IO
Imports System.Xml
Imports System.Text
Imports skysystem.common
Imports skysystem.common.SystemConst
Imports skysystem.common.MessageUtil
Imports System.Windows.Forms
Imports Devart.Data.Universal
Imports System.Globalization



'********************************************************************
'* ソースファイル名 : SystemUtil.vb
'* クラス名　　	    : SystemUtil
'* クラス説明　	    : システム共通ユーティリティー
'* 備考　           :
'* 作成  　         : 2007/07/08 駒方
'* 更新履歴         :
' 20090201_1 Komagta OpenFileDialogの改善(初期値のファイルパス表示)
'********************************************************************
''' <summary>
''' システム共通ユーティリティー
''' </summary>
''' <remarks></remarks>
Public Class SystemUtil

#Region "列挙体"
#Region "全角・半角"
    ''' <summary>
    ''' 全角・半角・混在の列挙対
    ''' </summary>
    ''' <remarks>全角・半角・混在の列挙対</remarks>
    Public Enum CHAR_SIZE

        FULLHALF = 0 '混在'
        FULL = 1 '全角
        HALF = 2 '半角
    End Enum
#End Region
#Region "ﾀﾞｲｱﾛｸﾞﾌｨﾙﾀｰ"
    ''' <summary>
    ''' ファイアログボックスのファイルフィルター"
    ''' </summary>
    ''' <remarks></remarks>
    Public Enum FILEKIND
        XLS = 0
        TXT = 1
        CSV = 2
        PDF = 3
        XLSX = 5
        ETC = 9
    End Enum
#End Region
#Region "DLL・EXE実行区分"
    Public Enum START_PG
        DLL     'DLL, 参照起動
        EXE     'EXE, 単独起動
    End Enum
#End Region
#Region "更新確認"
    Public Enum ACTION
        INS = 0 '新規モード
        UPD = 1 '修正モード
        RO = 2 '表示モード
        DEL = 3 '表示モード
        ERR = 9 'エラー等
    End Enum
#End Region
#Region "本支店"
    Public Enum KBHNS
        HNSYA = 0 '本社
        SISYA = 1 '新車
    End Enum
#End Region
#Region "チェック"
    ''' <summary>
    ''' チェックの有無
    ''' </summary>
    ''' <remarks></remarks>
    Public Enum CHK
        [FALSE] = 0 'なし
        [TRUE] = 1 'あり
    End Enum
#End Region
#Region "右か左か"
    Public Enum LorR
        [LEFT] = 0 'なし
        [RIGHT] = 1 'あり
    End Enum
#End Region

#End Region

#Region "★【Combo関連】"
    '#Region "コンボデータの生成"
    '    ''' <summary>
    '    ''' コンボデータの生成
    '    ''' </summary>
    '    ''' <param name="kbn"></param>
    '    ''' <returns></returns>
    '    ''' <remarks>NFUで使用・現在非共通のためPrivateへ変更</remarks>
    '    Private Shared Function CreateComboData(Optional ByVal kbn As Integer = Nothing) As DataTable
    '        Dim dtset As DataSet = New DataSet
    '        Dim dtt As DataTable
    '        Dim dtRow As DataRow
    '        Dim FCode As String = "CODE"
    '        Dim FName As String = "NAME"


    '        dtt = dtset.Tables.Add("TEMP")
    '        dtt.Columns.Add(FCode, Type.GetType("System.String"))
    '        dtt.Columns.Add(FName, Type.GetType("System.String"))

    '        'レコード追加
    '        dtRow = dtt.NewRow
    '        dtRow(FCode) = "0" : dtRow(FName) = "未発行" : dtt.Rows.Add(dtRow)

    '        dtRow = dtt.NewRow
    '        dtRow(FCode) = "1" : dtRow(FName) = "発行" : dtt.Rows.Add(dtRow)

    '        Return dtt

    '    End Function
    '#End Region '20071019_1
    '#Region "値ﾁｪｯｸ"
    '    ''' <summary>
    '    ''' コンボチェックチェック処理
    '    ''' </summary>
    '    ''' <param name="ctrlCombo">コンボコントロール</param>
    '    ''' <param name="blnPermitNULL">NULL許可有無(Optional, Default：許可しない) </param>
    '    ''' <returns>Boolean(True：正常　False：異常)</returns>
    '    ''' <remarks></remarks>
    '    Public Overloads Shared Function ChkCombo(ByVal ctrlCombo As IM.Combo, _
    '                                              Optional ByVal blnPermitNULL As Boolean = False) As Boolean
    '        'Dim blnResult As Boolean
    '        If Trim(ctrlCombo.Value) = "" Then
    '            'blnResult = blnPermitNULL
    '            Return blnPermitNULL
    '        Else
    '            If Not IsNothing(ctrlCombo.SelectedItem) Then
    '                'blnResult = True
    '                Return True
    '            Else
    '                If ctrlCombo.Value <> "" Then
    '                    Dim i As Integer
    '                    For i = 0 To ctrlCombo.Items.Count - 1
    '                        If GetCmbContent(ctrlCombo) = CStr(ctrlCombo.Items.Item(i).Value) Then
    '                            'blnResult = True
    '                            Return True
    '                        End If
    '                    Next
    '                End If
    '            End If
    '        End If
    '        'Return blnResult
    '        Return False
    '    End Function
    '    '-------------------------------------------------------------------------------------------------
    '    '引数：cmbCtrl           ( ComboControl：IM.Combo )
    '    '      blFlg_PermitNull  ( NULL許可有無：Boolean )
    '    '      strTitle          ( メッセージTitle：String )
    '    '      inDegit           ( 桁数：Integer(Default：1桁) )
    '    '      ocon              ( SqlConnection：MSG表示用 )
    '    '
    '    '備考：               【新規】：NULL許可(True) ／【変更】：NULL許可無し(False)
    '    '作成日  ：2006.06.06  黄
    '    '--------------------------------------------------------------------------------------------------
    '    Public Overloads Shared Function ChkCombo(ByVal cmbCtrl As IM.Combo, _
    '                                    ByVal blFlg_PermitNull As Boolean, _
    '                                    ByVal strTitle As String) As Boolean

    '        If Not ChkCombo(cmbCtrl, blFlg_PermitNull) Then
    '            If Trim(cmbCtrl.Value) = "" And blFlg_PermitNull = False Then

    '                ''DEL 20061018_1
    '                ''20061113_1　復帰
    '                '< MSG(0,11)：この項目は必須入力です。正しい値を入力してください。>
    '                'PBS_ShowMsg(ocon, PBCSTR_MSGCTG_MUST_INPUT, PBCSTR_MSGID_MUST_INPUT, strTitle & vbCrLf)
    '                ShowErrorMsg(PBCSTR_MSG_ERROR_1, strTitle & vbCrLf)
    '                Return False
    '            Else
    '                '< MSG(0,2)：この項目はリスト内にある項目から選択してください。>
    '                'PBS_ShowMsg(ocon, PBCSTR_MSGCTG_NO_LIST, PBCSTR_MSGID_NO_LIST, strTitle & vbCrLf)
    '                ShowErrorMsg(PBCSTR_MSG_ERROR_2, strTitle & vbCrLf)
    '                cmbCtrl.Text = ""
    '                Return False
    '            End If
    '        End If
    '        Return True
    '    End Function
    '    '-------------------------------------------------------------------------------------------------
    '    '引数：cmbCtrl           ( ComboControl：IM.Combo )
    '    '      strContent        ( 指定Content )
    '    '戻り値：Content存在=True
    '    '備考：   
    '    '作成日：2006.07.14     黄
    '    '修正日：2006.07.21     黄 (選択されたIndex番号返す)
    '    '--------------------------------------------------------------------------------------------------
    '    Public Shared Function ChkComboContent(ByVal cmbCtrl As IM.Combo, _
    '                                           ByVal strContent As String, _
    '                                           Optional ByRef index As Integer = 0) As Boolean
    '        Dim blnRtn As Boolean
    '        If strContent <> "" Then
    '            For i As Integer = 0 To cmbCtrl.Items.Count - 1
    '                If strContent = CStr(cmbCtrl.Items.Item(i).Value) Then
    '                    index = i       '← ADD 2006.07.21
    '                    blnRtn = True
    '                    Exit For
    '                End If
    '            Next
    '        End If
    '        Return blnRtn
    '    End Function
    '#End Region

    '#Region "値取得(Content・Description)"
    '    ' ------------------------------------------------------------------ 
    '    '　機能　：コンボテキストのContentを返す
    '    '
    '    '　引数　：コンボコントロール(cmbCtrl)
    '    '　戻り値：コンボテキスト中の'｜'前の文字を返す
    '    '
    '    '  作成日：2006.05.10　黄
    '    ' ------------------------------------------------------------------ 
    '    Public Overloads Shared Function GetCmbContent(ByVal cmbCtrl As IM.Combo) As String
    '        Dim intLength As Integer
    '        Dim strCmbVal, strContent As String
    '        Try
    '            If cmbCtrl.Value = "" Then Return ""
    '            strCmbVal = Trim(cmbCtrl.Value)

    '            intLength = InStr(strCmbVal, PBCSTR_VERTICAL)
    '            If intLength = 0 Then
    '                strContent = strCmbVal
    '            Else
    '                strContent = PBFSTR_MidB(strCmbVal, 1, intLength - 1)
    '            End If
    '            Return strContent
    '        Catch ex As Exception
    '            Throw ex
    '        End Try
    '    End Function
    '    ' ------------------------------------------------------------------ 
    '    '　機能　：コンボテキストのContentを返す
    '    '
    '    '　引数　：コンボコントロール(cmbCtrl)
    '    '          strValue(ﾃｷｽﾄ)
    '    '　戻り値：コンボテキスト中の'｜'前の文字を返す
    '    '
    '    '  作成日：2006.07.20　黄
    '    ' ------------------------------------------------------------------ 
    '    Public Overloads Shared Function GetCmbContent(ByVal cmbCtrl As IM.Combo, _
    '                                                   ByVal strValue As String) As String
    '        Dim intLength As Integer
    '        Dim strContent As String
    '        Try
    '            strContent = Trim(strValue)
    '            intLength = InStr(strContent, PBCSTR_VERTICAL)
    '            If intLength = 0 Then
    '                'Modify 2006.08.04
    '                'strContent = strContent
    '                strContent = Trim(strContent)
    '            Else
    '                'Modify 2006.08.04
    '                'strContent = PBFSTR_MidB(strContent, 1, intLength - 1)
    '                strContent = Trim(PBFSTR_MidB(strContent, 1, intLength - 1))
    '            End If
    '            If ChkComboContent(cmbCtrl, strContent) Then
    '                Return strContent
    '            Else
    '                Return ""
    '            End If
    '        Catch ex As Exception
    '            Throw ex
    '        End Try
    '    End Function
    ''' <summary>
    ''' コンボのテキストの"｜"右文字列を取得
    ''' </summary>
    ''' <param name="strValue"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overloads Shared Function GetCmbText(ByVal strValue As String) As String
        Dim intLength As Integer
        Dim strCmbVal As String = ""
        Dim strContent As String = ""
        Try

            intLength = InStr(strValue, PBCSTR_VERTICAL)
            If intLength = 0 Then
                strContent = strValue
            Else
                strContent = strValue.Substring(intLength)
            End If
            Return strContent
        Catch ex As Exception
            Throw ex
        End Try
    End Function
    ''' <summary>
    ''' コンボのテキストの"｜"左文字列を取得
    ''' </summary>
    ''' <param name="strValue"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overloads Shared Function GetCmbCode(ByVal strValue As String) As String
        Dim intLength As Integer
        Dim strCmbVal As String = ""
        Dim strContent As String = ""
        Try

            intLength = InStr(strValue, PBCSTR_VERTICAL)
            'intLength = strValue.Length - intLength

            If intLength = 0 Then
                strContent = strValue
            Else
                strContent = strValue.Substring(0, intLength - 1).Trim
            End If
            Return strContent
        Catch ex As Exception
            Throw ex
        End Try
    End Function
    '    ' ------------------------------------------------------------------ 
    '    '　機能　：コンボテキストのDescription(名称)を返す
    '    '
    '    '　引数　：コンボコントロール(cmbCtrl)
    '    '　戻り値：コンボテキスト中のDescription(説明)
    '    '
    '    '  作成日：2006.06.08　黄
    '    '  修正日：2006.08.28  黄
    '    ' ------------------------------------------------------------------ 
    '    Public Shared Function GetCmbDescription(ByVal cmbCtrl As IM.Combo, _
    '                                             Optional ByVal strContent As String = "") As String
    '        Dim strDescription As String = ""
    '        If cmbCtrl.Value <> "" OrElse strContent <> "" Then

    '            'ADD 2006.08.28
    '            If strContent = "" Then strContent = GetCmbContent(cmbCtrl)
    '            For i As Integer = 0 To cmbCtrl.Items.Count - 1
    '                'Modify 2006.08.28
    '                'If GetCmbContent(cmbCtrl) = CStr(cmbCtrl.Items.Item(i).Value) Then
    '                If strContent = CStr(cmbCtrl.Items.Item(i).Content) Then
    '                    strDescription = PBCStr(cmbCtrl.Items.Item(i).Description)
    '                End If
    '            Next
    '        End If
    '        Return strDescription
    '    End Function
    '#End Region
    '#Region "ｺﾝﾎﾞﾎﾞｯｸｽ(DropDownList)：値ﾁｪｯｸ・返し"
    '    ''' <summary>
    '    ''' コンボボックスのリスト名称を戻す
    '    ''' </summary>
    '    ''' <param name="ctrlCombo">対象のコンボボックスコントロール</param>
    '    ''' <param name="strValue">ターゲットコード(値)</param>
    '    ''' <returns></returns>
    '    ''' <remarks></remarks>
    '    Public Shared Function GetCmbListDescription(ByVal ctrlCombo As IM.Combo, _
    '                                                           Optional ByVal strValue As String = "") As String
    '        With ctrlCombo
    '            Dim strcontents As String = ""
    '            If strValue = "" Then strValue = .Text
    '            If strValue = "" Then Return ""

    '            For i As Integer = 0 To .Items.Count - 1
    '                If strValue = CStr(.Items.Item(i).Value) Then
    '                    strcontents = CStr(.Items.Item(i).Content)
    '                    Exit For
    '                End If
    '            Next
    '            Return strcontents
    '        End With
    '    End Function
    '#End Region
#End Region

#Region "★【XML関連】"
#Region "XMLファイル書き込み"
    Public Shared Sub PB_WriteXML(ByVal xmlPath As String, ByVal prmElement As String, ByVal prmValue As String)
        Try


            Dim domDoc As New XmlDocument
            Dim domNode As XmlNode

            'XML 形式の文字列データを設定する 
            domDoc.Load(xmlPath)

            '特定の要素にアクセスする 

            domNode = domDoc.SelectSingleNode(prmElement)
            domNode.InnerText = prmValue

            ''特定の属性にアクセスする 
            'domNode = domDoc.SelectSingleNode("//Item/@att")
            'Console.WriteLine("{0} => {1}", domNode.LocalName, domNode.Value)

            'ファイルとして保存する 

            domDoc.Save(xmlPath)



        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    ''--------------------------------------------------------
    ''  機能    ：XML文書を書き込み
    ''  引数    ：１．パス指定
    ''  　　　  ：２．なかった場合のデフォルト値
    ''  戻り値  ：検索結果
    ''  作成日  ：
    ''--------------------------------------------------------
    'Public Shared Sub PBS_AppendXML(ByVal xmlPath As String, ByVal prmElement As String, ByVal prmValue As String)

    '    Try
    '        Dim xmlDoc As New System.Xml.XmlDocument
    '        xmlDoc.Load(xmlPath)
    '        '要素を追加する
    '        Dim xmlRoot As System.Xml.XmlElement = xmlDoc.DocumentElement
    '        Dim xmlEle As System.Xml.XmlElement = xmlRoot.Item(prmElement)
    '        Dim xmlValue As System.Xml.XmlText
    '        '存在チェック
    '        Dim xmlList As System.Xml.XmlNodeList = xmlDoc.GetElementsByTagName(prmElement)
    '        If xmlList.Count > 0 Then
    '            '削除
    '            xmlRoot.RemoveChild(xmlEle)
    '        End If
    '        '追加
    '        xmlEle = xmlDoc.CreateElement(prmElement)
    '        xmlValue = xmlDoc.CreateTextNode(prmValue)
    '        xmlRoot.AppendChild(xmlEle)
    '        xmlEle.AppendChild(xmlValue)


    '        xmlDoc.Save(xmlPath)

    '    Catch ex As Exception
    '        Throw ex
    '    End Try
    'End Sub
#End Region
#Region "XMLファイル読み込み"
    '--------------------------------------------------------
    '  機能    ：XML文書を読み込み
    '  引数    ：１．パス指定
    '  　　　  ：２．なかった場合のデフォルト値
    '  戻り値  ：検索結果
    '  作成日  ：
    '--------------------------------------------------------
    Public Shared Function PB_ReadXML(ByVal Path As String, ByVal DefaultVal As String, ByVal xmlPath As String) As String

        Try
            Dim xmlDoc As XmlDocument = New XmlDocument
            xmlDoc.Load(xmlPath)
            Dim list As XmlNodeList = xmlDoc.SelectNodes(Path)
            Dim node As XmlNode
            If list.Count <= 0 Then
                Return DefaultVal
            End If

            node = list.Item(0)

            Return node.InnerText

        Catch ex As Exception
            Throw ex
        End Try
    End Function

#End Region

#Region "XMLファイル読み込み"
    '--------------------------------------------------------
    '  機能    ：XML文書を読み込み
    '  引数    ：１．パス指定
    '  　　　  ：２．なかった場合のデフォルト値
    '  戻り値  ：検索結果
    '  作成日  ：
    '--------------------------------------------------------
    Public Shared Function PB_ReadXmlNodeList(ByVal Path As String, ByVal xmlPath As String) As XmlNodeList

        Dim ary As New ArrayList

        Try
            Dim xmlDoc As XmlDocument = New XmlDocument
            xmlDoc.Load(xmlPath)
            Dim list As XmlNodeList = xmlDoc.SelectNodes(Path)
            Dim node As XmlNode = Nothing
            If list.Count <= 0 Then
                Return Nothing
            End If

            node = list.Item(0)


            Return node.ChildNodes

        Catch ex As Exception
            Throw ex
        End Try
    End Function


#End Region
#Region "個別呼込み"

#End Region

#End Region

#Region "★【データ型変換メッソド】"

    ''---------------------------------------------------------------------
    '' 機能    ：INPUTMAN Date日付セット(無しの時：クリア)
    '' 引数    ：1. Value(日付：YYYY/MM/DD(OracleDefaultValue), YY/MM/DD)
    ''           2. DateController(セットするコントローラー)
    '' 戻り値  ：無し
    '' 作成日  ：2005.12.20 黄
    ''---------------------------------------------------------------------
    'Public Shared Sub PBS_SetIMDate(ByVal objVal As Object, ByVal editDate As IM.Date)

    '    If IsNothing(objVal) OrElse IsDBNull(objVal) OrElse CStr(objVal) = "" Then
    '        editDate.Clear()
    '    Else
    '        editDate.Value = ToDateTimeEx(CStr(objVal), _
    '            System.Globalization.CultureInfo.CurrentCulture)
    '    End If
    'End Sub
    ''' <summary>
    ''' 値変換(String)
    ''' </summary>
    ''' <param name="objVal">変換する値</param>
    ''' <param name="rtnValue">空値の場合の戻り値</param>
    ''' <returns>String戻値</returns>
    ''' <remarks></remarks>
    Public Shared Function PBCStr(ByVal objVal As Object, _
                                  Optional ByVal rtnValue As String = "") As String

        If IsNothing(objVal) OrElse IsDBNull(objVal) OrElse CStr(objVal).Equals("") Then
            Return rtnValue
        Else
            Return CStr(objVal)
        End If

    End Function
    ''' <summary>
    ''' 値変換(Integer)
    ''' </summary>
    ''' <param name="objVal">変換する値</param>
    ''' <param name="rtnValue">空値の場合の戻り値</param>
    ''' <returns>Integer戻り値</returns>
    ''' <remarks></remarks>
    Public Shared Function PBCint(ByVal objVal As Object, _
                                  Optional ByVal rtnValue As Integer = 0) As Integer
        If IsNothing(objVal) OrElse IsDBNull(objVal) OrElse CStr(objVal).Equals("") Then
            Return rtnValue
        ElseIf IsNumeric(objVal) Then
            Return CInt(objVal)
        End If
    End Function

    ''' <summary>
    ''' 値変換(Boolean)
    ''' </summary>
    ''' <param name="objVal">変換する値</param>
    ''' <param name="rtnValue">空値の場合の戻り値</param>
    ''' <returns>Boolean戻り値</returns>
    ''' <remarks></remarks>
    Public Shared Function PBCBool(ByVal objVal As Object, _
                                  Optional ByVal rtnValue As Boolean = False) As Boolean
        If IsNothing(objVal) OrElse IsDBNull(objVal) OrElse CStr(objVal).Equals("") Then
            Return rtnValue
        Else
            Return CBool(objVal)
        End If
    End Function
    ''' <summary>
    ''' 値変換(Byte)
    ''' </summary>
    ''' <param name="objVal">変換する値</param>
    ''' <param name="rtnValue">空値の場合の戻り値</param>
    ''' <returns>Byte戻り値</returns>
    ''' <remarks></remarks>
    Public Function PBCbyt(ByVal objVal As Object, _
                                  Optional ByVal rtnValue As Byte = 0) As Byte
        If IsNothing(objVal) OrElse IsDBNull(objVal) OrElse CStr(objVal).Equals("") Then
            Return rtnValue
        ElseIf IsNumeric(objVal) Then
            Return CByte(objVal)
        End If
    End Function
    ''' <summary>
    ''' 値変換(Long)
    ''' </summary>
    ''' <param name="objVal">変換する値</param>
    ''' <param name="rtnValue">空値の場合の戻り値</param>
    ''' <returns>Long戻り値</returns>
    ''' <remarks></remarks>
    Public Shared Function PBClng(ByVal objVal As Object, _
                                  Optional ByVal rtnValue As Long = 0) As Long
        If IsNothing(objVal) OrElse IsDBNull(objVal) OrElse CStr(objVal).Equals("") Then
            Return rtnValue
        ElseIf IsNumeric(objVal) Then
            Return CLng(objVal)
        End If
    End Function

    ''' <summary>
    ''' 値変換(Decimal)
    ''' </summary>
    ''' <param name="objVal">変換する値</param>
    ''' <param name="rtnValue">空値の場合の戻り値</param>
    ''' <returns>Decimal戻り値</returns>
    ''' <remarks></remarks>
    Public Shared Function PBCdec(ByVal objVal As Object, _
                                  Optional ByVal rtnValue As Decimal = 0D) As Decimal
        If IsNothing(objVal) OrElse IsDBNull(objVal) OrElse CStr(objVal).Equals("") Then
            Return rtnValue
        ElseIf IsNumeric(objVal) Then
            Return CDec(objVal)
        End If
    End Function
    ''' <summary>
    ''' 値変換(Double)
    ''' </summary>
    ''' <param name="objVal">変換する値</param>
    ''' <param name="rtnValue">空値の場合の戻り値</param>
    ''' <returns>Double戻り値</returns>
    ''' <remarks></remarks>
    Public Shared Function PBCdbl(ByVal objVal As Object, _
                                  Optional ByVal rtnValue As Double = 0) As Double
        If IsNothing(objVal) OrElse IsDBNull(objVal) OrElse CStr(objVal).Equals("") Then
            Return rtnValue
        ElseIf IsNumeric(objVal) Then
            Return CDbl(objVal)
        End If
    End Function
    ''' <summary>
    '''  NULLチェック
    ''' </summary>
    ''' <param name="objVal">チェックする値</param>
    ''' <param name="bln">空値以外の場合の戻り値</param>
    ''' <returns>True：Null　False:Nukk以外</returns>
    ''' <remarks></remarks>
    Public Shared Function PB_ChkNUll(ByVal objVal As Object, _
                                   Optional ByVal bln As Boolean = False) As Boolean

        If IsNothing(objVal) OrElse IsDBNull(objVal) OrElse CStr(objVal).Equals("") Then
            Return True
        Else
            Return bln
        End If
    End Function

    '---------------------------------------------------------------------
    '  機能    ：ヌル(Nothing,DBNull, "")チェック
    '  引数    ：1．Object, 2.Byte(文字形・数字形区分) ← 0:数字 9:未設定(文字)
    '  戻り値  ：String(ヌルの場合"NULL"返す、ない場合は''をつける)
    '            Integer(そのまま返す)
    '  作成日  ：2006.01.17  黄
    '---------------------------------------------------------------------
    '文字形でも数字が代入された時は数字形に認識してるので。。。。
    '落ちる可能性が高まってるので。
    Public Shared Function PBFSTR_SetQTT(ByVal objVal As Object, _
                                    Optional ByVal byKBN As Byte = 9) As String

        If IsNothing(objVal) OrElse IsDBNull(objVal) OrElse CStr(objVal) = "" Then
            Return "NULL"
        Else
            Select Case byKBN
                Case 9    '文字形の場合
                    If IsDate(objVal) Then '20070308_1 どちらも遅い
                        'If PB_IsDate(CStr(objVal)) Then
                        Return "'" & CStr(objVal) & "'"
                    ElseIf IsNumeric(objVal) Then
                        Return "'" & addQuot(CStr(objVal)) & "'"
                    Else
                        Return "'" & addQuot(CStr(objVal)) & "'"
                    End If

                Case 0      '数字形の場合
                    If IsNumeric(objVal) Then
                        Return CStr(objVal)
                    Else
                        Return CStr(objVal)
                    End If
                Case Else
                    Return CStr(objVal)

            End Select
        End If
    End Function
    '---------------------------------------------------------------------
    '  機能    ：ヌル(Nothing,DBNull, "")をDecimalに変換
    '  引数    ：１．Object, (２．Decimal )
    '  戻り値  ：Decimal
    '  作成日  ：2006.07.25  F.Nishida
    '---------------------------------------------------------------------
    Public Shared Function PBCsng(ByVal objVal As Object, _
                                   Optional ByVal rtnValue As Single = 0) As Single
        If IsNothing(objVal) OrElse IsDBNull(objVal) OrElse CStr(objVal) = "" Then
            Return rtnValue
        ElseIf IsNumeric(objVal) Then
            Return CSng(objVal)
        End If
    End Function
    '---------------------------------------------------------------------
    '  機能    ：ヌル(Nothing,DBNull, "")をDateに変換
    '  引数    ：１．Object, (２．Data )
    '  戻り値  ：Date
    '  作成日  ：2007.08.01 加藤
    '---------------------------------------------------------------------
    Public Shared Function PBCDate(ByVal objVal As Object, _
                                   Optional ByVal rtnValue As Date = Nothing) As Date

        If IsNothing(objVal) OrElse IsDBNull(objVal) OrElse CStr(objVal) = "" Then
            Return rtnValue
        ElseIf IsDate(objVal) Then
            Return CDate(objVal)
        Else
            Return Nothing
        End If
    End Function

    '---------------------------------------------------------------------
    '  機能    ：ヌル(Nothing,DBNull, "")をDateに変換
    '  引数    ：１．Object, (２．Data )
    '  戻り値  ：Date
    '  作成日  ：2007.08.01 加藤
    '---------------------------------------------------------------------
    Public Shared Function PBCDateTime(ByVal objVal As Object, _
                                   Optional ByVal rtnValue As Date = Nothing) As Date

        If IsNothing(objVal) OrElse IsDBNull(objVal) OrElse CStr(objVal) = "" Then
            Return Nothing
        ElseIf IsDate(objVal) Then
            Return DateTime.Parse(PBCStr(objVal))
        End If
    End Function
#Region "シングルコーテーション追加"
    '------------------------------------------------------------------
    ' 機能         : シングルコーテーション追加
    '
    ' 返り値       : 正常終了 = 変換後の文字列
    '                異常終了 = ""
    '
    ' 引き数       : (IN) strVal  入力文字列
    '
    ' 機能説明     : 文字列中のシングルコーテーションを検索しシングルコーテーションを二重にする
    '
    ' 備考         : 2005.11.07 駒形
    '
    '------------------------------------------------------------------
    Public Shared Function addQuot(ByVal strVal As String) As String

        Dim intLocation As Integer        'シングルコーテーションの位置
        Dim strOutputVal As String
        Dim strInputVal As String

        ''出力値変数を初期化
        strOutputVal = ""

        ''入力値を入力値変数に移送
        strInputVal = strVal

        ''入力値変数のシングルコーテーションの位置を検査
        intLocation = InStr(strInputVal, "'")

        ''シングルコーテーションの位置が０より大きい間ループする。
        While intLocation > 0
            ''出力値変数に入力値変数のシングルコーテーションの位置までとシングルコーテーションを出力。
            strOutputVal = strOutputVal & Left$(strInputVal, intLocation) & "'"
            ''入力値変数から出力値変数に出力した文字列を削除する。
            strInputVal = Mid$(strInputVal, intLocation + 1, Len(strInputVal) - intLocation)
            ''入力値変数のシングルコーテーションの位置を検査
            intLocation = InStr(strInputVal, "'")
            ''ループ終了
        End While

        ''戻り値を設定
        Return strOutputVal & strInputVal

    End Function
#End Region


#End Region

#Region "★【データ取得関連】"
#Region "データセットより該当データを取得する"
    Public Shared Function GetOneItemData(ByVal dtTbl As DataTable, ByVal FldName As String, ByVal szWhere As String) As String
        Dim dtView As DataView

        Try
            dtView = New DataView(dtTbl, szWhere, "", DataViewRowState.CurrentRows)
            If dtView.Count <= 0 Then
                Return ""
            Else
                Return PBCStr(dtView.Item(0)(FldName))
            End If

        Catch ex As Exception
            'SkyLog.Error(ex.Message)
            Return ""
        End Try

    End Function
#End Region
#Region "データセットより問い合わせ結果(View)を取得する"
    Public Shared Function GetResultDataView(ByVal dtTbl As DataTable, ByVal szWhere As String) As DataView
        Dim dtView As DataView

        Try
            dtView = New DataView(dtTbl, szWhere, "", DataViewRowState.CurrentRows)
            If dtView.Count <= 0 Then
                Return Nothing
            Else
                Return dtView
            End If

        Catch ex As Exception
            'SkyLog.Error(ex.Message)
            Return Nothing
        End Try

    End Function
#End Region '20090729_1
#Region "データセット内よりデータの存在確認"
    Public Shared Function ExistValue(ByVal dtTbl As DataTable, ByVal szWhere As String) As Boolean
        Dim dtView As DataView

        Try
            dtView = New DataView(dtTbl, szWhere, "", DataViewRowState.CurrentRows)
            If dtView.Count <= 0 Then
                Return False
            Else
                Return True
            End If

        Catch ex As Exception
            'SkyLog.Error(ex.Message)
            Throw ex
        End Try

    End Function
#End Region
#Region "ArrayListをデータテーブルへ変換する"
    Public Shared Function GetdtFromArrayList(ByVal prmAryList As ArrayList, ByVal boHeader As Boolean) As DataTable
        Dim dt As New DataTable
        Dim dtRow As DataRow
        Dim arydt As New ArrayList

        arydt = prmAryList

        For i As Integer = 0 To prmAryList.Count - 1

            If boHeader Then
                '*---------------------
                'CSVの1行目をヘッダーに
                '*---------------------
                If i.Equals(0) Then
                    '
                    Dim fields As String()
                    fields = CType(arydt.Item(i), String())

                    For j As Integer = 0 To fields.Length - 1
                        Dim headName As String = fields(j).ToString
                        dt.Columns.Add(headName)
                    Next
                Else
                    ''dataTable の Row に1行分追加.
                    dtRow = dt.NewRow()
                    dtRow.ItemArray = CType(prmAryList.Item(i), [Object]())
                    dt.Rows.Add(dtRow)
                End If

            Else
                '*---------------------
                'すべてを明細扱い
                '*---------------------
                If i.Equals(0) Then
                    '
                    Dim fields As String()
                    fields = CType(arydt.Item(i), String())

                    For j As Integer = 0 To fields.Length - 1
                        dt.Columns.Add((j + 1).ToString)
                    Next
                End If


                ''dataTable の Row に1行分追加.
                dtRow = dt.NewRow()
                dtRow.ItemArray = CType(prmAryList.Item(i), [Object]())
                dt.Rows.Add(dtRow)
            End If

        Next

        Return dt

    End Function
#End Region '20100326_1
#End Region

#Region "★【消費税・端数処理】"
    '#Region "税率取得"
    '    ''' <summary>
    '    ''' 税率取得
    '    ''' </summary>
    '    ''' <param name="con"></param>
    '    ''' <param name="prmZeitp"></param>
    '    ''' <param name="prmBaseDate">基準日（yyyy/MM/dd形式）</param>
    '    ''' <returns></returns>
    '    ''' <remarks></remarks>
    '    Public Shared Function getZeiRt(con As uniConnection, prmZeitp As Integer, prmBaseDate As String,
    '                                                                        Optional ByVal tran As uniTransaction = Nothing) As Decimal
    '        Try

    '            Dim szSQL As String = ""

    '            szSQL = ""
    '            szSQL += " select ZEI_RT from M_ZEI  "
    '            szSQL += " where 1=1 "
    '            szSQL += " and  zei_tp = " & prmZeitp
    '            szSQL += " and aply_dt = (select max(aply_dt) from m_zei zei_max where zei_max.zei_tp=" & prmZeitp & "and TO_CHAR(zei_max.aply_dt,'YYYY/MM/DD') <='" & prmBaseDate & "') "

    '            Dim ary As New ArrayList
    '            ary = getAryDataDB(con, szSQL, tran)
    '            If ary.Count.Equals(0) Then
    '                Return 0
    '            Else
    '                If PBCint(ary.Item(0)).Equals(0) Then
    '                    Return 0
    '                Else
    '                    Return PBCdec(PBCint(ary.Item(0)) / 100)
    '                End If
    '            End If
    '        Catch ex As Exception
    '            'Throw ex
    '            Return 0
    '        End Try
    '    End Function
    '#End Region
    '#Region "税額取得"
    '    ''' <summary>
    '    ''' 消費税金額の取得
    '    ''' </summary>
    '    ''' <param name="prmItemKn">商品金額</param>
    '    ''' <param name="prmZeiRt">税率</param>
    '    ''' <param name="prmZeikKbn">税込区分</param>
    '    ''' <param name="prmHasu">端数処理区分</param>
    '    ''' <returns></returns>
    '    ''' <remarks></remarks>
    '    Public Shared Function GetZeiKn(ByVal prmItemKn As Integer, prmZeiRt As Decimal, prmZeikKbn As SystemConst.ZEIK, _
    '                                                    Optional ByVal prmHasu As SystemConst.Round = Round.Down, Optional prmKazeiKbn As SystemConst.KAZEI = KAZEI.KAZEI
    '                                                    ) As Decimal

    '        Try

    '            Dim rtnZeiKn As Decimal = 0

    '            If prmKazeiKbn = KAZEI.HIKAZEI Then
    '                rtnZeiKn = 0
    '            Else
    '                Select Case prmZeikKbn

    '                    Case SystemConst.ZEIK.ZEIOUT
    '                        ''税抜き (商品金額＊税率)端数処理
    '                        rtnZeiKn = SystemUtil.doCalHASU(PBCdec(prmItemKn * prmZeiRt), prmHasu)

    '                    Case SystemConst.ZEIK.ZEIIN
    '                        ''消費税 金額*(5/105) '20090907_1
    '                        ''税込み :商品金額ー(商品金額/1+税率)端数処理
    '                        ''rtnZeiKn = prmItemKn - SystemUtil.PBF_CalHASU(PBCdec(prmItemKn / (ZeiRt + 1)), prmHasu)
    '                        'rtnZeiKn = SystemUtil.PBF_CalHASU(prmItemKn - PBCdec(prmItemKn / (ZeiRt + 1)), prmHasu)
    '                        rtnZeiKn = SystemUtil.doCalHASU(PBCdec(prmItemKn * prmZeiRt / (prmZeiRt + 1)), prmHasu)

    '                    Case Else
    '                        rtnZeiKn = 0
    '                End Select
    '            End If


    '            Return rtnZeiKn

    '        Catch ex As Exception
    '            Throw ex
    '        End Try
    '    End Function
    '#End Region
#Region "端数処理区分"
    Public Shared Function doCalHASU(ByVal value As Decimal, _
                                Optional ByVal bytHASU As Round = Round.Half) As Decimal
        Dim tempValue As Decimal = value
        If value < 0 Then
            tempValue = 0 - value
        End If
        Dim result As Decimal
        Select Case bytHASU
            Case Round.UP       '切上げ
                result = Decimal.Truncate(tempValue)
                If result <> tempValue Then
                    result += 1D
                End If
            Case Round.Down      '切捨て
                result = Decimal.Truncate(tempValue)

            Case Round.Half     '四捨五入
                result = Decimal.Truncate(tempValue + 0.5D)
            Case Else
                Return value
        End Select

        If value < 0 Then
            result = 0 - result
        End If

        Return result
    End Function
#End Region
#Region "端数処理区分(Double)"
    ''' <summary>
    ''' 端数処理(Double)
    ''' </summary>
    ''' <param name="dblValue">丸め対象値</param>
    ''' <param name="intDigits">戻り値の有効桁数の精度</param>
    ''' <param name="bytHASU">端数処理区分</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function doCalHASU2(ByVal dblValue As Double, _
                                   Optional ByVal intDigits As Integer = 0, _
                                   Optional ByVal bytHASU As Round = Round.Half) As Double
        Dim dblSquare As Double = Math.Pow(10, intDigits)
        Select Case bytHASU
            Case Round.UP  '切上げ
                If dblValue > 0 Then
                    Return Math.Ceiling(dblValue * dblSquare) / dblSquare
                Else
                    Return Math.Floor(dblValue * dblSquare) / dblSquare
                End If

            Case Round.Down  '切捨て
                If dblValue > 0 Then
                    Return Math.Floor(dblValue * dblSquare) / dblSquare
                Else
                    Return Math.Ceiling(dblValue * dblSquare) / dblSquare
                End If

            Case Round.Half  '四捨五入
                If dblValue > 0 Then
                    Return Math.Floor((dblValue * dblSquare) + 0.5) / dblSquare
                Else
                    Return Math.Ceiling((dblValue * dblSquare) - 0.5) / dblSquare
                End If
            Case Else
                Return dblValue
        End Select
    End Function
#End Region

#End Region

#Region "★【日付関連】"
#Region "日付かどうかを調べる"
    ''' <summary>
    ''' "日付かどうかを調べる
    ''' </summary>
    ''' <param name="szObj">検証する値</param>
    ''' <returns>True：日付です　False：日付でない</returns>
    ''' <remarks></remarks>
    Public Shared Function PB_IsDate(ByVal szObj As String) As Boolean


        'DateTimeに変換できるか確かめる
        Try
            DateTime.Parse(szObj)
            Return True
        Catch
            Return False
        End Try
    End Function
#End Region
#Region "日付の生成(年月)"
    ''' <summary>
    ''' 日付の生成(年月)
    ''' </summary>
    ''' <param name="inY">年</param>
    ''' <param name="inM">月</param>
    ''' <param name="inD">日</param>
    ''' <param name="inAddDate">加算値</param>
    ''' <param name="interval">インターバル</param>
    ''' <param name="szFormat">戻り値のフォーマット</param>
    ''' <returns>日付(文字)</returns>
    ''' <remarks></remarks>
    Public Overloads Shared Function PBGetDate(ByVal inY As Integer, ByVal inM As Integer, _
                            Optional ByVal inD As Integer = 1, Optional ByVal inAddDate As Integer = 0, _
                            Optional ByVal interval As DateInterval = DateInterval.Day, Optional ByVal szFormat As String = "yyyy/MM/dd") As String

        Dim dtDate As Date
        Dim szDate As String

        Try


            szDate = inY & "/" & inM & "/" & inD
            dtDate = CDate(DateValue(szDate).ToString("yyyy/MM/dd"))
            ''InterValセット
            dtDate = DateAdd(interval, inAddDate, dtDate)
            ''フォーマット
            Return dtDate.ToString(szFormat)

        Catch ex As Exception
            Throw ex
        End Try
    End Function
#End Region
#Region "日付の生成(年月日)"
    ''' <summary>
    '''  日付の生成(年月日)
    ''' </summary>
    ''' <param name="szYMD">日付(yyyyMMdd形式)</param>
    ''' <param name="inAddDate">加算値</param>
    ''' <param name="interval">インターバル</param>
    ''' <param name="szFormat">戻り値のフォーマット</param>
    ''' <returns>日付(文字)</returns>
    ''' <remarks></remarks>
    Public Overloads Shared Function getDate(ByVal szYMD As String, Optional ByVal inAddDate As Integer = 0, _
                            Optional ByVal interval As DateInterval = DateInterval.Day, Optional ByVal szFormat As String = "yyyy/MM/dd") As String

        Dim dtDate As Date
        Dim szDate As String

        Try

            If szYMD.Equals("00010101") Then
                Return ""
            End If


            szDate = GetWantedByte(szYMD, 0, 4) & "/" & GetWantedByte(szYMD, 4, 2) & "/" & GetWantedByte(szYMD, 6, 2)
            dtDate = CDate(DateValue(szDate).ToString("yyyy/MM/dd"))
            ''InterValセット
            dtDate = DateAdd(interval, inAddDate, dtDate)
            ''フォーマット
            Return dtDate.ToString(szFormat)

        Catch ex As Exception
            Return ""
        End Try
    End Function
    ''' <summary>
    '''  日付の生成(年月日)
    ''' </summary>
    ''' <param name="prmDate">日付</param>
    ''' <param name="inAddDate">加算値</param>
    ''' <param name="interval">インターバル</param>
    ''' <param name="szFormat">戻り値のフォーマット</param>
    ''' <returns>日付(文字)</returns>
    ''' <remarks></remarks>
    Public Overloads Shared Function getDate(ByVal prmDate As Date, Optional ByVal inAddDate As Integer = 0, _
                            Optional ByVal interval As DateInterval = DateInterval.Day, Optional ByVal szFormat As String = "yyyy/MM/dd") As String

        Dim dtDate As Date

        Try

            If prmDate = Nothing Then
                Return ""
            End If

            ''InterValセット
            dtDate = DateAdd(interval, inAddDate, prmDate)
            ''フォーマット
            Return dtDate.ToString(szFormat)

        Catch ex As Exception
            Return ""
        End Try
    End Function
#End Region
#Region "日付の生成(yyyy/MM/dd→yyyyMMdd)"
    ''' <summary>
    ''' 日付の書式変換
    ''' </summary>
    ''' <param name="szYMD">変換する文字列</param>
    ''' <returns>変換後の日付文字列</returns>
    ''' <remarks>yyyy/MM/dd→yyyyMMdd</remarks>
    Public Overloads Shared Function PBGetCngDate(ByVal szYMD As String) As String

        Dim szDate As String = ""
        Dim dtDate As Date
        Try

            '日付か否かﾁｪｯｸ
            dtDate = CDate(DateValue(szYMD))


            szDate += GetWantedByte(szYMD, 0, 4)
            szDate += GetWantedByte(szYMD, 5, 2)
            szDate += GetWantedByte(szYMD, 8, 2)

            Return szDate

        Catch ex As Exception
            Return ""
        End Try
    End Function
#End Region
#Region "日付変換(yyyyMMdd→yyyy/MM/dd)"
    ''' <summary>
    ''' 日付変換(yyyyMMdd→yyyy/MM/dd)"
    ''' </summary>
    ''' <param name="szYMD"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overloads Shared Function getCngDate(ByVal szYMD As String, Optional formatString As String = "yyyy/MM/dd") As String

        Try
            Dim result As DateTime

            If DateTime.TryParseExact(szYMD, "yyyyMMdd", Nothing, DateTimeStyles.None, result) Then
                Return result.ToString(formatString)
            Else
                Return ""
            End If

        Catch ex As Exception
            Return ""
        End Try

    End Function
#End Region
#Region "月末を求める。"
    ''' <summary>
    ''' 月末を求める
    ''' </summary>
    ''' <param name="prmDate">yyyy/MM/dd形式</param>
    ''' <param name="szFormat">戻り値となる日付の書式</param>
    ''' <returns>日付の文字列</returns>
    ''' <remarks></remarks>
    Public Shared Function getLastDate(ByVal prmDate As String, Optional ByVal szFormat As String = "yyyy/MM/dd") As String

        Dim dtDate As Date

        Try



            dtDate = CDate(DateValue(prmDate).ToString("yyyy/MM/dd"))
            ''InterValセット'1ヶ月後
            dtDate = DateAdd(DateInterval.Month, 1, dtDate)
            '-1日で月末をセット
            dtDate = DateAdd(DateInterval.Day, -1, dtDate)

            ''フォーマット
            Return dtDate.ToString(szFormat)

        Catch ex As Exception
            Throw ex
        End Try
    End Function
    ''' <summary>
    ''' 月末を求める
    ''' </summary>
    ''' <param name="inY">年</param>
    ''' <param name="inM">月</param>
    ''' <param name="szFormat">戻り値となる日付の書式</param>
    ''' <returns>日付の文字列</returns>
    ''' <remarks></remarks>
    Public Shared Function PB_GetLastDate(ByVal inY As Integer, ByVal inM As Integer, Optional ByVal szFormat As String = "yyyy/MM/dd") As String

        Dim dtDate As Date
        Dim szDate As String

        Try


            szDate = inY & "/" & inM & "/" & "01"
            dtDate = CDate(DateValue(szDate).ToString("yyyy/MM/dd"))
            ''InterValセット'1ヶ月後
            dtDate = DateAdd(DateInterval.Month, 1, dtDate)
            '-1日で月末をセット
            dtDate = DateAdd(DateInterval.Day, -1, dtDate)

            ''フォーマット
            Return dtDate.ToString(szFormat)

        Catch ex As Exception
            Throw ex
        End Try
    End Function
#End Region
#Region "残り時間を求める"
    ''' <summary>
    ''' 残り時間を求める
    ''' </summary>
    ''' <param name="dtNow">現在の時刻</param>
    ''' <param name="dtTarget">ターゲットとなる時刻</param>
    ''' <param name="boGuide">戻り値の設定(True:ガイド False:時刻)</param>
    ''' <returns>String(時刻orガイドorﾌﾞﾗﾝｸ)</returns>
    ''' <remarks>
    ''' <para>boGuide = Ture の場合(ガイド)</para>
    ''' <para>1分以下の場合XX秒で戻す（0～59秒）</para>
    ''' <para>1分以上で1時間以下の場合、XX分で戻す（1～59分）</para>
    ''' <para>1時間以上で一日以下の場合、XX時間で戻す（1～24時間）</para>
    ''' <para>一日以上の場合、XX日で戻す（1日～）</para>
    ''' </remarks>
    Public Shared Function PBGetTimeRemit(ByVal dtNow As Date, ByVal dtTarget As Date, Optional ByVal boGuide As Boolean = True) As String

        Dim intTime As Integer
        Dim intHour As Integer = 0
        Dim intMinute As Integer = 0
        Dim intSecond As Integer = 0


        '現在の時刻からターゲットとなる時刻を引く
        intTime = PBCint(dtTarget.Subtract(dtNow).TotalSeconds)

        Select Case boGuide
            Case True
                '-----------------
                'ガイド有りの場合
                '-----------------
                If 0 < intTime And intTime < 60 Then
                    '1分以下の場合、XX秒で戻す
                    Return CStr(intTime) & "秒"

                ElseIf 60 <= intTime And intTime < 3600 Then
                    '1分以上で1時間以下、XX分で戻す
                    Return CStr(CInt(intTime / 60)) & "分"

                ElseIf 3600 <= intTime And intTime < 86400 Then
                    '1時間以上で一日以下の場合、XX時間で戻す
                    Return CStr(CInt(intTime / 3600)) & "時間"

                ElseIf 86400 <= intTime Then
                    '一日以上の場合、XX日で戻す
                    Return CStr(CInt(intTime / 86400)) & "日"

                Else

                    Return ""

                End If

            Case Else
                ''-----------------
                ''ガイド無しの場合
                ''-----------------
                If intTime > 0 Then

                    Return CStr(intTime)

                Else

                    Return ""

                End If

        End Select

    End Function

#End Region
#Region "和暦を求める"
    ''' <summary>
    ''' 和暦を求める
    ''' </summary>
    ''' <param name="prmDate">yyyy/MM/dd形式</param>
    ''' <param name="szFormat"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function getWareki(ByVal prmDate As Date, Optional ByVal szFormat As String = "ggyy/MM/dd") As String

        Dim szDate As String

        Try

            Dim culture As Globalization.CultureInfo = New Globalization.CultureInfo("ja-JP")
            culture.DateTimeFormat.Calendar = New System.Globalization.JapaneseCalendar

            szDate = prmDate.ToString(szFormat, culture)

            Return szDate


        Catch ex As Exception
            Return ""
        End Try

    End Function
    ''' <summary>
    ''' 和暦を求める
    ''' </summary>
    ''' <param name="prmDate">yyyy/MM/dd形式</param>
    ''' <param name="szFormat"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function getWareki(ByVal prmDate As String, Optional ByVal szFormat As String = "ggyy/MM/dd") As String

        Dim dtDate As Date
        Dim szDate As String

        dtDate = PBCDate(prmDate)

        Try

            Dim culture As Globalization.CultureInfo = New Globalization.CultureInfo("ja-JP")
            culture.DateTimeFormat.Calendar = New System.Globalization.JapaneseCalendar

            szDate = dtDate.ToString(szFormat, culture)

            Return szDate


        Catch ex As Exception
            Return ""
        End Try

    End Function
#End Region
#End Region

#Region "★ファイル関連"
#Region "ファイル関連"

#Region "保存ダイアログ"
    ''' <summary>
    ''' ファイル保存ダイアログボックスの表示
    ''' </summary>
    ''' <param name="szDefaltFileName">ファイル名</param>
    ''' <param name="szFileDir">ファイルディレクトリ</param>
    ''' <param name="szDefaultPath">初期ディレクトリ</param>
    ''' <param name="Filter">ファイル種類(初期値：XLS)</param>
    ''' <returns>Boolean(成功・失敗)</returns>
    ''' <remarks></remarks>
    Public Shared Function SaveFileDialog(ByRef szDefaltFileName As String, ByRef szFileDir As String, _
                               Optional ByVal szDefaultPath As String = "", Optional ByVal Filter As FILEKIND = FILEKIND.XLS) As Boolean

        Dim sfd As New SaveFileDialog
        'はじめのファイル名を指定する
        sfd.FileName = szDefaltFileName
        'はじめに表示されるフォルダを指定する
        If Not szDefaultPath.Equals("") Then
            sfd.InitialDirectory = szDefaultPath
        End If

        '[ファイルの種類]に表示される選択肢を指定する
        Select Case Filter
            Case FILEKIND.XLS
                sfd.Filter = "Excelファイル (*.xls)|*.xls"
            Case FILEKIND.XLSX
                sfd.Filter = "Excelファイル (*.xlsx)|*.xlsx"
            Case FILEKIND.TXT
                sfd.Filter = "テキストファイル (*.txt)|*.txt"
            Case FILEKIND.CSV
                sfd.Filter = "CSVファイル (*.csv)|*.csv"
            Case FILEKIND.PDF
                sfd.Filter = "PDFファイル (*.pdf)|*.pdf"
            Case FILEKIND.ETC
                sfd.Filter = "すべてのファイル (*.*)|*.*"
        End Select

        '[ファイルの種類]ではじめに
        '「すべてのファイル」が選択されているようにする
        sfd.FilterIndex = 2
        'タイトルを設定する
        sfd.Title = "保存先のファイルを選択してください"
        'ダイアログボックスを閉じる前に現在のディレクトリを復元するようにする
        sfd.RestoreDirectory = True
        '既に存在するファイル名を指定したとき警告する
        'デフォルトでTrueなので指定する必要はない
        sfd.OverwritePrompt = True
        '存在しないパスが指定されたとき警告を表示する
        'デフォルトでTrueなので指定する必要はない
        sfd.CheckPathExists = True
        '拡張子が指定されない場合に拡張子を設定するようにする
        'デフォルトでTrueなので指定する必要はない
        sfd.AddExtension = True

        'ダイアログを表示する
        If sfd.ShowDialog() = DialogResult.OK Then
            szFileDir = System.IO.Path.GetDirectoryName(sfd.FileName) & "\"
            szDefaltFileName = System.IO.Path.GetFileName(sfd.FileName)
            Return True
        Else
            Return False
        End If
    End Function
#End Region
#Region "ファイルOpenダイアログ"
    ''' <summary>
    ''' ファイルOpenダイアログボックスの表示
    ''' </summary>
    ''' <param name="szDefaultPath">初期ディレクトリ(省略時はC:\)</param>
    ''' <param name="Filter">ファイル種類(初期値：XLS)</param>
    ''' <param name="title">フォームタイトル</param>
    ''' <returns>ファイル名</returns>
    ''' <remarks>キャンセル時は空値を戻す</remarks>
    Public Shared Function OpenFileDialog(Optional ByVal szDefaultPath As String = "", _
                                         Optional ByVal Filter As FILEKIND = FILEKIND.XLS, _
                                          Optional ByVal FilterTxt As String = "", _
                                         Optional ByVal title As String = "") As String

        Using sfd As New OpenFileDialog
            'はじめのファイル名を指定する
            'sfd.FileName = szDefaltFileName
            'はじめに表示されるフォルダを指定する
            '20080201_1 Add
            If Not szDefaultPath.Equals("") Then
                sfd.InitialDirectory = szDefaultPath
            End If

            ''If szDefaultPath.Equals("") Then
            ''    sfd.InitialDirectory = "C:\"
            ''Else
            ''    sfd.InitialDirectory = szDefaultPath
            ''End If

            '[ファイルの種類]に表示される選択肢を指定する
            Select Case Filter
                Case FILEKIND.XLS
                    sfd.Filter = "Excelファイル (*.xls)|*.xls"
                Case FILEKIND.XLSX
                    sfd.Filter = "Excelファイル (*.xlsx)|*.xlsx"
                Case FILEKIND.TXT
                    sfd.Filter = "テキストファイル (*.txt)|*.txt"
                Case FILEKIND.CSV
                    sfd.Filter = "CSVファイル (*.csv)|*.csv"
                Case FILEKIND.ETC
                    sfd.Filter = FilterTxt
            End Select

            '[ファイルの種類]ではじめに
            '「すべてのファイル」が選択されているようにする
            sfd.FilterIndex = 2
            'タイトルを設定する
            If title.Equals("") Then
                sfd.Title = "保存先のファイルを選択してください"
            Else
                sfd.Title = title
            End If

            'ダイアログボックスを閉じる前に現在のディレクトリを復元するようにする
            sfd.RestoreDirectory = True
            '存在しないパスが指定されたとき警告を表示する
            'デフォルトでTrueなので指定する必要はない
            sfd.CheckPathExists = True

            'ダイアログを表示する
            If sfd.ShowDialog() = DialogResult.OK Then
                Return sfd.FileName
            Else
                Return ""
            End If
        End Using

    End Function
#End Region
#Region "フォルダダイアログ"
    ''' <summary>
    ''' フォルダの参照ダイアログ表示
    ''' </summary>
    ''' <param name="szDefaultPath">初期設定パス</param>
    ''' <param name="szTitle">ファイアログタイトル</param>
    ''' <returns>選択したフォルダ参照パス</returns>
    ''' <remarks>キャンセルの場合は空値</remarks>
    Public Shared Function FolderDialog(Optional ByVal szDefaultPath As String = "C:\", _
                                            Optional ByVal szTitle As String = "フォルダを参照してください") As String

        Dim szPath As String = ""
        Dim Dialog As New FolderBrowserDialog
        '初期参照Path
        Dialog.SelectedPath = szDefaultPath

        'ダイアログボックスに[新しいフォルダの作成]ボタンを表示しない場合は False 
        Dialog.ShowNewFolderButton = False
        'ダイアログタイトル
        Dialog.Description = szTitle

        If Dialog.ShowDialog() = DialogResult.OK Then
            'ファイルの取得
            szPath = Dialog.SelectedPath
        End If

        Return szPath

    End Function
#End Region
#Region "ファイル取得"
    ''' <summary>
    ''' 対象ディレクトリのファイル一覧を戻す
    ''' </summary>
    ''' <param name="szPath">対象ディレクトリパス</param>
    ''' <param name="arySerachPattarn">検索対象ファイル</param>
    ''' <returns>ファイルの一覧</returns>
    ''' <remarks></remarks>
    Public Shared Function GetFileList(ByVal szPath As String, Optional ByVal arySerachPattarn As ArrayList = Nothing) As ArrayList
        Dim aryextension As New ArrayList
        Dim szFile As String
        Dim aryFiles As New ArrayList


        If arySerachPattarn Is Nothing Then
            aryextension.Add("*.*")
        Else
            aryextension = arySerachPattarn
        End If

        'ファイルが存在すればそのまま返す
        If IO.File.Exists(szPath) Then
            aryFiles.Add(Path.GetFileName(szPath))
            Return aryFiles
        End If


        For i As Integer = 0 To aryextension.Count - 1
            'ファイル取得(トップディレクトリのみ)
            For Each szFile In Directory.GetFiles(szPath, PBCStr(aryextension.Item(i)), SearchOption.TopDirectoryOnly)
                aryFiles.Add(Path.GetFileName(szFile))
            Next
        Next


        Return aryFiles
    End Function
#End Region
#End Region
#End Region

#Region "★データ編集関連"
#Region "指定文字列のバイト長を戻す"
    Public Shared Function GetLengthASByte(ByVal prmVal As String) As Integer
        Return System.Text.Encoding.GetEncoding("Shift_JIS").GetByteCount(prmVal)

    End Function
#End Region
#Region "Mid関数のバイト版"
    '-----------------------------------------------------------------------------------
    '　機能　：文字数と位置をバイト数で指定して文字列を切り抜く
    '
    '　引数　：strVal(対象の文字列), 
    '          intStart(切り抜き開始位置。
    '                   全角文字を分割するよう位置が指定された場合、戻り値の文字列の先頭は意味不明の半角文字となる), 
    '          intLength(切り抜く文字列のバイト数)
    '
    '　戻り値：String(切り抜かれた文字列)
    '
    '　備考　：最後の１バイトが全角文字の半分になる場合、その１バイトは無視される。
    '-----------------------------------------------------------------------------------
    Public Shared Function PBFSTR_MidB(ByVal strVal As String, _
                            ByVal intStart As Integer, _
                            ByVal intLength As Integer) As String

        Try
            '*** 空文字に対しては常に空文字を返す
            If strVal = "" Then Return ""

            '*** intLengthのチェック
            'intLengthが0か、intStart以降のバイト数をオーバーする場合はintStart以降の全バイトが指定されたものとみなす。
            Dim intResetLength As Integer = System.Text.Encoding.GetEncoding("Shift_JIS").GetByteCount(strVal) - intStart + 1
            If intLength = 0 OrElse intLength > intResetLength Then
                intLength = intResetLength
            End If

            '*** 切り抜き
            Dim SJIS As System.Text.Encoding = System.Text.Encoding.GetEncoding("Shift-JIS")
            Dim bytBIG() As Byte = CType(Array.CreateInstance(GetType(Byte), intLength), Byte())
            Array.Copy(SJIS.GetBytes(strVal), intStart - 1, bytBIG, 0, intLength)


            Dim strNewVal As String = SJIS.GetString(bytBIG)

            '*** 切り抜いた結果、最後の１バイトが全角文字の半分だった場合、その半分は切り捨てる。
            Dim intResultLength As Integer = System.Text.Encoding.GetEncoding("Shift_JIS").GetByteCount(strNewVal) - intStart + 1

            If Asc(Strings.Right(strNewVal, 1)) = 0 Then
                'VB.NET2002,2003の場合、最後の１バイトが全角の半分の時
                Return strNewVal.Substring(0, strNewVal.Length - 1)

            ElseIf intLength = intResultLength - 1 Then
                'VB2005の場合で最後の１バイトが全角の半分の時
                Return strNewVal.Substring(0, strNewVal.Length - 1)

            Else
                'その他の場合
                Return strNewVal
            End If
        Catch ex As Exception
            Throw ex
        End Try

    End Function
#End Region
#Region "指定バイト文字列を取り出す"
    ''' <summary>
    ''' 指定バイト分文字列を取り出す
    ''' </summary>
    ''' <param name="strText">ターゲットの文字列</param>
    ''' <param name="intStart">開始位置</param>
    ''' <param name="intEnd">終了位置</param>
    ''' <param name="intMultiLine">改行排除 0:許可　1:排除</param>
    ''' <returns>編集後の文字列</returns>
    ''' <remarks>
    ''' ２Byteの文字の場合、指定Byte数より１Byte切って返す
    '''  Ex) test = "あかさ"
    ''' "あかさ" = PB_GetWantedString(test, 0, 5)
    ''' "あか" = PB_GetWantedString(test, 0, 4)
    '''  "かさ" = PB_GetWantedString(test, 1, 4)
    ''' </remarks>  
    Public Shared Function GetWantedByte(ByVal strText As String, _
                                         ByVal intStart As Integer, _
                                         ByVal intEnd As Integer, _
                                        Optional ByVal intMultiLine As Integer = 0) As String

        '指定バイト位置から指定バイト数分の文字列を取り出す関数
        Dim strJIS As System.Text.Encoding = System.Text.Encoding.GetEncoding("Shift_JIS")

        ''20061109_1 改行置換
        If intMultiLine <> 0 Then
            'strText = strText.Replace(ControlChars.CrLf, "")
            strText = strText.Replace(vbCrLf, "")
        End If

        If strText <> "" Then
            '指定文字列をバイト配列化
            Dim bytText() As Byte = strJIS.GetBytes(strText)
            Dim intSumText As Integer = strJIS.GetByteCount(strText)

            '引数のバイト数検証
            If intStart < 0 Or intEnd <= 0 Or intStart > intSumText Then Return ""

            'スタート文字をゲット、２Byteの場合はスタートバイト数 +1にする
            Dim strTemp As String = strJIS.GetString(bytText, 0, intStart)

            If intStart > 0 And strTemp.EndsWith(ControlChars.NullChar) Then
                intStart += 1   '開始位置が漢字の中なら次(前)の文字から開始
            End If

            If intStart + intEnd > intSumText Then    '文字長より多く取得しようとした場合
                intEnd = intSumText - intStart        '文字列の最後までの分とする
            End If

            '指定バイトの検証が問題無しの場合、取り出した文字を返す
            '2005と2003双方対応 20080806_1
            If strJIS.GetString(bytText, intStart, intEnd).EndsWith(ControlChars.NullChar) Or _
                                                strJIS.GetString(bytText, intStart, intEnd).EndsWith("・") Then
                Return strJIS.GetString(bytText, intStart, intEnd - 1)
            End If

            Return strJIS.GetString(bytText, intStart, intEnd)

            ''Return strJIS.GetString(bytText, intStart, intEnd).TrimEnd(ControlChars.NullChar)
        Else
            Return strText
        End If
    End Function
#End Region
#Region "指定文字前後の文字列取得"
    ' ------------------------------------------------------------------ 
    ' @(e) 
    ' 機能        : PBFSTR_GetWantedText 
    ' 返り値      : String(取り出した文字列)
    ' 
    ' 引き数      : strText:取り出したい文字列
    '               szLeft：どちら側の文字列を取得するか　True:左側　False:右側
    '               szTarget:指定文字
    '
    ' 機能説明    : 指定文字前後の文字列を取り出す
    ' 
    ''------------------------------------------------------------------
    Public Shared Function PBFSTR_GetWantedText(ByVal szText As String, Optional ByVal LorR As LorR = LorR.LEFT, _
                                                  Optional ByVal szTarget As String = " ") As String
        Try

            Dim intBarCnt As Integer

            intBarCnt = InStr(szText, szTarget)

            If intBarCnt > 1 Then
                If LorR = LorR.LEFT Then
                    Return Left(szText, intBarCnt - 1)
                Else
                    Return Right(szText, Len(szText) - intBarCnt)
                End If
            Else
                If szTarget = " " Then
                    intBarCnt = InStr(szText, "　")
                    If intBarCnt > 1 Then
                        If LorR = LorR.LEFT Then
                            Return Left(szText, intBarCnt - 1)
                        Else
                            Return Right(szText, Len(szText) - intBarCnt)
                        End If
                    End If
                    Return szText
                End If
                Return szText
            End If

        Catch ex As Exception
            'SkyLog.Error(ex.Message, ex)
            Return szText
        End Try
    End Function
#End Region
#Region "全角半角の判断"
    Public Shared Function ChkFullHalf(ByVal szText As String) As CHAR_SIZE

        Dim SJISEnc As Encoding = Encoding.GetEncoding("Shift_Jis")
        Dim inCnt As Integer = SJISEnc.GetByteCount(szText)

        '「GetByteCount(str)で取得したバイト数」と「str.Length * 2」が一致すれば、文字列はすべて全角
        '「GetByteCount(str)で取得したバイト数」と「str.Length」が一致すれば、文字列はすべて半角
        '「GetByteCount(str)で取得したバイト数 / 2」で余りが出た場合は、全角・半角混在です。

        If szText.Length = inCnt Then
            Return CHAR_SIZE.HALF

        ElseIf szText.Length * 2 = inCnt Then
            Return CHAR_SIZE.FULL

        Else
            Return CHAR_SIZE.FULLHALF
        End If

    End Function
#End Region
#Region "指定文字列から数字を除き文字のみを取得し、連結して戻す"
    Public Shared Function GetCharFromValue(ByVal prmValue As String) As String

        Dim CharVal As String = "" '取得した文字列
        Dim MaxLength As Integer = prmValue.Length '最大文字数

        For i As Integer = 0 To MaxLength - 1
            If Not IsNumeric(prmValue.Substring(i, 1)) Then
                CharVal += PBCStr(prmValue.Substring(i, 1))
            End If
        Next

        Return CharVal
    End Function
#End Region
#Region "指定文字列を除いて値を戻す"
    Public Shared Function RemoveFromValue(ByVal prmValue As String, ByVal prmtCar As String) As String

        Dim CharVal As String = "" '取得した文字列
        Dim MaxLength As Integer = prmValue.Length '最大文字数

        For i As Integer = 0 To MaxLength - 1
            If prmValue.Substring(i, 1) <> prmtCar Then
                CharVal += PBCStr(prmValue.Substring(i, 1))
            End If
        Next

        Return CharVal
    End Function
#End Region
#Region "改行コードを置換する"
    '*************************************************************
    '* 機能     : 改行コードを置換する
    '* 返り値   : prmbaseText : 基本となる文字列
    '* 引き数   : prmReplacement : 置換後の文字列
    '* 機能説明 : 　　　　　　
    '* 作成     : 
    '* 更新履歴 : 20090407_1 加藤   Lfのケースしか改行の変換ができていなかったため修正
    '*************************************************************
    Public Shared Function doReplaceLine(ByVal prmbaseText As String, Optional ByVal prmReplacement As String = " ") As String


        If IsNothing(prmbaseText) OrElse IsDBNull(prmbaseText) OrElse CStr(prmbaseText).Equals("") Then
            Return ""
        End If

        Dim rtn As String = ""

        rtn = Replace(prmbaseText, ControlChars.CrLf, prmReplacement) 'キャリッジリターン文字とラインフィード文字
        rtn = Replace(rtn, ControlChars.Cr, prmReplacement) 'キャリッジリターン文字
        rtn = Replace(rtn, ControlChars.Lf, prmReplacement) 'ラインフィード文字

        Return rtn

    End Function
#End Region '20091207_1
#Region "ArrayListからIN句を生成"
    '*************************************************************
    '* 機能     : PBF_SQLIN
    '* 返り値   : CharVal : 文字列
    '* 引き数   : prmAry : ArrayList
    '*            prmText : XXXX.XXXXX
    '* 機能説明 : 画面のチェック
    '* 備考     : AryからSQLのIN句を作成(ORACLE10gのIN句は1000までなのでIN句を分割する)
    '*　　　　　　prmKey IN ('XXX','XXX','XXX', …) or prmKey IN ('XXX','XXX','XXX', …)
    '* 作成     : 
    '* 更新履歴 :
    '*************************************************************
    Public Shared Function PbfCreateSqlIN(ByVal prmAry As ArrayList, ByVal prmText As String) As String
        Try

            Dim intAry As Integer 'IN句の必要個数
            Dim iCount As Integer '行数(prmAry)
            Dim CharVal As String = "" '取得した文字列
            Dim arytemp As New ArrayList '一時Ary
            Dim Max_Count As Integer = 999  ' AryTblMax値

            '--------------
            'IN句の必要個数を求める
            '--------------
            If PBCint(prmAry.Count) > Max_Count Then
                '999件以上の場合、999で割ることでIN句が何個必要なのかを計算
                intAry = PBCint(doCalHASU2(PBCdbl(prmAry.Count / Max_Count), 0, Round.Down))
                '剰余が0以外の場合は剰余の分のIN句が必要なので+1をする
                If PBCint(prmAry.Count) Mod Max_Count <> 0 Then intAry = intAry + 1
            Else
                '999件未満の場合は1
                intAry = 1
            End If

            For j As Integer = 1 To intAry
                '--------------
                'IN句の生成
                '--------------
                '999未満になるようprmAryを分割
                For i As Integer = 1 To Max_Count
                    If iCount < PBCint(prmAry.Count) Then
                        arytemp.Add(prmAry(iCount)) '一時Aryにセット
                        iCount += 1
                    Else
                        i = Max_Count
                    End If
                Next

                CharVal += prmText & " IN ("

                'ｺﾝﾏをつける  'XXX','XXX','XXX', …
                For iCnt As Integer = 0 To arytemp.Count - 1
                    If iCnt = 0 Then
                        CharVal += "'" & PBCStr(arytemp.Item(iCnt)) & "'"
                    Else
                        CharVal += ",'" & PBCStr(arytemp.Item(iCnt)) & "'"
                    End If
                Next

                If j <> intAry Then
                    CharVal += ") OR " '最終でない場合ORを付ける
                Else
                    CharVal += ")"
                End If

                arytemp.Clear() '一時Aryクリア

            Next

            Return CharVal

        Catch ex As Exception
            Throw ex
            Return ""
        End Try
    End Function
#End Region
#Region "ダブルクォーテーションで括る"
    Public Shared Function setDoubleQuotes(field As String) As String
        If field.IndexOf(""""c) > -1 Then
            '"を""とする
            field = field.Replace("""", """""")
        End If
        Return """" & field & """"
    End Function
#End Region
#End Region

#Region "暗号化"
    ''' <summary>
    ''' 文字列を暗号化する
    ''' </summary>
    ''' <param name="sourceString">暗号化する文字列</param>
    ''' <param name="password">暗号化に使用するパスワード</param>
    ''' <returns>暗号化された文字列</returns>
    Public Shared Function doEncrypt(ByVal sourceString As String, _
                                         ByVal password As String) As String

        Try

            'RijndaelManagedオブジェクトを作成
            Dim rijndael As New System.Security.Cryptography.RijndaelManaged()

            'パスワードから共有キーと初期化ベクタを作成
            Dim key As Byte() = Nothing
            Dim iv As Byte() = Nothing
            GenerateKeyFromPassword(password, rijndael.KeySize, key, rijndael.BlockSize, iv)
            rijndael.Key = key
            rijndael.IV = iv

            '文字列をバイト型配列に変換する
            Dim strBytes As Byte() = System.Text.Encoding.UTF8.GetBytes(sourceString)

            '対称暗号化オブジェクトの作成
            Dim encryptor As System.Security.Cryptography.ICryptoTransform = _
                rijndael.CreateEncryptor()
            'バイト型配列を暗号化する
            Dim encBytes As Byte() = encryptor.TransformFinalBlock(strBytes, 0, strBytes.Length)
            '閉じる
            encryptor.Dispose()

            'バイト型配列を文字列に変換して返す
            Return System.Convert.ToBase64String(encBytes)


        Catch ex As Exception
            Throw ex
        End Try
    End Function

    ''' <summary>
    ''' 暗号化された文字列を復号化する
    ''' </summary>
    ''' <param name="sourceString">暗号化された文字列</param>
    ''' <param name="password">暗号化に使用したパスワード</param>
    ''' <returns>復号化された文字列</returns>
    Public Shared Function doDecrypt(ByVal sourceString As String, _
                                         ByVal password As String) As String

        Try

            'RijndaelManagedオブジェクトを作成
            Dim rijndael As New System.Security.Cryptography.RijndaelManaged()

            'パスワードから共有キーと初期化ベクタを作成
            Dim key As Byte() = Nothing
            Dim iv As Byte() = Nothing
            GenerateKeyFromPassword(password, rijndael.KeySize, key, rijndael.BlockSize, iv)
            rijndael.Key = key
            rijndael.IV = iv

            '文字列をバイト型配列に戻す
            Dim strBytes As Byte() = System.Convert.FromBase64String(sourceString)

            '対称暗号化オブジェクトの作成
            Dim decryptor As System.Security.Cryptography.ICryptoTransform = _
                rijndael.CreateDecryptor()
            'バイト型配列を復号化する
            '復号化に失敗すると例外CryptographicExceptionが発生
            Dim decBytes As Byte() = decryptor.TransformFinalBlock(strBytes, 0, strBytes.Length)
            '閉じる
            decryptor.Dispose()

            'バイト型配列を文字列に戻して返す
            Return System.Text.Encoding.UTF8.GetString(decBytes)

        Catch ex As Exception
            Throw ex
        End Try
    End Function
    ''' <summary>
    ''' パスワードから共有キーと初期化ベクタを生成する
    ''' </summary>
    ''' <param name="password">基になるパスワード</param>
    ''' <param name="keySize">共有キーのサイズ（ビット）</param>
    ''' <param name="key">作成された共有キー</param>
    ''' <param name="blockSize">初期化ベクタのサイズ（ビット）</param>
    ''' <param name="iv">作成された初期化ベクタ</param>
    Private Shared Sub GenerateKeyFromPassword(ByVal password As String, _
                                               ByVal keySize As Integer, _
                                               ByRef key As Byte(), _
                                               ByVal blockSize As Integer, _
                                               ByRef iv As Byte())
        'パスワードから共有キーと初期化ベクタを作成する
        'saltを決める
        Dim salt As Byte() = System.Text.Encoding.UTF8.GetBytes("saltは必ず8バイト以上")
        'Rfc2898DeriveBytesオブジェクトを作成する
        Dim deriveBytes As New System.Security.Cryptography.Rfc2898DeriveBytes( _
            password, salt)
        '.NET Framework 1.1以下の時は、PasswordDeriveBytesを使用する
        'Dim deriveBytes As New System.Security.Cryptography.PasswordDeriveBytes( _
        '    password, salt)

        '反復処理回数を指定する デフォルトで1000回
        deriveBytes.IterationCount = 1000

        '共有キーと初期化ベクタを生成する
        key = deriveBytes.GetBytes(keySize \ 8)
        iv = deriveBytes.GetBytes(blockSize \ 8)
    End Sub

#End Region

#Region "★ガイド用 "
    Public Shared Function GUID_RowsCount(ByVal iCnt As Integer) As String
        Return "検索結果：" & iCnt & "件ありました。"
    End Function
    Public Shared Function GUID_RegUserInfo(ByVal InUser As String, ByVal InDate As String, ByVal UpUser As String, ByVal Update As String) As String
        Return "[作成者]" & InUser & "(" & InDate & ")   [更新者]" & UpUser & "(" & Update & ") "
    End Function
#End Region

    '#Region "データ取得(arrlyList)"
    '    '---------------------------------------------------------
    '    '　機能：データゲット(Return ArrayList)
    '    '
    '    '　引数　：Connection, SQL文, Optional(Transaction)
    '    '　戻り値：ArrayList(ゲットしたもの)
    '    '---------------------------------------------------------
    '    Private Shared Function getAryDataDB(ByVal ocon As uniConnection, ByVal SQL As String, _
    '                                        Optional ByVal tran As uniTransaction = Nothing) As ArrayList
    '        Dim ocd As New uniCommand
    '        Dim odr As UniDataReader
    '        Dim arlData As New ArrayList

    '        Try

    '            ocd.Connection = ocon
    '            ocd.CommandText = SQL

    '            If Not tran Is Nothing Then
    '                ocd.Transaction = tran
    '            End If

    '            odr = ocd.ExecuteReader

    '            While (odr.Read)
    '                For i As Integer = 0 To odr.FieldCount - 1
    '                    With arlData
    '                        .Add(PBCStr(odr.Item(i)))
    '                    End With
    '                Next
    '            End While

    '            odr.Close()
    '            Return arlData

    '        Catch ex As Exception
    '            Throw ex
    '        Finally

    '        End Try
    '    End Function
    '#End Region

#Region "保管"
#Region "日付の生成(年月日)"
    ''Public Overloads Shared Function PBDataAdd(ByVal szYMD As String, Optional ByVal inAddDate As Integer = 0, _
    ''                        Optional ByVal interval As DateInterval = DateInterval.Day, Optional ByVal szFormat As String = "yyyy/MM/dd") As String

    ''    Dim dtDate As Date

    ''    Try

    ''        dtDate = CDate(DateValue(szYMD).ToString("yyyy/MM/dd"))
    ''        ''InterValセット
    ''        dtDate = DateAdd(interval, inAddDate, dtDate)
    ''        ''フォーマット
    ''        Return dtDate.ToString(szFormat)

    ''    Catch ex As Exception
    ''        Return ""
    ''    End Try
    ''End Function
#End Region
#Region "日付の妥当性チェック"
    'Public Function PBFBL_CheckDay(ByVal intY As Integer, ByVal intM As Integer, ByVal intD As Integer) As Boolean
    '    If (DateTime.MinValue.Year > intY) OrElse (intY > DateTime.MaxValue.Year) Then
    '        Return False
    '    End If

    '    If (DateTime.MinValue.Month > intM) OrElse (intM > DateTime.MaxValue.Month) Then
    '        Return False
    '    End If

    '    Dim iLastDay As Integer = DateTime.DaysInMonth(intY, intM)
    '    If (DateTime.MinValue.Day > intD) OrElse (intD > iLastDay) Then
    '        Return False
    '    End If

    '    Return True
    'End Function
#End Region
#Region "消費税率取得"
    ''税率：税適用開始日 ～ 税適用終了日 間の税率取得
    'Public Function PBF_GetRTZEI(ByVal strDTZEI As String, ByVal con As SqlConnection) As Decimal
    '    Dim strSQL As String
    '    strSQL = ""
    '    strSQL = strSQL & " SELECT TO_CHAR(ZEI_RTZEI, '0.00') "
    '    strSQL = strSQL & " FROM M_ZEI "
    '    'strSQL = strSQL & " WHERE TO_CHAR(SYSDATE, 'YYYY/MM/DD') "
    '    strSQL = strSQL & " WHERE '" & strDTZEI & "'"
    '    strSQL = strSQL & "       BETWEEN TO_CHAR(ZEI_DTST, 'YYYY/MM/DD') "
    '    strSQL = strSQL & "       AND TO_CHAR(ZEI_DTED, 'YYYY/MM/DD') "

    '    Return PBFDEC_RtnDec(PBFSTR_GetOneDataDB(con, strSQL))
    'End Function
#End Region
#Region "消費税金額計算 (仕入先)"
    ' -------------------------------------------------------------------------------------------
    ' 機能        : PBF_CalSIRZEI 
    ' 
    ' 返り値      : Decimal(算出金額)
    ' 
    '引数         ：strCDSIR：仕入先M.仕入先CD
    '               decKNSIR：仕入金額(買金額)
    '               intSURYO：数量
    '               decRTZEI：消費税率 
    '               con : SqlConnection
    '               tran : SqlTransaction
    ' 機能説明    : 
    ' 備考        : 
    ''-------------------------------------------------------------------------------------------
    'Public Function PBF_CalSIRZEI(ByVal strCDSIR As String, ByVal lngKNSIR As Long, _
    '                              ByVal lngSURYO As Long, ByVal decRTZEI As Decimal, _
    '                              ByVal con As SqlConnection, _
    '                              Optional ByVal tran As SqlTransaction = Nothing) As Decimal


    '    Dim bytKBHASU As Byte   '仕入先M.端数処理区分
    '    Dim strSQL As String

    '    Try
    '        strSQL = ""
    '        strSQL = strSQL & " SELECT SIR_KBHASU "
    '        strSQL = strSQL & " FROM M_SIR "
    '        strSQL = strSQL & " WHERE SIR_CDSIR =  " & strCDSIR

    '        bytKBHASU = PBCbyt(PBFSTR_GetOneDataDB(con, strSQL, tran))

    '        '仕入金額 * 消費税率 
    '        Return PBF_CalHASU(lngKNSIR * decRTZEI, bytKBHASU)

    '    Catch ex As Exception
    '        SkyLog.Error(ex.Message, ex)
    '    End Try
    'End Function
#End Region
#Region "消費税金額計算 (得意先)"
    '' -------------------------------------------------------------------------------------------
    '' 機能        : PBF_CalTOKZEI 
    '' 
    '' 返り値      : Decimal(算出金額)
    '' 
    ''引数         ：strCDTOK：得意先M.得意先CD
    ''               decKNTOK：売上金額
    ''               intSURYO：数量
    ''               decRTZEI：消費税率 
    ''               con : SqlConnection
    ''               tran : SqlTransaction
    '' 機能説明    : 
    '' 備考        : 
    ' ''-------------------------------------------------------------------------------------------
    'Public Function PBF_CalTOKZEI(ByVal strCDTOK As String, ByVal lngKNTOK As Long, _
    '                              ByVal lngSURYO As Long, ByVal decRTZEI As Decimal, _
    '                              ByVal con As SqlConnection, _
    '                              Optional ByVal tran As SqlTransaction = Nothing) As Decimal


    '    Dim bytKBHASU As Byte   '得意先M.端数処理区分
    '    Dim strSQL As String

    '    Try
    '        strSQL = ""
    '        strSQL = strSQL & " SELECT TOK_KBHASU "
    '        strSQL = strSQL & " FROM M_TOK "
    '        strSQL = strSQL & " WHERE TOK_CDTOK =  " & strCDTOK

    '        '20060818_1 端数処理は四捨五入
    '        '' bytKBHASU = PBCbyt(PBFSTR_GetOneDataDB(con, strSQL, tran))
    '        bytKBHASU = 2

    '        '売上金額 * 消費税率 
    '        Return PBF_CalHASU(lngKNTOK * decRTZEI, bytKBHASU)

    '    Catch ex As Exception
    '        SkyLog.Error(ex.Message, ex)
    '    End Try
    'End Function
#End Region
    '---------------------------------------------------------------------
    '  機能    ：DateをDateFormatに変換
    '  引数    ：１．Object, (２．String )
    '  戻り値  ：String
    '  作成日  ：2006.07.25  F.Nishida
    '---------------------------------------------------------------------
    'Public Function PBFSTR_RtnDTE(ByVal objVal As Object, _
    ''                               Optional ByVal strFormat As String = "") As String
    '    '    Dim arrTemp() As String
    '    '    If IsNothing(objVal) OrElse IsDBNull(objVal) OrElse CStr(objVal) = "" Then
    '    '        Return ""
    '    '    ElseIf IsDate(objVal) Then
    '    '        If strFormat = "" Or strFormat = "1" Then
    '    '            arrTemp = CStr(objVal).Split(CChar("/"))
    '    '            Return arrTemp(0) & "年" & arrTemp(1) & "月" & arrTemp(2) & "日"
    '    '        End If
    '    '    Else
    '    '        Return ""
    '    '    End If
    'End Function
    '---------------------------------------------------------------------
    '  機能    ：DateをDateFormatに変換
    '  引数    ：１．Object, (２．String )
    '  戻り値  ：String
    '  作成日  ：2006.07.25  F.Nishida
    '---------------------------------------------------------------------
    ''Public Function PBFSTR_IsDATE(ByVal objVal As Object, _
    ''                               Optional ByVal strFormat As String = "yyyyMMdd") As String
    ''    Dim arrTemp() As String
    ''    If IsNothing(objVal) OrElse IsDBNull(objVal) OrElse CStr(objVal) = "" Then
    ''        Return ""
    ''    ElseIf IsDate(objVal) Then
    ''        Return CDate(objVal).ToString(strFormat)
    ''    Else
    ''        Return ""
    ''    End If
    ''End Function
#End Region
End Class
