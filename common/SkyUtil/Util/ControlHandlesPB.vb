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
'*　機能　　：共通ハンドルクラス
'*            <Button・Combo・Enterコントロールイベント制御クラス>
'*　作成日　：2006.04.25    黄
'*
'*　＜変更内容＞
'*  2006.05.24    黄    コンボボックス空行挿入
'*  2006.06.27    黄    コンボボックスContents 文字列skylog(targetCombo_ValueChanged)
'*  2006.06.29    黄    InputMan NumberｺﾝﾄﾛｰﾗのNegativeValueskylog(targetIMNumber_ValueChanged)
'*  20060630　    駒形  ｺﾝﾎﾞﾎﾞｯｸｽ初期化(PBS_SetComboDS)
'*  2006.07.05    黄    ｺﾝﾎﾞﾎﾞｯｸｽｸﾘｱ(PBS_SetComboDS) 
'*                      StackOverflowException発生防止(targetCombo_ValueChanged)
'*  2006.07.10    黄    DropDownStyle(初期値：DropDown), Field代用名称(strAS)
'*                      ChangedComboValue(Combo値変更時、Event発生させる) 
'*  2006.07.11    黄    LinkLabel 使用可･不可により、Label状態変更
'*  2006.07.13    黄    Contentの小文字skylog(targetCombo_ValueChanged)ｺﾒﾝﾄ処理
'*  2006.07.19    黄    ImeMode追加(PBS_SetComboDS), LINK_LABELのTabStop設定(Default:False)
'*  2006.07.26    黄    PBS_SetComboDS：ｺﾝﾎﾞｾｯﾄし、MaxDropDownItemsの不具合修正
'*  2006.07.28    黄    targetCombo_ValueChangedｲﾍﾞﾝﾄ二回発生の不具合修正
'*  20060807_1    駒形  PBS_SetComboDS DEL
'*  2006.09.06    黄    ｺﾝﾎﾞﾎﾞｯｸｽWidth調整(PBS_SetComboDS, intMinus追加)
'*  2006.09.11    黄    ｺﾝﾎﾞﾎﾞｯｸｽDropDownList設定
'*  20060921_1    黄    PBS_SetComboDS(Tran追加)
'*  20061122_1    駒形  OrderBy句の追加
'*********************************************************************************************
''' <summary>
''' 共通ハンドルクラス
''' </summary>
''' <remarks></remarks>
Public Class ControlHandlesPB

#Region "Private変数"
    Private WithEvents targetModeButton As Button   'モードボタン
    Private targetModeLabel As Label                'モードのラベル

    Private WithEvents targetCtrl As Control        'Enterキーを受け入れるコントロール
    Private startCtrl As Control                    'フォーカスの移動先となるコントロール
    Private baseForm As Form                        'コントロールを保持しているフォーム

    'Private WithEvents targetCombo As IM.Combo      'ターゲットのコンボボックス
    'Private strPattern As String                    'コンボボックスPatternセット

    'Private WithEvents targetIMNumber As IM.Number  'HighlightTextｾｯﾄｺﾝﾄﾛｰﾗ(
    'Private WithEvents targetIMEdit As IM.Edit      '   (Default=Trueですが、うまくｾｯﾄできなかったので、強制的ｾｯﾄのため。)

    Private WithEvents targetStbar As StatusBar      ''ステイタスバー 20061225_1 

    Private oda As SqlDataAdapter
    Private ocd As SqlCommand
    Private dts As DataSet
    Private blnFlgInit As Boolean = False           'セットコンボ初期化フラグ
    Private strSQL As String

    Private bytHighNega As Byte                     'IM.Number型のｲﾍﾞﾝﾄ制御(0：HighLightText不具合skylog, 1：NegativeColorskylog)

    '↓ADD 2006.07.10 (ｺﾝﾎﾞﾎﾞｯｸｽ値が変更された時、ﾌｫｰﾑ側で処理したい場合)
    Public Event ChangedComboValue(ByVal sender As Object, ByVal strKey As String)

    '↓ADD 2006.07.11
    Private WithEvents lbl_LinkLabel As LinkLabel
    Private lbl_TitleLabel As Label
    Private fontStyleLeave As FontStyle     'MouseLeave時のFontStyle(Title文字太さ：入力系･マスタ系)

#End Region


#Region "コンストラクタ"
    ' ------------------------------------------------------------------ 
    '　機能：ステータスバー 作成者：更新者の追加
    '　引数：frmStbar
    ' ------------------------------------------------------------------ 
    Public Sub New(ByVal frmStbar As StatusBar)
        Me.targetStbar = frmStbar

    End Sub '20061225_1

    ' ------------------------------------------------------------------ 
    '　機能：Edit・NumberのImputManｺﾝﾄﾛｰﾗの不具合対策(HighLightText)
    '　引数：IMNumber・IMEdit
    ' ------------------------------------------------------------------ 
    Public Sub New(ByVal frmIMNumber As IM.Number, Optional ByVal highNega As Byte = 0)
        Me.targetIMNumber = frmIMNumber
        Me.bytHighNega = highNega       'HighLight(0)・Negative区分ﾌﾗｸﾞ(1)
    End Sub
    Public Sub New(ByVal frmIMEdit As IM.Edit)
        Me.targetIMEdit = frmIMEdit
    End Sub

    ' ------------------------------------------------------------------ 
    '　機能：ボタンイベント発生用(Enabled=True・Falseにより、ラベルの状態変更)
    '  引数：frmButton (状態変更により、ラベル変更)
    '        frmLabel (ボタンによって、状態変更)
    ' ------------------------------------------------------------------ 
    Public Sub New(ByVal frmButton As Button, _
                   ByVal frmLabel As Label)
        Me.targetModeButton = frmButton
        Me.targetModeLabel = frmLabel
    End Sub

    ' ------------------------------------------------------------------ 
    '　機能：キーEnterハンドル用
    '  引数：targetCtrl (Enterキーを受け入れるControl)
    '        startCtrl (フォーカス移動先のControl)
    '        baseForm (Controlを保持するフォーム)
    ' ------------------------------------------------------------------ 
    Public Sub New(ByVal targetCtrl As Control, _
                   ByVal startCtrl As Control, _
                   ByVal baseForm As Form)
        Me.targetCtrl = targetCtrl
        Me.startCtrl = startCtrl
        Me.baseForm = baseForm
    End Sub

    ' ------------------------------------------------------------------ 
    '　機能：コンボチェンジハンドル用
    '  引数：targetCombo (コンボValueChanged発生Control)
    ' ------------------------------------------------------------------ 
    Public Sub New(ByVal targetCombo As IM.Combo)
        Me.targetCombo = targetCombo
        ''Me.blnFlgSpace = blnSpace
    End Sub


    '↓ADD 2006.07.11
    ' ------------------------------------------------------------------------- 
    '　機能：LinkLabelハンドル用
    '  引数：lbl_LinkLabel (MouseEnter･MouseLeave･EnabledChanged発生Control)
    '        lbl_TitleLabel(LinkLabelｺﾝﾄﾛｰﾙがEnabled=Falseの場合、代用ｾｯﾄ)
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
        Me.lbl_LinkLabel.TabStop = blnTabStop   '← ADD 2006.07.19
        Me.fontStyleLeave = style
    End Sub
#End Region


#Region "★※メッソド※★"

#Region "有効なコントロールにフォーカスを移動する"
    ' ------------------------------------------------------------------ 
    '　機能：次のコントロールへ移行
    '
    ' ------------------------------------------------------------------ 
    Private Sub PRS_NextControl()
        Dim nextCtrl As Control = startCtrl
        Do
            If (TypeOf nextCtrl Is RadioButton) Then
                'ラジオボタンの時は,チェックされているものに対して,フォーカスをセットする
                If CType(nextCtrl, RadioButton).Checked Then
                    'ラベル以外のフォーカスを受け入れるコントロールの場合フォーカス移動
                    nextCtrl.Focus()
                    Exit Do
                End If
            ElseIf Not (TypeOf nextCtrl Is Label) And nextCtrl.Visible And nextCtrl.Enabled Then
                'ラベル以外のフォーカスを受け入れるコントロールの場合フォーカス移動
                nextCtrl.Focus()
                Exit Do
            End If
            nextCtrl = baseForm.GetNextControl(nextCtrl, True)
        Loop Until nextCtrl Is startCtrl
    End Sub
#End Region

#End Region

#Region "コンボボックス値セット：バインド"
    '********************************************************************
    '* 機能　　　: コンボボックスバインド処理
    '* 返り値　　: なし
    '* 引き数　　: target           -in GrapeCity.Win.Input.Combo   対象コンボボックス
    '*             dtTbl            -in DataTable                   データテーブル
    '*             valueMember      -in String                      value格納カラム名
    '*             displayMember    -in String                      名称格納カラム名
    '*             addSpaceItem     -in Boolean                     空白行追加フラグ(True..追加する、False..追加しない)
    '* 機能説明　:
    '* 備考　    :
    '* 作成  　  : 2007/04/11 Hoshiya
    '* 更新履歴  : 20071025_1 ドロップダウンリスト 自動調整
    '********************************************************************
    ''' <summary>
    '''  コンボボックスバインド処理
    ''' </summary>
    ''' <param name="con"></param>
    ''' <param name="strTable"></param>
    ''' <param name="valueMember"></param>
    ''' <param name="displayMember"></param>
    ''' <param name="strWhere"></param>
    ''' <param name="szOrder"></param>
    ''' <param name="blnSpace"></param>
    ''' <param name="tran"></param>
    ''' <remarks>SkyBaseComboを使用しているため、現在未使用</remarks>
    Public Sub BindCmbData(ByVal con As SqlConnection, ByVal strTable As String, _
                         ByVal valueMember As String, ByVal displayMember As String, _
                         Optional ByVal strWhere As String = "", Optional ByVal szOrder As String = Nothing, _
                         Optional ByVal blnSpace As Boolean = True, Optional ByVal tran As SqlTransaction = Nothing)

        blnFlgInit = False   'セットコンボ初期化フラグ(False)
        Dim strAS As String = "NM" 'As句


        'SQL文作成
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

        'コマンド作成
        oda = New SqlDataAdapter
        dts = New DataSet

        'Modify 20060921_1
        'ocd = New OracleCommand(strSQL, con)
        ocd = New SqlCommand(strSQL, con, tran)

        oda.SelectCommand = ocd
        oda.Fill(dts, strTable)

        '** ADD 2006.07.05
        targetCombo.Items.Clear()  'ｾｯﾄする前初期化


        If dts.Tables(0).Rows.Count > 0 Then

            '初期化
            targetCombo.DataSource = Nothing
            targetCombo.Items.Clear()


            Dim dtTbl_Clone As DataTable = dts.Tables(0).Copy()

            'ValuMenbの最大文字数を取得
            Dim iMaxLength As Integer = 0
            For i As Integer = 0 To dtTbl_Clone.Rows.Count - 1
                If PBCStr(dtTbl_Clone.Rows.Item(i)(valueMember)).Length > iMaxLength Then
                    iMaxLength = PBCStr(dtTbl_Clone.Rows.Item(i)(valueMember)).Length
                End If
            Next

            '配列クラスに格納
            Dim cmbAry As New ArrayList

            '空項目を追加するか？
            If blnSpace Then
                cmbAry.Add(New cConf("", "", 0))
            End If

            For i As Integer = 0 To dtTbl_Clone.Rows.Count - 1
                cmbAry.Add(New cConf(PBCStr(dtTbl_Clone.Rows.Item(i)(valueMember)), _
                                                        PBCStr(dtTbl_Clone.Rows.Item(i)(strAS)), iMaxLength))
            Next

            'データソース設定
            targetCombo.DataSource = cmbAry
            targetCombo.ValueMember = "ValueData"
            targetCombo.DisplayMember = "DisplayDataEdting"

            '自動調整 20071025_1
            targetCombo.DropDownAutoSize = True
            '候補取得
            targetCombo.AutoSelect = True

        End If
    End Sub
#End Region

#Region "SetStBar ステイタスバー 作成者、更新者のセット"
    ''' <summary>
    ''' ステイタスバー 作成者、更新者のセット
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

            'コマンド作成
            oda = New SqlDataAdapter
            dts = New DataSet
            ocd = New SqlCommand(szSQL, con, tran)

            oda.SelectCommand = ocd
            oda.Fill(dts)

            If dts.Tables(0).Rows.Count > 0 Then

                With dts.Tables(0)

                    stBarText.Append("【新規作成】")
                    stBarText.Append(PBCStr(.Rows(0)("IN_NAME")))
                    stBarText.Append("(" & PBCStr(.Rows(0)("IN_DATE")) & ")")
                    stBarText.Append("　【最終更新】")
                    stBarText.Append(PBCStr(.Rows(0)("UP_NAME")))
                    stBarText.Append("(" & PBCStr(.Rows(0)("UP_DATE")) & ")")


                    ''ステータスバーにセット
                    Me.targetStbar.Panels(0).Text = stBarText.ToString

                End With
            End If


        Catch ex As Exception
            Throw ex
        End Try

    End Sub
#End Region '20061225_1

#Region "☆※イベント※☆"

#Region "EnabledChangedイベント：ボタン状態イベント"
    ' ------------------------------------------------------------------ 
    '
    '　機能：モードによるボタンの状態変更の場合、ラベル状態変更
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

#Region "KeyDownイベント：EnterKeyHandles"
    ' ------------------------------------------------------------------ 
    '　機能：フォーカス移動
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

#Region "ValueChangedイベント：コンボチェンジイベント"
    ' DEL 20080331_1 Leaveイベントへ移行
    '------------------------------------------------------------------ 
    '　機能：コンボテキストセットイベント(Content & ｜ & Desciption)
    '
    '　備考：
    ' ------------------------------------------------------------------ 
    Private Sub targetCombo_ValueChanged(ByVal sender As Object, _
                                         ByVal e As System.EventArgs) Handles targetCombo.ValueChanged

        'Dim intLength As Integer

        'コンボボックス値セット中、イベント発生しないように。
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
                    '    RaiseEvent ChagedComboValue(sender, strContent) '← ADD 2006.07.10 
                    '    Exit For    '← ADD 2006.07.05 StackOverflowException 発生防止
                    'End If

                    If strContent.ToUpper = CStr(targetCombo.Items.Item(i).Content) OrElse _
                        strContent.ToLower = CStr(targetCombo.Items.Item(i).Content) Then

                        blnFlgInit = False  '← ADD 2006.07.28(ｲﾍﾞﾝﾄ二回発生防止)

                        strDesciption = PBCStr(targetCombo.Items.Item(i).Description)
                        targetCombo.Value = CStr(targetCombo.Items.Item(i).Content) & PBCSTR_VERTICAL & strDesciption

                        blnFlgInit = True  '← ADD 2006.07.28

                        RaiseEvent ChangedComboValue(targetCombo, strContent) '← ADD 2006.07.10 
                        Exit For    '← ADD 2006.07.05 StackOverflowException 発生防止
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
                RaiseEvent ChangedComboValue(targetCombo, "") '← ADD 2006.07.14 
                Exit Sub
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
#End Region

#Region "ValueChangedイベント：NumberｺﾝﾄﾛｰﾙのNegativeValueにskylog"
    Private Sub targetIMNumber_ValueChanged(ByVal sender As Object, _
                                            ByVal e As System.EventArgs) Handles targetIMNumber.ValueChanged
        '*** NegativeValueskylogの場合のみ処理を行わせる
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

#Region "GotFocusイベント：選択された状態(HighlightText状態)"
    Private Sub targetIMNumber_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) _
                                        Handles targetIMNumber.GotFocus

        '*** HighLight不具合の場合のみ処理を行わせる
        If bytHighNega = 0 Then
            '*** 文字が選択された状態
            Dim iLength As Integer = targetIMNumber.Text.IndexOf(targetIMNumber.Text)
            If iLength > -1 Then
                targetIMNumber.SelectionStart = iLength
                targetIMNumber.SelectionLength = targetIMNumber.Text.Length
            End If
        End If
    End Sub
    Private Sub targetIMEdit_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) _
                                  Handles targetIMEdit.GotFocus

        '*** 文字が選択された状態
        Dim iLength As Integer = targetIMEdit.Text.IndexOf(targetIMEdit.Text)
        If iLength > -1 Then
            targetIMEdit.SelectionStart = iLength
            targetIMEdit.SelectionLength = targetIMEdit.Text.Length
        End If
    End Sub
#End Region

#Region "LinkLabelイベント"
#Region "MouseEnter(LinkLabel)イベント：LinkLabelのフォト変更"
    'マウスがテキスト上にある場合の処理
    Private Sub LinkLabel_MouseEnter(ByVal sender As Object, _
                                     ByVal e As System.EventArgs) Handles lbl_LinkLabel.MouseEnter

        Dim link As LinkLabel = CType(sender, LinkLabel)
        link.Font = New Font(link.Font, FontStyle.Bold)     'フォントを太字にする 
    End Sub
#End Region
#Region "MouseLeave(LinkLabel)イベント：LinkLabelのフォト変更"
    'マウスがテキスト上から離れた場合の処理 
    Private Sub LinkLabel_MouseLeave(ByVal sender As Object, _
                                     ByVal e As System.EventArgs) Handles lbl_LinkLabel.MouseLeave

        Dim link As LinkLabel = CType(sender, LinkLabel)
        link.Font = New Font(link.Font, fontStyleLeave)     'フォントｾｯﾄ(入力系・マスタ系)
    End Sub
#End Region
#Region "EnabledChanged(LinkLabel)イベント：代用Label表示･非表示"
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

#Region "MouseHover,Leaveｲﾍﾞﾝﾄ"
#Region "MouseHoverイベント：共通処理"
    'Friend Sub MouseHoverEvent(ByVal sender As System.Object, ByVal e As System.EventArgs) _
    '                        Handles targetStbar.MouseHover
    '    If sender Is targetStbar Then
    '        Dim asm As System.Reflection.Assembly = System.Reflection.Assembly.GetExecutingAssembly()
    '        Dim ver As System.Version = asm.GetName().Version  'バージョンの取得
    '        Me.targetStbar.Panels(3).Text = ver.ToString
    '    End If

    'End Sub
#End Region

#Region "MouseLeaveイベント：共通処理"
    'Friend Sub MouseLeaveEvent(ByVal sender As System.Object, ByVal e As System.EventArgs) _
    '                        Handles targetStbar.MouseLeave
    '    'If sender Is targetStbar Then
    '    '    Me.targetStbar.Panels(3).Text = stu_EMP.PGMID
    '    'End If

    'End Sub
#End Region

#End Region '20070328_1

#End Region

#Region "保留"

    ''DEL 20060807_1
    '#Region "コンボボックス値セット"
    '    ' ------------------------------------------------------------------------------ 
    '    '　機能：コンボのDataSource値セット
    '    '　      (スペースプラグにより、最初行にスペース行挿入・無しにする)
    '    '  引数：コネクション(con)
    '    '        Table名(strTable), キー(strKey)
    '    '        詳細(strDescription)
    '    '        Optional：Where条件(strWhere)・DropDownStyle(bytDropDownStyle)
    '    '                  ALias(strAS, 二つ以上のFieldを結合する場合)
    '    '                  ImeMode(ime, 文字入力可能の場合)
    '    '                  MaxDropDownItems(最大Listを超えた場合、ｴﾗｰ発生)
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
    '            blnFlgInit = False   'セットコンボ初期化フラグ(False)

    '            'SQL文作成
    '            strSQL = ""
    '            strSQL = strSQL & " SELECT " & strKey & "," & strDescription
    '            strSQL = strSQL & " FROM " & strTable & " "
    '            If strWhere <> "" Then
    '                strSQL = strSQL & " WHERE " & strWhere
    '            End If
    '            strSQL = strSQL & " ORDER BY " & strKey

    '            con = PBFCON_ChkConnection(con)

    '            'コマンド作成
    '            oda = New sqlDataReader
    '            dts = New DataSet
    '            ocd = New sqlCommand(strSQL, con)
    '            oda.SelectCommand = ocd
    '            oda.Fill(dts, strTable)

    '            '** 20060630
    '            targetCombo.Items.Clear()  'ｾｯﾄする前初期化

    '            '** ADD 2006.07.10
    '            If strAS <> "" Then strDescription = strAS

    '            If dts.Tables(0).Rows.Count > 0 Then
    '                With targetCombo

    '                    If blnFlgSpace Then
    '                        '■DataSourceにスペース行追加の場合

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
    '                        '■DataSourceにスペース行追加しないの場合

    '                        'パラメーター設定
    '                        .DataSource = dts.Tables(0)

    '                        '説明文として表示するデータソースのプロパティを示す文字列を設定します。
    '                        .DisplayMember = strKey

    '                        '値として扱うデータソースのプロパティを示す文字列を設定します。
    '                        .ValueMember = strKey

    '                        '説明文として表示するデータソースのプロパティを示す文字列を設定
    '                        .DescriptionMember = strDescription
    '                    End If


    '                    '##<< 共通設定 >>##
    '                    .AutoSelect = True
    '                    .HighlightText = IM.HighlightText.All
    '                    '.ImeMode = ImeMode.Disable
    '                    .ImeMode = ime   '← Modify 2006.07.19
    '                    .ListBoxStyle = IM.ListBoxStyle.TextWithDescription
    '                    .TextBoxStyle = IM.TextBoxStyle.TextOnly
    '                    .TextHAlign = IM.AlignHorizontal.Left
    '                    .TextVAlign = IM.AlignVertical.Middle
    '                    .DropDownWidth = .Width
    '                    .ImageWidth = 0
    '                    .DropDownStyle = bytDropDownStyle               '← ADD 2006.07.10
    '                    .ShowScrollBar = True                           '← ADD 2006.07.28
    '                    .ScrollBarMode = IM.ScrollBarMode.Automatic     '← ADD 2006.07.28

    '                    Dim maxDigit As Integer 'Contentの最大桁数取得
    '                    For i As Integer = 0 To .Items.Count - 1
    '                        If maxDigit < CStr(.Items.Item(i).Content).Length Then
    '                            maxDigit = CStr(.Items.Item(i).Content).Length
    '                        End If
    '                    Next

    '                    '↓↓↓ADD 2006.07.26
    '                    'ドロップダウン部分に表示される項目の最大数を取得または設定
    '                    '.MaxDropDownItems = .Items.Count + 1
    '                    If .Items.Count > iMaxDownItems Then
    '                        .MaxDropDownItems = iMaxDownItems + 1

    '                        '↓↓Modify 2006.07.28 
    '                        'Contentの幅部分で、少し見えなくなる不具合発生skylog
    '                        If .DropDownWidth > (maxDigit * 10) Then

    '                            'Content((最大桁数 + 1) *12)  ｜ Descriptionサイズｾｯﾄ
    '                            .DescriptionWidth = .DropDownWidth - ((maxDigit + 1) * 12)
    '                        Else
    '                            .DescriptionWidth = .DropDownWidth - CInt(.DropDownWidth / 4)
    '                        End If
    '                    Else
    '                        .MaxDropDownItems = .Items.Count + 1

    '                        If .DropDownWidth > (maxDigit * 10) Then

    '                            'Content((最大桁数 + 1) *10)  ｜ Descriptionサイズｾｯﾄ
    '                            .DescriptionWidth = .DropDownWidth - ((maxDigit + 1) * 10)
    '                        Else
    '                            .DescriptionWidth = .DropDownWidth - CInt(.DropDownWidth / 4)
    '                        End If
    '                    End If

    '                    'Comment 2006.07.28 (↑移動)
    '                    ''Content((最大桁数 + 1) *10)  ｜ Descriptionサイズｾｯﾄ
    '                    'If .DropDownWidth > (maxDigit * 10) Then
    '                    '    .DescriptionWidth = .DropDownWidth - ((maxDigit + 1) * 10)
    '                    'Else
    '                    '    .DescriptionWidth = .DropDownWidth - CInt(.DropDownWidth / 4)
    '                    'End If

    '                End With
    '            End If

    '            '★初期化処理完了(セットコンボ初期化フラグ=True)
    '            blnFlgInit = True

    '        Catch ex As Exception
    '            SkyLog.Error(ex.Message, ex)
    '        End Try
    '    End Sub
    ''#End Region
    '#Region "KeyPressイベント：数字、BackSpace以外のものは入力不可"
    '    Private Sub targetCombo_KeyPress(ByVal sender As Object, _
    '                                     ByVal e As System.Windows.Forms.KeyPressEventArgs) _
    '                                     Handles targetCombo.KeyPress
    '        If ControlFlg <> 2 Then Exit Sub
    '        '数字、BackSpace以外のものは入力不可にする
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
    '*　機能　　：配列格納クラス
    '*            
    '*　
    '*********************************************************************************************
    Private DisMeb As String
    Private DescMeb As String
    Private ValMeb As String
    Private pMaxLength As Integer

#Region "コンストラクタ"
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

#Region "プロパティ"
#Region "DisplayMemberEditing "
    ReadOnly Property DisplayDataEdting() As String
        Get
            If ValMeb.Equals("") Then
                Return ""
            Else
                Return ValMeb.PadRight(pMaxLength) & "｜" & DisMeb
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
