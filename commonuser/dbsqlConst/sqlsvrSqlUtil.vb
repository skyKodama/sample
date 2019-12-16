Imports skysystem.common
Imports System.Text

Public Class sqlsvrSqlUtil
    '***************************************************************
    '*  SQLSERVER 各種定義取得用クラス
    '*            そのまま返すとビューなど見づらいので加工して返す 
    '***************************************************************

#Region "GetTableDifine"
    ''' <summary>
    ''' テーブル定義取得
    ''' 引数：
    ''' TableName:テーブル名
    ''' ObjectKbn:テーブル=0 | ビュー=1 | その他の場合テーブルと同じ判定
    ''' </summary>
    ''' <param name="TableName"></param>
    ''' <param name="ObjectKbn">テーブル=0 | ビュー=1 | その他の場合テーブルと同じ判定</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function getTableDesign(TableName As String, ObjectKbn As Integer) As String
        Dim sb As New System.Text.StringBuilder

        sb.AppendLine(" select ")
        sb.AppendLine("     ep.value  as COMMENT ") '説明
        sb.AppendLine("     ,c.name    as COLUMN_NAME ")  '列名
        sb.AppendLine("     ,tp.name as typeName ")  'データ型

        '桁数
        '文字列の場合・・・桁数が-1ならば「max」と判定、その他は、(SJISの場合)1/2にして判定。内部的にはバイト長で持っている？
        '  ※nvarcharしか追加していないので、char等も必要であれば追加する必要有
        'Decimalの場合・・・整数部と小数部の桁数をカンマで結合
        'Numericの場合・・・整数部と小数部の桁数をカンマで結合
        'その他・・・桁数が-1であれば「max」と判定、その他は最大桁数をそのまま出力
        sb.AppendLine("     ,case tp.name ")
        sb.AppendLine("    when 'nvarchar' then case c.max_length when -1 then 'max' else convert(varchar,c.max_length / 2) end ")
        sb.AppendLine("    when 'decimal' then CONVERT(varchar,c.precision)  + ',' + convert(varchar,c.scale) ")
        sb.AppendLine("    when 'numeric' then CONVERT(varchar,c.precision)  + ',' + convert(varchar,c.scale) ")
        sb.AppendLine("    else case c.max_length when -1 then 'max' else CONVERT(varchar,c.max_length) end  end as max_length ")

        sb.AppendLine("    ,CASE ISNULL(key_const.name,'FALSE')WHEN 'FALSE' THEN '' ELSE 'PK' END AS constraint_name ") 'PK判定
        sb.AppendLine("    ,CASE c.is_nullable WHEN 0 THEN 'FALSE' ELSE 'TRUE' END as is_nullable ") 'NULL許可判定
        sb.AppendLine("    ,'(' + convert(varchar,idc.seed_value) + ',' + convert(varchar,idc.increment_value) + ')' as identity_Val  ") 'IDENTIFY内容
        sb.AppendLine("    ,cc.definition ") 'definition(計算式)内容 

        sb.AppendLine(" from ")

        If ObjectKbn = 1 Then
            'VIEWの場合
            sb.AppendLine("      sys.views t ")
        Else
            'TABLEの場合
            sb.AppendLine("      sys.tables t ")
        End If

        sb.AppendLine(" left join  sys.columns c ")
        sb.AppendLine("      on t.object_id = c.object_id     ")
        sb.AppendLine(" left join  sys.extended_properties ep ")
        sb.AppendLine("      on c.object_id = ep.major_id ")
        sb.AppendLine("     and c.column_id = ep.minor_id ")
        sb.AppendLine(" left join sys.index_columns idx_cols ")
        sb.AppendLine("      on idx_cols. object_id = c.object_id ")
        sb.AppendLine("     AND idx_cols.column_id = c.column_id ")
        sb.AppendLine("  ")
        sb.AppendLine(" left join sys.key_constraints key_const ")
        sb.AppendLine("      on t.object_id = key_const.parent_object_id ")
        sb.AppendLine("     AND idx_cols.index_id = key_const.unique_index_id ")
        sb.AppendLine("     AND key_const.type = 'PK' ")
        sb.AppendLine(" left join sys.types tp ")
        sb.AppendLine("     on c.user_type_id = tp.user_type_id ")
        sb.AppendLine(" left join sys.identity_columns idc ")
        sb.AppendLine("     on idc.object_id = c.object_id  ")
        sb.AppendLine("    and idc.column_id = c.column_id  ")
        sb.AppendLine("  ")
        sb.AppendLine(" left join sys.computed_columns  cc ")
        sb.AppendLine("        on c.object_id = cc.object_id  ")
        sb.AppendLine("   and c.column_id = cc.column_id  ")
        sb.AppendLine(" where ")
        sb.AppendLine("     t.name = '" & TableName & "' ")
        sb.AppendLine(" order by ")
        sb.AppendLine("     c.column_id ")

        Return sb.ToString
    End Function

#End Region

#Region "ストアドプロシージャ定義取得"
    Public Shared Function GetProcDifine(dbc As DBconnection, ProcName As String) As String
        Try

            Dim sb As New StringBuilder
            sb.AppendLine("EXEC sp_helptext " & ProcName)

            Dim rtnDt As DataTable = DBUtil.GetDtDataDB(dbc.rtncon, sb.ToString)

            If rtnDt.Rows.Count = 0 Then Return ""

            sb = New StringBuilder
            For i As Integer = 0 To rtnDt.Rows.Count - 1
                sb.Append(rtnDt.Rows(i)("Text").ToString)
            Next

            Return sb.ToString
        Catch ex As Exception
            Throw ex
        End Try
    End Function
#End Region


#Region "ビュー定義取得"
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="dbc"></param>
    ''' <param name="ProcName"></param>
    ''' <returns></returns>
    Public Shared Function GetViewDifine(dbc As DBconnection, ProcName As String) As String
        Try
            '平文の取得
            Dim viewText As String = GetProcDifine(dbc, ProcName)
            viewText = Replace(viewText, vbCrLf, " ")

            'クォーテーション間、ダブルクォーテーション間の半角スペースをエスケープ文字へ変換
            viewText = replaceSpace(viewText)

            '下記ルールだと実現できないことに気が付いた・・・
            'スペースでSepalateしてTrimで結果を比較、のちに結合した方が良いかも
            'もしスペースが連続する様なら除外する
            '空白スペースを文字と文字の間に入れるようなビューだと成立しないのでNG？⇒ちょっと考える(例えばクォートとクォートの間にあるスペースは置き換えるなど)
            Dim sepText As String() = viewText.Split(" "c)
            Dim flgAftReturn As Integer = 0
            Dim iLevel As Integer = 0
            Dim strStock As String = ""
            Dim rtnText As New StringBuilder
            Dim temp As String = ""

            For Each rowText As String In sepText
                If strStock <> "" Then
                    temp = strStock & " " & rowText
                Else
                    temp = rowText
                End If


                Select Case chkViewTextAdd(temp, flgAftReturn, iLevel)
                    Case 0
                        'そのまま追加
                        rtnText.Append(" " & temp)
                        flgAftReturn = 0

                        strStock = ""
                    Case 1
                        '改行して追加
                        rtnText.Append(vbCrLf & StrDup(iLevel, vbTab) & temp)
                        flgAftReturn = 1
                        strStock = ""
                    Case 2
                        '直前にレベル追加
                        iLevel += 1
                        rtnText.Append(vbCrLf & StrDup(iLevel, vbTab) & temp)
                        flgAftReturn = 1
                        strStock = ""
                    Case 3
                        '直後にレベル追加
                        rtnText.Append(vbCrLf & StrDup(iLevel, vbTab) & temp)
                        iLevel += 1
                        flgAftReturn = 1
                        strStock = ""
                    Case 4
                        '直前にレベル減算
                        If iLevel > 0 Then
                            iLevel -= 1
                        End If
                        rtnText.Append(vbCrLf & StrDup(iLevel, vbTab) & temp)
                        flgAftReturn = 1
                        strStock = ""
                    Case 5
                        '直後にレベル追加
                        rtnText.Append(vbCrLf & StrDup(iLevel, vbTab) & temp)
                        If iLevel > 0 Then
                            iLevel -= 1
                        End If
                        flgAftReturn = 1
                        strStock = ""
                    Case 9
                        strStock = temp
                End Select

            Next

            Return rtnText.ToString.Replace("{sp}", " ")

        Catch ex As Exception
            Throw ex
        End Try
    End Function

    ''' <summary>
    ''' インデックスを追加する用。レベル制御は上層で行う
    ''' </summary>
    ''' <param name="prmAftReturn">直前改行フラグ⇒直前で改行しているなら、自分は改行しなくてもよい</param>
    ''' <param name="tempText"></param>
    ''' <returns>0:そのまま追加　1:自分の直前に改行追加　2：自分の直前でレベル加算　3：自分の直後でレベル加算　4：自分の直前でレベル減算　5：自分の直後でレベル減算 9:次の句と合わせて判定(Joinなど)</returns>
    Private Shared Function chkViewTextAdd(ByVal tempText As String, prmAftReturn As Integer, Optional ByVal prmLevel As Integer = 0) As Integer

        Try
            Dim rtnVal As Integer = 0
            'If prmAftReturn = 1 Then Return 0

            Select Case tempText.ToUpper.Trim
                Case "DISTINCT", "AS"
                    rtnVal = 0
                Case "ON"
                    rtnVal = 2
                Case "AND"
                    rtnVal = 1
                Case "SELECT", "("
                    rtnVal = 3
                Case ")", "FROM", "LEFT JOIN", "LEFT OUTER JOIN", "RIGHT JOIN", "RIGHT OUTER JOIN", "INNER JOIN", "JOIN", "WHERE"
                    rtnVal = 4

                Case "LEFT", "RIGHT", "INNER", "FULL", "OUTER", "GROUP", "LEFT OUTER", "RIGHT OUTER", "FULL OUTER"
                    rtnVal = 9
                Case Else
                    rtnVal = 1
            End Select

            Return rtnVal
        Catch ex As Exception
            Throw ex
        End Try
    End Function


    Private Shared Function replaceSpace(prmBaseText As String) As String
        Try
            Dim tempText As String = prmBaseText
            Dim tempIdx As Integer = prmBaseText.Length
            Dim idxStart As Integer = 0

            Dim FlgQuort As Boolean = False
            Dim FlgDQuort As Boolean = False

            Dim idxSp As Integer = prmBaseText.Length
            Dim idxQut As Integer = prmBaseText.Length
            Dim idxDqt As Integer = prmBaseText.Length


            While tempIdx >= 0
                

                idxSp = prmBaseText.Length
                idxQut = prmBaseText.Length
                idxDqt = prmBaseText.Length


                idxSp = tempText.IndexOf(" ", idxStart)
                idxQut = tempText.IndexOf("'", idxStart)
                idxDqt = tempText.IndexOf("""", idxStart)

                '最小値を取る(Minだと-1が紛れたときに処理がおかしくなる)
                'スペースは無条件で入れてよい。
                tempIdx = idxSp

                'Idxが取れており、現在最小より小さいか、Idxが取れていない場合置き換え
                If idxQut >= 0 And (tempIdx > idxQut Or tempIdx < 0) Then
                    tempIdx = idxQut
                End If

                If idxDqt >= 0 And (tempIdx > idxDqt Or tempIdx < 0) Then
                    tempIdx = idxDqt
                End If

                If tempIdx >= 0 Then

                    idxStart = tempIdx + 1

                    '最小値がクォートの場合、フラグ反転
                    If tempIdx = idxQut Then
                        FlgQuort = Not FlgQuort
                    End If

                    '最小値がダブルクォートの場合、フラグ反転
                    If tempIdx = idxDqt Then
                        FlgQuort = Not FlgQuort
                    End If


                    '最小値がスペースの場合、フラグを確認し、どちらかのフラグが立っていた場合にスペースを置き換え
                    If tempIdx = idxSp And (FlgQuort Or FlgDQuort) Then
                        tempText = tempText.Remove(tempIdx, 1).Insert(tempIdx, "{sp}")
                    End If
                End If






            End While


            Return tempText


        Catch ex As Exception

        End Try
    End Function
#End Region

End Class
