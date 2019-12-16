
Imports skysystem.common
Imports skysystem.common.SystemUtil
Imports Infragistics.Win
Imports Devart.Data.Universal
Public Class systemValueList

    Protected Shared prdt As New DataTable

    ''' <summary>
    ''' ショートカット
    ''' </summary>
    ''' <param name="con"></param>
    ''' <param name="prmSpace"></param>
    ''' <param name="tran"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function vlSortCut(con As UniConnection, Optional prmSpace As Boolean = True, Optional tran As UniTransaction = Nothing) As ValueList

 

        Dim vl As New ValueList
        Dim valNum As Integer

        vl.ValueListItems.Clear()
        '空項目を追加するか？
        If prmSpace Then
            vl.ValueListItems.Add("", "")
        End If

        For Each valNum In System.Enum.GetValues(GetType(Shortcut))

            vl.ValueListItems.Add(valNum, System.Enum.GetName(GetType(Shortcut), valNum))

        Next

        Return vl

    End Function
    ''' <summary>
    ''' 名称マスタ取得
    ''' </summary>
    ''' <param name="con"></param>
    ''' <param name="prmSpace"></param>
    ''' <param name="tran"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function vlName(con As UniConnection, prmKey As String, Optional bar As String = "｜",
                                         Optional prmSpace As Boolean = True, Optional tran As UniTransaction = Nothing) As ValueList

        Dim vl As New ValueList
        Dim szsql As String = ""
        szsql = "select nam_cdnam,nam_name from m_sys_name "
        szsql += String.Format(" where nam_cdkey = '{0}'", prmKey)

        szsql += " order by nam_sort,nam_cdnam"


        prdt = DBUtil.GetDtDataDB(con, szsql, tran)

        If prdt.Rows.Count > 0 Then

            ''最大文字数を取得
            Dim iMaxLen As Integer = getMaxLength("nam_cdnam")

            '空項目を追加するか？
            Dim j As Integer = 0
            If prmSpace Then
                vl.ValueListItems.Add("", "")
                j += 1
            End If


            For Each dr As DataRow In prdt.Rows

                vl.ValueListItems.Add(PBCStr(dr.Item("nam_cdnam")),
                                       PBCStr(dr.Item("nam_cdnam")).PadRight(iMaxLen) & "｜" & PBCStr(dr.Item("nam_name")))

                j += 1

            Next

        End If

        Return vl

    End Function
    ''' <summary>
    ''' 名称マスタ取得/ｺｰﾄﾞ表示なし
    ''' </summary>
    ''' <param name="con"></param>
    ''' <param name="prmSpace"></param>
    ''' <param name="tran"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function vlNameNotCode(con As UniConnection, prmKey As String,
                                         Optional prmSpace As Boolean = True, Optional tran As UniTransaction = Nothing) As ValueList

        Dim vl As New ValueList
        Dim szsql As String = ""
        szsql = "select nam_cdnam,nam_name from m_sys_name "
        szsql += String.Format(" where nam_cdkey = '{0}'", prmKey)

        szsql += " order by nam_sort,nam_cdnam"


        prdt = DBUtil.GetDtDataDB(con, szsql, tran)

        If prdt.Rows.Count > 0 Then

            ''最大文字数を取得
            Dim iMaxLen As Integer = getMaxLength("nam_cdnam")

            '空項目を追加するか？
            Dim j As Integer = 0
            If prmSpace Then
                vl.ValueListItems.Add("", "")
                j += 1
            End If


            For Each dr As DataRow In prdt.Rows

                vl.ValueListItems.Add(PBCStr(dr.Item("nam_cdnam")),
                                              PBCStr(dr.Item("nam_name")))
                'PBCStr(dr.Item("item_code")).PadRight(iMaxLen) & "｜" & PBCStr(dr.Item("cust_name")))
                j += 1

            Next

        End If

        Return vl

    End Function
    ''' <summary>
    ''' 最大文字数の取得
    ''' </summary>
    ''' <param name="prmNameField"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Shared Function getMaxLength(prmNameField As String) As Integer
        Dim iMaxLength As Integer = 0
        For Each dr As DataRow In prdt.Rows
            If PBCStr(dr.Item(prmNameField)).Length > iMaxLength Then
                iMaxLength = PBCStr(dr.Item(prmNameField)).Length
            End If
        Next

        Return iMaxLength

    End Function
End Class
