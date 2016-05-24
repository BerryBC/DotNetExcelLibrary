Imports System.Data.OleDb
Imports Excel = Microsoft.Office.Interop.Excel
Imports Microsoft.Office
Imports System.Runtime.InteropServices

Public Class LoadExcel
    Implements IDisposable

    Public Property LastException As Exception

    Private strFileName As String
    Public strSheets As New List(Of String)



    Public Sub New()
    End Sub
    ''' <summary>
    ''' File to get information from
    ''' </summary>
    ''' <param name="FileName"></param>
    ''' <remarks>
    ''' The caller is responsible to ensure the file exists.
    ''' </remarks>
    Public Sub New(ByVal FileName As String)
        strFileName = FileName
    End Sub


    Private Function ConnectionString(ByVal FileName As String) As String
        Dim Builder As New OleDb.OleDbConnectionStringBuilder
        Try


            If IO.Path.GetExtension(FileName).ToUpper = ".XLS" Then
                Builder.Provider = "Microsoft.Jet.OLEDB.4.0"
                Builder.Add("Extended Properties", "Excel 8.0;IMEX=6;HDR=Yes;")
            ElseIf IO.Path.GetExtension(FileName).ToUpper = ".XLSX" Then
                Builder.Provider = "Microsoft.ACE.OLEDB.12.0"
                Builder.Add("Extended Properties", "Excel 12.0;IMEX=6;HDR=Yes;")
            End If

            Builder.DataSource = FileName

            Return Builder.ToString
        Catch ex As Exception
            Throw New Exception(ex.Message, ex)
        End Try
        Return ""

    End Function




    ''' <summary>
    ''' Retrieve worksheet and name range names.
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetInformation() As Boolean
        Dim Success As Boolean = True
        Dim xlApp As Excel.Application = Nothing
        Dim xlWorkBooks As Excel.Workbooks = Nothing
        Dim xlWorkBook As Excel.Workbook = Nothing
        Dim xlWorkSheets As Excel.Sheets = Nothing

        Try


            If Not IO.File.Exists(strFileName) Then
                Dim ex As New Exception("Failed to locate '" & strFileName & "'")
                _LastException = ex
                Throw ex
            End If

            strSheets.Clear()



            xlApp = New Excel.Application
            xlApp.DisplayAlerts = False
            xlWorkBooks = xlApp.Workbooks
            xlWorkBook = xlWorkBooks.Open(strFileName)


            xlWorkSheets = xlWorkBook.Sheets

            For x As Integer = 1 To xlWorkSheets.Count
                Dim Sheet1 As Excel.Worksheet = CType(xlWorkSheets(x), Excel.Worksheet)
                strSheets.Add(Sheet1.Name)
                Runtime.InteropServices.Marshal.FinalReleaseComObject(Sheet1)
                Sheet1 = Nothing
            Next

            xlWorkBook.Close()
            xlApp.UserControl = True
            xlApp.Quit()

        Catch ex As Exception
            Throw New Exception(ex.Message, ex)
        Finally

            If Not xlWorkSheets Is Nothing Then
                Marshal.FinalReleaseComObject(xlWorkSheets)
                xlWorkSheets = Nothing
            End If

            If Not xlWorkBook Is Nothing Then
                Marshal.FinalReleaseComObject(xlWorkBook)
                xlWorkBook = Nothing
            End If

            If Not xlWorkBooks Is Nothing Then
                Marshal.FinalReleaseComObject(xlWorkBooks)
                xlWorkBooks = Nothing
            End If

            If Not xlApp Is Nothing Then
                Marshal.FinalReleaseComObject(xlApp)
                xlApp = Nothing
            End If
        End Try

        Return Success

    End Function


    Public Function SaveToExistsFile(ByRef dtFromData As DataTable, ByRef strSheetName As String) As Boolean
        Dim Success As Boolean = True
        Dim i As Integer
        Dim Connection_String As String = ""
        Dim intHowManyCol As Integer
        Dim strTmpValues As String = ""
        Dim cnConnection As OleDb.OleDbConnection
        Dim adaptData As New OleDbDataAdapter
        Dim dtData As DataTable

        Try

            If dtFromData.Rows.Count > 0 Then



                Connection_String = ConnectionString(strFileName)

                cnConnection = New OleDb.OleDbConnection(ConnectionString(strFileName))
                cnConnection.Open()



                '看看数据源中的列数
                intHowManyCol = dtFromData.Columns.Count

                '创建插入语句的语句(对应变量)
                For i = 1 To intHowManyCol
                    strTmpValues = strTmpValues & " @T" & i & " ,"

                Next i
                strTmpValues = strTmpValues.Substring(0, strTmpValues.Length - 1)



                '--------------弄个新的Datatable----------
                dtData = ReturnNewNormalDT(dtFromData, dtFromData)


                '对应数据源与SQL语句
                'Dim cmd As New OleDbCommand("INSERT INTO  [" & strSheetName & "$]  VALUES (" & strTmpValues & ")", cnConnection)

                Dim cmd = New OleDbCommand("INSERT INTO  [" & strSheetName & "$]  VALUES (" & strTmpValues & ")", cnConnection)
                With cmd
                    .CommandType = CommandType.Text
                    For i = 1 To intHowManyCol
                        .Parameters.Add(New OleDb.OleDbParameter("@T" & i, TranType2OLE(dtData.Columns(i - 1).DataType.ToString)))
                        .Parameters("@T" & i).SourceColumn = dtData.Columns(i - 1).ColumnName
                    Next i

                End With
                '插入!!
                adaptData.InsertCommand = cmd
                Dim builder As New OleDbCommandBuilder(adaptData)
                builder.QuotePrefix = "["
                builder.QuoteSuffix = "]"
                adaptData.Update(dtData)
                cnConnection.Close()


            End If




        Catch ex As Exception
            Throw New Exception(ex.Message, ex)
        Finally

            If cnConnection IsNot Nothing Then
                If ((cnConnection.State = ConnectionState.Open)) Then
                    cnConnection.Close()
                End If

            End If

        End Try
        Return Success
    End Function





    Public Function SaveAsNewFromDatatable(ByRef dtFromData As DataTable, ByRef strSheetName As String) As Boolean
        Dim Success As Boolean = True
        Dim xlApp As Excel.Application = Nothing
        Dim xlWorkBook As Excel.Workbook = Nothing
        Dim xlWorkSheet As Excel.Worksheet = Nothing
        Dim dcCol As DataColumn
        Dim i As Integer
        Dim Connection_String As String = ""
        Dim intHowManyCol As Integer
        Dim strTmpValues As String = ""
        Dim cnConnection As OleDb.OleDbConnection
        Dim adaptData As New OleDbDataAdapter
        Dim dtData As DataTable
        Dim drDataOrg As DataRow

        Try
            xlApp = New Excel.Application
            xlApp.DisplayAlerts = False
            xlWorkBook = xlApp.Workbooks.Add
            xlWorkSheet = xlWorkBook.Sheets(1)

            If dtFromData.Rows.Count > 0 Then
                drDataOrg = dtFromData.Rows(0)
            End If


            i = 1
            For Each dcCol In dtFromData.Columns
                xlWorkSheet.Cells(1, i) = dcCol.Caption
                If dtFromData.Rows.Count > 0 Then
                    xlWorkSheet.Cells(2, i) = drDataOrg.Item(i - 1)
                End If
                i += 1
            Next
            xlWorkBook.SaveAs(strFileName)
            xlWorkSheet = Nothing
            xlWorkBook.Close()
            xlWorkBook = Nothing
            xlApp.Quit()
            xlApp = Nothing


            If dtFromData.Rows.Count > 0 Then



                Connection_String = ConnectionString(strFileName)

                cnConnection = New OleDb.OleDbConnection(ConnectionString(strFileName))
                cnConnection.Open()



                '看看数据源中的列数
                intHowManyCol = dtFromData.Columns.Count

                '创建插入语句的语句(对应变量)
                For i = 1 To intHowManyCol
                    strTmpValues = strTmpValues & " @T" & i & " ,"

                Next i
                strTmpValues = strTmpValues.Substring(0, strTmpValues.Length - 1)



                '--------------弄个新的Datatable----------
                dtData = ReturnNewNormalDT(dtFromData, dtFromData)


                '对应数据源与SQL语句
                'Dim cmd As New OleDbCommand("INSERT INTO  [" & strSheetName & "$]  VALUES (" & strTmpValues & ")", cnConnection)

                Dim cmd = New OleDbCommand("INSERT INTO  [" & strSheetName & "$]  VALUES (" & strTmpValues & ")", cnConnection)
                With cmd
                    .CommandType = CommandType.Text
                    For i = 1 To intHowManyCol
                        .Parameters.Add(New OleDb.OleDbParameter("@T" & i, TranType2OLE(dtData.Columns(i - 1).DataType.ToString)))
                        .Parameters("@T" & i).SourceColumn = dtData.Columns(i - 1).ColumnName
                    Next i

                End With
                '插入!!
                adaptData.InsertCommand = cmd
                Dim builder As New OleDbCommandBuilder(adaptData)
                builder.QuotePrefix = "["
                builder.QuoteSuffix = "]"
                adaptData.Update(dtData)
                cnConnection.Close()


            End If

            xlApp = New Excel.Application
            xlApp.DisplayAlerts = False
            xlWorkBook = xlApp.Workbooks.Open(strFileName)
            xlWorkSheet = xlWorkBook.Sheets(1)

            xlWorkSheet.Rows(2).Delete
            xlWorkBook.Save()
            xlWorkSheet = Nothing
            xlWorkBook.Close()
            xlWorkBook = Nothing
            xlApp.Quit()
            xlApp = Nothing



        Catch ex As Exception
            Throw New Exception(ex.Message, ex)
        Finally

            If Not xlWorkSheet Is Nothing Then
                Marshal.FinalReleaseComObject(xlWorkSheet)
                xlWorkSheet = Nothing
            End If

            If Not xlWorkBook Is Nothing Then
                Marshal.FinalReleaseComObject(xlWorkBook)
                xlWorkBook = Nothing
            End If

            If Not xlApp Is Nothing Then
                Marshal.FinalReleaseComObject(xlApp)
                xlApp = Nothing
            End If
            If cnConnection IsNot Nothing Then
                If ((cnConnection.State = ConnectionState.Open)) Then
                    cnConnection.Close()
                End If

            End If

        End Try
        Return Success
    End Function

    ''' <summary>
    ''' 根据给出的格式表返回根据格式表格式的数据
    ''' </summary>
    ''' <param name="dtData">数据表</param>
    ''' <param name="dtFormat">格式表</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function ReturnNewNormalDT(ByVal dtData As DataTable, ByVal dtFormat As DataTable) As DataTable
        Dim intHowManyCol As Integer
        Dim i As Integer
        Dim j As Integer
        Dim dtDataNew As New DataTable
        Dim drTmp As DataRow
        Dim intTmpListOfTitle() As Integer

        Try


            ReDim intTmpListOfTitle(0 To (dtData.Columns.Count - 1))


            '读取格式表的列数
            intHowManyCol = dtFormat.Columns.Count

            '创建新数据表的列标题以及数据类型
            For i = 1 To intHowManyCol
                dtDataNew.Columns.Add(New System.Data.DataColumn(dtFormat.Columns(i - 1).ColumnName, dtFormat.Columns(i - 1).DataType))
            Next i


            intHowManyCol = dtData.Columns.Count

            For i = 1 To intHowManyCol
                intTmpListOfTitle(i - 1) = dtFormat.Columns(dtData.Columns(i - 1).ColumnName).Ordinal
            Next i




            '每行数据表的来读
            For i = 1 To dtData.Rows.Count
                drTmp = dtDataNew.NewRow
                For j = 1 To dtData.Columns.Count
                    Try
                        If dtData.Rows(i - 1).Item(j - 1).ToString = "" Or dtData.Rows(i - 1).Item(j - 1).ToString = "#DIV/0" Then
                            drTmp(intTmpListOfTitle(j - 1)) = DBNull.Value
                        Else
                            drTmp(intTmpListOfTitle(j - 1)) = dtData.Rows(i - 1).Item(j - 1)
                        End If
                    Catch
                        drTmp(intTmpListOfTitle(j - 1)) = DBNull.Value
                    End Try
                Next j
                dtDataNew.Rows.Add(drTmp)
            Next i
        Catch ex As Exception
            Throw New Exception(ex.Message, ex)
        End Try
        Return dtDataNew
    End Function


    ''' <summary>
    ''' 把.Net数据类型编程OLE数据类型
    ''' </summary>
    ''' <param name="strTable"></param>
    ''' <returns>返回OLE的数据类型</returns>
    Private Function TranType2OLE(ByVal strTable As String) As OleDbType
        If strTable = "System.Int32" Then
            Return OleDb.OleDbType.Integer
        ElseIf strTable = "System.String" Then
            Return OleDb.OleDbType.VarChar
        ElseIf strTable = "System.Double" Then
            Return OleDb.OleDbType.Double
        ElseIf strTable = "System.DateTime" Then
            Return OleDb.OleDbType.Date
        ElseIf strTable = "System.Single" Then
            Return OleDb.OleDbType.Single
        End If
        Return OleDb.OleDbType.Error
    End Function



    Public Function GetData(ByVal strTableName As String) As DataTable
        Dim dtExl As New DataTable
        Dim strSelectStatement As String

        Dim Connection_String As String = ""
        Connection_String = ConnectionString(strFileName)
        strSelectStatement = "SELECT * FROM [" & strTableName & "$]"

        Try
            Using cn As New OleDbConnection With {.ConnectionString = Connection_String}
                Using cmd As New OleDbCommand With {.Connection = cn, .CommandText = strSelectStatement}
                    cn.Open()
                    dtExl.Load(cmd.ExecuteReader)
                End Using
            End Using
        Catch ex As Exception
            Throw New Exception(ex.Message, ex)
        End Try

        Return dtExl
    End Function


    Public Function ReturnFormat(ByVal strTable As String) As DataTable
        Dim dtExl As New DataTable
        Dim strSelectStatement As String

        Dim Connection_String As String = ""
        Connection_String = ConnectionString(strFileName)
        strSelectStatement = "SELECT top 1 * FROM [" & strTable & "$]"

        Try
            Using cn As New OleDbConnection With {.ConnectionString = Connection_String}
                Using cmd As New OleDbCommand With {.Connection = cn, .CommandText = strSelectStatement}
                    cn.Open()
                    dtExl.Load(cmd.ExecuteReader)
                End Using
            End Using
        Catch ex As Exception
            Throw New Exception(ex.Message, ex)
        End Try

        Return dtExl
    End Function


#Region "IDisposable Support"
    Private disposedValue As Boolean ' 检测冗余的调用

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                ' TODO:  释放托管状态(托管对象)。
            End If

            ' TODO:  释放非托管资源(非托管对象)并重写下面的 Finalize()。
            ' TODO:  将大型字段设置为 null。
        End If
        Me.disposedValue = True
    End Sub

    ' TODO:  仅当上面的 Dispose(ByVal disposing As Boolean)具有释放非托管资源的代码时重写 Finalize()。
    'Protected Overrides Sub Finalize()
    '    ' 不要更改此代码。    请将清理代码放入上面的 Dispose(ByVal disposing As Boolean)中。
    '    Dispose(False)
    '    MyBase.Finalize()
    'End Sub

    ' Visual Basic 添加此代码是为了正确实现可处置模式。
    Public Sub Dispose() Implements IDisposable.Dispose
        ' 不要更改此代码。    请将清理代码放入上面的 Dispose (disposing As Boolean)中。
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub
#End Region

End Class
