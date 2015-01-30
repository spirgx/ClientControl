Module SubProgram

#Region "共用變數/常數宣告區"

    Public gDBConn As System.Data.SqlClient.SqlConnection   '資料庫連線物件
    Public gDBConnString As System.Data.SqlClient.SqlConnectionStringBuilder   '資料庫連線字串物件
    Public UserName As String = "None"
    Public Permission As Integer = 0
    Public QueryStr As String
    Public QuerySet As System.Data.DataSet
    Public QueryAdapt As System.Data.SqlClient.SqlDataAdapter

#End Region

#Region "限制輸入數字並只到小數點後位"
    Public Function OutputFloat(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs, ByVal num As Integer) As Boolean
        '輸入為數字，小數點或backspace  
        If Char.IsDigit(e.KeyChar) Or e.KeyChar = "." Or e.KeyChar = Chr(8) Then 'Backspace: 
            '輸入的小數點為唯一  
            If e.KeyChar = "." And InStr(CType(sender, System.Windows.Forms.TextBox).Text, ".") > 0 Then
                e.Handled = True
            Else
                '小數最多2位  
                If e.KeyChar <> Chr(8) And InStr(CType(sender, System.Windows.Forms.TextBox).Text, ".") > 0 Then
                    Dim sAry() As String = CType(sender, System.Windows.Forms.TextBox).Text.Split(".")
                    If sAry(1).Length >= num Then
                        e.Handled = True
                    Else
                        e.Handled = False
                    End If
                Else
                    e.Handled = False
                End If
            End If
            '輸入的負號是否在第一位  
        ElseIf e.KeyChar = "-" And CType(sender, System.Windows.Forms.TextBox).Text = "" Then
            e.Handled = False
        Else
            e.Handled = True
        End If
    End Function
#End Region

#Region "與資料庫連線相關程式段"
    Public Function ConnectToDatabase() As Boolean
        '===================================================================
        '程式段目的：連線至SQL Server資料庫
        '連線設定來源：設定檔
        '設定參數：WindowsAuthentication -> SQL認證模式，True為Windows認證。
        '設定參數：SQLServerLocation -> SQL Server位址
        '設定參數：SQLUserID -> 登入SQL Server的帳號
        '設定參數：SQLUserPassword -> 登入SQL Server的密碼
        '設定參數：DatabaseName -> 連線資料庫名稱
        '===================================================================
        Dim ReturnFlag As Boolean = False
        Try
            gDBConnString = New System.Data.SqlClient.SqlConnectionStringBuilder()
            With gDBConnString
                .DataSource = Trim(My.Settings.SQLServerLocation.ToString)
                If CBool(My.Settings.WindowsAuthentication.ToString) Then
                    .IntegratedSecurity = True
                Else
                    .IntegratedSecurity = False
                    .UserID = Trim(My.Settings.SQLUserID.ToString)
                    .Password = Trim(My.Settings.SQLUserPassword.ToString)
                End If
                .InitialCatalog = Trim(My.Settings.DatabaseName.ToString)
            End With

            gDBConn = New System.Data.SqlClient.SqlConnection(gDBConnString.ConnectionString)
            gDBConn.Open()
            ReturnFlag = True
        Catch ex As Exception
            MsgBox("資料庫連線失敗", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "錯誤")
            Exit Function
        End Try
        ConnectToDatabase = ReturnFlag
    End Function

    Public Function DisconnectFromDatabase() As Boolean
        '===================================================================
        '程式段目的：與資料庫斷線並將資源回收
        '===================================================================
        Dim ReturnFlag As Boolean = False
        Try
            If gDBConn.State Then
                gDBConnString = Nothing
                gDBConn.Close()
                gDBConn.Dispose()
            End If
            ReturnFlag = True
        Catch ex As Exception
            MsgBox("資料庫結束連線失敗", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "錯誤")
        End Try

        GC.Collect()
        DisconnectFromDatabase = ReturnFlag
    End Function
#End Region

End Module