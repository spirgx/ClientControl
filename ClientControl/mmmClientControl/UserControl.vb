Public Class UserControl

    Dim QueryDate As String
    Public xlApp As Microsoft.Office.Interop.Excel.Application
    Public xlBook As Microsoft.Office.Interop.Excel.Workbook
    Public xlSheet As Microsoft.Office.Interop.Excel.Worksheet
    Public xlRange As Microsoft.Office.Interop.Excel.Range

    Private Sub UserControl_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        If ConnectToDatabase() Then
            DisconnectFromDatabase()
        End If

    End Sub

    Private Sub FitWetTextBox_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles FitWetTextBox.KeyPress
        OutputFloat(sender, e, 3)
    End Sub

    Private Sub FitMaterialTextBox_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles FitMaterialTextBox.KeyPress
        OutputFloat(sender, e, 3)
    End Sub

    Private Sub FitInsertButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FitInsertButton.Click

        ConnectToDatabase()

        QueryDate = ProductDateTextBox.Text

        Dim IdQueryStr As String
        Dim IdQuerySet As System.Data.DataSet
        Dim IdQueryAdapt As System.Data.SqlClient.SqlDataAdapter
        Dim IdPercent As String = ""
        Dim IdPrice As String = ""

        IdQueryStr = "Select * From RMDatabase Where [New ID] = '" & FitStockNoTextBox.Text.Trim & "'"
        IdQueryAdapt = New System.Data.SqlClient.SqlDataAdapter(IdQueryStr, gDBConn)
        IdQuerySet = New System.Data.DataSet
        IdQueryAdapt.Fill(IdQuerySet, "DataSet")

        If IdQuerySet.Tables("DataSet").Rows.Count > 0 Then
            IdPercent = Trim(IdQuerySet.Tables("DataSet").Rows(0).Item(4).ToString)
            IdPrice = Trim(IdQuerySet.Tables("DataSet").Rows(0).Item(5).ToString)
        Else
            MsgBox("Stock No. isn't find.", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Insert")
            Exit Sub
        End If

        If MsgBox("Insert？", MsgBoxStyle.Question + MsgBoxStyle.YesNo, "Question") = MsgBoxResult.Yes Then

            QueryStr = "INSERT INTO FiberInputTarget( ProductNo, ID, [Material %], [Kg/sq M], [Foul. solid%], [Foul. Cost (NT/Kg)] )" & _
                        "VALUES ( '" & ProductNoTextBox.Text.Trim & _
                        "', '" & FitStockNoTextBox.Text.Trim & _
                        "', '" & FitMaterialTextBox.Text.Trim & _
                        "', '" & Double.Parse(FitWetTextBox.Text.Trim) * (Double.Parse(FitMaterialTextBox.Text.Trim) / 100.0) & _
                        "', '" & (Double.Parse(FitMaterialTextBox.Text.Trim) / 100.0) * Double.Parse(IdPercent) & _
                        "', '" & Double.Parse(FitWetTextBox.Text.Trim) * (Double.Parse(FitMaterialTextBox.Text.Trim) / 100.0) * Double.Parse(IdPrice) & "')"

            QueryAdapt = New System.Data.SqlClient.SqlDataAdapter(QueryStr, gDBConn)
            QuerySet = New System.Data.DataSet
            QueryAdapt.Fill(QuerySet, "FiberInputTarget")

            QueryStr = "Select FiberInputTarget.ID, RMDatabase.[RM Description] ,FiberInputTarget.[Material %], FiberInputTarget.[Kg/sq M], RMDatabase.[Solid%], " & _
                        "FiberInputTarget.[Foul. solid%], RMDatabase.[Price (NT/Kg)], FiberInputTarget.[Foul. Cost (NT/Kg)] From RMDatabase, FiberInputTarget " & _
                        "Where RMDatabase.[New ID] = FiberInputTarget.ID And ProductNo = '" & ProductNoTextBox.Text.Trim & "'"
            QueryAdapt = New System.Data.SqlClient.SqlDataAdapter(QueryStr, gDBConn)
            QuerySet = New System.Data.DataSet
            QueryAdapt.Fill(QuerySet, "DataSet1")

            Dim dr As DataRow
            dr = QuerySet.Tables("DataSet1").NewRow

            dr.Item(0) = "Total"
            dr.Item("Material %") = QuerySet.Tables("DataSet1").Compute("SUM([Material %])", "True")
            dr.Item("Kg/sq M") = QuerySet.Tables("DataSet1").Compute("SUM([Kg/sq M])", "True")
            dr.Item("Foul. solid%") = QuerySet.Tables("DataSet1").Compute("SUM([Foul. solid%])", "True")
            dr.Item("Foul. Cost (NT/Kg)") = QuerySet.Tables("DataSet1").Compute("SUM([Foul. Cost (NT/Kg)])", "True")

            QuerySet.Tables("DataSet1").Rows.Add(dr)

            FitDataGridView.DataSource = QuerySet.Tables("DataSet1")
            FitDataGridView.Columns(0).Width = 150
            FitDataGridView.Columns(1).Width = 450
            FitDataGridView.Columns(2).Width = 150
            FitDataGridView.Columns(3).Width = 150
            FitDataGridView.Columns(4).Width = 150
            FitDataGridView.Columns(5).Width = 150
            FitDataGridView.Columns(6).Width = 150
            FitDataGridView.Columns(7).Width = 150

        End If

        DisconnectFromDatabase()

    End Sub

    Private Sub ProductSearchButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ProductSearchButton.Click

        ConnectToDatabase()

        QueryDate = ProductDateTextBox.Text

        QueryStr = "Select FiberInputTarget.ID, RMDatabase.[RM Description] ,FiberInputTarget.[Material %], FiberInputTarget.[Kg/sq M], RMDatabase.[Solid%], " & _
                        "FiberInputTarget.[Foul. solid%], RMDatabase.[Price (NT/Kg)], FiberInputTarget.[Foul. Cost (NT/Kg)] From RMDatabase, FiberInputTarget " & _
                        "Where RMDatabase.[New ID] = FiberInputTarget.ID And ProductNo = '" & ProductNoTextBox.Text.Trim & "'"
        'If QueryDate <> "" Then
        'QueryStr = QueryStr & " And ProductDate = '" & QueryDate & "'"
        'End If

        QueryAdapt = New System.Data.SqlClient.SqlDataAdapter(QueryStr, gDBConn)
        QuerySet = New System.Data.DataSet
        QueryAdapt.Fill(QuerySet, "DataSet1")

        Dim dr As DataRow
        dr = QuerySet.Tables("DataSet1").NewRow

        dr.Item(0) = "Total"
        dr.Item("Material %") = QuerySet.Tables("DataSet1").Compute("SUM([Material %])", "True")
        dr.Item("Kg/sq M") = QuerySet.Tables("DataSet1").Compute("SUM([Kg/sq M])", "True")
        dr.Item("Foul. solid%") = QuerySet.Tables("DataSet1").Compute("SUM([Foul. solid%])", "True")
        dr.Item("Foul. Cost (NT/Kg)") = QuerySet.Tables("DataSet1").Compute("SUM([Foul. Cost (NT/Kg)])", "True")


        QuerySet.Tables("DataSet1").Rows.Add(dr)

        FitDataGridView.DataSource = QuerySet.Tables("DataSet1")
        FitDataGridView.Columns(0).Width = 150
        FitDataGridView.Columns(1).Width = 450
        FitDataGridView.Columns(2).Width = 150
        FitDataGridView.Columns(3).Width = 150
        FitDataGridView.Columns(4).Width = 150
        FitDataGridView.Columns(5).Width = 150
        FitDataGridView.Columns(6).Width = 150
        FitDataGridView.Columns(7).Width = 150


        QueryStr = "Select RollCoatTarget.ID, RMDatabase.[RM Description] ,RollCoatTarget.[Material %], RollCoatTarget.[Kg/sq M], RMDatabase.[Solid%], " & _
                        "RollCoatTarget.[Foul. solid%], RMDatabase.[Price (NT/Kg)], RollCoatTarget.[Foul. Cost (NT/Kg)] From RMDatabase, RollCoatTarget " & _
                        "Where RMDatabase.[New ID] = RollCoatTarget.ID And ProductNo = '" & ProductNoTextBox.Text.Trim & "'"
        'If QueryDate <> "" Then
        'QueryStr = QueryStr & " And ProductDate = '" & QueryDate & "'"
        'End If

        QueryAdapt = New System.Data.SqlClient.SqlDataAdapter(QueryStr, gDBConn)
        QuerySet = New System.Data.DataSet
        QueryAdapt.Fill(QuerySet, "DataSet2")

        dr = QuerySet.Tables("DataSet2").NewRow

        dr.Item(0) = "Total"
        dr.Item("Material %") = QuerySet.Tables("DataSet2").Compute("SUM([Material %])", "True")
        dr.Item("Kg/sq M") = QuerySet.Tables("DataSet2").Compute("SUM([Kg/sq M])", "True")
        dr.Item("Foul. solid%") = QuerySet.Tables("DataSet2").Compute("SUM([Foul. solid%])", "True")
        dr.Item("Foul. Cost (NT/Kg)") = QuerySet.Tables("DataSet2").Compute("SUM([Foul. Cost (NT/Kg)])", "True")


        QuerySet.Tables("DataSet2").Rows.Add(dr)

        RctDataGridView.DataSource = QuerySet.Tables("DataSet2")
        RctDataGridView.Columns(0).Width = 150
        RctDataGridView.Columns(1).Width = 450
        RctDataGridView.Columns(2).Width = 150
        RctDataGridView.Columns(3).Width = 150
        RctDataGridView.Columns(4).Width = 150
        RctDataGridView.Columns(5).Width = 150
        RctDataGridView.Columns(6).Width = 150
        RctDataGridView.Columns(7).Width = 150



        QueryStr = "Select MakeCoat1Target.ID, RMDatabase.[RM Description] ,MakeCoat1Target.[Material %], MakeCoat1Target.[Kg/sq M], RMDatabase.[Solid%], " & _
                    "MakeCoat1Target.[Foul. solid%], RMDatabase.[Price (NT/Kg)], MakeCoat1Target.[Foul. Cost (NT/Kg)] From RMDatabase, MakeCoat1Target " & _
                    "Where RMDatabase.[New ID] = MakeCoat1Target.ID And ProductNo = '" & ProductNoTextBox.Text.Trim & "'"
        'If QueryDate <> "" Then
        'QueryStr = QueryStr & " And ProductDate = '" & QueryDate & "'"
        'End If

        QueryAdapt = New System.Data.SqlClient.SqlDataAdapter(QueryStr, gDBConn)
        QuerySet = New System.Data.DataSet
        QueryAdapt.Fill(QuerySet, "DataSet3")

        dr = QuerySet.Tables("DataSet3").NewRow

        dr.Item(0) = "Total"
        dr.Item("Material %") = QuerySet.Tables("DataSet3").Compute("SUM([Material %])", "True")
        dr.Item("Kg/sq M") = QuerySet.Tables("DataSet3").Compute("SUM([Kg/sq M])", "True")
        dr.Item("Foul. solid%") = QuerySet.Tables("DataSet3").Compute("SUM([Foul. solid%])", "True")
        dr.Item("Foul. Cost (NT/Kg)") = QuerySet.Tables("DataSet3").Compute("SUM([Foul. Cost (NT/Kg)])", "True")


        QuerySet.Tables("DataSet3").Rows.Add(dr)


        Mc1tDataGridView.DataSource = QuerySet.Tables("DataSet3")
        Mc1tDataGridView.Columns(0).Width = 150
        Mc1tDataGridView.Columns(1).Width = 450
        Mc1tDataGridView.Columns(2).Width = 150
        Mc1tDataGridView.Columns(3).Width = 150
        Mc1tDataGridView.Columns(4).Width = 150
        Mc1tDataGridView.Columns(5).Width = 150
        Mc1tDataGridView.Columns(6).Width = 150
        Mc1tDataGridView.Columns(7).Width = 150




        QueryStr = "Select SprayCoatTarget.ID, RMDatabase.[RM Description] ,SprayCoatTarget.[Material %], SprayCoatTarget.[Kg/sq M], RMDatabase.[Solid%], " & _
                    "SprayCoatTarget.[Foul. solid%], RMDatabase.[Price (NT/Kg)], SprayCoatTarget.[Foul. Cost (NT/Kg)] From RMDatabase, SprayCoatTarget " & _
                    "Where RMDatabase.[New ID] = SprayCoatTarget.ID And ProductNo = '" & ProductNoTextBox.Text.Trim & "'"
        'If QueryDate <> "" Then
        'QueryStr = QueryStr & " And ProductDate = '" & QueryDate & "'"
        'End If

        QueryAdapt = New System.Data.SqlClient.SqlDataAdapter(QueryStr, gDBConn)
        QuerySet = New System.Data.DataSet
        QueryAdapt.Fill(QuerySet, "DataSet4")

        dr = QuerySet.Tables("DataSet4").NewRow

        dr.Item(0) = "Total"
        dr.Item("Material %") = QuerySet.Tables("DataSet4").Compute("SUM([Material %])", "True")
        dr.Item("Kg/sq M") = QuerySet.Tables("DataSet4").Compute("SUM([Kg/sq M])", "True")
        dr.Item("Foul. solid%") = QuerySet.Tables("DataSet4").Compute("SUM([Foul. solid%])", "True")
        dr.Item("Foul. Cost (NT/Kg)") = QuerySet.Tables("DataSet4").Compute("SUM([Foul. Cost (NT/Kg)])", "True")


        QuerySet.Tables("DataSet4").Rows.Add(dr)

        SctDataGridView.DataSource = QuerySet.Tables("DataSet4")
        SctDataGridView.Columns(0).Width = 150
        SctDataGridView.Columns(1).Width = 450
        SctDataGridView.Columns(2).Width = 150
        SctDataGridView.Columns(3).Width = 150
        SctDataGridView.Columns(4).Width = 150
        SctDataGridView.Columns(5).Width = 150
        SctDataGridView.Columns(6).Width = 150
        SctDataGridView.Columns(7).Width = 150


        QueryStr = "Select SprayCoat2Target.ID, RMDatabase.[RM Description] ,SprayCoat2Target.[Material %], SprayCoat2Target.[Kg/sq M], RMDatabase.[Solid%], " & _
                        "SprayCoat2Target.[Foul. solid%], RMDatabase.[Price (NT/Kg)], SprayCoat2Target.[Foul. Cost (NT/Kg)] From RMDatabase, SprayCoat2Target " & _
                        "Where RMDatabase.[New ID] = SprayCoat2Target.ID And ProductNo = '" & ProductNoTextBox.Text.Trim & "'"
        'If QueryDate <> "" Then
        'QueryStr = QueryStr & " And ProductDate = '" & QueryDate & "'"
        'End If

        QueryAdapt = New System.Data.SqlClient.SqlDataAdapter(QueryStr, gDBConn)
        QuerySet = New System.Data.DataSet
        QueryAdapt.Fill(QuerySet, "DataSet5")

        dr = QuerySet.Tables("DataSet5").NewRow

        dr.Item(0) = "Total"
        dr.Item("Material %") = QuerySet.Tables("DataSet5").Compute("SUM([Material %])", "True")
        dr.Item("Kg/sq M") = QuerySet.Tables("DataSet5").Compute("SUM([Kg/sq M])", "True")
        dr.Item("Foul. solid%") = QuerySet.Tables("DataSet5").Compute("SUM([Foul. solid%])", "True")
        dr.Item("Foul. Cost (NT/Kg)") = QuerySet.Tables("DataSet5").Compute("SUM([Foul. Cost (NT/Kg)])", "True")


        QuerySet.Tables("DataSet5").Rows.Add(dr)

        Sc2tDataGridView.DataSource = QuerySet.Tables("DataSet5")
        Sc2tDataGridView.Columns(0).Width = 150
        Sc2tDataGridView.Columns(1).Width = 450
        Sc2tDataGridView.Columns(2).Width = 150
        Sc2tDataGridView.Columns(3).Width = 150
        Sc2tDataGridView.Columns(4).Width = 150
        Sc2tDataGridView.Columns(5).Width = 150
        Sc2tDataGridView.Columns(6).Width = 150
        Sc2tDataGridView.Columns(7).Width = 150



        QueryStr = "Select SizeCoatTarget.ID, RMDatabase.[RM Description] ,SizeCoatTarget.[Material %], SizeCoatTarget.[Kg/sq M], RMDatabase.[Solid%], " & _
                        "SizeCoatTarget.[Foul. solid%], RMDatabase.[Price (NT/Kg)], SizeCoatTarget.[Foul. Cost (NT/Kg)] From RMDatabase, SizeCoatTarget " & _
                        "Where RMDatabase.[New ID] = SizeCoatTarget.ID And ProductNo = '" & ProductNoTextBox.Text.Trim & "'"
        'If QueryDate <> "" Then
        'QueryStr = QueryStr & " And ProductDate = '" & QueryDate & "'"
        'End If

        QueryAdapt = New System.Data.SqlClient.SqlDataAdapter(QueryStr, gDBConn)
        QuerySet = New System.Data.DataSet
        QueryAdapt.Fill(QuerySet, "DataSet6")

        dr = QuerySet.Tables("DataSet6").NewRow

        dr.Item(0) = "Total"
        dr.Item("Material %") = QuerySet.Tables("DataSet6").Compute("SUM([Material %])", "True")
        dr.Item("Kg/sq M") = QuerySet.Tables("DataSet6").Compute("SUM([Kg/sq M])", "True")
        dr.Item("Foul. solid%") = QuerySet.Tables("DataSet6").Compute("SUM([Foul. solid%])", "True")
        dr.Item("Foul. Cost (NT/Kg)") = QuerySet.Tables("DataSet6").Compute("SUM([Foul. Cost (NT/Kg)])", "True")


        QuerySet.Tables("DataSet6").Rows.Add(dr)


        SizeCTDataGridView.DataSource = QuerySet.Tables("DataSet6")
        SizeCTDataGridView.Columns(0).Width = 150
        SizeCTDataGridView.Columns(1).Width = 450
        SizeCTDataGridView.Columns(2).Width = 150
        SizeCTDataGridView.Columns(3).Width = 150
        SizeCTDataGridView.Columns(4).Width = 150
        SizeCTDataGridView.Columns(5).Width = 150
        SizeCTDataGridView.Columns(6).Width = 150
        SizeCTDataGridView.Columns(7).Width = 150


        QueryStr = "Select MineralCoatTarget.ID, RMDatabase.[RM Description] ,MineralCoatTarget.[Material %], MineralCoatTarget.[Kg/sq M], RMDatabase.[Solid%], " & _
                       "MineralCoatTarget.[Foul. solid%], RMDatabase.[Price (NT/Kg)], MineralCoatTarget.[Foul. Cost (NT/Kg)] From RMDatabase, MineralCoatTarget " & _
                       "Where RMDatabase.[New ID] = MineralCoatTarget.ID And ProductNo = '" & ProductNoTextBox.Text.Trim & "'"
        'If QueryDate <> "" Then
        'QueryStr = QueryStr & " And ProductDate = '" & QueryDate & "'"
        'End If

        QueryAdapt = New System.Data.SqlClient.SqlDataAdapter(QueryStr, gDBConn)
        QuerySet = New System.Data.DataSet
        QueryAdapt.Fill(QuerySet, "DataSet7")

        dr = QuerySet.Tables("DataSet7").NewRow

        dr.Item(0) = "Total"
        dr.Item("Material %") = QuerySet.Tables("DataSet7").Compute("SUM([Material %])", "True")
        dr.Item("Kg/sq M") = QuerySet.Tables("DataSet7").Compute("SUM([Kg/sq M])", "True")
        dr.Item("Foul. solid%") = QuerySet.Tables("DataSet7").Compute("SUM([Foul. solid%])", "True")
        dr.Item("Foul. Cost (NT/Kg)") = QuerySet.Tables("DataSet7").Compute("SUM([Foul. Cost (NT/Kg)])", "True")


        QuerySet.Tables("DataSet7").Rows.Add(dr)

        MctDataGridView.DataSource = QuerySet.Tables("DataSet7")
        MctDataGridView.Columns(0).Width = 150
        MctDataGridView.Columns(1).Width = 450
        MctDataGridView.Columns(2).Width = 150
        MctDataGridView.Columns(3).Width = 150
        MctDataGridView.Columns(4).Width = 150
        MctDataGridView.Columns(5).Width = 150
        MctDataGridView.Columns(6).Width = 150
        MctDataGridView.Columns(7).Width = 150

        DisconnectFromDatabase()

    End Sub

    Private Sub SettingFitButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        If FitWetTextBox.Enabled = False Then
            FitWetTextBox.Enabled = True
            FitWetTextBox.Focus()
            Exit Sub
        End If

        If FitWetTextBox.Enabled = True Then
            FitWetTextBox.Enabled = False
            Exit Sub
        End If
    End Sub

    Private Sub RctInsertButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RctInsertButton.Click

        ConnectToDatabase()

        QueryDate = ProductDateTextBox.Text

        Dim IdQueryStr As String
        Dim IdQuerySet As System.Data.DataSet
        Dim IdQueryAdapt As System.Data.SqlClient.SqlDataAdapter
        Dim IdPercent As String = ""
        Dim IdPrice As String = ""

        IdQueryStr = "Select * From RMDatabase Where [New ID] = '" & RctStockNoTextBox.Text.Trim & "'"
        IdQueryAdapt = New System.Data.SqlClient.SqlDataAdapter(IdQueryStr, gDBConn)
        IdQuerySet = New System.Data.DataSet
        IdQueryAdapt.Fill(IdQuerySet, "DataSet")

        If IdQuerySet.Tables("DataSet").Rows.Count > 0 Then
            IdPercent = Trim(IdQuerySet.Tables("DataSet").Rows(0).Item(4).ToString)
            IdPrice = Trim(IdQuerySet.Tables("DataSet").Rows(0).Item(5).ToString)
        Else
            MsgBox("Stock No. isn't find.", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Insert")
            Exit Sub
        End If

        If MsgBox("Insert？", MsgBoxStyle.Question + MsgBoxStyle.YesNo, "Question") = MsgBoxResult.Yes Then

            QueryStr = "INSERT INTO RollCoatTarget( ProductNo, ID, [Material %], [Kg/sq M], [Foul. solid%], [Foul. Cost (NT/Kg)] )" & _
                        "VALUES ( '" & ProductNoTextBox.Text.Trim & _
                        "', '" & RctStockNoTextBox.Text.Trim & _
                        "', '" & RctMaterialTextBox.Text.Trim & _
                        "', '" & Double.Parse(RctWetTextBox.Text.Trim) * (Double.Parse(RctMaterialTextBox.Text.Trim) / 100.0) & _
                        "', '" & (Double.Parse(RctMaterialTextBox.Text.Trim) / 100.0) * Double.Parse(IdPercent) & _
                        "', '" & Double.Parse(RctWetTextBox.Text.Trim) * (Double.Parse(RctMaterialTextBox.Text.Trim) / 100.0) * Double.Parse(IdPrice) & "')"

            QueryAdapt = New System.Data.SqlClient.SqlDataAdapter(QueryStr, gDBConn)
            QuerySet = New System.Data.DataSet
            QueryAdapt.Fill(QuerySet, "RollCoatTarget")

            QueryStr = "Select RollCoatTarget.ID, RMDatabase.[RM Description] ,RollCoatTarget.[Material %], RollCoatTarget.[Kg/sq M], RMDatabase.[Solid%], " & _
                        "RollCoatTarget.[Foul. solid%], RMDatabase.[Price (NT/Kg)], RollCoatTarget.[Foul. Cost (NT/Kg)] From RMDatabase, RollCoatTarget " & _
                        "Where RMDatabase.[New ID] = RollCoatTarget.ID And ProductNo = '" & ProductNoTextBox.Text.Trim & "'"
            QueryAdapt = New System.Data.SqlClient.SqlDataAdapter(QueryStr, gDBConn)
            QuerySet = New System.Data.DataSet
            QueryAdapt.Fill(QuerySet, "DataSet1")

            Dim dr As DataRow
            dr = QuerySet.Tables("DataSet1").NewRow

            dr.Item(0) = "Total"
            dr.Item("Material %") = QuerySet.Tables("DataSet1").Compute("SUM([Material %])", "True")
            dr.Item("Kg/sq M") = QuerySet.Tables("DataSet1").Compute("SUM([Kg/sq M])", "True")
            dr.Item("Foul. solid%") = QuerySet.Tables("DataSet1").Compute("SUM([Foul. solid%])", "True")
            dr.Item("Foul. Cost (NT/Kg)") = QuerySet.Tables("DataSet1").Compute("SUM([Foul. Cost (NT/Kg)])", "True")

            QuerySet.Tables("DataSet1").Rows.Add(dr)

            RctDataGridView.DataSource = QuerySet.Tables("DataSet1")
            RctDataGridView.Columns(0).Width = 150
            RctDataGridView.Columns(1).Width = 450
            RctDataGridView.Columns(2).Width = 150
            RctDataGridView.Columns(3).Width = 150
            RctDataGridView.Columns(4).Width = 150
            RctDataGridView.Columns(5).Width = 150
            RctDataGridView.Columns(6).Width = 150
            RctDataGridView.Columns(7).Width = 150


        End If

        DisconnectFromDatabase()

    End Sub

    Private Sub Mc1tInsertButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Mc1tInsertButton.Click

        ConnectToDatabase()

        QueryDate = ProductDateTextBox.Text


        Dim IdQueryStr As String
        Dim IdQuerySet As System.Data.DataSet
        Dim IdQueryAdapt As System.Data.SqlClient.SqlDataAdapter

        Dim IdPercent As String = ""
        Dim IdPrice As String = ""

        IdQueryStr = "Select * From RMDatabase Where [New ID] = '" & Mc1tStockNoTextBox.Text.Trim & "'"
        IdQueryAdapt = New System.Data.SqlClient.SqlDataAdapter(IdQueryStr, gDBConn)
        IdQuerySet = New System.Data.DataSet
        IdQueryAdapt.Fill(IdQuerySet, "DataSet")

        If IdQuerySet.Tables("DataSet").Rows.Count > 0 Then
            IdPercent = Trim(IdQuerySet.Tables("DataSet").Rows(0).Item(4).ToString)
            IdPrice = Trim(IdQuerySet.Tables("DataSet").Rows(0).Item(5).ToString)
        Else
            MsgBox("Stock No. isn't find.", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Insert")
            Exit Sub
        End If

        If MsgBox("Insert？", MsgBoxStyle.Question + MsgBoxStyle.YesNo, "Question") = MsgBoxResult.Yes Then

            QueryStr = "INSERT INTO MakeCoat1Target( ProductNo, ID, [Material %], [Kg/sq M], [Foul. solid%], [Foul. Cost (NT/Kg)] )" & _
                        "VALUES ( '" & ProductNoTextBox.Text.Trim & _
                        "', '" & Mc1tStockNoTextBox.Text.Trim & _
                        "', '" & Mc1tMaterialTextBox.Text.Trim & _
                        "', '" & Double.Parse(Mc1tWetTextBox.Text.Trim) * (Double.Parse(Mc1tMaterialTextBox.Text.Trim) / 100.0) & _
                        "', '" & (Double.Parse(Mc1tMaterialTextBox.Text.Trim) / 100.0) * Double.Parse(IdPercent) & _
                        "', '" & Double.Parse(Mc1tWetTextBox.Text.Trim) * (Double.Parse(Mc1tMaterialTextBox.Text.Trim) / 100.0) * Double.Parse(IdPrice) & "')"

            QueryAdapt = New System.Data.SqlClient.SqlDataAdapter(QueryStr, gDBConn)
            QuerySet = New System.Data.DataSet
            QueryAdapt.Fill(QuerySet, "MakeCoat1Target")

            QueryStr = "Select MakeCoat1Target.ID, RMDatabase.[RM Description] ,MakeCoat1Target.[Material %], MakeCoat1Target.[Kg/sq M], RMDatabase.[Solid%], " & _
                        "MakeCoat1Target.[Foul. solid%], RMDatabase.[Price (NT/Kg)], MakeCoat1Target.[Foul. Cost (NT/Kg)] From RMDatabase, MakeCoat1Target " & _
                        "Where RMDatabase.[New ID] = MakeCoat1Target.ID And ProductNo = '" & ProductNoTextBox.Text.Trim & "'"
            QueryAdapt = New System.Data.SqlClient.SqlDataAdapter(QueryStr, gDBConn)
            QuerySet = New System.Data.DataSet
            QueryAdapt.Fill(QuerySet, "DataSet1")

            Dim dr As DataRow
            dr = QuerySet.Tables("DataSet1").NewRow

            dr.Item(0) = "Total"
            dr.Item("Material %") = QuerySet.Tables("DataSet1").Compute("SUM([Material %])", "True")
            dr.Item("Kg/sq M") = QuerySet.Tables("DataSet1").Compute("SUM([Kg/sq M])", "True")
            dr.Item("Foul. solid%") = QuerySet.Tables("DataSet1").Compute("SUM([Foul. solid%])", "True")
            dr.Item("Foul. Cost (NT/Kg)") = QuerySet.Tables("DataSet1").Compute("SUM([Foul. Cost (NT/Kg)])", "True")

            QuerySet.Tables("DataSet1").Rows.Add(dr)

            Mc1tDataGridView.DataSource = QuerySet.Tables("DataSet1")
            Mc1tDataGridView.Columns(0).Width = 150
            Mc1tDataGridView.Columns(1).Width = 450
            Mc1tDataGridView.Columns(2).Width = 150
            Mc1tDataGridView.Columns(3).Width = 150
            Mc1tDataGridView.Columns(4).Width = 150
            Mc1tDataGridView.Columns(5).Width = 150
            Mc1tDataGridView.Columns(6).Width = 150
            Mc1tDataGridView.Columns(7).Width = 150

        End If

        DisconnectFromDatabase()

    End Sub

    Private Sub SctInsertButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SctInsertButton.Click

        ConnectToDatabase()

        QueryDate = ProductDateTextBox.Text

        Dim IdQueryStr As String
        Dim IdQuerySet As System.Data.DataSet
        Dim IdQueryAdapt As System.Data.SqlClient.SqlDataAdapter
        Dim IdPercent As String = ""
        Dim IdPrice As String = ""

        IdQueryStr = "Select * From RMDatabase Where [New ID] = '" & SctStockNoTextBox.Text.Trim & "'"
        IdQueryAdapt = New System.Data.SqlClient.SqlDataAdapter(IdQueryStr, gDBConn)
        IdQuerySet = New System.Data.DataSet
        IdQueryAdapt.Fill(IdQuerySet, "DataSet")

        If IdQuerySet.Tables("DataSet").Rows.Count > 0 Then
            IdPercent = Trim(IdQuerySet.Tables("DataSet").Rows(0).Item(4).ToString)
            IdPrice = Trim(IdQuerySet.Tables("DataSet").Rows(0).Item(5).ToString)
        Else
            MsgBox("Stock No. isn't find.", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Insert")
            Exit Sub
        End If

        If MsgBox("Insert？", MsgBoxStyle.Question + MsgBoxStyle.YesNo, "Question") = MsgBoxResult.Yes Then

            QueryStr = "INSERT INTO SprayCoatTarget( ProductNo, ID, [Material %], [Kg/sq M], [Foul. solid%], [Foul. Cost (NT/Kg)] )" & _
                        "VALUES ( '" & ProductNoTextBox.Text.Trim & _
                        "', '" & SctStockNoTextBox.Text.Trim & _
                        "', '" & SctMaterialTextBox.Text.Trim & _
                        "', '" & Double.Parse(SctWetTextBox.Text.Trim) * (Double.Parse(SctMaterialTextBox.Text.Trim) / 100.0) & _
                        "', '" & (Double.Parse(SctMaterialTextBox.Text.Trim) / 100.0) * Double.Parse(IdPercent) & _
                        "', '" & Double.Parse(SctWetTextBox.Text.Trim) * (Double.Parse(SctMaterialTextBox.Text.Trim) / 100.0) * Double.Parse(IdPrice) & "')"

            QueryAdapt = New System.Data.SqlClient.SqlDataAdapter(QueryStr, gDBConn)
            QuerySet = New System.Data.DataSet
            QueryAdapt.Fill(QuerySet, "SprayCoatTarget")

            QueryStr = "Select SprayCoatTarget.ID, RMDatabase.[RM Description] ,SprayCoatTarget.[Material %], SprayCoatTarget.[Kg/sq M], RMDatabase.[Solid%], " & _
                        "SprayCoatTarget.[Foul. solid%], RMDatabase.[Price (NT/Kg)], SprayCoatTarget.[Foul. Cost (NT/Kg)] From RMDatabase, SprayCoatTarget " & _
                        "Where RMDatabase.[New ID] = SprayCoatTarget.ID And ProductNo = '" & ProductNoTextBox.Text.Trim & "'"
            QueryAdapt = New System.Data.SqlClient.SqlDataAdapter(QueryStr, gDBConn)
            QuerySet = New System.Data.DataSet
            QueryAdapt.Fill(QuerySet, "DataSet1")

            Dim dr As DataRow
            dr = QuerySet.Tables("DataSet1").NewRow

            dr.Item(0) = "Total"
            dr.Item("Material %") = QuerySet.Tables("DataSet1").Compute("SUM([Material %])", "True")
            dr.Item("Kg/sq M") = QuerySet.Tables("DataSet1").Compute("SUM([Kg/sq M])", "True")
            dr.Item("Foul. solid%") = QuerySet.Tables("DataSet1").Compute("SUM([Foul. solid%])", "True")
            dr.Item("Foul. Cost (NT/Kg)") = QuerySet.Tables("DataSet1").Compute("SUM([Foul. Cost (NT/Kg)])", "True")

            QuerySet.Tables("DataSet1").Rows.Add(dr)

            SctDataGridView.DataSource = QuerySet.Tables("DataSet1")
            SctDataGridView.Columns(0).Width = 150
            SctDataGridView.Columns(1).Width = 450
            SctDataGridView.Columns(2).Width = 150
            SctDataGridView.Columns(3).Width = 150
            SctDataGridView.Columns(4).Width = 150
            SctDataGridView.Columns(5).Width = 150
            SctDataGridView.Columns(6).Width = 150
            SctDataGridView.Columns(7).Width = 150

        End If

        DisconnectFromDatabase()

    End Sub

    Private Sub Sc2tInsertButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Sc2tInsertButton.Click

        ConnectToDatabase()

        QueryDate = ProductDateTextBox.Text

        Dim IdQueryStr As String
        Dim IdQuerySet As System.Data.DataSet
        Dim IdQueryAdapt As System.Data.SqlClient.SqlDataAdapter
        Dim IdPercent As String = ""
        Dim IdPrice As String = ""

        IdQueryStr = "Select * From RMDatabase Where [New ID] = '" & Sc2tStockNoTextBox.Text.Trim & "'"
        IdQueryAdapt = New System.Data.SqlClient.SqlDataAdapter(IdQueryStr, gDBConn)
        IdQuerySet = New System.Data.DataSet
        IdQueryAdapt.Fill(IdQuerySet, "DataSet")

        If IdQuerySet.Tables("DataSet").Rows.Count > 0 Then
            IdPercent = Trim(IdQuerySet.Tables("DataSet").Rows(0).Item(4).ToString)
            IdPrice = Trim(IdQuerySet.Tables("DataSet").Rows(0).Item(5).ToString)
        Else
            MsgBox("Stock No. isn't find.", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Insert")
            Exit Sub
        End If

        If MsgBox("Insert？", MsgBoxStyle.Question + MsgBoxStyle.YesNo, "Question") = MsgBoxResult.Yes Then

            QueryStr = "INSERT INTO SprayCoat2Target( ProductNo, ID, [Material %], [Kg/sq M], [Foul. solid%], [Foul. Cost (NT/Kg)] )" & _
                        "VALUES ( '" & ProductNoTextBox.Text.Trim & _
                        "', '" & Sc2tStockNoTextBox.Text.Trim & _
                        "', '" & Sc2tMaterialTextBox.Text.Trim & _
                        "', '" & Double.Parse(Sc2tWetTextBox.Text.Trim) * (Double.Parse(Sc2tMaterialTextBox.Text.Trim) / 100.0) & _
                        "', '" & (Double.Parse(Sc2tMaterialTextBox.Text.Trim) / 100.0) * Double.Parse(IdPercent) & _
                        "', '" & Double.Parse(Sc2tWetTextBox.Text.Trim) * (Double.Parse(Sc2tMaterialTextBox.Text.Trim) / 100.0) * Double.Parse(IdPrice) & "')"

            QueryAdapt = New System.Data.SqlClient.SqlDataAdapter(QueryStr, gDBConn)
            QuerySet = New System.Data.DataSet
            QueryAdapt.Fill(QuerySet, "SprayCoat2Target")

            QueryStr = "Select SprayCoat2Target.ID, RMDatabase.[RM Description] ,SprayCoat2Target.[Material %], SprayCoat2Target.[Kg/sq M], RMDatabase.[Solid%], " & _
                        "SprayCoat2Target.[Foul. solid%], RMDatabase.[Price (NT/Kg)], SprayCoat2Target.[Foul. Cost (NT/Kg)] From RMDatabase, SprayCoat2Target " & _
                        "Where RMDatabase.[New ID] = SprayCoat2Target.ID And ProductNo = '" & ProductNoTextBox.Text.Trim & "'"
            QueryAdapt = New System.Data.SqlClient.SqlDataAdapter(QueryStr, gDBConn)
            QuerySet = New System.Data.DataSet
            QueryAdapt.Fill(QuerySet, "DataSet1")

            Dim dr As DataRow
            dr = QuerySet.Tables("DataSet1").NewRow

            dr.Item(0) = "Total"
            dr.Item("Material %") = QuerySet.Tables("DataSet1").Compute("SUM([Material %])", "True")
            dr.Item("Kg/sq M") = QuerySet.Tables("DataSet1").Compute("SUM([Kg/sq M])", "True")
            dr.Item("Foul. solid%") = QuerySet.Tables("DataSet1").Compute("SUM([Foul. solid%])", "True")
            dr.Item("Foul. Cost (NT/Kg)") = QuerySet.Tables("DataSet1").Compute("SUM([Foul. Cost (NT/Kg)])", "True")

            QuerySet.Tables("DataSet1").Rows.Add(dr)

            Sc2tDataGridView.DataSource = QuerySet.Tables("DataSet1")
            Sc2tDataGridView.Columns(0).Width = 150
            Sc2tDataGridView.Columns(1).Width = 450
            Sc2tDataGridView.Columns(2).Width = 150
            Sc2tDataGridView.Columns(3).Width = 150
            Sc2tDataGridView.Columns(4).Width = 150
            Sc2tDataGridView.Columns(5).Width = 150
            Sc2tDataGridView.Columns(6).Width = 150
            Sc2tDataGridView.Columns(7).Width = 150

        End If

        DisconnectFromDatabase()

    End Sub

    Private Sub SizeCTInsertButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SizeCTInsertButton.Click

        ConnectToDatabase()

        QueryDate = ProductDateTextBox.Text

        Dim IdQueryStr As String
        Dim IdQuerySet As System.Data.DataSet
        Dim IdQueryAdapt As System.Data.SqlClient.SqlDataAdapter
        Dim IdPercent As String = ""
        Dim IdPrice As String = ""

        IdQueryStr = "Select * From RMDatabase Where [New ID] = '" & SizeCTStockNoTextBox.Text.Trim & "'"
        IdQueryAdapt = New System.Data.SqlClient.SqlDataAdapter(IdQueryStr, gDBConn)
        IdQuerySet = New System.Data.DataSet
        IdQueryAdapt.Fill(IdQuerySet, "DataSet")

        If IdQuerySet.Tables("DataSet").Rows.Count > 0 Then
            IdPercent = Trim(IdQuerySet.Tables("DataSet").Rows(0).Item(4).ToString)
            IdPrice = Trim(IdQuerySet.Tables("DataSet").Rows(0).Item(5).ToString)
        Else
            MsgBox("Stock No. isn't find.", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Insert")
            Exit Sub
        End If

        If MsgBox("Insert？", MsgBoxStyle.Question + MsgBoxStyle.YesNo, "Question") = MsgBoxResult.Yes Then

            QueryStr = "INSERT INTO SizeCoatTarget( ProductNo, ID, [Material %], [Kg/sq M], [Foul. solid%], [Foul. Cost (NT/Kg)] )" & _
                        "VALUES ( '" & ProductNoTextBox.Text.Trim & _
                        "', '" & SizeCTStockNoTextBox.Text.Trim & _
                        "', '" & SizeCTMaterialTextBox.Text.Trim & _
                        "', '" & Double.Parse(SizeCTWetTextBox.Text.Trim) * (Double.Parse(SizeCTMaterialTextBox.Text.Trim) / 100.0) & _
                        "', '" & (Double.Parse(SizeCTMaterialTextBox.Text.Trim) / 100.0) * Double.Parse(IdPercent) & _
                        "', '" & Double.Parse(SizeCTWetTextBox.Text.Trim) * (Double.Parse(SizeCTMaterialTextBox.Text.Trim) / 100.0) * Double.Parse(IdPrice) & "')"

            QueryAdapt = New System.Data.SqlClient.SqlDataAdapter(QueryStr, gDBConn)
            QuerySet = New System.Data.DataSet
            QueryAdapt.Fill(QuerySet, "SizeCoatTarget")

            QueryStr = "Select SizeCoatTarget.ID, RMDatabase.[RM Description] ,SizeCoatTarget.[Material %], SizeCoatTarget.[Kg/sq M], RMDatabase.[Solid%], " & _
                        "SizeCoatTarget.[Foul. solid%], RMDatabase.[Price (NT/Kg)], SizeCoatTarget.[Foul. Cost (NT/Kg)] From RMDatabase, SizeCoatTarget " & _
                        "Where RMDatabase.[New ID] = SizeCoatTarget.ID And ProductNo = '" & ProductNoTextBox.Text.Trim & "'"
            QueryAdapt = New System.Data.SqlClient.SqlDataAdapter(QueryStr, gDBConn)
            QuerySet = New System.Data.DataSet
            QueryAdapt.Fill(QuerySet, "DataSet1")

            Dim dr As DataRow
            dr = QuerySet.Tables("DataSet1").NewRow

            dr.Item(0) = "Total"
            dr.Item("Material %") = QuerySet.Tables("DataSet1").Compute("SUM([Material %])", "True")
            dr.Item("Kg/sq M") = QuerySet.Tables("DataSet1").Compute("SUM([Kg/sq M])", "True")
            dr.Item("Foul. solid%") = QuerySet.Tables("DataSet1").Compute("SUM([Foul. solid%])", "True")
            dr.Item("Foul. Cost (NT/Kg)") = QuerySet.Tables("DataSet1").Compute("SUM([Foul. Cost (NT/Kg)])", "True")

            QuerySet.Tables("DataSet1").Rows.Add(dr)

            SizeCTDataGridView.DataSource = QuerySet.Tables("DataSet1")
            SizeCTDataGridView.Columns(0).Width = 150
            SizeCTDataGridView.Columns(1).Width = 450
            SizeCTDataGridView.Columns(2).Width = 150
            SizeCTDataGridView.Columns(3).Width = 150
            SizeCTDataGridView.Columns(4).Width = 150
            SizeCTDataGridView.Columns(5).Width = 150
            SizeCTDataGridView.Columns(6).Width = 150
            SizeCTDataGridView.Columns(7).Width = 150

        End If

        DisconnectFromDatabase()

    End Sub



    Private Sub MctInsertButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MctInsertButton.Click

        ConnectToDatabase()

        QueryDate = ProductDateTextBox.Text

        Dim IdQueryStr As String
        Dim IdQuerySet As System.Data.DataSet
        Dim IdQueryAdapt As System.Data.SqlClient.SqlDataAdapter
        Dim IdPercent As String = ""
        Dim IdPrice As String = ""

        IdQueryStr = "Select * From RMDatabase Where [New ID] = '" & MctStockNoTextBox.Text.Trim & "'"
        IdQueryAdapt = New System.Data.SqlClient.SqlDataAdapter(IdQueryStr, gDBConn)
        IdQuerySet = New System.Data.DataSet
        IdQueryAdapt.Fill(IdQuerySet, "DataSet")

        If IdQuerySet.Tables("DataSet").Rows.Count > 0 Then
            IdPercent = Trim(IdQuerySet.Tables("DataSet").Rows(0).Item(4).ToString)
            IdPrice = Trim(IdQuerySet.Tables("DataSet").Rows(0).Item(5).ToString)
        Else
            MsgBox("Stock No. isn't find.", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Insert")
            Exit Sub
        End If

        If MsgBox("Insert？", MsgBoxStyle.Question + MsgBoxStyle.YesNo, "Question") = MsgBoxResult.Yes Then

            QueryStr = "INSERT INTO MineralCoatTarget( ProductNo,ID, [Material %], [Kg/sq M], [Foul. solid%], [Foul. Cost (NT/Kg)] )" & _
                        "VALUES ( '" & ProductNoTextBox.Text.Trim & _
                        "', '" & MctStockNoTextBox.Text.Trim & _
                        "', '" & MctMaterialTextBox.Text.Trim & _
                        "', '" & Double.Parse(MctWetTextBox.Text.Trim) * (Double.Parse(MctMaterialTextBox.Text.Trim) / 100.0) & _
                        "', '" & (Double.Parse(MctMaterialTextBox.Text.Trim) / 100.0) * Double.Parse(IdPercent) & _
                        "', '" & Double.Parse(MctWetTextBox.Text.Trim) * (Double.Parse(MctMaterialTextBox.Text.Trim) / 100.0) * Double.Parse(IdPrice) & "')"

            QueryAdapt = New System.Data.SqlClient.SqlDataAdapter(QueryStr, gDBConn)
            QuerySet = New System.Data.DataSet
            QueryAdapt.Fill(QuerySet, "FiberInputTarget")

            QueryStr = "Select MineralCoatTarget.ID, RMDatabase.[RM Description] ,MineralCoatTarget.[Material %], MineralCoatTarget.[Kg/sq M], RMDatabase.[Solid%], " & _
                        "MineralCoatTarget.[Foul. solid%], RMDatabase.[Price (NT/Kg)], MineralCoatTarget.[Foul. Cost (NT/Kg)] From RMDatabase, MineralCoatTarget " & _
                        "Where RMDatabase.[New ID] = MineralCoatTarget.ID And ProductNo = '" & ProductNoTextBox.Text.Trim & "'"
            QueryAdapt = New System.Data.SqlClient.SqlDataAdapter(QueryStr, gDBConn)
            QuerySet = New System.Data.DataSet
            QueryAdapt.Fill(QuerySet, "DataSet1")

            Dim dr As DataRow
            dr = QuerySet.Tables("DataSet1").NewRow

            dr.Item(0) = "Total"
            dr.Item("Material %") = QuerySet.Tables("DataSet1").Compute("SUM([Material %])", "True")
            dr.Item("Kg/sq M") = QuerySet.Tables("DataSet1").Compute("SUM([Kg/sq M])", "True")
            dr.Item("Foul. solid%") = QuerySet.Tables("DataSet1").Compute("SUM([Foul. solid%])", "True")
            dr.Item("Foul. Cost (NT/Kg)") = QuerySet.Tables("DataSet1").Compute("SUM([Foul. Cost (NT/Kg)])", "True")

            QuerySet.Tables("DataSet1").Rows.Add(dr)

            MctDataGridView.DataSource = QuerySet.Tables("DataSet1")
            MctDataGridView.Columns(0).Width = 150
            MctDataGridView.Columns(1).Width = 450
            MctDataGridView.Columns(2).Width = 150
            MctDataGridView.Columns(3).Width = 150
            MctDataGridView.Columns(4).Width = 150
            MctDataGridView.Columns(5).Width = 150
            MctDataGridView.Columns(6).Width = 150
            MctDataGridView.Columns(7).Width = 150

        End If

        DisconnectFromDatabase()

    End Sub

    Private Sub SpecSearchButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SpecSearchButton.Click
        ConnectToDatabase()

        QueryDate = ProductDateTextBox.Text

        QueryStr = "Select * From SpecData Where TypeNo = '" & SpecNoTextBox.Text.Trim & "'"
        QueryAdapt = New System.Data.SqlClient.SqlDataAdapter(QueryStr, gDBConn)
        QuerySet = New System.Data.DataSet
        QueryAdapt.Fill(QuerySet, "SPEC")

        If QuerySet.Tables("SPEC").Rows.Count > 0 Then
            SPECNameLabel.Text = Trim(QuerySet.Tables("SPEC").Rows(0).Item("TypeName").ToString)
            SPECColorLabel.Text = Trim(QuerySet.Tables("SPEC").Rows(0).Item("FontColor").ToString)
            ForwardElongationLabel.Text = Trim(QuerySet.Tables("SPEC").Rows(0).Item("ForwardElongation").ToString)
            BackwordElongationLabel.Text = Trim(QuerySet.Tables("SPEC").Rows(0).Item("BackwordElongation").ToString)
            JumboSizeLabel.Text = Trim(QuerySet.Tables("SPEC").Rows(0).Item("JumboSize").ToString)
            JumboParameterLabel.Text = Trim(QuerySet.Tables("SPEC").Rows(0).Item("JumboParameter").ToString)


        Else
            MsgBox("SPEC No. isn't find.", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Search SPEC")
            Exit Sub
        End If






    End Sub

End Class
