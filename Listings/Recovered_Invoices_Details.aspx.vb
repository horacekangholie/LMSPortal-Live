
Imports AjaxControlToolkit.HtmlEditor.ToolbarButtons
Imports NPOI.SS.Formula.Functions

Partial Class Listings_Recovered_Invoices_Details
    Inherits LMSPortalBaseCode

    Dim Currency As String
    Dim TotalAmount As Decimal

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not Me.Page.User.Identity.IsAuthenticated AndAlso Session("Login_Status") <> "Logged in" Then
            FormsAuthentication.RedirectToLoginPage()
        End If

        If Request.QueryString("Invoice_No") <> "" And Not Request.QueryString("Invoice_No") Is Nothing Then
            LB_PageTitle.Text = "Billed Item Details"
            PopulateGridViewData()
        Else
            Response.Redirect("~/Listings/Recovered_Invoices.aspx")
        End If

        '' Postback link for downloading the invoice file
        RefLink.PostBackUrl = "~/Download/DownloadFile.aspx?Inv_Ref_No=" & Request.QueryString("Invoice_No")
    End Sub

    Protected Sub PopulateGridViewData()
        Try
            Dim sqlStr() As String = {" SELECT PO_No, Invoice_No, Invoice_Date, Item_Code, Description, Currency, SUM(Amount) As Amount " &
                                      " FROM I_DB_Recovered_Invoice WHERE Invoice_No = '" & Request.QueryString("Invoice_No") & "' " &
                                      " GROUP BY PO_No, Invoice_No, Invoice_Date, Item_Code, Description, Currency "}

            BuildGridView(GridView1, "GridView1", "Invoice_No")
            GridView1.DataSource = GetDataTable(sqlStr(0))
            GridView1.DataBind()
        Catch ex As Exception
            Response.Write("Error:  " & ex.Message)
        End Try
    End Sub

    Protected Sub BuildGridView(ByVal ControlObj As Object, ByVal ControlName As String, ByVal DataKeyName As String)
        Dim GridViewObj As GridView = CType(ControlObj, GridView)

        '' GridView Properties
        GridViewObj.ID = ControlName
        GridViewObj.AutoGenerateColumns = False
        GridViewObj.CellPadding = 4
        GridViewObj.Font.Size = 10
        GridViewObj.GridLines = GridLines.None
        GridViewObj.ShowFooter = True
        GridViewObj.ShowHeaderWhenEmpty = True
        GridViewObj.DataKeyNames = New String() {DataKeyName}
        GridViewObj.CssClass = "table table-bordered"
        GridViewObj.Style.Add("width", "99.3%")

        '' Header Style
        GridViewObj.HeaderStyle.CssClass = "table-secondary"
        GridViewObj.HeaderStyle.Font.Bold = True
        GridViewObj.HeaderStyle.VerticalAlign = VerticalAlign.Top

        '' Row Style
        GridViewObj.RowStyle.CssClass = "Default"
        GridViewObj.RowStyle.VerticalAlign = VerticalAlign.Middle

        '' Footer Style
        GridViewObj.FooterStyle.CssClass = "table-active"

        '' Pager Style
        GridViewObj.PagerSettings.Mode = PagerButtons.NumericFirstLast
        GridViewObj.PagerSettings.FirstPageText = "First"
        GridViewObj.PagerSettings.LastPageText = "Last"
        GridViewObj.PagerSettings.PageButtonCount = "5"
        GridViewObj.PagerStyle.HorizontalAlign = HorizontalAlign.Center
        GridViewObj.PagerStyle.CssClass = "pagination-ys"

        '' Empty Data Template
        GridViewObj.EmptyDataText = "No records found."

        '' Define each Gridview
        Select Case ControlName
            Case "GridView1"
                '' Build GridView Content
                GridViewObj.AllowPaging = True
                GridViewObj.PageSize = 50
                GridViewObj.Columns.Clear()
                Dim ColData() As String = {"Invoice_Date", "PO_No", "Item_Code", "Description", "Currency", "Amount"}
                Dim ColSize() As Integer = {50, 50, 100, 400, 50, 50}

                For i = 0 To ColData.Length - 1
                    Dim Bfield As BoundField = New BoundField()
                    Bfield.DataField = ColData(i)
                    Bfield.HeaderText = ColData(i).Replace("_", " ")
                    Bfield.HeaderStyle.Width = ColSize(i)
                    If Bfield.HeaderText.Contains("Date") Then
                        Bfield.DataFormatString = "{0:dd MMM yy}"
                    End If
                    If Bfield.HeaderText.Contains("Amount") Then
                        Bfield.DataFormatString = "{0:#,##0.00}"
                    End If
                    Bfield.HeaderStyle.Wrap = False
                    Bfield.ItemStyle.Wrap = False
                    GridViewObj.Columns.Add(Bfield)
                Next

                '' Add templatefield for Edit icon
                Dim TField As TemplateField = New TemplateField()
                TField.HeaderStyle.Width = Unit.Percentage(2)
                TField.ItemStyle.Wrap = False
                TField.ItemTemplate = New GridViewItemTemplateControl()
                GridViewObj.Columns.Add(TField)

        End Select
    End Sub



    '' Gridview control
    Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As GridViewRowEventArgs) Handles GridView1.RowDataBound
        If e.Row.RowType = DataControlRowType.Header Then
            e.Row.Cells(GetColumnIndexByName(e.Row, "Currency")).Style.Add("text-align", "right !important")
            e.Row.Cells(GetColumnIndexByName(e.Row, "Amount")).Style.Add("text-align", "right !important")

        ElseIf e.Row.RowType = DataControlRowType.DataRow Then
            '' Create Edit button at the last column
            Dim drv As System.Data.DataRowView = e.Row.DataItem

            Dim CtrlCellIndex As Integer = e.Row.Cells.Count - 1
            Dim EditLinkButton As LinkButton = TryCast(e.Row.Cells(CtrlCellIndex).Controls(0), LinkButton)
            EditLinkButton.CommandArgument = e.Row.RowIndex & "|" & drv("Invoice_No") & "|" & drv("Invoice_Date") & "|" & drv("Item_Code") & "|" & drv("Currency") & "|" & drv("Amount") & "|" & drv("PO_No")
            EditLinkButton.CausesValidation = False
            AddHandler EditLinkButton.Click, AddressOf Edit_BilledItem_Click

            EditLinkButton.Style.Add("margin-right", "5px")   '' add separator between button

            Dim DeleteLinkButton As LinkButton = TryCast(e.Row.Cells(CtrlCellIndex).Controls(1), LinkButton)
            DeleteLinkButton.CommandArgument = e.Row.RowIndex & "|" & drv("Invoice_No") & "|" & drv("Invoice_Date") & "|" & drv("Item_Code") & "|" & drv("Currency") & "|" & drv("Amount") & "|" & drv("PO_No")
            DeleteLinkButton.CausesValidation = False
            DeleteLinkButton.OnClientClick = "return confirm('Are you sure to delete record?')"
            AddHandler DeleteLinkButton.Click, AddressOf Delete_BilledItem_Click


            Dim lockedCodes As String() = {"DMC004", "DMC005", "DMC013"}
            Dim isLockedItem As Boolean = lockedCodes.Contains(drv("Item_Code").ToString())

            '' Lock record if invoice has been recovered
            If Not isLockedItem Then
                EditLinkButton.Text = "<i class='bi bi-pencil-fill'></i>"
                EditLinkButton.CssClass = "btn btn-xs btn-info"
                EditLinkButton.Enabled = True

                DeleteLinkButton.Text = "<i class='bi bi-trash'></i>"
                DeleteLinkButton.CssClass = "btn btn-xs btn-danger"
                DeleteLinkButton.Enabled = True
            Else
                EditLinkButton.Text = "<i class='bi bi-lock'></i>"
                EditLinkButton.CssClass = "btn btn-xs btn-light disabled"
                EditLinkButton.ToolTip = "Item Locked"
                EditLinkButton.Enabled = False

                DeleteLinkButton.Text = "<i class='bi bi-lock'></i>"
                DeleteLinkButton.CssClass = "btn btn-xs btn-light disabled"
                DeleteLinkButton.Enabled = False
                DeleteLinkButton.Visible = False
            End If

            '' Total Amount
            TotalAmount += CDec(DataBinder.Eval(e.Row.DataItem, "Amount"))

            '' Currency
            Currency = drv("Currency")  '' Get currency from here to prevent heavy load using Get_Value from DB

            e.Row.Cells(GetColumnIndexByName(e.Row, "Currency")).Style.Add("text-align", "right !important")
            e.Row.Cells(GetColumnIndexByName(e.Row, "Amount")).Style.Add("text-align", "right !important")

        ElseIf e.Row.RowType = DataControlRowType.Footer Then
            e.Row.Cells(GetColumnIndexByName(e.Row, "Description")).Text = "Total"
            e.Row.Cells(GetColumnIndexByName(e.Row, "Description")).HorizontalAlign = HorizontalAlign.Right
            e.Row.Cells(GetColumnIndexByName(e.Row, "Description")).Style.Add("padding-right", "30px")
            e.Row.Cells(GetColumnIndexByName(e.Row, "Currency")).Text = Currency
            e.Row.Cells(GetColumnIndexByName(e.Row, "Amount")).Text = TotalAmount.ToString("#,##0.00")

            e.Row.Cells(GetColumnIndexByName(e.Row, "Currency")).Style.Add("text-align", "right !important")
            e.Row.Cells(GetColumnIndexByName(e.Row, "Amount")).Style.Add("text-align", "right !important")
        End If

    End Sub

    Protected Sub GridView1_PageIndexChanging(sender As Object, e As GridViewPageEventArgs) Handles GridView1.PageIndexChanging
        GridView1.PageIndex = e.NewPageIndex
        PopulateGridViewData()
    End Sub



    '' Modal control
    Protected Sub DDL_Item_Code_Load(sender As Object, e As EventArgs) Handles DDL_Item_Code.Load
        If Not IsPostBack Then
            Try
                BindDropDownList_Custom_Default_Value(DDL_Item_Code, GetSQL(DDL_Item_Code), "Item Description", "Item Code", "", "-1", True)
            Catch ex As Exception
                Response.Write("Error: " & ex.Message)
            End Try
        End If
    End Sub

    Protected Sub DDL_Currency_Load(sender As Object, e As EventArgs) Handles DDL_Currency.Load
        If Not IsPostBack Then
            Try
                BindDropDownList_Custom_Default_Value(DDL_Currency, GetSQL(DDL_Currency), "Currency", "Currency", "", "-1", True)
            Catch ex As Exception
                Response.Write("Error: " & ex.Message)
            End Try
        End If
    End Sub

    Protected Sub DDL_PO_No_Load(sender As Object, e As EventArgs) Handles DDL_PO_No.Load
        If Not IsPostBack Then
            Try
                Dim Invoice_No As String = Request.QueryString("Invoice_No")
                BindDropDownList_Custom_Default_Value(DDL_PO_No, GetSQL(DDL_PO_No, Invoice_No), "PO_No", "PO_No", "", "-1", True)
            Catch ex As Exception
                Response.Write("Error: " & ex.Message)
            End Try
        End If
    End Sub

    Protected Sub Add_BilledItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnAddBilledItem.Click
        ModalHeaderBilledItem.Text = "Add Bill Item"
        btnSaveBilledItem.Text = "Save"
        btnCancelBilledItem.Text = "Cancel"

        ' Initialiaze the hidden fields
        Dim HiddenFields As Array = {TB_Selected_Row_Index, TB_Selected_Invoice_No, TB_Selected_Invoice_Date, TB_Selected_Item_Code, TB_Selected_Currency, TB_Selected_Amount, TB_Selected_PO_No}

        For i = 0 To HiddenFields.Length - 1
            HiddenFields(i).Text = String.Empty
        Next

        ' Reinitialize field based on new/edit mode
        ReInitializeField(btnSaveBilledItem.Text)

    End Sub

    Protected Sub Edit_BilledItem_Click(ByVal sender As Object, ByVal e As EventArgs)
        ModalHeaderBilledItem.Text = "Update Bill Item"
        btnSaveBilledItem.Text = "Update"
        btnCancelBilledItem.Text = "Cancel"

        '' Get row command argument
        Dim EditLinkButton As LinkButton = TryCast(sender, LinkButton)
        Dim EditLinkButtonCommandArgument As Array = Split(EditLinkButton.CommandArgument, "|")

        Dim HiddenFields As Array = {TB_Selected_Row_Index, TB_Selected_Invoice_No, TB_Selected_Invoice_Date, TB_Selected_Item_Code, TB_Selected_Currency, TB_Selected_Amount, TB_Selected_PO_No}

        '' file value to hidden fields
        For i = 0 To EditLinkButtonCommandArgument.Length - 1
            HiddenFields(i).Text = EditLinkButtonCommandArgument(i)
        Next

        ' Reinitialize field based on new/edit mode
        ReInitializeField(btnSaveBilledItem.Text)

    End Sub

    Protected Sub ReInitializeField(ByVal currmode As String)
        Dim oMode As String = IIf(currmode = "Save", "New", "Edit")
        Dim modalFields As New List(Of Control) From {TB_Invoice_No, TB_Invoice_Date, DDL_Item_Code, DDL_PO_No, DDL_Currency, TB_Amount, CompareValidator_DDL_PO_No}

        For Each ctrl As Control In modalFields
            Select Case True
                Case TypeOf ctrl Is TextBox
                    Dim tb As TextBox = CType(ctrl, TextBox)
                    Select Case tb.ID
                        Case "TB_Invoice_No"
                            tb.Text = IIf(oMode = "New", Request.QueryString("Invoice_No"), TB_Selected_Invoice_No.Text)
                            tb.Enabled = False

                        Case "TB_Invoice_Date"
                            tb.Text = If(oMode = "New",
                                         CDate(Get_Value("SELECT TOP 1 Invoice_Date FROM I_DB_Recovered_Invoice WHERE Invoice_No = '" & Request.QueryString("Invoice_No") & "' ", "Invoice_Date")).ToString("yyyy-MM-dd"),
                                         CDate(TB_Selected_Invoice_Date.Text).ToString("yyyy-MM-dd"))
                            tb.Enabled = False
                            tb.TextMode = TextBoxMode.SingleLine

                        Case "TB_Amount"
                            tb.Text = If(oMode = "New", String.Empty, Math.Round(CDec(TB_Selected_Amount.Text), 2).ToString("0.00"))
                            tb.Enabled = True

                    End Select
                Case TypeOf ctrl Is DropDownList
                    Dim ddl As DropDownList = CType(ctrl, DropDownList)
                    Select Case ddl.ID
                        Case "DDL_Item_Code"
                            ddl.SelectedIndex = IIf(oMode = "New", -1, ddl.Items.IndexOf(ddl.Items.FindByValue(TB_Selected_Item_Code.Text)))
                            ddl.Enabled = (oMode = "New")

                        Case "DDL_Currency"
                            Dim CurrencySelected As String = Get_Value("SELECT TOP 1 Currency FROM I_DB_Recovered_Invoice WHERE Invoice_No = '" & Request.QueryString("Invoice_No") & "' ", "Currency")
                            ddl.SelectedIndex = IIf(oMode = "New", -1, ddl.Items.IndexOf(ddl.Items.FindByValue(CurrencySelected)))
                            ddl.Enabled = True

                        Case "DDL_PO_No"
                            Dim PO_No_Selected As String = Get_Value("SELECT DISTINCT PO_No FROM I_DB_Recovered_Invoice WHERE Invoice_No = '" & Request.QueryString("Invoice_No") & "' AND PO_No != '' ", "PO_No")
                            ddl.SelectedIndex = IIf(oMode = "New", -1, ddl.Items.IndexOf(ddl.Items.FindByValue(PO_No_Selected)))
                            ddl.Enabled = (oMode = "New" And DDL_PO_No.Items.Cast(Of ListItem)().Count(Function(i) i.Value <> "-1") > 0)

                    End Select

                Case TypeOf ctrl Is CompareValidator
                    Dim cpv As CompareValidator = CType(ctrl, CompareValidator)
                    cpv.Enabled = IIf(DDL_PO_No.Items.Cast(Of ListItem)().Count(Function(i) i.Value <> "-1") > 0, True, False)

            End Select
        Next

        popupBilledItem.Show()
    End Sub

    Protected Sub Save_BilledItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnSaveBilledItem.Click
        Dim Invoice_No As String = Request.QueryString("Invoice_No")
        Dim Invoice_Date As TextBox = pnlAddEditBilledItem.FindControl("TB_Invoice_Date")
        Dim Item_Code As DropDownList = pnlAddEditBilledItem.FindControl("DDL_Item_Code")
        Dim Currency As DropDownList = pnlAddEditBilledItem.FindControl("DDL_Currency")
        Dim Amount As TextBox = pnlAddEditBilledItem.FindControl("TB_Amount")
        Dim PO_No As String = IIf(btnSaveBilledItem.Text = "Save", DDL_PO_No.SelectedItem.Text, TB_Selected_PO_No.Text)


        '' Handle invoice data format
        Dim formattedInvoiceDate As String = ""
        If Not String.IsNullOrWhiteSpace(Invoice_Date.Text) Then
            Dim parsedDate As DateTime
            If DateTime.TryParse(Invoice_Date.Text, parsedDate) Then
                formattedInvoiceDate = parsedDate.ToString("yyyy-MM-dd")
            Else
                Throw New Exception("Invalid date format.")
            End If
        End If

        Try
            Dim sqlStr As String = "EXEC SP_CRUD_Recovered_Invoice_Bill_Items N'" & Invoice_No &
                                                                          "', N'" & formattedInvoiceDate &
                                                                          "', N'" & Item_Code.SelectedValue &
                                                                          "', N'" & Currency.SelectedValue &
                                                                          "', N'" & Amount.Text &
                                                                          "', N'" & PO_No & "' "

            RunSQL(sqlStr)
        Catch ex As Exception
            Response.Write("Error:  " & ex.Message)
        End Try

        HandleSplitRouting()
        PopulateGridViewData()
    End Sub

    Protected Sub Delete_BilledItem_Click(ByVal sender As Object, ByVal e As EventArgs)
        '' Get row command argument
        Dim EditLinkButton As LinkButton = TryCast(sender, LinkButton)
        Dim EditLinkButtonCommandArgument As Array = Split(EditLinkButton.CommandArgument, "|")

        Dim HiddenFields As Array = {TB_Selected_Row_Index, TB_Selected_Invoice_No, TB_Selected_Invoice_Date, TB_Selected_Item_Code, TB_Selected_Currency, TB_Selected_Amount, TB_Selected_PO_No}

        '' file value to hidden fields
        For i = 0 To EditLinkButtonCommandArgument.Length - 1
            HiddenFields(i).Text = EditLinkButtonCommandArgument(i)
        Next

        Dim PO_No As String = TB_Selected_PO_No.Text
        Dim Invoice_No As String = TB_Selected_Invoice_No.Text
        Dim Item_Code As String = TB_Selected_Item_Code.Text
        Try
            Dim sqlStr As String = "DELETE FROM DB_Recovered_Invoice WHERE Invoice_No ='" & Invoice_No & "' AND Item_Code = '" & Item_Code & "' AND PO_No LIKE '" & TB_Selected_PO_No.Text & "' "
            RunSQL(sqlStr)
        Catch ex As Exception
            Response.Write("Error:  " & ex.Message)
        End Try

        HandleSplitRouting()
        PopulateGridViewData()
    End Sub

    Protected Sub HandleSplitRouting()
        If Request.QueryString("Subpage") Is Nothing Then
            Response.Redirect("~/Listings/Recovered_Invoices_Details.aspx?Invoice_No=" & Request.QueryString("Invoice_No"))
        Else
            Response.Redirect(String.Format("/Listings/Recovered_Invoices_Details.aspx?Invoice_No={0}&Subpage={1}", Request.QueryString("Invoice_No"), "1"))
        End If
    End Sub


    '' common
    Protected Function GetSQL(ByVal controlObj As Object, Optional ByVal keyword As String = Nothing, Optional ByVal filter As String = Nothing) As String
        Dim sqlStr As String = Nothing
        Dim objID As String = String.Empty
        Dim Search_Keyword As String = EscapeChar(keyword)

        ' Check if the object has an ID property or field
        Dim idProperty = controlObj.GetType().GetProperty("ID")
        If idProperty IsNot Nothing AndAlso idProperty.PropertyType Is GetType(String) Then
            objID = idProperty.GetValue(controlObj)?.ToString()
        End If

        ' Form SQL String based on control object ID
        Select Case objID
            Case "DDL_Item_Code"
                sqlStr = " SELECT Value_1 As [Item Code], Value_1 + ' - ' + Value_2 AS [Item Description] " &
                         " FROM DB_Lookup " &
                         " WHERE Lookup_Name = 'Bill Items' " &
                         "   AND Value_1 NOT IN ('DMC004', 'DMC005', 'DMC013', '00000', '00002') " &
                         "   AND Value_4 NOT IN ('Module Licence Key') " &
                         " ORDER BY Value_4, Value_1 "

            Case "DDL_Currency"
                sqlStr = " SELECT DISTINCT(Value_3) AS Currency FROM DB_Lookup WHERE Lookup_Name = 'Country' AND Value_3 in ('SGD', 'USD', 'EUR') "

            Case "DDL_PO_No"
                sqlStr = " SELECT DISTINCT PO_No FROM I_DB_Recovered_Invoice " &
                         " WHERE Invoice_No = '" & Search_Keyword & "' AND PO_No != '' " &
                         " ORDER BY PO_No"

        End Select

        Return sqlStr
    End Function


    '' Bottom control button
    Protected Sub BT_Close_Click(ByVal sender As Object, ByVal e As EventArgs) Handles BT_Close.Click
        If Request.QueryString("Subpage") Is Nothing Then
            Response.Redirect("~/Listings/Recovered_Invoices.aspx")
        Else
            Response.Write("<script>window.close();</script>")
        End If
    End Sub

End Class
