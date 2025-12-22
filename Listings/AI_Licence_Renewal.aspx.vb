
Partial Class Listings_AI_Licence_Renewal
    Inherits LMSPortalBaseCode

    Dim PageTitle As String = "AI Licence Renewal"

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        LB_PageTitle.Text = PageTitle

        If Not Me.Page.User.Identity.IsAuthenticated AndAlso Session("Login_Status") <> "Logged in" Then
            FormsAuthentication.RedirectToLoginPage()
        End If

        PopulateGridViewData(TB_Search.Text)
    End Sub

    Protected Sub PopulateGridViewData(Optional ByVal TB_Search As String = "")
        Dim keyword As String = EscapeChar(TB_Search)
        Try
            Dim sqlStr = "SELECT [UID], [Customer ID], [Customer], [PO No], [PO Date], [Invoice No], [Invoice Date], [Renewal Date] " &
                         "     , dbo.Get_Licence_Inv_Currency([Invoice No]) As Currency " &
                         "     , dbo.Get_Licence_Inv_Amount([Invoice No], [PO No]) AS Amount " &
                         "FROM R_AI_Licence_Renewal " &
                         "WHERE ([Customer] LIKE '%" & keyword & "%' OR [PO No] LIKE '%" & keyword & "%') " &
                         "GROUP BY [UID], [Customer ID], [Customer], [PO No], [PO Date], [Invoice No], [Invoice Date], [Renewal Date] " &
                         "ORDER BY CASE [Invoice No] WHEN 'NA' THEN 2 ELSE 1 END, UID DESC "

            BuildGridView()

            ''Response.Write(sqlStr & "<br>")
            GridView1.DataSource = GetDataTable(sqlStr)
            GridView1.DataBind()
        Catch ex As Exception
            Response.Write("ERROR - PopulateGridViewData: " & ex.Message)
        End Try

        '' Get the total record
        If keyword <> "" AndAlso CInt(GridView1.Rows.Count.ToString()) > 0 Then
            FooterTotalCount.Visible = True
            FooterTotalCount.Text = "Record(s) found: " & GridView1.Rows.Count.ToString()
        Else
            FooterTotalCount.Visible = False
        End If
    End Sub

    Protected Sub BuildGridView()
        '' GridView Properties
        GridView1.AutoGenerateColumns = False
        GridView1.AllowPaging = True
        GridView1.PageSize = 15
        GridView1.CellPadding = 4
        GridView1.Font.Size = 10
        GridView1.GridLines = GridLines.None
        GridView1.ShowFooter = False
        GridView1.ShowHeaderWhenEmpty = True
        GridView1.DataKeyNames = New String() {"UID"}
        GridView1.CssClass = "table table-bordered"

        '' Header Style
        GridView1.HeaderStyle.CssClass = "table-primary"
        GridView1.HeaderStyle.Font.Bold = True
        GridView1.HeaderStyle.VerticalAlign = VerticalAlign.Top

        '' Row Style
        GridView1.RowStyle.CssClass = "Default"
        GridView1.RowStyle.VerticalAlign = VerticalAlign.Middle

        '' Footer Style
        GridView1.FooterStyle.CssClass = "table-active"

        '' Pager Style
        GridView1.PagerSettings.Mode = PagerButtons.NumericFirstLast
        GridView1.PagerSettings.FirstPageText = "First"
        GridView1.PagerSettings.LastPageText = "Last"
        GridView1.PagerSettings.PageButtonCount = "10"
        GridView1.PagerStyle.HorizontalAlign = HorizontalAlign.Center
        GridView1.PagerStyle.CssClass = "pagination-ys"

        '' Empty Data Template
        GridView1.EmptyDataText = "No records found."

        GridView1.Columns.Clear()

        Dim ColData() As String = {"UID", "Customer", "PO No", "PO Date", "Invoice No", "Invoice Date", "Currency", "Amount", "Renewal Date"}
        Dim ColSize() As Integer = {100, 200, 100, 50, 100, 50, 50, 50, 50}

        '' add template field for the nested gridview
        Dim Expandfield As TemplateField = New TemplateField()
        Expandfield.ItemTemplate = New AIRenewalNestedGridViewItemTemplate()
        Expandfield.HeaderStyle.Width = Unit.Percentage(1)
        GridView1.Columns.Add(Expandfield)

        For i = 0 To ColData.Length - 1
            Dim Bfield As BoundField = New BoundField()
            Bfield.DataField = ColData(i)
            Bfield.HeaderText = Replace(ColData(i), "_", " ")
            Bfield.HeaderStyle.Width = ColSize(i)
            If Bfield.HeaderText.Contains("Date") Then
                Bfield.DataFormatString = "{0:dd MMM yy}"
            End If
            If Bfield.HeaderText.Contains("Amount") Then
                Bfield.DataFormatString = "{0:#,##0.00}"
            End If
            Bfield.HeaderStyle.Wrap = False
            Bfield.ItemStyle.Wrap = False
            GridView1.Columns.Add(Bfield)
        Next

        '' Add template field for the delete button
        Dim TField As TemplateField = New TemplateField()
        TField.HeaderStyle.Width = Unit.Percentage(2)
        TField.ItemStyle.Wrap = False
        TField.ItemTemplate = New GridViewItemTemplateControl()
        GridView1.Columns.Add(TField)

    End Sub


    '' Gridview control
    Protected Sub GridView1_PageIndexChanging(ByVal sender As Object, ByVal e As GridViewPageEventArgs) Handles GridView1.PageIndexChanging
        GridView1.PageIndex = e.NewPageIndex
        PopulateGridViewData(TB_Search.Text)
    End Sub

    Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As GridViewRowEventArgs) Handles GridView1.RowDataBound
        Dim GridViewObj As GridView = CType(sender, GridView)
        GridViewObj.ShowFooter = False

        If e.Row.RowType = DataControlRowType.DataRow Then
            Dim UID As String = GridViewObj.DataKeys(e.Row.RowIndex).Value.ToString()
            Dim PO_No As String = e.Row.Cells(GetColumnIndexByName(e.Row, "PO No")).Text
            Dim PO_Date As String = e.Row.Cells(GetColumnIndexByName(e.Row, "PO Date")).Text
            Dim AILicence As GridView = TryCast(e.Row.FindControl("gvAILicenceList"), GridView)

            If Replace(PO_Date, "&nbsp;", "") = "" Then
                PO_Date = ""
            Else
                CDate(PO_Date).ToString("yyyy-MM-dd")
            End If

            Dim query As String = " SELECT * FROM R_AI_Licence_Renewal " &
                                  " WHERE [UID] = '" & UID & "'  AND [PO No] = '" & PO_No & "' "
            If PO_No <> "NA" Then
                query += " AND [PO Date] = CAST('" & PO_Date & "' As date) "
            End If

            Try
                AILicence.DataSource = GetDataTable(query)
                AILicence.DataBind()
            Catch ex As Exception
                Response.Write("Error:  " & ex.Message)
            End Try

            '' Get Data row details
            Dim drv As System.Data.DataRowView = e.Row.DataItem

            '' Invoice Download Link
            Dim InvoiceDownloadLink As HyperLink = New HyperLink()
            InvoiceDownloadLink.ID = "lnkDownload"

            '' Link to recovored invoice subpage and refresh page
            Dim RecoveredInvoiceLink As HyperLink = New HyperLink()
            RecoveredInvoiceLink.ID = "lnkRecoveredInvoice"

            Dim RefreshInvoiceLink As HyperLink = New HyperLink()
            RefreshInvoiceLink.ID = "lnkRefreshInvoice"

            '' Validate the invoice format
            Dim isInvoiceFormatMatch As Boolean = Regex.IsMatch(drv("Invoice No"), "^TWS/", RegexOptions.IgnoreCase)

            If drv("Invoice No") <> "" Then
                If isInvoiceFormatMatch Then
                    e.Row.Cells(GetColumnIndexByName(e.Row, "Invoice No")).Controls.Add(InvoiceDownloadLink)
                    InvoiceDownloadLink.Text = drv("Invoice No")
                    InvoiceDownloadLink.NavigateUrl = String.Format("/Download/DownloadFile.aspx?Inv_Ref_No={0}", drv("Invoice No"))
                    InvoiceDownloadLink.Target = "_blank"

                    e.Row.Cells(GetColumnIndexByName(e.Row, "Invoice No")).Controls.Add(New LiteralControl("<span class='text-muted mx-2'></span>"))

                    e.Row.Cells(GetColumnIndexByName(e.Row, "Invoice No")).Controls.Add(RecoveredInvoiceLink)
                    RecoveredInvoiceLink.Text = "<i class='bi bi-box-arrow-up-right'></i>"
                    RecoveredInvoiceLink.NavigateUrl = String.Format("/Listings/Recovered_Invoices_Details.aspx?Invoice_No={0}&Subpage={1}", drv("Invoice No"), "1")
                    RecoveredInvoiceLink.ToolTip = "Edit"
                    RecoveredInvoiceLink.CssClass = "btn btn-xs btn-light border"
                    RecoveredInvoiceLink.Target = "_blank"

                    e.Row.Cells(GetColumnIndexByName(e.Row, "Invoice No")).Controls.Add(New LiteralControl("<span class='text-muted mx-1'></span>"))

                    e.Row.Cells(GetColumnIndexByName(e.Row, "Invoice No")).Controls.Add(RefreshInvoiceLink)
                    RefreshInvoiceLink.Text = "<i class='bi bi-arrow-repeat'></i>"
                    RefreshInvoiceLink.NavigateUrl = ResolveUrl(Request.Url.PathAndQuery)
                    RefreshInvoiceLink.ToolTip = "Refresh"
                    RefreshInvoiceLink.CssClass = "btn btn-xs btn-light border"
                Else
                    e.Row.Cells(GetColumnIndexByName(e.Row, "Invoice No")).Text = drv("Invoice No").ToString.ToUpper()
                    e.Row.Cells(GetColumnIndexByName(e.Row, "Invoice No")).Style.Add("font-style", "italic")
                    e.Row.Cells(GetColumnIndexByName(e.Row, "Invoice No")).Style.Add("color", "#999999")
                End If
            End If


            '' Edit Button
            Dim EditctrlCellIndex As Integer = e.Row.Cells.Count - 1
            Dim EditLinkButton As LinkButton = TryCast(e.Row.Cells(EditctrlCellIndex).Controls(0), LinkButton)
            EditLinkButton.CommandArgument = drv("UID")
            EditLinkButton.CommandName = drv("Customer ID") & "|" & drv("PO No")
            EditLinkButton.CausesValidation = False
            AddHandler EditLinkButton.Click, AddressOf Edit_AILicenceRenewal_Click

            EditLinkButton.Style.Add("margin-right", "5px")   '' add separator between button

            '' Void Button
            Dim VoidLinkButton As LinkButton = TryCast(e.Row.Cells(EditctrlCellIndex).Controls(1), LinkButton)
            VoidLinkButton.CommandArgument = drv("UID")
            VoidLinkButton.CommandName = drv("Customer ID") & "|" & drv("PO No")
            VoidLinkButton.Attributes.Add("onclick", "javascript:if (confirm('You are about to void PO No. " & drv("PO No") & "\nClick OK to proceed.\nOtherwise, click Cancel')){return true;} else {return false;}")
            VoidLinkButton.CausesValidation = False
            AddHandler VoidLinkButton.Click, AddressOf VoidAILicenceRenewal_Click

            '' Reset assigned Invoice section
            Dim ResetLinkButton As LinkButton = TryCast(e.Row.Cells(EditctrlCellIndex).Controls(2), LinkButton)
            ResetLinkButton.CommandArgument = drv("UID") & "|" & drv("Customer ID") & "|" & drv("PO No") & "|" & drv("Invoice No")
            ResetLinkButton.Attributes.Add("onclick", "javascript:if (confirm('You are about to reset the " & drv("Invoice No") & " assigned.\nClick OK to proceed.\nOtherwise, click Cancel')){return true;} else {return false;}")
            ResetLinkButton.CausesValidation = False
            AddHandler ResetLinkButton.Click, AddressOf ResetInvoiceAssigned_Click


            Dim renewDate As DateTime
            Dim renewDateText As String = CDate(drv("Renewal Date")).ToString("yyyy-MM-dd")

            '' Lock record if invoice has been recovered
            If drv("Invoice No") = "" And drv("Invoice No") <> UCase("Cancelled") Then
                EditLinkButton.Text = "<i class='bi bi-pencil-fill'></i>"
                EditLinkButton.CssClass = "btn btn-xs btn-info"
                EditLinkButton.Enabled = True

                VoidLinkButton.Text = "<i class='bi bi-x-circle'></i>"
                VoidLinkButton.CssClass = "btn btn-xs btn-warning"
                VoidLinkButton.Enabled = True
            Else
                EditLinkButton.Text = "<i class='bi bi-lock'></i>"
                EditLinkButton.CssClass = "btn btn-xs btn-light disabled"
                EditLinkButton.ToolTip = "Item Locked"
                EditLinkButton.Enabled = False

                If drv("Invoice No") <> UCase("Cancelled") Then
                    If DateTime.TryParse(renewDateText, renewDate) Then
                        If renewDate > DateTime.Today.AddMonths(-3) Then
                            ResetLinkButton.Text = "<i class='bi bi-arrow-90deg-left'></i>"
                            ResetLinkButton.CssClass = "btn btn-xs btn-secondary"
                            ResetLinkButton.Visible = IIf(drv("PO No") <> "NA", True, False)
                        Else
                            ResetLinkButton.Visible = False
                        End If
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub GridView1_RowCreated(sender As Object, e As GridViewRowEventArgs) Handles GridView1.RowCreated
        ' Call javascript function for GridView Row highlight effect
        If e.Row.RowType = DataControlRowType.DataRow Then
            e.Row.Attributes.Add("OnMouseOver", "javascript:SetMouseOver(this);")
            e.Row.Attributes.Add("OnMouseOut", "javascript:SetMouseOut(this);")
        End If
    End Sub


    '' control button events
    Protected Sub Edit_AILicenceRenewal_Click(ByVal sender As Object, ByVal e As EventArgs)
        ModalHeaderAILicenceRenewal.Text = "Recover Invoice"
        btnUpdateAILicenceRenewal.Text = "Update"
        btnCancelAILIcenceRenewal.Text = "Cancel"

        Dim EditLinkButton As LinkButton = TryCast(sender, LinkButton)
        TB_UID.Text = EditLinkButton.CommandArgument

        Dim EditCommandNameArray As String() = Split(EditLinkButton.CommandName, "|")
        TB_Customer_ID.Text = EditCommandNameArray(0)
        TB_PO_No.Text = EditCommandNameArray(1)

        TB_SO_No.Text = Get_Value("SELECT ISNULL(SO_No, '') AS SO_No FROM DB_SO_No_By_PO WHERE Customer_ID ='" & TB_Customer_ID.Text & "' AND PO_No = '" & TB_PO_No.Text & "' ", "SO_No")
        TB_SO_No.Enabled = IIf(Len(TB_SO_No.Text) <= 0, True, False)
        RequiredField_TB_SO_No.Enabled = IIf(Len(TB_SO_No.Text) <= 0, True, False)

        TB_Invoice_No.Text = String.Empty
        TB_Invoice_Date.Text = String.Empty
        RequiredField_TB_Invoice_No.Enabled = True
        RequiredField_TB_Invoice_Date.Enabled = True

        popupAILicenceRenewal.Show()
    End Sub

    Protected Sub Update_AILicenceRenewal_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnUpdateAILicenceRenewal.Click
        Dim UID As TextBox = pnlUpdateAILicenceRenewal.FindControl("TB_UID")
        Dim Customer_ID As TextBox = pnlUpdateAILicenceRenewal.FindControl("TB_Customer_ID")
        Dim PO_No As TextBox = pnlUpdateAILicenceRenewal.FindControl("TB_PO_No")
        Dim SO_NO As TextBox = pnlUpdateAILicenceRenewal.FindControl("TB_SO_No")
        Dim Invoice_No As TextBox = pnlUpdateAILicenceRenewal.FindControl("TB_Invoice_No")
        Dim Invoice_Date As TextBox = pnlUpdateAILicenceRenewal.FindControl("TB_Invoice_Date")

        Try
            Dim sqlStr As String = "UPDATE LMS_AI_Licence_Renewal SET Invoice_No = '" & Invoice_No.Text & "', Invoice_Date = '" & Invoice_Date.Text & "' " &
                                   "WHERE Renewal_UID = '" & UID.Text & "' "

            Dim sqlStr1 As String = "EXEC SP_Insert_AI_Licence_Renewal_Recovered_Invoice_Items '" & UID.Text & "', '" & PO_No.Text & "' "

            Dim sqlStr2 As String = "UPDATE DB_SO_No_By_PO SET SO_No = '" & Trim(SO_NO.Text) & "' WHERE Customer_ID ='" & Customer_ID.Text & "' AND PO_No = '" & PO_No.Text & "' "

            RunSQL(sqlStr)
            RunSQL(sqlStr1)
            RunSQL(sqlStr2)
        Catch ex As Exception
            Response.Write("ERROR - Update_AILicenceRenewal_Click: " & ex.Message)
        End Try

        PopulateGridViewData(TB_Search.Text)
    End Sub

    Protected Sub Cancel_AILIcenceRenewal_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnCancelAILIcenceRenewal.Click
        PopulateGridViewData(TB_Search.Text)
    End Sub

    Protected Sub VoidAILicenceRenewal_Click(ByVal sender As Object, ByVal e As EventArgs)
        Dim VoidLinkButton As LinkButton = TryCast(sender, LinkButton)
        Dim UID As String = VoidLinkButton.CommandArgument

        Try
            Dim sqlStr As String = "UPDATE LMS_AI_Licence_Renewal SET Is_Cancelled = 1 WHERE Renewal_UID = '" & UID & "' "
            RunSQL(sqlStr)
        Catch ex As Exception
            Response.Write("ERROR: " & ex.Message)
        End Try

        PopulateGridViewData(TB_Search.Text)
    End Sub

    Protected Sub ResetInvoiceAssigned_Click(ByVal sender As Object, ByVal e As EventArgs)
        Dim ResetLinkButton As LinkButton = TryCast(sender, LinkButton)
        Dim ResetLinkButtonArray As String() = Split(ResetLinkButton.CommandArgument, "|")
        Dim Module_Type As String = "AI Renewal"
        Dim UID As String = ResetLinkButtonArray(0)
        Dim Customer_ID As String = ResetLinkButtonArray(1)
        Dim PO_No As String = ResetLinkButtonArray(2)
        Dim Invoice_No As String = ResetLinkButtonArray(3)

        Try
            Dim sqlStr As String = "EXEC SP_Reset_Invoice_Assigned N'" & Module_Type &
                                                               "', '" & Customer_ID &
                                                               "', '" & PO_No &
                                                               "', '" & Invoice_No & "' "
            RunSQL(sqlStr)
        Catch ex As Exception
            Response.Write("ERROR: " & ex.Message)
        End Try

        PopulateGridViewData(TB_Search.Text)
    End Sub


    Protected Sub BT_Search_Click(ByVal sender As Object, ByVal e As EventArgs) Handles BT_Search.Click
        PopulateGridViewData(TB_Search.Text)
    End Sub


End Class
