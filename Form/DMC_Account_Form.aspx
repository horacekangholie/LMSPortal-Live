﻿<%@ Page Title="LMSPortal - Register DMC Account" 
    Language="VB" 
    MasterPageFile="~/MasterPages/Frontend.master" 
    AutoEventWireup="false" 
    CodeFile="DMC_Account_Form.aspx.vb" 
    MaintainScrollPositionOnPostback ="true"
    ValidateRequest="false"
    EnableEventValidation="false"
    Inherits="Form_DMC_Account_Form" %>

<%@ Register Assembly="AjaxControlToolkit" Namespace="AjaxControlToolkit" TagPrefix="cc1" %>

<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="Server">
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">

    <!-- Style sheet and Script -->
    <%--<script src="https://ajax.googleapis.com/ajax/libs/jquery/3.4.1/jquery.min.js"></script>--%>
    <script src="https://maxcdn.bootstrapcdn.com/bootstrap/3.4.1/js/bootstrap.min.js"></script>
    <link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/3.4.1/css/bootstrap.min.css">
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.7.2/font/bootstrap-icons.css" />
    <link rel="stylesheet" href="/Content/ContentPage.css" />

    <script src="../Scripts/jquery.blockUI.js" type="text/javascript"></script>

    <script type="text/javascript">
        function BlockUI(elementID) {
            var prm = Sys.WebForms.PageRequestManager.getInstance();
            prm.add_beginRequest(function () {
                $("#" + elementID).block({
                    message: '<table align = "center"><tr><td>' +
                        '<img src="../Images/loadingAnim.gif"/></td></tr></table>',
                    css: {},
                    overlayCSS: {
                        backgroundColor: '#000000', opacity: 0.6
                    }
                });
            });
            prm.add_endRequest(function () {
                $("#" + elementID).unblock();
            });
        }
        $(document).ready(function () {
            BlockUI("<%=pnlAddEditHQ.ClientID %>");
            $.blockUI.defaults.css = {};
        });
        $(document).ready(function () {
            BlockUI("<%=pnlAddEditStore.ClientID %>");
            $.blockUI.defaults.css = {};
        });
        $(document).ready(function () {
            BlockUI("<%=pnlAddEditUser.ClientID %>");
            $.blockUI.defaults.css = {};
        });
        $(document).ready(function () {
            BlockUI("<%=pnlAddEditSubscription.ClientID %>");
            $.blockUI.defaults.css = {};
        });
        $(document).ready(function () {
            BlockUI("<%=pnlAddEditSubscriptionBatch.ClientID %>");
            $.blockUI.defaults.css = {};
        });
        $(document).ready(function () {
            BlockUI("<%=pnlAddEditSubscriptionBatch.ClientID %>");
            $.blockUI.defaults.css = {};
        });
        $(document).ready(function () {
            BlockUI("<%=pnlAddEditNotes.ClientID %>");
            $.blockUI.defaults.css = {};
        });
        function Hidepopup() {
            $find("popup").hide();
            return false;
        }
    </script>

    <style>
        .modalBackground {
            display: block;
            background-color: Black;
            filter: blur(0px);
            opacity: 0.8;
        }

        .modalPopup {
            position: absolute;
            background-color: #FFFFFF;
            border-radius: 5px;
            padding: 0px;
            width:800px;
            height: auto;
        }

        .modalPopupTable {
            display: block;
            padding: 0px 10px 20px 20px;
            table-layout: fixed;
        }
    </style>

    <style>
        .right-stick-textbox {
            border-top-right-radius: 0px;
            border-bottom-right-radius: 0px;
        }
        .left-stick-button {
            border-top-left-radius: 0px;
            border-bottom-left-radius: 0px;
        }

        .default-button-new-product-type{
            width: 100px;
        }

        .guided-text-dropdownlist{
            margin-left: 10px;
            margin-right: 20px;
            font-weight:bold;
            color:#337ab7;
        }

        /*Style for label convert to asp:lable*/
        .label-asp{
            display: inline-block;
            max-width: 100%;
            margin-bottom: 5px;
            font-weight: 700;
            box-sizing: border-box;
        }

        .tab-box{
            padding: 10px 0px 100px 10px;
            border: solid 1px;
            border-color: #fff #dee2e6 #dee2e6;
        }

        .selectedRow {
            background-color: #eeeeee;
        }

        .unselectedRow{
            background-color:transparent;
        }

        /* CSS for the up and down arrows */
        .sort-arrow {
            font-size: 12px; /* You can adjust this value to change the icon size */
        }

        .sort-arrow.asc::after {
            content: '\25B2'; /* Up-arrow character */
        }

        .sort-arrow.desc::after {
            content: '\25BC'; /* Down-arrow character */
        }

        /* CSS to position the arrow next to the header text */
        th {
            position: relative;
            white-space: nowrap; /* Prevent header text from wrapping */
            cursor: pointer; /* Add a pointer cursor to indicate the header is sortable */
        }

        th span.sort-arrow {
            margin-left: 5px; /* Adjust this value as needed to create space between header text and arrow */
        }
    </style>

    <style type="text/css">
        .sorted-column {
            background-color: #ffffe6;
        }
    </style>

    <style>
        .ftpFont{
            font-family: Menlo,Monaco,Consolas,Liberation Mono,courier New,monospace;
            font-size: 15px;
            background-color: #fbfcfc;
        }
    </style>

</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="PageContent" Runat="Server">
    <!-- Page Title -->
    <h2 class="mb-4">
        <asp:Label ID="LB_PageTitle" runat="server"></asp:Label></h2>
    <hr />

    <!-- Form Page Content -->
    <div class="card border-light mb-3">
        <div class="card-header"><h4 class="card-title">Customer Details</h4></div>
        <div class="card-body">
            <asp:FormView ID="FormView1" runat="server" Width="100%">
                <ItemTemplate>
                    <table border="0" style="width: 100%">
                        <tr style="vertical-align: top">
                            <td style="width: 50%; padding: 0px 10px 10px 0px">
                                <label>Name</label>
                                <asp:TextBox ID="TB_Name" Text='<%# Eval("Name") %>' runat="server" CssClass="form-control" Enabled="false" />
                            </td>
                            <td style="width: 20%; padding: 0px 10px 10px 0px">
                                <label>Code</label>
                                <asp:TextBox ID="TB_Code" Text='<%# Eval("Distributor_Code") %>' runat="server" CssClass="form-control" Enabled="false" />
                            </td>
                            <td style="width: 30%; padding: 0px 10px 10px 0px">
                                <label>Country</label>
                                <asp:Label ID="LB_Country" Text='<%# Eval("Country") %>' runat="server" Visible="false" />
                                <asp:DropDownList ID="DDL_Country" runat="server" CssClass="form-control" Enabled="false"
                                    DataSourceID="Country_SqlDataSource"
                                    DataTextField="Country"
                                    DataValueField="Country"
                                    OnDataBound="DDL_Country_DataBound">
                                </asp:DropDownList>
                                <asp:SqlDataSource ID="Country_SqlDataSource" runat="server"
                                    ConnectionString="<%$ ConnectionStrings:lmsConnectionString %>"
                                    SelectCommand="SELECT Value_1 AS Country FROM DB_Lookup WHERE Lookup_Name = 'Country' ORDER BY Value_1"></asp:SqlDataSource>
                            </td>
                        </tr>
                        <tr style="vertical-align: top">
                            <td style="width: 50%; padding: 0px 10px 10px 0px" rowspan="2">
                                <label style="display: block">Address</label>
                                <asp:TextBox ID="TB_Address" Text='<%# Eval("Address") %>' runat="server" TextMode="MultiLine" Enabled="false" BackColor="#eeeeee" />
                            </td>
                            <td style="width: 20%; padding: 0px 10px 10px 0px">
                                <label>Type</label>
                                <asp:Label ID="LB_Type" Text='<%# Eval("Type") %>' runat="server" Visible="false" />
                                <asp:DropDownList ID="DDL_Type" runat="server" CssClass="form-control" Enabled="false"
                                    DataSourceID="Type_SqlDataSource"
                                    DataTextField="Type"
                                    DataValueField="Type"
                                    OnDataBound="DDL_Type_DataBound"
                                    AutoPostBack="true"
                                    OnSelectedIndexChanged="DDL_Type_SelectedIndexChanged">
                                </asp:DropDownList>
                                <asp:SqlDataSource ID="Type_SqlDataSource" runat="server"
                                    ConnectionString="<%$ ConnectionStrings:lmsConnectionString %>"
                                    SelectCommand="SELECT Value_1 AS Type FROM DB_Lookup WHERE Lookup_Name = 'Customer Type'"></asp:SqlDataSource>
                            </td>
                            <td style="width: 30%; padding: 0px 10px 10px 0px">
                                <label>Group</label>
                                <asp:Label ID="LB_Group_ID" Text='<%# Eval("Group_ID") %>' runat="server" Visible="false" />
                                <asp:DropDownList ID="DDL_Group_ID" runat="server" CssClass="form-control" Enabled="false"
                                    DataSourceID="Group_SqlDataSource"
                                    DataTextField="Name"
                                    DataValueField="Group_ID"
                                    OnDataBound="DDL_Group_ID_DataBound">
                                </asp:DropDownList>
                                <asp:SqlDataSource ID="Group_SqlDataSource" runat="server"
                                    ConnectionString="<%$ ConnectionStrings:lmsConnectionString %>"
                                    SelectCommand="SELECT Group_ID, Name FROM Master_Customer_Group"></asp:SqlDataSource>
                            </td>
                        </tr>
                        <tr style="vertical-align: bottom">
                            <td style="width: 50%; padding: 0px 10px 10px 0px" colspan="2">
                                <label id="lbl_By_Distributor" runat="server">By Distributor</label>
                                <asp:Label ID="LB_By_Distributor" Text='<%# Eval("By_Distributor") %>' runat="server" Visible="false" />
                                <asp:DropDownList ID="DDL_By_Distributor" runat="server"
                                    DataSourceID="By_Distributor_SqlDataSource" CssClass="form-control" Enabled="false"
                                    DataTextField="Name"
                                    DataValueField="Customer_ID"
                                    OnDataBound="DDL_By_Distributor_DataBound"
                                    AppendDataBoundItems="True">
                                    <asp:ListItem Text="" Value="" />
                                </asp:DropDownList>
                                <asp:SqlDataSource ID="By_Distributor_SqlDataSource" runat="server"
                                    ConnectionString="<%$ ConnectionStrings:lmsConnectionString %>"
                                    SelectCommand="SELECT DISTINCT Customer_ID, Name FROM Master_Customer WHERE Type NOT IN ('Customer', 'DIGI') AND Name NOT IN ('') ORDER BY Name"></asp:SqlDataSource>
                            </td>
                        </tr>
                        <tr style="vertical-align: top">
                            <td style="width: 50%; padding: 0px 10px 10px 0px">
                                <label>Contact Person</label>
                                <asp:TextBox ID="TB_Contact_Person" Text='<%# Eval("Contact_Person") %>' runat="server" CssClass="form-control" Enabled="false" />
                            </td>
                            <td style="width: 20%; padding: 0px 10px 10px 0px">
                                <label>Phone</label>
                                <asp:TextBox ID="TB_Phone" Text='<%# Eval("Phone") %>' runat="server" CssClass="form-control" Enabled="false" />
                            </td>
                            <td style="width: 30%; padding: 0px 10px 10px 0px">
                                <label>Email</label>
                                <asp:TextBox ID="TB_Email" Text='<%# Eval("Email") %>' runat="server" CssClass="form-control" Enabled="false" />
                            </td>
                        </tr>
                    </table>
                </ItemTemplate>
            </asp:FormView>
        </div>
    </div>

    <asp:ScriptManager ID="ScriptManager1" runat="server"></asp:ScriptManager>

    <!-- DMC Account Content -->
    <div id="DMCAccountContent" runat="server" style="margin-top: 40px">
        <ul class='nav nav-tabs' id="myTab" role="tablist">
            <li id="tab_link1" class="nav-item" role="presentation"><a id="link1" class="nav-link" data-toggle="tab" href="#dmcheadquarter" aria-selected="false" tabindex="-1" role="tab" style="font-size: 15px; width: 180px; text-align: center">DMC Headquarter</a></li>
            <li id="tab_link2" class="nav-item" role="presentation"><a id="link2" class="nav-link active" data-toggle="tab" href="#dmcstore" aria-selected="true" role="tab"  style="font-size: 15px; width: 180px; text-align: center">DMC Store</a></li>
            <li id="tab_link3" class="nav-item" role="presentation"><a id="link3" class="nav-link" data-toggle="tab" href="#dmcuser" aria-selected="false" tabindex="-1" role="tab" style="font-size: 15px; width: 180px; text-align: center">DMC User</a></li>
            <li id="tab_link4" class="nav-item" role="presentation"><a id="link4" class="nav-link" data-toggle="tab" href="#dmcsubscription" aria-selected="false" tabindex="-1" role="tab" style="font-size: 15px; width: 180px; text-align: center">DMC Subscription</a></li>
        </ul>
        <div class="tab-content" id="myTabContent">
            <div class="tab-pane" id="dmcheadquarter" role="tabpanel">
                <!-- DMC Headquarter Details -->
                <asp:UpdatePanel ID="UpdatePanel1" runat="server">
                    <ContentTemplate>
                        <div style="padding: 20px 0px 40px 10px; border: solid 1px; border-color: #fff #dee2e6 #dee2e6">
                            <asp:GridView ID="GridView1" runat="server"></asp:GridView>
                            <asp:Button ID="btnAddHQ" Text="Add New" runat="server" CssClass="btn btn-sm btn-info" CausesValidation="false" />
                        </div>

                        <!-- Modal body -->
                        <asp:Panel ID="pnlAddEditHQ" runat="server" CssClass="modalPopup">
                            <h3 style="padding: 10px 0px 20px 20px">
                                <asp:Label ID="ModalHeaderHQ" runat="server"></asp:Label></h3>
                            <table border="0" class="modalPopupTable" style="width: 100%">
                                <tr style="vertical-align: top">
                                    <td style="width: 300px; padding: 0px 10px 10px 0px">
                                        <label>Headquarter ID</label>
                                        <asp:TextBox ID="TB_Headquarter_ID" runat="server" CssClass="form-control"></asp:TextBox>
                                    </td>
                                    <td style="width: 500px; padding: 0px 10px 10px 0px">
                                        <label>Headquarter Name</label>
                                        <asp:TextBox ID="TB_Name" runat="server" CssClass="form-control"></asp:TextBox>
                                        <asp:RequiredFieldValidator ID="RequiredField_TB_Name" runat="server" ValidationGroup="Headquarter"
                                            ErrorMessage="This field cannot be blank" CssClass="invalid-feedback"
                                            Display="Dynamic"
                                            SetFocusOnError="True"
                                            ControlToValidate="TB_Name">
                                        </asp:RequiredFieldValidator>
                                    </td>
                                </tr>
                            </table>
                            <table border="0" class="modalPopupTable" style="width: 100%">
                                <tr style="vertical-align: top">
                                    <td style="width: 300px; padding: 0px 10px 10px 0px">
                                        <label>Status</label>
                                        <asp:DropDownList ID="DDL_HQ_Status" runat="server" CssClass="form-control">
                                            <asp:ListItem Value="1" Text="Active"></asp:ListItem>
                                            <asp:ListItem Value="0" Text="Inactive"></asp:ListItem>
                                        </asp:DropDownList>
                                    </td>
                                    <td style="width: 500px; padding: 0px 10px 10px 0px">
                                        <label>Requested By</label>
                                        <asp:DropDownList ID="DDL_Sales_Representative" runat="server" CssClass="form-control" AppendDataBoundItems="true">
                                            <asp:ListItem Value="-1" Text="Please select"></asp:ListItem>
                                        </asp:DropDownList>
                                        <asp:CompareValidator ID="CompareValidator_DDL_Sales_Representative" runat="server" ValidationGroup="Headquarter"
                                            ErrorMessage="Please select Requestor" CssClass="invalid-feedback"
                                            Display="Dynamic"
                                            SetFocusOnError="True"
                                            Operator="NotEqual" ValueToCompare="-1"
                                            ControlToValidate="DDL_Sales_Representative" Type="String">
                                        </asp:CompareValidator>
                                    </td>
                                </tr>
                            </table>
                            <!-- hidden field -->
                            <div>
                                <asp:TextBox ID="TB_Selected_HQ_Row_Index" runat="server" CssClass="form-control" Visible="false"></asp:TextBox>
                                <asp:TextBox ID="TB_Selected_HQ_ID" runat="server" CssClass="form-control" Visible="false"></asp:TextBox>
                                <asp:TextBox ID="TB_Selected_HQ_Name" runat="server" CssClass="form-control" Visible="false"></asp:TextBox>
                                <asp:TextBox ID="TB_Selected_HQ_Status" runat="server" CssClass="form-control" Visible="false"></asp:TextBox>
                                <asp:TextBox ID="TB_Selected_HQ_Requested_By" runat="server" CssClass="form-control" Visible="false"></asp:TextBox>
                            </div>
                            <table border="0" class="modalPopupTable" style="width: 100%">
                                <tr style="vertical-align: bottom">
                                    <td style="width: 800px; padding: 10px 10px 10px 0px" colspan="2">
                                        <asp:Button ID="btnSaveHQ" runat="server" CssClass="btn btn-sm btn-default" ValidationGroup="Headquarter" />
                                        <asp:Button ID="btnCancelHQ" runat="server" CssClass="btn btn-sm btn-default" CausesValidation="false" OnClientClick="return Hidepopup()" />
                                    </td>
                                </tr>
                            </table>
                        </asp:Panel>

                        <asp:LinkButton ID="lnkFakeHQ" runat="server"></asp:LinkButton>
                        <cc1:ModalPopupExtender ID="popupHQ" runat="server" DropShadow="false"
                            PopupControlID="pnlAddEditHQ" TargetControlID="lnkFakeHQ"
                            BackgroundCssClass="modalBackground">
                        </cc1:ModalPopupExtender>
                    </ContentTemplate>

                    <Triggers>
                        <asp:AsyncPostBackTrigger ControlID="GridView1" />
                        <asp:AsyncPostBackTrigger ControlID="btnSaveHQ" />
                    </Triggers>

                </asp:UpdatePanel>
            </div>
            <div class="tab-pane active" id="dmcstore" role="tabpanel">
                <!-- DMC Store Details -->
                <asp:UpdatePanel ID="UpdatePanel2" runat="server">
                    <ContentTemplate>
                        <div style="padding: 20px 0px 40px 10px; border: solid 1px; border-color: #fff #dee2e6 #dee2e6">
                            <asp:GridView ID="GridView2" runat="server"></asp:GridView>
                            <asp:Button ID="btnAddStore" Text="Add New" runat="server" CssClass="btn btn-sm btn-info" CausesValidation="false" />
                            <div class="alert alert-dismissible alert-secondary" style="margin-top: 15px; margin-right: 10px">
                                <h5 class="alert-heading">Store status:</h5>
                                <h6>
                                    <p><b>Active</b> - Current active store.</p>
                                    <p><b>Suspended</b> - Expired store, but still within 30-day period, charges applied for account re-instatement.</p>
                                    <p><b>Closed</b> - Account permanently closed, no reinstatement allowed.</p>
                                </h6>
                            </div>
                        </div>

                        <!-- Modal body -->
                        <asp:Panel ID="pnlAddEditStore" runat="server" CssClass="modalPopup">
                            <h3 style="padding: 10px 0px 20px 20px">
                                <asp:Label ID="ModalHeaderStore" runat="server"></asp:Label></h3>
                            <table border="0" class="modalPopupTable" style="width: 100%">
                                <tr style="vertical-align: top">
                                    <td style="width: 800px; padding: 0px 10px 20px 0px" colspan="2">
                                        <label>Headquarter</label>
                                        <asp:DropDownList ID="DDL_Store_Headquarter" runat="server" CssClass="form-control" AppendDataBoundItems="true" AutoPostBack="true">
                                            <asp:ListItem Value="-1" Text="Please select"></asp:ListItem>
                                        </asp:DropDownList>
                                        <asp:CompareValidator ID="CompareValidator_DDL_Store_Headquarter" runat="server" ValidationGroup="Store"
                                            ErrorMessage="Please select Headquarter" CssClass="invalid-feedback"
                                            Display="Dynamic"
                                            SetFocusOnError="True"
                                            Operator="NotEqual" ValueToCompare="-1"
                                            ControlToValidate="DDL_Store_Headquarter" Type="String">
                                        </asp:CompareValidator>
                                    </td>
                                </tr>
                                <tr style="vertical-align: top">
                                    <td style="width: 400px; padding: 0px 10px 20px 0px">
                                        <label>Store No</label>
                                        <asp:TextBox ID="TB_Store_No" runat="server" CssClass="form-control"></asp:TextBox>
                                    </td>
                                    <td style="width: 400px; padding: 0px 10px 20px 0px">
                                        <label>Public IP</label>
                                        <asp:TextBox ID="TB_Public_IP" runat="server" CssClass="form-control" placeholder="172.168.0.0"></asp:TextBox>
                                        <asp:RequiredFieldValidator ID="Required_TB_Public_IP" runat="server" ValidationGroup="Store"
                                            ErrorMessage="This field cannot be blank" CssClass="invalid-feedback"
                                            Display="Dynamic"
                                            SetFocusOnError="True"
                                            ControlToValidate="TB_Public_IP">
                                        </asp:RequiredFieldValidator>
                                    </td>
                                </tr>
                                <tr style="vertical-align: top">
                                    <td style="width: 400px; padding: 0px 10px 20px 0px">
                                        <label>Store Name</label>
                                        <asp:TextBox ID="TB_Store_Name" runat="server" CssClass="form-control"></asp:TextBox>
                                        <asp:RequiredFieldValidator ID="RequiredField_TB_Store_Name" runat="server" ValidationGroup="Store"
                                            ErrorMessage="This field cannot be blank" CssClass="invalid-feedback"
                                            Display="Dynamic"
                                            SetFocusOnError="True"
                                            ControlToValidate="TB_Store_Name">
                                        </asp:RequiredFieldValidator>
                                    </td>
                                    <td style="width: 400px; padding: 0px 10px 20px 0px">
                                        <label>FTP User</label>
                                        <asp:TextBox ID="TB_FTP_User" runat="server" CssClass="form-control"></asp:TextBox>
                                        <asp:RequiredFieldValidator ID="Required_TB_FTP_User" runat="server" ValidationGroup="Store"
                                            ErrorMessage="This field cannot be blank" CssClass="invalid-feedback"
                                            Display="Dynamic"
                                            SetFocusOnError="True"
                                            ControlToValidate="TB_FTP_User">
                                        </asp:RequiredFieldValidator>
                                    </td>
                                </tr>
                                <tr style="vertical-align: top">
                                    <td style="width: 400px; padding: 0px 10px 20px 0px">
                                        <label>Banner</label>
                                        <asp:TextBox ID="TB_Banner" runat="server" CssClass="form-control"></asp:TextBox>
                                        <asp:RequiredFieldValidator ID="Required_TB_Banner" runat="server" ValidationGroup="Store"
                                            ErrorMessage="This field cannot be blank" CssClass="invalid-feedback"
                                            Display="Dynamic"
                                            SetFocusOnError="True"
                                            ControlToValidate="TB_Banner">
                                        </asp:RequiredFieldValidator>
                                    </td>
                                    <td style="width: 400px; padding: 0px 10px 20px 0px">
                                        <label>FTP Password</label>
                                        <asp:TextBox ID="TB_FTP_Password" runat="server" CssClass="form-control"></asp:TextBox>
                                        <asp:RequiredFieldValidator ID="Required_TB_FTP_Password" runat="server" ValidationGroup="Store"
                                            ErrorMessage="This field cannot be blank" CssClass="invalid-feedback"
                                            Display="Dynamic"
                                            SetFocusOnError="True"
                                            ControlToValidate="TB_FTP_Password">
                                        </asp:RequiredFieldValidator>
                                    </td>
                                </tr>
                                <tr style="vertical-align: top">
                                    <td style="width: 400px; padding: 0px 10px 20px 0px">
                                        <label>Zone</label>
                                        <asp:TextBox ID="TB_Zone" runat="server" CssClass="form-control" placeholder="(optional)"></asp:TextBox>
                                    </td>
                                    <td style="width: 400px; padding: 0px 10px 20px 0px">
                                        <label>Account Type</label>
                                        <asp:DropDownList ID="DDL_Type_Code" runat="server" CssClass="form-control" AppendDataBoundItems="true" AutoPostBack="true">
                                            <asp:ListItem Value="-1" Text="Please select"></asp:ListItem>
                                        </asp:DropDownList>
                                        <asp:CompareValidator ID="CompareValidator_DDL_Type_Code" runat="server" ValidationGroup="Store"
                                            ErrorMessage="Please select account type" CssClass="invalid-feedback"
                                            Display="Dynamic"
                                            SetFocusOnError="True"
                                            Operator="NotEqual" ValueToCompare="-1"
                                            ControlToValidate="DDL_Type_Code" Type="String">
                                        </asp:CompareValidator>
                                    </td>
                                </tr>
                                <tr id="trTrialAccountStatusRow" runat="server" style="vertical-align: top">
                                    <td style="width: 400px; padding: 0px 10px 20px 0px">
                                        <label>Status</label>
                                        <asp:DropDownList ID="DDL_Store_Status" runat="server" CssClass="form-control" AutoPostBack="true">
                                            <asp:ListItem Value="1" Text="Active"></asp:ListItem>
                                            <asp:ListItem Value="0" Text="Suspended"></asp:ListItem>
                                            <%--<asp:ListItem Value="0" Text="Closed"></asp:ListItem>--%>
                                        </asp:DropDownList>
                                    </td>
                                    <td style="width: 400px; padding: 0px 10px 20px 0px">
                                        <label>Expired On</label>
                                        <asp:TextBox ID="TB_Store_End_Date" runat="server" Text="" CssClass="form-control" TextMode="Date"></asp:TextBox>
                                    </td>
                                </tr>
                            </table>
                            <!-- hidden field -->
                            <div>
                                <asp:TextBox ID="TB_Selected_Store_Row_Index" runat="server" CssClass="form-control" Visible="false"></asp:TextBox>
                                <asp:TextBox ID="TB_Selected_Store_HQ_ID" runat="server" CssClass="form-control" Visible="false"></asp:TextBox>
                                <asp:TextBox ID="TB_Selected_Store_No" runat="server" CssClass="form-control" Visible="false"></asp:TextBox>
                                <asp:TextBox ID="TB_Selected_Store_Name" runat="server" CssClass="form-control" Visible="false"></asp:TextBox>
                                <asp:TextBox ID="TB_Selected_Store_Banner" runat="server" CssClass="form-control" Visible="false"></asp:TextBox>
                                <asp:TextBox ID="TB_Selected_Store_Zone" runat="server" CssClass="form-control" Visible="false"></asp:TextBox>
                                <asp:TextBox ID="TB_Selected_Store_Public_IP" runat="server" CssClass="form-control" Visible="false"></asp:TextBox>
                                <asp:TextBox ID="TB_Selected_Store_FTP_User" runat="server" CssClass="form-control" Visible="false"></asp:TextBox>
                                <asp:TextBox ID="TB_Selected_Store_FTP_Password" runat="server" CssClass="form-control" Visible="false"></asp:TextBox>
                                <asp:TextBox ID="TB_Selected_Store_Account_Type_Code" runat="server" CssClass="form-control" Visible="false"></asp:TextBox>
                                <asp:TextBox ID="TB_Selected_Store_Account_Status" runat="server" CssClass="form-control" Visible="false"></asp:TextBox>
                                <asp:TextBox ID="TB_Selected_Store_Account_End_Date" runat="server" CssClass="form-control" Visible="false"></asp:TextBox>
                                <asp:TextBox ID="TB_Selected_Store_ID" runat="server" CssClass="form-control" Visible="false"></asp:TextBox>
                            </div>
                            <table border="0" class="modalPopupTable" style="width: 100%">
                                <tr style="vertical-align: bottom">
                                    <td style="width: 800px; padding: 10px 10px 10px 0px" colspan="2">
                                        <asp:Button ID="btnSaveStore" runat="server" CssClass="btn btn-sm btn-default" ValidationGroup="Store" />
                                        <asp:Button ID="btnCancelStore" runat="server" CssClass="btn btn-sm btn-default" CausesValidation="false" OnClientClick="return Hidepopup()" />
                                    </td>
                                </tr>
                            </table>
                        </asp:Panel>
                        <asp:LinkButton ID="lnkFakeStore" runat="server"></asp:LinkButton>
                        <cc1:ModalPopupExtender ID="popupStore" runat="server" DropShadow="false"
                            PopupControlID="pnlAddEditStore" TargetControlID="lnkFakeStore"
                            BackgroundCssClass="modalBackground">
                        </cc1:ModalPopupExtender>
                    </ContentTemplate>

                    <Triggers>
                        <asp:AsyncPostBackTrigger ControlID="GridView2" />
                        <asp:AsyncPostBackTrigger ControlID="btnSaveStore" />
                    </Triggers>

                </asp:UpdatePanel>
            </div>
            <div class="tab-pane" id="dmcuser" role="tabpanel">
                <!-- DMC User Details -->
                <asp:UpdatePanel ID="UpdatePanel3" runat="server">
                    <ContentTemplate>
                        <div style="padding: 20px 0px 40px 10px; border: solid 1px; border-color: #fff #dee2e6 #dee2e6">
                            <asp:GridView ID="GridView3" runat="server"></asp:GridView>
                            <asp:Button ID="btnAddUser" Text="Add New" runat="server" CausesValidation="false" CssClass="btn btn-sm btn-info" />
                        </div>

                        <!-- Modal body -->
                        <asp:Panel ID="pnlAddEditUser" runat="server" CssClass="modalPopup">
                            <h3 style="padding: 10px 0px 20px 20px">
                                <asp:Label ID="ModalHeaderUser" runat="server"></asp:Label></h3>
                            <table border="0" class="modalPopupTable" style="width: 100%">
                                <tr style="vertical-align: top">
                                    <td style="width: 800px; padding: 0px 10px 20px 0px" colspan="2">
                                        <label>Headquarter</label>
                                        <asp:DropDownList ID="DDL_User_Headquarter" runat="server" CssClass="form-control" AppendDataBoundItems="true" AutoPostBack="true">
                                            <asp:ListItem Value="-1" Text="Please select"></asp:ListItem>
                                        </asp:DropDownList>
                                        <asp:CompareValidator ID="CompareValidator_DDL_User_Headquarter" runat="server" ValidationGroup="User"
                                            ErrorMessage="Please select Headquarter" CssClass="invalid-feedback"
                                            Display="Dynamic"
                                            SetFocusOnError="True"
                                            Operator="NotEqual" ValueToCompare="-1"
                                            ControlToValidate="DDL_User_Headquarter" Type="String">
                                        </asp:CompareValidator>
                                    </td>
                                </tr>
                                <tr style="vertical-align: top">
                                    <td style="width: 400px; padding: 0px 10px 20px 0px">
                                        <label>Username</label>
                                        <asp:TextBox ID="TB_Username" runat="server" CssClass="form-control"></asp:TextBox>
                                    </td>
                                    <td style="width: 400px; padding: 0px 10px 20px 0px">
                                        <label>Password</label>
                                        <asp:TextBox ID="TB_Password" runat="server" CssClass="form-control"></asp:TextBox>
                                        <asp:RequiredFieldValidator ID="Required_TB_Password" runat="server" ValidationGroup="User"
                                            ErrorMessage="This field cannot be blank" CssClass="invalid-feedback"
                                            Display="Dynamic"
                                            SetFocusOnError="True"
                                            ControlToValidate="TB_Password">
                                        </asp:RequiredFieldValidator>
                                    </td>
                                </tr>
                                <tr style="vertical-align: top">
                                    <td style="width: 800px; padding: 0px 10px 20px 0px" colspan="2">
                                        <label>Email</label>
                                        <asp:TextBox ID="TB_Email" runat="server" CssClass="form-control"></asp:TextBox>
                                        <asp:RequiredFieldValidator ID="Required_TB_Email" runat="server" ValidationGroup="User"
                                            ErrorMessage="This field cannot be blank" CssClass="invalid-feedback"
                                            Display="Dynamic"
                                            SetFocusOnError="True"
                                            ControlToValidate="TB_Email">
                                        </asp:RequiredFieldValidator>
                                        <asp:RegularExpressionValidator ID="RegEx_TB_Email" runat="server" ValidationGroup="User"
                                            ErrorMessage="Invalid email address" CssClass="invalid-feedback"
                                            Display="Dynamic"
                                            SetFocusOnError="true"
                                            ControlToValidate="TB_Email"
                                            ValidationExpression="^([\w-\.]+)@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.)|(([\w-]+\.)+))([a-zA-Z]{2,4}|[0-9]{1,3})(\]?)$" />
                                    </td>
                                </tr>
                                <tr style="vertical-align: top">
                                    <td style="width: 400px; padding: 0px 10px 20px 0px">
                                        <label>Device Type</label>
                                        <asp:DropDownList ID="DDL_Device_Type" runat="server" CssClass="form-control" AppendDataBoundItems="true">
                                            <asp:ListItem Value="-1" Text="Please select"></asp:ListItem>
                                        </asp:DropDownList>
                                        <asp:CompareValidator ID="CompareValidator_DDL_Device_Type" runat="server" ValidationGroup="User"
                                            ErrorMessage="Please select device type" CssClass="invalid-feedback"
                                            Display="Dynamic"
                                            SetFocusOnError="True"
                                            Operator="NotEqual" ValueToCompare="-1"
                                            ControlToValidate="DDL_Device_Type" Type="String">
                                        </asp:CompareValidator>
                                    </td>
                                    <td style="width: 400px; padding: 0px 10px 20px 0px"></td>
                                </tr>
                                <tr id="trUserAccountStatusRow" runat="server" style="vertical-align: top">
                                    <td style="width: 400px; padding: 0px 10px 20px 0px">
                                        <label>Status</label>
                                        <asp:DropDownList ID="DDL_User_Status" runat="server" CssClass="form-control" AutoPostBack="true">
                                            <asp:ListItem Value="1" Text="Active"></asp:ListItem>
                                            <asp:ListItem Value="0" Text="Suspended"></asp:ListItem>
                                            <%--<asp:ListItem Value="0" Text="Closed"></asp:ListItem>--%>
                                        </asp:DropDownList>
                                    </td>
                                    <td style="width: 400px; padding: 0px 10px 20px 0px">
                                        <label>Expired On</label>
                                        <asp:TextBox ID="TB_User_End_Date" runat="server" CssClass="form-control" TextMode="Date"></asp:TextBox>
                                        <asp:RequiredFieldValidator ID="Required_TB_User_End_Date" runat="server" ValidationGroup="User"
                                            ErrorMessage="Field cannot be blank if status is not Active" CssClass="invalid-feedback" Enabled="false"
                                            Display="Dynamic"
                                            SetFocusOnError="True"
                                            ControlToValidate="TB_User_End_Date">
                                        </asp:RequiredFieldValidator>
                                    </td>
                                </tr>
                            </table>
                            <!-- hidden field -->
                            <div>
                                <asp:TextBox ID="TB_Selected_User_Row_Index" runat="server" CssClass="form-control" Visible="false"></asp:TextBox>
                                <asp:TextBox ID="TB_Selected_User_HQ_ID" runat="server" CssClass="form-control" Visible="false"></asp:TextBox>
                                <asp:TextBox ID="TB_Selected_User_Username" runat="server" CssClass="form-control" Visible="false"></asp:TextBox>
                                <asp:TextBox ID="TB_Selected_User_Password" runat="server" CssClass="form-control" Visible="false"></asp:TextBox>
                                <asp:TextBox ID="TB_Selected_User_Email" runat="server" CssClass="form-control" Visible="false"></asp:TextBox>
                                <asp:TextBox ID="TB_Selected_User_Device_Type" runat="server" CssClass="form-control" Visible="false"></asp:TextBox>
                                <asp:TextBox ID="TB_Selected_User_Account_Status" runat="server" CssClass="form-control" Visible="false"></asp:TextBox>
                                <asp:TextBox ID="TB_Selected_User_Account_End_Date" runat="server" CssClass="form-control" Visible="false"></asp:TextBox>
                            </div>
                            <table border="0" class="modalPopupTable" style="width: 100%">
                                <tr style="vertical-align: bottom">
                                    <td style="width: 800px; padding: 10px 10px 10px 0px" colspan="2">
                                        <asp:Button ID="btnSaveUser" runat="server" CssClass="btn btn-sm btn-default" ValidationGroup="User" />
                                        <asp:Button ID="btnCancelUser" runat="server" CssClass="btn btn-sm btn-default" CausesValidation="false" OnClientClick="return Hidepopup()" />
                                    </td>
                                </tr>
                            </table>
                        </asp:Panel>
                        <asp:LinkButton ID="lnkFakeUser" runat="server"></asp:LinkButton>
                        <cc1:ModalPopupExtender ID="popupUser" runat="server" DropShadow="false"
                            PopupControlID="pnlAddEditUser" TargetControlID="lnkFakeUser"
                            BackgroundCssClass="modalBackground">
                        </cc1:ModalPopupExtender>
                    </ContentTemplate>

                    <Triggers>
                        <asp:AsyncPostBackTrigger ControlID="GridView3" />
                        <asp:AsyncPostBackTrigger ControlID="btnSaveUser" />
                    </Triggers>

                </asp:UpdatePanel>
            </div>
            <div class="tab-pane" id="dmcsubscription" role="tabpanel">
                <!-- DMC Subscription Details -->
                <asp:UpdatePanel ID="UpdatePanel4" runat="server">
                    <ContentTemplate>
                        <div style="padding: 20px 0px 40px 10px; border: solid 1px; border-color: #fff #dee2e6 #dee2e6">
                            <!-- Dropdownlist for Store -->
                            <div class="form-group form-inline" style="float: left; margin-left: 0px; margin-top: 0px">
                                <asp:Label ID="LB_Store_Contract_List" runat="server" Text="Filter By Store" CssClass="guided-text-dropdownlist"></asp:Label>
                                <asp:DropDownList ID="DDL_Store_Contract_List" runat="server" CssClass="form-control" Width="400px" AutoPostBack="true"
                                    AppendDataBoundItems="true">
                                </asp:DropDownList>
                            </div>
                            
                            <asp:GridView ID="GridView4" runat="server"></asp:GridView>
                            <asp:Button ID="btnAddSubscription" Text="Add New" runat="server" CausesValidation="false" CssClass="btn btn-sm btn-info" />
                            <asp:Button ID="btnAddSubscriptionBatch" Text="Add New (Batch)" runat="server" CausesValidation="false" CssClass="btn btn-sm btn-info" />
                            <div class="alert alert-dismissible alert-info" style="margin-top: 15px; margin-right: 10px">
                                <h5 class="alert-heading">Note:</h5>
                                <h6>
                                    <p>List of DMC Subscription record for the past 3 years</p>
                                </h6>
                            </div>
                        </div>

                        <!-- Modal body -->
                        <asp:Panel ID="pnlAddEditSubscription" runat="server" CssClass="modalPopup">
                            <h3 style="padding: 10px 0px 20px 20px"><asp:Label ID="ModalHeaderSubscription" runat="server"></asp:Label></h3>
                            <table border="0" class="modalPopupTable" style="width: 100%">
                                <tr style="vertical-align: top">
                                    <td style="width: 400px; padding: 0px 10px 20px 0px" colspan="2">
                                        <label>Headquarter</label>
                                        <asp:DropDownList ID="DDL_Subscription_Headquarter" runat="server" CssClass="form-control" AppendDataBoundItems="true" AutoPostBack="true">
                                            <asp:ListItem Value="-1" Text="Please select"></asp:ListItem>
                                        </asp:DropDownList>
                                        <asp:CompareValidator ID="CompareValidator_DDL_Subscription_Headquarter" runat="server" ValidationGroup="SubscriptionLineItems"
                                            ErrorMessage="Please select Headquarter" CssClass="invalid-feedback"
                                            Display="Dynamic"
                                            SetFocusOnError="True"
                                            Operator="NotEqual" ValueToCompare="-1"
                                            ControlToValidate="DDL_Subscription_Headquarter" Type="String">
                                        </asp:CompareValidator>
                                    </td>
                                    <td style="width: 400px; padding: 0px 10px 20px 0px" colspan="2">
                                        <label>Store</label>
                                        <asp:DropDownList ID="DDL_Subscription_Store" runat="server" CssClass="form-control" AppendDataBoundItems="true" AutoPostBack="true">
                                            <asp:ListItem Value="-1" Text="Please select"></asp:ListItem>
                                        </asp:DropDownList>
                                        <asp:CompareValidator ID="CompareValidator_DDL_Subscription_Store" runat="server" ValidationGroup="SubscriptionLineItems"
                                            ErrorMessage="Please select store" CssClass="invalid-feedback"
                                            Display="Dynamic"
                                            SetFocusOnError="True"
                                            Operator="NotEqual" ValueToCompare="-1"
                                            ControlToValidate="DDL_Subscription_Store" Type="String">
                                        </asp:CompareValidator>
                                    </td>
                                </tr>
                                <tr style="vertical-align: top">
                                    <td style="width: 400px; padding: 0px 10px 20px 0px" colspan="2">
                                        <label>Duration</label>
                                        <asp:DropDownList ID="DDL_Subscription_Duration" runat="server" CssClass="form-control" AppendDataBoundItems="true">
                                            <asp:ListItem Value="-1" Text="Please select"></asp:ListItem>
                                        </asp:DropDownList>
                                        <asp:CompareValidator ID="CompareValidator_DDL_Subscription_Duration" runat="server" ValidationGroup="SubscriptionLineItems"
                                            ErrorMessage="Please select duration" CssClass="invalid-feedback"
                                            Display="Dynamic"
                                            SetFocusOnError="True"
                                            Operator="NotEqual" ValueToCompare="-1"
                                            ControlToValidate="DDL_Subscription_Duration" Type="String">
                                        </asp:CompareValidator>
                                    </td>
                                    <td style="width: 200px; padding: 0px 10px 20px 0px">
                                        <label>Currency</label>
                                        <asp:DropDownList ID="DDL_Subscription_Currency" runat="server" CssClass="form-control" AppendDataBoundItems="true">
                                            <asp:ListItem Value="-1" Text="Please select"></asp:ListItem>
                                        </asp:DropDownList>
                                        <asp:CompareValidator ID="CompareValidator_DDL_Subscription_Currency" runat="server" ValidationGroup="SubscriptionLineItems"
                                            ErrorMessage="Please select currency" CssClass="invalid-feedback"
                                            Display="Dynamic"
                                            SetFocusOnError="True"
                                            Operator="NotEqual" ValueToCompare="-1"
                                            ControlToValidate="DDL_Subscription_Currency" Type="String">
                                        </asp:CompareValidator>
                                    </td>
                                    <td style="width: 200px; padding: 0px 10px 20px 0px">
                                        <label>Fee</label>
                                        <asp:TextBox ID="TB_Subscription_Fee" runat="server" CssClass="form-control" placeholder="0.00"></asp:TextBox>
                                        <asp:RequiredFieldValidator ID="RequiredField_TB_Subscription_Fee" runat="server" ValidationGroup="SubscriptionLineItems"
                                            ErrorMessage="This field cannot be blank" CssClass="invalid-feedback"
                                            Display="Dynamic"
                                            SetFocusOnError="True"
                                            ControlToValidate="TB_Subscription_Fee">
                                        </asp:RequiredFieldValidator>
                                        <asp:RegularExpressionValidator ID="RegEx_TB_Subscription_Fee" runat="server" ValidationGroup="SubscriptionLineItems"
                                            ErrorMessage="Fee must be numeric value" CssClass="invalid-feedback"
                                            Display="Dynamic"
                                            SetFocusOnError="True"
                                            ControlToValidate="TB_Subscription_Fee"
                                            ValidationExpression="^[+-]?(?:\d+\.?\d*|\d*\.?\d+)[\r\n]*$">
                                        </asp:RegularExpressionValidator>
                                    </td>
                                </tr>
                                <tr style="vertical-align: top">
                                    <td style="width: 200px; padding: 0px 10px 20px 0px">
                                        <label>Payment Method</label>
                                        <asp:DropDownList ID="DDL_Subscription_Payment_Method" runat="server" CssClass="form-control"></asp:DropDownList>
                                    </td>
                                    <td style="width: 200px; padding: 0px 10px 20px 0px">
                                        <label>Payment Mode</label>
                                        <asp:DropDownList ID="DDL_Subscription_Payment_Mode" runat="server" CssClass="form-control"></asp:DropDownList>
                                    </td>
                                    <td style="width: 400px; padding: 0px 10px 20px 0px" colspan="2">
                                        <label>Subscriber Group</label>
                                        <asp:DropDownList ID="DDL_Subscription_Subscriber_Group" runat="server" CssClass="form-control" AppendDataBoundItems="true">
                                            <asp:ListItem Value="-1" Text="Please select"></asp:ListItem>
                                        </asp:DropDownList>
                                        <asp:CompareValidator ID="CompareValidator_DDL_Subscription_Subscriber_Group" runat="server" ValidationGroup="SubscriptionLineItems"
                                            ErrorMessage="Please select subscriber group" CssClass="invalid-feedback"
                                            Display="Dynamic"
                                            SetFocusOnError="True"
                                            Operator="NotEqual" ValueToCompare="-1"
                                            ControlToValidate="DDL_Subscription_Subscriber_Group" Type="String">
                                        </asp:CompareValidator>
                                    </td>
                                </tr>
                                <tr style="vertical-align: top">
                                    <td style="width: 800px; padding: 0px 10px 10px 0px" colspan="4">
                                        <asp:Button ID="btnAddLineItems" runat="server" Text="Add" CssClass="btn btn-sm btn-info" ValidationGroup="SubscriptionLineItems" />
                                        <asp:Button ID="btnClearLineItems" runat="server" Text="Clear" CssClass="btn btn-sm btn-info" CausesValidation="false" />
                                        <div class="text text-info" style="margin-top: 10px">Add store(s) to renew subscription</div>
                                    </td>
                                </tr>
                                <tr style="vertical-align: top">
                                    <td style="width: 800px; padding: 0px 10px 0px 0px" colspan="4">
                                        <div class="GridList">
                                            <asp:GridView ID="GridView_Store_List" runat="server"
                                                AutoGenerateColumns="false"
                                                AllowPaging="false"
                                                CellPadding="4"
                                                Font-Size="10"
                                                GridLines="None"
                                                ShowHeaderWhenEmpty="true"
                                                CssClass="table table-bordered">
                                                <HeaderStyle CssClass="table-secondary" Font-Bold="true" VerticalAlign="Top" />
                                                <RowStyle CssClass="Default" VerticalAlign="Middle" />
                                                <Columns>
                                                    <asp:BoundField DataField="Headquarter_ID" HeaderText="HQ ID" SortExpression="Headquarter_ID" />
                                                    <asp:BoundField DataField="Store_ID" HeaderText="Store ID" SortExpression="Store_ID" Visible="false" />
                                                    <asp:BoundField DataField="Store_No" HeaderText="Store No" SortExpression="Store_No" />
                                                    <asp:BoundField DataField="Duration" HeaderText="Duration" SortExpression="Duration" />
                                                    <asp:BoundField DataField="Currency" HeaderText="Currency" SortExpression="Currency" />
                                                    <asp:BoundField DataField="Fee" HeaderText="Fee" SortExpression="Fee" DataFormatString="{0:#0.00}" />
                                                    <asp:BoundField DataField="Payment_Method" HeaderText="Payment" SortExpression="Payment_Method" />
                                                    <asp:BoundField DataField="Payment_Mode" HeaderText="Mode" SortExpression="Payment_Mode" />
                                                    <asp:BoundField DataField="Subscriber_Group" HeaderText="Group" SortExpression="Subscriber_Group" />
                                                    <asp:TemplateField ItemStyle-Width="0px" ItemStyle-Wrap="false">
                                                        <ItemTemplate>
                                                            <asp:LinkButton ID="DeleteLinkButton" runat="server"
                                                                Text="<i class='bi bi-trash'></i>"
                                                                CssClass="btn btn-xs btn-danger"
                                                                CausesValidation="false"
                                                                CommandArgument='<%# (CType(Container, GridViewRow)).RowIndex & "|" & Eval("Headquarter_ID") & "|" & Eval("Store_ID") %>'
                                                                OnClick="DeleteLineItem_Click"
                                                                OnClientClick="return confirm('Confirm to delete?')">      
                                                            </asp:LinkButton>
                                                        </ItemTemplate>
                                                    </asp:TemplateField>
                                                </Columns>
                                            </asp:GridView>
                                        </div>
                                    </td>
                                </tr>
                                <tr id="subscriptionlistboxerrormsg" runat="server" style="vertical-align: top">
                                    <td style="width: 800px; padding: 0px 0px 0px 0px" colspan="4">
                                        <span style="font-size: 12px; color: #dc3545">Please add subscription store</span>
                                    </td>
                                </tr>
                            </table>
                            <!-- hidden field -->
                            <div>
                                <asp:TextBox ID="TB_Selected_Subscription_Row_Index" runat="server" CssClass="form-control" Visible="false"></asp:TextBox>
                                <asp:TextBox ID="TB_Selected_Subscription_ID" runat="server" CssClass="form-control" Visible="false"></asp:TextBox>
                            </div>
                            <table border="0" class="modalPopupTable" style="width: 100%">
                                <tr style="vertical-align: bottom">
                                    <td style="width: 800px; padding: 10px 10px 10px 0px" colspan="4">
                                        <asp:Button ID="btnSaveSubscription" runat="server" CssClass="btn btn-sm btn-default" ValidationGroup="Subscription" />
                                        <asp:Button ID="btnCancelSubscription" runat="server" CssClass="btn btn-sm btn-default" CausesValidation="false" OnClientClick="return Hidepopup()" />
                                    </td>
                                </tr>
                            </table>
                        </asp:Panel>
                        <asp:LinkButton ID="lnkFakeSubscription" runat="server"></asp:LinkButton>
                        <cc1:ModalPopupExtender ID="popupSubscription" runat="server" DropShadow="false"
                            PopupControlID="pnlAddEditSubscription" TargetControlID="lnkFakeSubscription"
                            BackgroundCssClass="modalBackground">
                        </cc1:ModalPopupExtender>


                        <!-- Modal body - Add Subscription Batch -->
                        <asp:Panel ID="pnlAddEditSubscriptionBatch" runat="server" CssClass="modalPopup">
                            <h3 style="padding: 10px 0px 20px 20px"><asp:Label ID="ModalHeaderSubscriptionBatch" runat="server"></asp:Label></h3>
                            <table border="0" class="modalPopupTable" style="width: 100%">
                                <tr style="vertical-align: top">
                                    <td style="width: 400px; padding: 0px 10px 20px 0px" colspan="2">
                                        <label>Headquarter</label>
                                        <asp:DropDownList ID="DDL_Subscription_Batch_Headquarter" runat="server" CssClass="form-control" AppendDataBoundItems="true">
                                            <asp:ListItem Value="-1" Text="Please select"></asp:ListItem>
                                        </asp:DropDownList>
                                        <asp:CompareValidator ID="CompareValidator_DDL_Subscription_Batch_Headquarter" runat="server" ValidationGroup="SubscriptionBatch"
                                            ErrorMessage="Please select Headquarter" CssClass="invalid-feedback"
                                            Display="Dynamic"
                                            SetFocusOnError="True"
                                            Operator="NotEqual" ValueToCompare="-1"
                                            ControlToValidate="DDL_Subscription_Batch_Headquarter" Type="String">
                                        </asp:CompareValidator>
                                    </td>
                                    <td style="width: 400px; padding: 0px 10px 20px 0px" colspan="2">
                                        <label>Store</label>
                                        <asp:TextBox ID="TB_Subscription_Batch_Store" runat="server" CssClass="form-control" Text="All active Billed stores" Enabled="false"></asp:TextBox>
                                    </td>
                                </tr>
                                <tr style="vertical-align: top">
                                    <td style="width: 400px; padding: 0px 10px 20px 0px" colspan="2">
                                        <label>Duration</label>
                                        <asp:DropDownList ID="DDL_Subscription_Batch_Duration" runat="server" CssClass="form-control" AppendDataBoundItems="true">
                                            <asp:ListItem Value="-1" Text="Please select"></asp:ListItem>
                                        </asp:DropDownList>
                                        <asp:CompareValidator ID="CompareValidator_DDL_Subscription_Batch_Duration" runat="server" ValidationGroup="SubscriptionBatch"
                                            ErrorMessage="Please select duration" CssClass="invalid-feedback"
                                            Display="Dynamic"
                                            SetFocusOnError="True"
                                            Operator="NotEqual" ValueToCompare="-1"
                                            ControlToValidate="DDL_Subscription_Batch_Duration" Type="String">
                                        </asp:CompareValidator>
                                    </td>
                                    <td style="width: 200px; padding: 0px 10px 20px 0px">
                                        <label>Currency</label>
                                        <asp:DropDownList ID="DDL_Subscription_Batch_Currency" runat="server" CssClass="form-control" AppendDataBoundItems="true">
                                            <asp:ListItem Value="-1" Text="Please select"></asp:ListItem>
                                        </asp:DropDownList>
                                        <asp:CompareValidator ID="CompareValidator_DDL_Subscription_Batch_Currency" runat="server" ValidationGroup="SubscriptionBatch"
                                            ErrorMessage="Please select currency" CssClass="invalid-feedback"
                                            Display="Dynamic"
                                            SetFocusOnError="True"
                                            Operator="NotEqual" ValueToCompare="-1"
                                            ControlToValidate="DDL_Subscription_Batch_Currency" Type="String">
                                        </asp:CompareValidator>
                                    </td>
                                    <td style="width: 200px; padding: 0px 10px 20px 0px">
                                        <label>Fee</label>
                                        <asp:TextBox ID="TB_Subscription_Batch_Fee" runat="server" CssClass="form-control" placeholder="0.00"></asp:TextBox>
                                        <asp:RequiredFieldValidator ID="RequiredField_TB_Subscriptio_Batch_Fee" runat="server" ValidationGroup="SubscriptionBatch"
                                            ErrorMessage="This field cannot be blank" CssClass="invalid-feedback"
                                            Display="Dynamic"
                                            SetFocusOnError="True"
                                            ControlToValidate="TB_Subscription_Batch_Fee">
                                        </asp:RequiredFieldValidator>
                                        <asp:RegularExpressionValidator ID="RegEx_TB_Subscription_Batch_Fee" runat="server" ValidationGroup="SubscriptionBatch"
                                            ErrorMessage="Fee must be numeric value" CssClass="invalid-feedback"
                                            Display="Dynamic"
                                            SetFocusOnError="True"
                                            ControlToValidate="TB_Subscription_Batch_Fee"
                                            ValidationExpression="^[+-]?(?:\d+\.?\d*|\d*\.?\d+)[\r\n]*$">
                                        </asp:RegularExpressionValidator>
                                    </td>
                                </tr>
                                <tr style="vertical-align: top">
                                    <td style="width: 200px; padding: 0px 10px 20px 0px">
                                        <label>Payment Method</label>
                                        <asp:DropDownList ID="DDL_Subscription_Batch_Payment_Method" runat="server" CssClass="form-control"></asp:DropDownList>
                                    </td>
                                    <td style="width: 200px; padding: 0px 10px 20px 0px">
                                        <label>Payment Mode</label>
                                        <asp:DropDownList ID="DDL_Subscription_Batch_Payment_Mode" runat="server" CssClass="form-control"></asp:DropDownList>
                                    </td>
                                    <td style="width: 400px; padding: 0px 10px 20px 0px" colspan="2">
                                        <label>Subscriber Group</label>
                                        <asp:DropDownList ID="DDL_Subscription_Batch_Subscriber_Group" runat="server" CssClass="form-control" AppendDataBoundItems="true">
                                            <asp:ListItem Value="-1" Text="Please select"></asp:ListItem>
                                        </asp:DropDownList>
                                        <asp:CompareValidator ID="CompareValidator_DDL_Subscription_Batch_Subscriber_Group" runat="server" ValidationGroup="SubscriptionBatch"
                                            ErrorMessage="Please select subscriber group" CssClass="invalid-feedback"
                                            Display="Dynamic"
                                            SetFocusOnError="True"
                                            Operator="NotEqual" ValueToCompare="-1"
                                            ControlToValidate="DDL_Subscription_Batch_Subscriber_Group" Type="String">
                                        </asp:CompareValidator>
                                    </td>
                                </tr>
                            </table>
                            <table border="0" class="modalPopupTable" style="width: 100%">
                                <tr style="vertical-align: bottom">
                                    <td style="width: 800px; padding: 10px 10px 10px 0px" colspan="4">
                                        <asp:Button ID="btnSaveSubscriptionBatch" runat="server" CssClass="btn btn-sm btn-default" ValidationGroup="SubscriptionBatch" />
                                        <asp:Button ID="btnCancelSubscriptionBatch" runat="server" CssClass="btn btn-sm btn-default" CausesValidation="false" OnClientClick="return Hidepopup()" />
                                    </td>
                                </tr>
                            </table>
                        </asp:Panel>
                        <asp:LinkButton ID="lnkFakeSubscriptionBatch" runat="server"></asp:LinkButton>
                        <cc1:ModalPopupExtender ID="popupSubscriptionBatch" runat="server" DropShadow="false"
                            PopupControlID="pnlAddEditSubscriptionBatch" TargetControlID="lnkFakeSubscriptionBatch"
                            BackgroundCssClass="modalBackground">
                        </cc1:ModalPopupExtender>


                        <!-- Modal body - Adjust Subscription Period -->
                        <asp:Panel ID="pnlAdjustSubscriptionPeriod" runat="server" CssClass="modalPopup">
                            <h3 style="padding: 10px 0px 20px 20px"><asp:Label ID="ModalHeaderAdjustSubscriptionPeriod" runat="server"></asp:Label></h3>
                            <table border="0" class="modalPopupTable" style="width: 100%">
                                <tr style="vertical-align: top">
                                    <td style="width: 400px; padding: 0px 10px 20px 0px" colspan="2">
                                        <label>New Start Date</label>
                                        <asp:TextBox ID="TB_New_Start_Date" runat="server" CssClass="form-control" TextMode="Date"></asp:TextBox>
                                        <asp:RequiredFieldValidator ID="Required_TB_New_Start_Date" runat="server" ValidationGroup="AdjustSubscriptionPeriod"
                                            ErrorMessage="Field cannot be blank" CssClass="invalid-feedback"
                                            Display="Dynamic"
                                            SetFocusOnError="True"
                                            ControlToValidate="TB_New_Start_Date">
                                        </asp:RequiredFieldValidator>
                                    </td>
                                    <td style="width: 400px; padding: 0px 10px 20px 0px" colspan="2">
                                        <label>Duration</label>
                                        <asp:DropDownList ID="DDL_Adjust_Subscription_Duration" runat="server" CssClass="form-control" AppendDataBoundItems="true">
                                            <asp:ListItem Value="-1" Text="Please select"></asp:ListItem>
                                        </asp:DropDownList>
                                        <asp:CompareValidator ID="CompareValidator_DDL_Adjust_Subscription_Duration" runat="server" ValidationGroup="AdjustSubscriptionPeriod"
                                            ErrorMessage="Please select duration" CssClass="invalid-feedback"
                                            Display="Dynamic"
                                            SetFocusOnError="True"
                                            Operator="NotEqual" ValueToCompare="-1"
                                            ControlToValidate="DDL_Adjust_Subscription_Duration" Type="String">
                                        </asp:CompareValidator>
                                    </td>
                                </tr>
                            </table>
                            <table border="0" class="modalPopupTable" style="width: 100%">
                                <tr style="vertical-align: bottom">
                                    <td style="width: 800px; padding: 10px 10px 10px 0px" colspan="4">
                                        <asp:Button ID="btnSaveAdjustSubscriptionPeriod" runat="server" CssClass="btn btn-sm btn-default" ValidationGroup="AdjustSubscriptionPeriod" />
                                        <asp:Button ID="btnCancelAdjustSubscriptionPeriod" runat="server" CssClass="btn btn-sm btn-default" CausesValidation="false" OnClientClick="return Hidepopup()" />
                                    </td>
                                </tr>
                            </table>
                            <!-- hidden field -->
                            <div>
                                <asp:TextBox ID="TB_Selected_Adjust_Subscription_Row_Index" runat="server" CssClass="form-control" Visible="false"></asp:TextBox>
                                <asp:TextBox ID="TB_Selected_Adjust_Subscription_ID" runat="server" CssClass="form-control" Visible="false"></asp:TextBox>
                            </div>
                        </asp:Panel>
                        <asp:LinkButton ID="lnkFakeAdjustSubscriptionPeriod" runat="server"></asp:LinkButton>
                        <cc1:ModalPopupExtender ID="popupAdjustSubscriptionPeriod" runat="server" DropShadow="false"
                            PopupControlID="pnlAdjustSubscriptionPeriod" TargetControlID="lnkFakeAdjustSubscriptionPeriod"
                            BackgroundCssClass="modalBackground">
                        </cc1:ModalPopupExtender>

                    </ContentTemplate>

                    <Triggers>
                        <asp:AsyncPostBackTrigger ControlID="GridView4" />
                        <asp:AsyncPostBackTrigger ControlID="btnSaveSubscription" />
                        <asp:AsyncPostBackTrigger ControlID="btnSaveSubscriptionBatch" />
                        <asp:AsyncPostBackTrigger ControlID="btnSaveAdjustSubscriptionPeriod" />
                    </Triggers>

                </asp:UpdatePanel>
            </div>
        </div>
    </div>

    <!-- Account Notes -->
    <div id="DMCAccountNotes" runat="server" style="margin-top: 40px">
        <asp:UpdatePanel ID="UpdatePanel5" runat="server">
            <ContentTemplate>
                <div class="card border-0 mb-3">
                    <div style="padding-left: 15px; padding-right: 12px">
                        <h4 class="card-title">Notes</h4>
                    </div>
                    <div class="card-body">
                        <asp:GridView ID="GridView5" runat="server"></asp:GridView>
                        <asp:Button ID="btnAddNotes" Text="Add New" runat="server" CausesValidation="false" CssClass="btn btn-sm btn-info" />
                    </div>
                </div>

                <!-- Modal body -->
                <asp:Panel ID="pnlAddEditNotes" runat="server" CssClass="modalPopup">
                    <h3 style="padding: 10px 0px 20px 20px">
                        <asp:Label ID="ModalHeaderNotes" runat="server"></asp:Label></h3>
                    <table border="0" class="modalPopupTable" style="width: 100%">
                        <tr style="vertical-align: top">
                            <td style="width: 800px; padding: 0px 10px 20px 0px">
                                <label>Notes</label>
                                <asp:TextBox ID="TB_Notes" runat="server" TextMode="MultiLine" placeholder="Type note here.."></asp:TextBox>
                                <asp:RequiredFieldValidator ID="RequiredField_TB_Notes" runat="server" ValidationGroup="Notes"
                                    ErrorMessage="This field cannot be blank" CssClass="invalid-feedback"
                                    Display="Dynamic"
                                    SetFocusOnError="True"
                                    ControlToValidate="TB_Notes">
                                </asp:RequiredFieldValidator>
                            </td>
                        </tr>
                        <tr style="vertical-align: bottom">
                            <td style="width: 800px; padding: 10px 10px 10px 0px" colspan="4">
                                <asp:Button ID="btnSaveNotes" runat="server" CssClass="btn btn-sm btn-default" ValidationGroup="Notes" CommandArgument="Create" />
                                <asp:Button ID="btnCancelNotes" runat="server" CssClass="btn btn-sm btn-default" CausesValidation="false" OnClientClick="return Hidepopup()" />
                            </td>
                        </tr>
                    </table>
                </asp:Panel>
                <asp:LinkButton ID="lnkFakeNotes" runat="server"></asp:LinkButton>
                <cc1:ModalPopupExtender ID="popupNotes" runat="server" DropShadow="false"
                    PopupControlID="pnlAddEditNotes" TargetControlID="lnkFakeNotes"
                    BackgroundCssClass="modalBackground">
                </cc1:ModalPopupExtender>
            </ContentTemplate>

            <Triggers>
                <asp:AsyncPostBackTrigger ControlID="GridView5" />
                <asp:AsyncPostBackTrigger ControlID="btnSaveNotes" />
            </Triggers>

        </asp:UpdatePanel>
    </div>


    <hr style="margin-top: 100px"/>

    <!-- Control Button -->
    <asp:LinkButton ID="BT_Close" Text="Close" CommandName="Close" runat="server" CssClass="btn btn-default" CausesValidation="false"></asp:LinkButton>


    <!-- javascript for Nested GridView effect -->
    <script src="/Scripts/jquery-1.3.2.min.js"></script> <!-- need to use jquery-1.3.2.min.js -->
    <script type="text/javascript">
        $("[src*=plus]").live("click", function () {
            $(this).closest("tr").after("<tr><td></td><td colspan = '999'>" + $(this).next().html() + "</td></tr>")
            $(this).attr("src", "/Images/minus-square-solid-24.png");
        });
        $("[src*=minus]").live("click", function () {
            $(this).attr("src", "/Images/plus-square-solid-24.png");
            $(this).closest("tr").next().remove();
        });
    </script>

    <!-- Keep where last tab selected -->
    <script>
        $(document).ready(function () {
            $('a[data-toggle="tab"]').on('show.bs.tab', function (e) {
                localStorage.setItem('activeTab', $(e.target).attr('href'));
                localStorage.setItem('activeLink', $(e.target).attr('id'));
            });

            var activeTab = localStorage.getItem('activeTab');
            //var activeLink = localStorage.getItem('activeLink');

            if (activeTab) {
                $('#myTab a[href="' + activeTab + '"]').tab('show');  // Show current tab
                //var element = document.getElementById("tab_" + activeLink);
                //element.classList.add("active");
            }
        });
    </script>

</asp:Content>

