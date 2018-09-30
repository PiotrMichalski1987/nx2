<%@ Page Title="" Language="C#" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeBehind="UpdateDB.aspx.cs" Inherits="WrkWebApp.Presentation_Layer.UpdateDB" %>
<asp:Content ID="Content1" ContentPlaceHolderID="MainContent" runat="server">

    <p style="height: 338px">



        <asp:TextBox ID="txtOvertimeRate" runat="server" style="z-index: 1; position: absolute; top: 240px; left: 204px; width: 62px"></asp:TextBox>
        <asp:TextBox ID="txtStandardRate" runat="server" style="z-index: 1; position: absolute; top: 267px; left: 204px; width: 60px; height: 17px"></asp:TextBox>
        <asp:Label ID="lblStandardRate" runat="server" style="z-index: 1; position: absolute; top: 271px; left: 98px; width: 113px; height: 16px; bottom: 80px" Text="Standard Rate"></asp:Label>
        <asp:Button ID="btnConfirm" runat="server" OnClick="btnConfirm_Click" style="z-index: 1; position: absolute; top: 296px; left: 95px; height: 23px" Text="Confirm" />
        <asp:Label ID="lblSetRatesFor" runat="server" style="z-index: 1; position: absolute; top: 146px; left: 90px" Text="Set Rate(s) for:"></asp:Label>
        <asp:DropDownList ID="ddlStRateFor" runat="server" OnSelectedIndexChanged="ddlStRateFor_SelectedIndexChanged" style="z-index: 1; position: absolute; top: 143px; left: 202px; height: 16px">
            <asp:ListItem>All</asp:ListItem>
            <asp:ListItem>Individual</asp:ListItem>
        </asp:DropDownList>
        <asp:Label ID="lblOverTime" runat="server" style="z-index: 1; position: absolute; top: 236px; left: 98px" Text="Overtime rate:"></asp:Label>
        <asp:DropDownList ID="ddlSelectDriver" runat="server" style="z-index: 1; position: absolute; top: 179px; left: 201px">
        </asp:DropDownList>
        <asp:Label ID="lblSelectDriver" runat="server" style="z-index: 1; position: absolute; top: 182px; left: 94px" Text="Select Driver"></asp:Label>
        <asp:Label ID="lblSelectType" runat="server" style="z-index: 1; position: absolute; top: 211px; left: 112px" Text="Select Type"></asp:Label>
        <asp:DropDownList ID="ddlSelectType" runat="server" style="z-index: 1; position: absolute; top: 208px; left: 202px">
        </asp:DropDownList>
        <asp:Label ID="lblInfo" runat="server" style="z-index: 1; position: absolute; top: 302px; left: 202px"></asp:Label>



    </p>
</asp:Content>
