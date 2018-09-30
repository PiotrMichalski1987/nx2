<%@ Page Title="" Language="C#" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeBehind="ProduceReportsScreen.aspx.cs" Inherits="WrkWebApp.Presentation_Layer.ProduceReportsScreen" %>
<asp:Content ID="Content1" ContentPlaceHolderID="MainContent" runat="server">
    <p style="height: 511px">

        <asp:DropDownList ID="ddlSelectReport" runat="server" style="z-index: 1; position: absolute; top: 261px; left: 37px; width: 203px; height: 20px">
        </asp:DropDownList>
        <asp:Label ID="lblSelectReport" runat="server" style="z-index: 1; position: absolute; top: 229px; left: 33px" Text="Select Report Type"></asp:Label>
        <asp:Label ID="lblSelectRange" runat="server" style="z-index: 1; position: absolute; top: 313px; left: 40px" Text="Select Range"></asp:Label>
        <asp:Calendar ID="cldFrom" runat="server" style="z-index: 1; width: 259px; height: 188px; position: absolute; top: 338px; left: 39px"></asp:Calendar>
        <asp:Calendar ID="cldTo" runat="server" style="z-index: 1; width: 259px; height: 188px; position: absolute; top: 339px; left: 307px"></asp:Calendar>
        <asp:Button ID="btnConfrm" runat="server" style="z-index: 1; position: absolute; top: 535px; left: 41px; width: 100px" Text="Confirm" OnClick="btnConfrm_Click" />
        <asp:Label ID="lblInfo" runat="server" style="z-index: 1; position: absolute; top: 538px; left: 149px"></asp:Label>

    </p>
</asp:Content>
