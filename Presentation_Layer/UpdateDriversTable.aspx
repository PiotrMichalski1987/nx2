<%@ Page Title="" Language="C#" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeBehind="UpdateDriversTable.aspx.cs" Inherits="WrkWebApp.Presentation_Layer.UpdateDriversTable" %>
<asp:Content ID="Content1" ContentPlaceHolderID="MainContent" runat="server">
    <p style="height: 338px">
    <asp:Button ID="btnUpload" runat="server" OnClick="btnUpload_Click" style="z-index: 1; position: absolute; top: 202px; left: 35px; width: 124px; height: 25px" Text="Upload" />
    <asp:FileUpload ID="flUpl" runat="server" style="z-index: 1; width: 222px; height: 22px; position: absolute; top: 173px; left: 37px" />
    <asp:Label ID="lblInfo" runat="server" style="z-index: 1; position: absolute; top: 202px; left: 167px"></asp:Label>
    <asp:GridView ID="GridView1" runat="server" style="z-index: 1; width: 193px; height: 139px; position: absolute; top: 180px; left: 772px">
    </asp:GridView>
    <%--</p>--%>
        <asp:Button ID="btnUploadClient" runat="server" OnClick="btnUploadClient_Click" style="z-index: 1; position: absolute; top: 319px; left: 42px; width: 143px; height: 25px" Text="Upload Client Names" />
        <asp:FileUpload ID="flUplClient" runat="server" style="z-index: 1; width: 222px; height: 22px; position: absolute; top: 289px; left: 38px" />
        <asp:Label ID="lblInfoClient" runat="server" style="z-index: 1; position: absolute; top: 319px; left: 214px"></asp:Label>

</asp:Content>
