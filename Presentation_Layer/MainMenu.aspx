<%@ Page Title="" Language="C#" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeBehind="MainMenu.aspx.cs" Inherits="WrkWebApp.Presentation_Layer.MainMenu" %>
<asp:Content ID="Content1" ContentPlaceHolderID="MainContent" runat="server">
    
    
    
    <p style="margin-top: 0; height: 361px;">
        
        <asp:Button ID="btnUploadReports" runat="server" style="z-index: 1;position: absolute;  top: 285px; left: 501px; height: 25px; width: 143px" Text="Upload Reports" OnClick="btnUploadReports_Click"  />
        <asp:Button ID="btnUpdateDb" runat="server" style="z-index: 1; position: absolute; top: 197px; left: 457px; width: 140px; height: 25px" Text="Update Data" OnClick="btnUpdateDb_Click" />
        <asp:Button ID="btnAnalyse2" runat="server" style="z-index: 1; position: absolute; top: 334px; left: 493px; width: 138px; height: 24px" Text="Analyse2" OnClick="btnAnalyse2_Click" />
        <asp:Button ID="btnProduceReports2" runat="server"  style="z-index: 1; position: absolute; top: 151px; left: 445px; width: 139px; height: 26px" Text="Produce Repors2" OnClick="btnProduceReports2_Click" />
        <asp:Button ID="btnUpdateDriversTable" runat="server" OnClick="btnUpdateDriversTable_Click" style="z-index: 1; position: absolute; top: 245px; left: 440px; width: 140px; height: 26px" Text="Update Drivers Table" />
        <asp:ImageButton ID="ImageButton2" runat="server" Height="50px" ImageUrl="~/Content/buttondriversTable.png" style="z-index: 1; position: absolute; top: 233px; left: 15px" Width="150px" />
        <asp:ImageButton ID="ImageButton3" runat="server" Height="50px" ImageUrl="~/Content/buttonProduceReports.png" style="z-index: 1; position: absolute; top: 170px; left: 16px" Width="150px" />
        <asp:ImageButton ID="ImageButton4" runat="server" Height="50px" ImageUrl="~/Content/buttonUpdateData.png" style="z-index: 1; position: absolute; top: 296px; left: 16px" Width="150px" />
        <asp:ImageButton ID="ImageButton5" runat="server" Height="50px" ImageUrl="~/Content/buttonUploadRep.png" style="z-index: 1; position: absolute; top: 107px; left: 17px" Width="150px" />
        <asp:ImageButton ID="ImageButton1" runat="server" BorderColor="#FF3300" ImageUrl="~/Content/buttonAnalyse.png" style="z-index: 1; position: absolute; top: 356px; left: 15px; width: 150px; height: 50px; right: 863px" />
    </p>

       
</asp:Content>
