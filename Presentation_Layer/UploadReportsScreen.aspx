<%@ Page Title="" Language="C#" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeBehind="UploadReportsScreen.aspx.cs" Inherits="WrkWebApp.Presentation_Layer.UploadReportsScreen" %>
<asp:Content ID="Content1" ContentPlaceHolderID="MainContent" runat="server">
    <p <%--<%--style="height: 338px"--%>--%>>
         <asp:FileUpload ID="FlUplUpload" runat="server" style="z-index: 1; width: 222px; height: 22px; position: absolute; top: 215px; left: 24px" />
         <asp:Label ID="lblName" runat="server" style="z-index: 1; position: absolute; top: 112px; left: 27px; width: 171px; height: 27px" Text="Upload Excel Reports"></asp:Label>
         <asp:Label ID="lblSelect" runat="server" style="z-index: 1; position: absolute; top: 189px; left: 27px; width: 123px" Text="Select File 1"></asp:Label>
         
         <asp:Label ID="lblInfo" runat="server" style="z-index: 1; position: absolute; top: 249px; left: 129px"></asp:Label>
         <asp:Button ID="btnCnfr" runat="server" OnClick="btnCnfr_Click" style="z-index: 1; position: absolute; top: 284px; left: 22px; height: 18px" Text="Cnfr" />
           <asp:GridView ID="GridView1" runat="server" style="z-index: 1; width: 425px; height: 181px; position: absolute; top: 349px; left: 265px" OnSelectedIndexChanged="GridView1_SelectedIndexChanged"> 
         </asp:GridView> 

       
        <asp:Calendar ID="cldSelectDate" runat="server" style="z-index: 1; width: 252px; height: 196px; position: absolute; top: 123px; left: 273px"></asp:Calendar>

       
    </p>
   
</asp:Content>
