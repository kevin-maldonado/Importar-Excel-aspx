<%@ Page Title="" Language="C#" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeBehind="ComboBox.aspx.cs" Inherits="WebEjemplo.ComboBox" %>
<asp:Content ID="Content1" ContentPlaceHolderID="MainContent" runat="server">
   
    <h3>Checkbox</h3>
    <div aria-autocomplete="both">
    <asp:ListBox ID="lstServicios" runat="server" SelectionMode="Multiple">
        <asp:ListItem Text="Servicio 1" Value="1" />
        <asp:ListItem Text="Servicio 2" Value="2" />
        <asp:ListItem Text="Servicio 3" Value="3" />
        <asp:ListItem Text="Servicio 4" Value="4" />
        <asp:ListItem Text="Servicio 5" Value="5" />
    </asp:ListBox>
    <asp:Button Text="Submit" runat="server" OnClick="Submit" />
     </div>               
</asp:Content>
