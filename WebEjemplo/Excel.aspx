<%@ Page Title="" Language="C#" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeBehind="Excel.aspx.cs" Inherits="WebEjemplo.Excel" %>
<asp:Content ID="Content1" ContentPlaceHolderID="MainContent" runat="server">
    
    
    <h3>Importaciones en Excel</h3>
   
   
    <div aria-autocomplete="both">
    
        
    
    </div>
    <div class="input-group-append">
        <asp:FileUpload ID="FileUpload1" runat="server" Width="600px" />
        <asp:Button ID="Button1" runat="server" onclick="Button1_Click" Text="Subir" />
    
        
    </div>
    <br />
    <div>
        <asp:Button ID="Button2" runat="server" Text="Cargar Datos-Tabla" 
            onclick="Button2_Click" />
    </div>

    <asp:Label ID="lblOculto" runat="server" Text="" Visible="false"></asp:Label>
    <br />
    <asp:Label ID="Label1" runat="server" Text=""></asp:Label>
    <div class="table-responsive">
        <asp:GridView ID="GridView1" runat="server" Height="251px" Width="100%" CssClass="table table-border table-hover table-sm miEstilo">
    </asp:GridView>
        
    </div>
    <br />
    <div>
        <asp:Button ID="Button3" runat="server" Text="Confirmar Excel" Width="100%" onclick="Button3_Click" />
    </div>
    <br />
    <div class="table-responsive">
        <asp:GridView ID="GridView2" runat="server" Height="251px" CellPadding="4" Width="100%" CssClass="table table-border table-hover table-sm miEstilo" ForeColor="#333333" GridLines="None">
            <AlternatingRowStyle BackColor="White" />
            <EditRowStyle BackColor="#2461BF" />
        <FooterStyle BackColor="#507CD1" ForeColor="White" Font-Bold="True" />
        <HeaderStyle BackColor="#507CD1" Font-Bold="True" ForeColor="White" />
        <PagerStyle ForeColor="White" HorizontalAlign="Center" BackColor="#2461BF" />
        <RowStyle BackColor="#EFF3FB" />
        <SelectedRowStyle BackColor="#D1DDF1" Font-Bold="True" ForeColor="#333333" />
        <SortedAscendingCellStyle BackColor="#F5F7FB" />
        <SortedAscendingHeaderStyle BackColor="#6D95E1" />
        <SortedDescendingCellStyle BackColor="#E9EBEF" />
        <SortedDescendingHeaderStyle BackColor="#4870BE" />
    </asp:GridView>
    </div>


    
</asp:Content>
