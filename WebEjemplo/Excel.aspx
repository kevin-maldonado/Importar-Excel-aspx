﻿<%@ Page Title="" Language="C#" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeBehind="Excel.aspx.cs" Inherits="WebEjemplo.Excel" %>
<asp:Content ID="Content1" ContentPlaceHolderID="MainContent" runat="server">
 
    <h3>Importaciones en Excel</h3>
    <div aria-autocomplete="both">

    </div>
    <div class="input-group-append">
        <asp:FileUpload ID="FileUpload1" runat="server" Width="600px" />
        <asp:Button ID="Btn_Subir" runat="server" onclick="Subir_Onclick" Text="Subir" />
    </div>
    <asp:Label ID="Lbl_Oculto" runat="server" Text="" Visible="false"></asp:Label>
    <br />
    <asp:Label ID="Lbl_Mensaje" runat="server" Text=""></asp:Label>
    <br />
    <div class="table-responsive">
        <asp:GridView ID="GridView1" runat="server" Height="251px" Width="100%" CssClass="table table-border table-hover table-sm miEstilo" OnRowDataBound="grdDatos_RowDataBound" >
        </asp:GridView>
    </div>
    <br />
    <div>
        <asp:Button ID="Btn_Guardar" runat="server" Text="Guardar Datos" Width="100%" onclick="Guardar_Onclick" Visible="true" />
    </div>
    <br />
   
   
 


    
</asp:Content>
