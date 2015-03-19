<%@ Page Title="" Language="C#" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeBehind="Usuarios.aspx.cs" Inherits="Rodas.Usuarios" %>
<%@ MasterType  virtualPath="~/Site.master"%> 
<asp:Content ID="Content1" ContentPlaceHolderID="HeadContent" runat="server">
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="FeaturedContent" runat="server">
</asp:Content>
<asp:Content ID="Content3" ContentPlaceHolderID="MainContent" runat="server">

    <asp:Button ID="Button1" runat="server" OnClick="Button1_Click" Text="Nuevo Usuario" />

    <asp:Label ID="Label1" runat="server" Text="Label" Visible="False"></asp:Label>

</asp:Content>
