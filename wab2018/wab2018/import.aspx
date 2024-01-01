<%@ Page Title="" Language="C#" MasterPageFile="~/Master.Master" AutoEventWireup="true" CodeBehind="import.aspx.cs" Inherits="wab2018.import" %>
<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="server">
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
       <div id ="mainWindow" class="newPage">  Biegli 

    <asp:Button ID="Button1" runat="server" onclick="Button1_Click" 
    Text="Importuj dane!" CssClass="button_" Width="136px" />
    &nbsp;<asp:Button ID="Button2" runat="server" onclick="Button2_Click" 
    Text="Quick" CssClass="button_" Width="136px" Visible="False" />
    <asp:Panel ID="Panel1" runat="server" Height="400px" ScrollBars="Vertical">
        <asp:GridView ID="GridView1" runat="server" 
    AutoGenerateColumns="False" CellPadding="2" 
            ForeColor="#333333" GridLines="None" Width="100%" AllowPaging="True">
            <AlternatingRowStyle BackColor="White" />
            <EditRowStyle BackColor="#2461BF" />
            <FooterStyle BackColor="#507CD1" Font-Bold="True" ForeColor="White" />
            <HeaderStyle BackColor="#507CD1" Font-Bold="True" ForeColor="White" />
            <PagerStyle BackColor="#2461BF" ForeColor="White" HorizontalAlign="Center" />
            <RowStyle BackColor="#EFF3FB" />
            <SelectedRowStyle BackColor="#D1DDF1" Font-Bold="True" ForeColor="#333333" />
            <SortedAscendingCellStyle BackColor="#F5F7FB" />
            <SortedAscendingHeaderStyle BackColor="#6D95E1" />
            <SortedDescendingCellStyle BackColor="#E9EBEF" />
            <SortedDescendingHeaderStyle BackColor="#4870BE" />
        </asp:GridView>
        <br />
        <asp:SqlDataSource ID="SqlDataSource2" runat="server" 
    ConnectionString="<%$ ConnectionStrings:wap %>" 
    SelectCommand="SELECT * FROM tbl_main ORDER BY [ident]" ViewStateMode="Disabled"></asp:SqlDataSource>
    </asp:Panel>
   
    <asp:Label ID="Label1" runat="server"></asp:Label>
    <br />
    <asp:Label ID="Label2" runat="server"></asp:Label>
<br />
<asp:Label ID="Label3" runat="server"></asp:Label>
    <br />
    <asp:TextBox ID="TextBox1" runat="server" BackColor="#CCCCCC" Height="138px" 
        TextMode="MultiLine" Width="100%"></asp:TextBox>
    <br />
    </div>

</asp:Content>
