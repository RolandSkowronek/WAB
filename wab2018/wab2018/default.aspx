<%@ Page Title="" Language="C#" MasterPageFile="~/Master.Master" AutoEventWireup="true" CodeBehind="default.aspx.cs" Inherits="wab2018._default" %>
<%@ Register assembly="DevExpress.Web.v17.1, Version=17.1.17.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.Web" tagprefix="dx" %>

<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="server">
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
      <div id ="mainWindow" style="background-color:white; min-height:800px;" >

   <asp:Panel ID="Panel1" runat="server"  Width="100%">
    
        <h2> Obciążenie biegłych</h2>
   
        <table style="width:100%;">
            <tr>
                <td width="20%" valign="top">
                    Wyświetlanie po&nbsp;
                    <asp:CheckBox ID="CheckBox1" runat="server" AutoPostBack="True" Checked="True" oncheckedchanged="CheckBox1_CheckedChanged" Text="Dziedzina :" />
                    &nbsp;&nbsp;&nbsp;
                    <asp:DropDownList ID="DropDownList2" runat="server" AutoPostBack="True" DataSourceID="grupy" DataTextField="Nazwa" DataValueField="ident" onselectedindexchanged="DropDownList2_SelectedIndexChanged">
                    </asp:DropDownList>
                    &nbsp;&nbsp;
                    <asp:CheckBox ID="CheckBox2" runat="server" AutoPostBack="True" oncheckedchanged="CheckBox2_CheckedChanged" Text="Specjalizacja" />
                    &nbsp;&nbsp;&nbsp;
                    <asp:DropDownList ID="DropDownList1" runat="server" AutoPostBack="True" DataSourceID="specjalizacje" DataTextField="nazwa" DataValueField="id_" Enabled="False" onselectedindexchanged="DropDownList1_SelectedIndexChanged">
                    </asp:DropDownList>
                </td>
            </tr>
        </table>
    
    <asp:SqlDataSource ID="widok_glowny" runat="server" 
    ConnectionString="<%$ ConnectionStrings:wap %>" 
    
            SelectCommand="SELECT DISTINCT nazwisko + ' ' + imie AS biegly, stan, stan_inne, zwrocone, zwrocone_inne, opinie, przeterminowane, grzywna, data_poczatkowa, data_koncowa, nazwisko FROM obciazenia ORDER BY stan,stan_inne, zwrocone">
    </asp:SqlDataSource>
        <asp:SqlDataSource ID="grupy" runat="server" ConnectionString="<%$ ConnectionStrings:wap %>" SelectCommand="SELECT DISTINCT [ident], [Nazwa] FROM [glo_grupy_specjalizacji] ORDER BY [Nazwa]"></asp:SqlDataSource>
        <asp:SqlDataSource ID="specjalizacje" runat="server" 
            ConnectionString="<%$ ConnectionStrings:wap %>" 
            
            SelectCommand="SELECT DISTINCT id_, nazwa, grupa FROM glo_specjalizacje WHERE (grupa = @grupa) ORDER BY nazwa">
            <SelectParameters>
                <asp:SessionParameter Name="grupa" SessionField="grupa" />
            </SelectParameters>
        </asp:SqlDataSource>

        <asp:Panel ID="Panel2" runat="server" Height="100%" 
            Width="100%">
            <asp:GridView ID="GridView1" runat="server" 
                AutoGenerateColumns="False" CellPadding="3" DataSourceID="widok_glowny" 
                ForeColor="#333333" GridLines="None" Width="95%" 
                onsorted="GridView1_Sorted" OnRowCreated="GridView1_RowCreated1" ShowHeader="False">
                <AlternatingRowStyle BackColor="White" />
                <Columns>
                    <asp:BoundField DataField="biegly" 
                        HeaderText="biegly" 
                        SortExpression="biegly" ReadOnly="True">
                    <ItemStyle HorizontalAlign="Left" />
                    </asp:BoundField>
                    <asp:BoundField DataField="stan" HeaderText="stan" SortExpression="stan">
                    </asp:BoundField>
                    <asp:BoundField DataField="stan_inne" HeaderText="stan_inne" SortExpression="stan_inne">
                    </asp:BoundField>
                    <asp:BoundField DataField="zwrocone" HeaderText="zwrocone" SortExpression="zwrocone">
                    </asp:BoundField>
                    <asp:BoundField DataField="zwrocone_inne" HeaderText="zwrocone_inne" SortExpression="zwrocone_inne" />
                    <asp:BoundField DataField="opinie" HeaderText="opinie" SortExpression="opinie" />
                    <asp:BoundField DataField="przeterminowane" HeaderText="przeterminowane" SortExpression="przeterminowane" />
                    <asp:BoundField DataField="grzywna" HeaderText="grzywna" SortExpression="grzywna" />
                    <asp:TemplateField HeaderText="daty" SortExpression="data_poczatkowa">
                        <EditItemTemplate>
                            <asp:TextBox ID="TextBox1" runat="server" Text='<%# Bind("data_poczatkowa") %>'></asp:TextBox>
                        </EditItemTemplate>
                        <ItemTemplate>
                            <asp:Label ID="Label1" runat="server" Text='<%# Bind("data_poczatkowa", "{0:d}") %>'></asp:Label>
                            &nbsp;-
                            <asp:Label ID="Label3" runat="server" Text='<%# Eval("data_koncowa", "{0:d}") %>'></asp:Label>
                        </ItemTemplate>
                    </asp:TemplateField>
                </Columns>
                <EditRowStyle BackColor="#2461BF" />
                <FooterStyle BackColor="#507CD1" Font-Bold="True" ForeColor="White" />
                <HeaderStyle BackColor="#507CD1" Font-Bold="True" ForeColor="White" />
                <PagerStyle BackColor="#2461BF" ForeColor="White" HorizontalAlign="Center" />
                <RowStyle BackColor="#BECEEF" HorizontalAlign="Center" />
                <SelectedRowStyle BackColor="#D1DDF1" Font-Bold="True" ForeColor="#333333" />
                <SortedAscendingCellStyle BackColor="#BECEEF" />
                <SortedAscendingHeaderStyle BackColor="#6D95E1" />
                <SortedDescendingCellStyle BackColor="#BECEEF" />
                <SortedDescendingHeaderStyle BackColor="#4870BE" />
            </asp:GridView>
        </asp:Panel>
</asp:Panel>
              </div>

</asp:Content>
