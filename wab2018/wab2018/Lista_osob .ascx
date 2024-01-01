<%@ Control Language="C#" AutoEventWireup="true" CodeBehind="Lista_osob .ascx.cs" Inherits="wab.Lista_osob" %>
<link href="Styles/Site.css" rel="stylesheet" type="text/css" />
<table style="width:100%;">
    <tr>
        <td>
            <asp:TextBox ID="TextBox1" runat="server" Width="241px"></asp:TextBox>
            <asp:LinkButton ID="LinkButton1" runat="server" onclick="LinkButton1_Click" 
                CausesValidation="False" CssClass="button_">Wybierz</asp:LinkButton>
            <asp:HiddenField ID="HiddenField2" runat="server" />
        </td>
        <td>
            &nbsp;</td>
    </tr>
</table>

   <asp:Panel ID="Panel3" runat="server" CssClass="Panel_2_2" Visible="False">
        
        <asp:Panel ID="Panel4" runat="server" CssClass="Panel_1_1" style=" height:400px; width:700px">
            
            <div style="background-color: #3399FF; width: 100%; top:0px;">
           
                <table style="width:100%;">
                    <tr>
                        <td>
                            <b>Wybór osoby</b></td>
                        <td>
                            &nbsp;</td>
                        <td style="float: right">
                            <asp:ImageButton ID="ImageButton1" runat="server" CausesValidation="False" 
                                ImageUrl="~/img/zamknij.jpg" onclick="ImageButton1_Click" />
                        </td>
                    </tr>
                </table>
               </div>
         
            <table style="width:100%;">
                <tr>
                    <td>
                        <asp:CheckBox ID="CheckBox1" runat="server" AutoPostBack="True" 
                            oncheckedchanged="CheckBox1_CheckedChanged" />
                    </td>
                    <td>
                        <asp:Label ID="Label1" runat="server" Text="Specjalizacja:"></asp:Label>
                    </td>
                    <td>
                        <asp:DropDownList ID="DropDownList1" runat="server" 
                            DataSourceID="specjalizacje" DataTextField="nazwa" DataValueField="id_" 
                            AutoPostBack="True" onselectedindexchanged="DropDownList1_SelectedIndexChanged">
                        </asp:DropDownList>
                        <asp:SqlDataSource ID="specjalizacje" runat="server" 
                            ConnectionString="<%$ ConnectionStrings:wap %>" 
                            SelectCommand="SELECT DISTINCT [id_], [nazwa] FROM [glo_specjalizacje] ORDER BY [nazwa]">
                        </asp:SqlDataSource>
                    </td>
                </tr>
            </table>
            <asp:Panel ID="Panel5" runat="server" Height="350px" ScrollBars="Vertical" 
                Width="100%">
                  <asp:Panel ID="Panel1" runat="server" ScrollBars="Vertical" Height="400px" 
                        Width="700px">

                         <asp:GridView ID="GridView1" runat="server" 
    AutoGenerateColumns="False" CellPadding="1" DataSourceID="lista_osob_x" 
    ForeColor="#333333" GridLines="Horizontal" DataKeyNames="ident,nazwisko,imie" 
                            onselectedindexchanged="GridView1_SelectedIndexChanged" Width="100%" 
                            ShowHeaderWhenEmpty="True" AllowSorting="True" CellSpacing="1">
                            <AlternatingRowStyle BackColor="White" />
                            <Columns>
                                <asp:CommandField ShowSelectButton="True" SelectText="Wybierz" 
                                    ButtonType="Image" SelectImageUrl="~/img/details.jpg"  >
                                <ItemStyle Width="36px" />
                                </asp:CommandField>
                                <asp:BoundField DataField="imie" HeaderText="Imie" SortExpression="imie" />
                                <asp:BoundField DataField="nazwisko" HeaderText="Nazwisko" 
                                    SortExpression="nazwisko" />
                                <asp:TemplateField HeaderText="Miejscowość" SortExpression="kod_poczt">
                                    <EditItemTemplate>
                                        <asp:TextBox ID="TextBox1" runat="server" Text='<%# Bind("kod_poczt") %>'></asp:TextBox>
                                    </EditItemTemplate>
                                    <ItemTemplate>
                                        <asp:Label ID="Label4" runat="server" Text='<%# Eval("ulica") %>'></asp:Label>
                                        <br />
                                        <asp:Label ID="Label1" runat="server" CssClass="bold" 
                                            Text='<%# Bind("kod_poczt") %>'></asp:Label>
                                        &nbsp;<asp:Label ID="Label2" runat="server" CssClass="bold" 
                                            Text='<%# Eval("miejscowosc") %>'></asp:Label>
                                    </ItemTemplate>
                                </asp:TemplateField>
                                <asp:BoundField DataField="data_poczatkowa" DataFormatString="{0:dd.MM.yyyy}" 
                                    HeaderText="data_poczatkowa" SortExpression="data_poczatkowa" />
                                <asp:BoundField DataField="data_koncowa" DataFormatString="{0:dd.MM.yyyy}" 
                                    HeaderText="data_koncowa" SortExpression="data_koncowa" />
                                <asp:TemplateField>
                                    <ItemTemplate>
                                        <asp:HiddenField ID="HiddenField1" runat="server" 
                                            Value='<%# Eval("ident") %>' />
                                        <asp:GridView ID="GridView6" runat="server" AutoGenerateColumns="False" 
                                            BackColor="White" BorderColor="#E7E7FF" BorderStyle="None" BorderWidth="1px" 
                                            CellPadding="0" DataKeyNames="nazwa" DataSourceID="krotka_lista" 
                                            Font-Bold="True" Font-Size="X-Small" GridLines="Horizontal" ShowHeader="False" 
                                            Width="100%">
                                            <AlternatingRowStyle BackColor="#F7F7F7" />
                                            <Columns>
                                                <asp:BoundField DataField="nazwa" HeaderText="nazwa" ReadOnly="True" 
                                                    SortExpression="nazwa" />
                                            </Columns>
                                            <FooterStyle BackColor="#B5C7DE" ForeColor="#4A3C8C" />
                                            <HeaderStyle BackColor="#4A3C8C" Font-Bold="True" ForeColor="#F7F7F7" />
                                            <PagerStyle BackColor="#E7E7FF" ForeColor="#4A3C8C" HorizontalAlign="Right" />
                                            <RowStyle BackColor="#E7E7FF" ForeColor="#4A3C8C" />
                                            <SelectedRowStyle BackColor="#738A9C" Font-Bold="True" ForeColor="#F7F7F7" />
                                            <SortedAscendingCellStyle BackColor="#F4F4FD" />
                                            <SortedAscendingHeaderStyle BackColor="#5A4C9D" />
                                            <SortedDescendingCellStyle BackColor="#D8D8F0" />
                                            <SortedDescendingHeaderStyle BackColor="#3E3277" />
                                        </asp:GridView>
                                        <asp:SqlDataSource ID="krotka_lista" runat="server" 
                                            ConnectionString="<%$ ConnectionStrings:wap %>" SelectCommand="SELECT     TOP (2)  glo_specjalizacje.nazwa
FROM         glo_specjalizacje INNER JOIN
                      tbl_specjalizacje_osob ON glo_specjalizacje.id_ = tbl_specjalizacje_osob.id_specjalizacji
WHERE     (tbl_specjalizacje_osob.id_osoby = @id_osoby)
ORDER BY tbl_specjalizacje_osob.id_">
                                            <SelectParameters>
                                                <asp:ControlParameter ControlID="HiddenField1" Name="id_osoby" 
                                                    PropertyName="Value" />
                                            </SelectParameters>
                                        </asp:SqlDataSource>
                                    </ItemTemplate>
                                </asp:TemplateField>
                            </Columns>
                            <EditRowStyle BackColor="#2461BF" />
                            <FooterStyle BackColor="#507CD1" Font-Bold="True" 
        ForeColor="White" />
                            <HeaderStyle BackColor="#507CD1" Font-Bold="True" 
        ForeColor="White" />
                            <PagerStyle BackColor="#2461BF" ForeColor="White" 
        HorizontalAlign="Center" />
                            <RowStyle BackColor="#EFF3FB" />
                            <SelectedRowStyle BackColor="#D1DDF1" Font-Bold="True" 
        ForeColor="#333333" />
                            <SortedAscendingCellStyle BackColor="#F5F7FB" />
                            <SortedAscendingHeaderStyle BackColor="#6D95E1" />
                            <SortedDescendingCellStyle BackColor="#E9EBEF" />
                            <SortedDescendingHeaderStyle BackColor="#4870BE" />
                        </asp:GridView>
                            <asp:SqlDataSource ID="lista_osob_x" runat="server" 
                        ConnectionString="<%$ ConnectionStrings:wap %>" 
                        DeleteCommand="delete from tbl_osoby where ident is null" 
                        SelectCommand="SELECT [imie], [nazwisko], [ulica], [kod_poczt], [miejscowosc], [data_poczatkowa], [data_koncowa], [ident] FROM [View_lista_osob_aktywnych]">
                    </asp:SqlDataSource>
                    </asp:Panel>
                    <br />
            </asp:Panel>
         
            <br />
        </asp:Panel>
        <br />
    </asp:Panel>
