<%@ Control Language="C#" AutoEventWireup="true" CodeBehind="daneStatystyczne.ascx.cs" Inherits="wab2018.daneStatystyczne" %>
     <div style="min-height:275px;"><asp:DropDownList ID="DropDownList3" runat="server" AutoPostBack="True" OnSelectedIndexChanged="changeQuerry" DataSourceID="kwerendyStatystyczne0" DataTextField="Nazwa" DataValueField="kwerenda" meta:resourcekey="DropDownList3Resource1" ViewStateMode="Enabled"></asp:DropDownList>
                    <asp:GridView ID="GridView1" runat="server" Css meta:resourcekey="GridView1Resource1" CellPadding="4" ForeColor="#333333" GridLines="Vertical" CellSpacing="2" ShowHeaderWhenEmpty="True" Width="100%">
                        <AlternatingRowStyle BackColor="White" />
                        <EditRowStyle BackColor="#2461BF" />
                        <FooterStyle BackColor="#507CD1" Font-Bold="True" ForeColor="White" />
                        <HeaderStyle BackColor="#507CD1" Font-Bold="True" ForeColor="White" />
                        <PagerStyle BackColor="#2461BF" ForeColor="White" HorizontalAlign="Center" />
                        <RowStyle BackColor="#EFF3FB" CssClass="wciecie" HorizontalAlign="Center" />
                        <SelectedRowStyle BackColor="#D1DDF1" Font-Bold="True" ForeColor="#333333" />
                        <SortedAscendingCellStyle BackColor="#F5F7FB" />
                        <SortedAscendingHeaderStyle BackColor="#6D95E1" />
                        <SortedDescendingCellStyle BackColor="#E9EBEF" />
                        <SortedDescendingHeaderStyle BackColor="#4870BE" />
         </asp:GridView>
                                <asp:SqlDataSource ID="kwerendyStatystyczne0" runat="server" ConnectionString="<%$ ConnectionStrings:wap %>" SelectCommand="SELECT Nazwa, kwerenda, id_ FROM dane_statystyczne WHERE (czy_us &lt;&gt; 1)  ORDER BY Nazwa">
                                </asp:SqlDataSource>
                            </div>

