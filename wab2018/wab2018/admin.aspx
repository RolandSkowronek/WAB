<%@ Page Title="" Language="C#" MasterPageFile="~/Master.Master" AutoEventWireup="true" CodeBehind="admin.aspx.cs" Inherits="wab2018.admin" %>
<%@ Register assembly="DevExpress.Web.v17.1, Version=17.1.17.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.Web" tagprefix="dx" %>
<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="server">
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
     <style type="text/css">
        .style1
        {
            height: 21px;
        }
        .style2
        {
            width: 266px;
        }
        .style3
        {

            height: 21px;
            width: 266px;
        }
        .style4
        {
            width: 145px;
        }
    .auto-style1 {
        width: 686px;
    }
    .auto-style2 {
        width: 625px;
    }
        .auto-style4 {
            width: 536px;
        }
        .auto-style5 {
            width: 70%;
        }
        .auto-style6 {
            -moz-box-shadow: inset 0px 1px 0px 0px #ffffff;
            -webkit-box-shadow: inset 0px 1px 0px 0px #ffffff;
            box-shadow: inset 0px 1px 0px 0px #ffffff;
            filter: progid:DXImageTransform.Microsoft.gradient(startColorstr='#ededed', endColorstr='#dfdfdf');
            background-color: #ededed;
            -webkit-border-top-left-radius: 6px;
            -moz-border-radius-topleft: 6px;
            border-top-left-radius: 6px;
            -webkit-border-top-right-radius: 6px;
            -moz-border-radius-topright: 6px;
            border-top-right-radius: 6px;
            -webkit-border-bottom-right-radius: 6px;
            -moz-border-radius-bottomright: 6px;
            border-bottom-right-radius: 6px;
            -webkit-border-bottom-left-radius: 6px;
            -moz-border-radius-bottomleft: 6px;
            border-bottom-left-radius: 6px;
            text-indent: 0;
            border: 1px solid #dcdcdc;
            display: inline-block;
            color: #777777;
            font-family: arial;
            font-size: 15px;
            font-weight: bold;
            font-style: normal;
            line-height: 30px;
            text-decoration: none;
            text-align: center;
            text-shadow: 1px 1px 0px #ffffff;
        }
        </style>
    <div id ="mainWindow" class="newPage"> 
      <asp:Panel ID="Panel1" runat="server">
        <table style="width:100%;">
            <tr>
                <td>
                <asp:LinkButton ID="LinkButton2" runat="server" CssClass="button_" 
                        Width="125px" onclick="LinkButton2_Click" CausesValidation="False">Specjalizacje</asp:LinkButton>

                   
                </td>
                <td>
                    <asp:LinkButton ID="LinkButton25" runat="server" CausesValidation="False" 
                        CssClass="button_" onclick="LinkButton25_Click" Width="125px">Gr. specjalizacji</asp:LinkButton>
                </td>
                <td>
                    <asp:LinkButton ID="LinkButton4" runat="server" CssClass="button_" 
                        onclick="LinkButton4_Click" Width="125px" CausesValidation="False">Pełna lista osób</asp:LinkButton>
                </td>
                <td>
                    <asp:LinkButton ID="LinkButton3" runat="server" CssClass="button_" 
                        Width="125px" onclick="LinkButton3_Click" CausesValidation="False">Lista obciążeń</asp:LinkButton>
                </td>
                <td>
                    <asp:LinkButton ID="LinkButton5" runat="server" CssClass="button_" 
                        Width="125px" onclick="LinkButton5_Click" CausesValidation="False">Modyfikacje osób</asp:LinkButton>
                </td>
                <td>
                    <asp:LinkButton ID="LinkButton6" runat="server" CssClass="button_" 
                        Width="125px" onclick="LinkButton6_Click" CausesValidation="False" 
                        Visible="False">Mod. obciążeń</asp:LinkButton>
                </td>
                <td>
                     <asp:LinkButton ID="LinkButton1" runat="server" CssClass="button_" 
                        onclick="LinkButton1_Click" Width="125px" CausesValidation="False">Użytkownicy</asp:LinkButton>
                </td>
                <td >
                    <asp:LinkButton ID="LinkButton12" runat="server" CausesValidation="False" 
                        CssClass="button_" Width="125px" onclick="LinkButton12_Click">Bazy danych</asp:LinkButton>
                </td>
            </tr>
        </table>
    </asp:Panel>
 <br />
    <asp:Panel ID="Panel2" runat="server" CssClass="admin_panel">
        <h4>
            <asp:Label ID="Label2" runat="server" Text="Użytkownicy"></asp:Label>
        </h4>
  <asp:LinkButton ID="LinkButton7" runat="server" CssClass="button_" 
                        Width="200px" onclick="LinkButton7_Click">Dodaj użytkownika</asp:LinkButton>
  <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False" 
                            CellPadding="2" DataKeyNames="id_,user_,pasword,imie,nazwisko,rola" DataSourceID="uzytkownicy" 
                            ForeColor="#333333" GridLines="None" Width="90%" 
                            onselectedindexchanged="GridView1_SelectedIndexChanged" PageSize="15">
                            <AlternatingRowStyle BackColor="White" />
                            <Columns>
                                <asp:CommandField ButtonType="Image" SelectImageUrl="~/img/details.jpg" 
                                    ShowSelectButton="True" />
                                <asp:BoundField DataField="user_" HeaderText="Użytkownik" ReadOnly="True" 
                                    SortExpression="user_" />
                                <asp:BoundField DataField="imie" HeaderText="imie" SortExpression="imie" />
                                <asp:BoundField DataField="nazwisko" HeaderText="nazwisko" 
                                    SortExpression="nazwisko" />
                                <asp:BoundField DataField="nazwa" HeaderText="Rola" SortExpression="nazwa" />
                            </Columns>
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
                        <asp:SqlDataSource ID="uzytkownicy" runat="server" 
                            ConnectionString="<%$ ConnectionStrings:wap %>" 
                            SelectCommand="SELECT     tbl_users_.id_, tbl_users_.user_, tbl_users_.pasword, tbl_users_.imie, tbl_users_.nazwisko, glo_role_userow.nazwa, tbl_users_.rola
FROM         tbl_users_ LEFT OUTER JOIN
                      glo_role_userow ON tbl_users_.rola = glo_role_userow.id_
order by 2">
                        </asp:SqlDataSource>
  

      
        <br />
    </asp:Panel>
     
   <asp:Panel ID="Panel4" runat="server" CssClass="admin_panel" Visible="False">
       <h4>
           Pełna lista osób
       </h4>
     
                        <asp:GridView ID="GridView2" runat="server" AutoGenerateColumns="False" 
                            CellPadding="4" DataSourceID="osoby_full" 
           ForeColor="#333333" GridLines="Vertical" Width="100%">
                            <AlternatingRowStyle BackColor="White" />
                            <Columns>
                                <asp:BoundField DataField="imie" HeaderText="Imie" SortExpression="imie" />
                                <asp:BoundField DataField="nazwisko" HeaderText="Nazwisko" 
                                    SortExpression="nazwisko" />
                                <asp:BoundField DataField="ulica" HeaderText="Ulica" SortExpression="ulica" />
                                <asp:TemplateField HeaderText="Miejscowość" SortExpression="kod_poczt">
                                    <EditItemTemplate>
                                        <asp:TextBox ID="TextBox1" runat="server" Text='<%# Bind("kod_poczt") %>'></asp:TextBox>
                                    </EditItemTemplate>
                                    <ItemTemplate>
                                        <asp:Label ID="Label1" runat="server" Text='<%# Bind("kod_poczt") %>'></asp:Label>
                                    </ItemTemplate>
                                </asp:TemplateField>
                                <asp:BoundField DataField="data_poczatkowa" HeaderText="Data poczatkowa" 
                                    SortExpression="data_poczatkowa" DataFormatString="{0:dd.MM.yyyy}" />
                                <asp:BoundField DataField="data_koncowa" HeaderText="data koncowa" 
                                    SortExpression="data_koncowa" DataFormatString="{0:dd.MM.yyyy}" />
                                <asp:BoundField DataField="data_kreacji" HeaderText="data kreacji" 
                                    SortExpression="data_kreacji" DataFormatString="{0:dd.MM.yyyy}" />
                                <asp:BoundField DataField="kreator" HeaderText="kreator" 
                                    SortExpression="kreator" />
                                <asp:BoundField DataField="czyus" HeaderText="Usunięty" 
                                    SortExpression="czyus" />
                                <asp:BoundField DataField="d_usuniecia" DataFormatString="{0:dd.MM.yyyy}" 
                                    HeaderText="data usunięcia" SortExpression="d_usuniecia" />
                                <asp:BoundField DataField="destruktor" HeaderText="Usuwający" 
                                    SortExpression="destruktor" />
                                <asp:BoundField DataField="Czy zawieszony" HeaderText="Czy zawieszony" SortExpression="Czy zawieszony" />
                            </Columns>
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
                        <asp:SqlDataSource ID="osoby_full" runat="server" 
                            ConnectionString="<%$ ConnectionStrings:wap %>" SelectCommand="SELECT TOP (200) tbl_osoby.imie, tbl_osoby.nazwisko, tbl_osoby.ulica, tbl_osoby.kod_poczt, tbl_osoby.miejscowosc, tbl_osoby.data_poczatkowa, tbl_osoby.data_koncowa, tbl_osoby.data_kreacji, tbl_osoby.pesel, tbl_osoby.czyus, tbl_osoby.d_usuniecia, tbl_osoby.przyczyna_usuniecia, tbl_users__1.user_ AS destruktor, tbl_users_.user_ AS kreator, tbl_osoby.czy_zaw AS [Czy zawieszony] FROM tbl_osoby LEFT OUTER JOIN tbl_users_ ON tbl_osoby.id_kreatora = tbl_users_.id_ LEFT OUTER JOIN tbl_users_ AS tbl_users__1 ON tbl_osoby.id_usuwajacego = tbl_users__1.id_">
                        </asp:SqlDataSource>
                    </asp:Panel>
   
    <asp:Panel ID="Panel5" runat="server" Height="400px" ScrollBars="Vertical" 
        Visible="False" CssClass="admin_panel">
        <h4>
            <asp:Label ID="Label1" runat="server" Text="Pełna lista obciążeń"></asp:Label>
        </h4>
       
     
        <asp:GridView ID="GridView3" runat="server" AutoGenerateColumns="False" 
            CellPadding="4" DataSourceID="obciazenia_full" ForeColor="#333333" 
            GridLines="None" Width="100%">
            <AlternatingRowStyle BackColor="White" />
            <Columns>
                <asp:BoundField DataField="instytucja" HeaderText="instytucja" 
                    SortExpression="instytucja" />
                <asp:BoundField DataField="sygnatura" HeaderText="sygnatura" 
                    SortExpression="sygnatura" />
                <asp:BoundField DataField="data_wprowadzenia" HeaderText="data_wprowadzenia" 
                    SortExpression="data_wprowadzenia" />
                <asp:BoundField DataField="data_otrzymania" HeaderText="data_otrzymania" 
                    SortExpression="data_otrzymania" />
                <asp:BoundField DataField="termin" HeaderText="termin" 
                    SortExpression="termin" />
                <asp:BoundField DataField="imie" HeaderText="imie" SortExpression="imie" />
                <asp:BoundField DataField="nazwisko" HeaderText="nazwisko" 
                    SortExpression="nazwisko" />
            </Columns>
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
        <asp:SqlDataSource ID="obciazenia_full" runat="server" 
            ConnectionString="<%$ ConnectionStrings:wap %>" SelectCommand="SELECT     distinct tbl_obciazenia.instytucja, tbl_obciazenia.sygnatura, tbl_obciazenia.data_wprowadzenia, tbl_obciazenia.data_otrzymania, tbl_obciazenia.termin, tbl_osoby.imie, 
                      tbl_osoby.nazwisko
FROM         tbl_obciazenia LEFT OUTER JOIN
                      tbl_osoby ON tbl_obciazenia.id_osoby = tbl_osoby.ident">
        </asp:SqlDataSource>
       
    </asp:Panel>
      <asp:Panel ID="Panel6" runat="server" Height="400px" ScrollBars="Vertical" 
        Visible="False" CssClass="admin_panel">
          <h4>
              <asp:Label ID="Label3" runat="server" Text="Modyfikacje użytkowników"></asp:Label>
          </h4>
      
          <asp:GridView ID="GridView4" runat="server" AutoGenerateColumns="False" 
              CellPadding="4" DataSourceID="modyfikacje_osob" ForeColor="#333333" 
              GridLines="None" Width="100%">
              <AlternatingRowStyle BackColor="White" />
              <Columns>
                  <asp:BoundField DataField="data_modyfikacji" DataFormatString="{0:dd.MM.yyyy}" 
                      HeaderText="Data modyfikacji" SortExpression="data_modyfikacji" />
                  <asp:TemplateField HeaderText="Imię" SortExpression="imie_org">
                      <EditItemTemplate>
                          <asp:TextBox ID="TextBox1" runat="server" Text='<%# Bind("imie_org") %>'></asp:TextBox>
                      </EditItemTemplate>
                      <ItemTemplate>
                          <asp:Label ID="Label1" runat="server" Text='<%# Bind("imie_org") %>'></asp:Label>
                          <br />
                          <asp:Label ID="Label2" runat="server" Text='<%# Bind("imie_mod") %>'></asp:Label>
                      </ItemTemplate>
                  </asp:TemplateField>
                  <asp:TemplateField HeaderText="Nazwisko" SortExpression="nazwisko_org">
                      <EditItemTemplate>
                          <asp:TextBox ID="TextBox2" runat="server" Text='<%# Bind("nazwisko_org") %>'></asp:TextBox>
                      </EditItemTemplate>
                      <ItemTemplate>
                          <asp:Label ID="Label2" runat="server" Text='<%# Bind("nazwisko_org") %>'></asp:Label>
                          <br />
                          <asp:Label ID="Label8" runat="server" Text='<%# Eval("nazwisko_mod") %>'></asp:Label>
                      </ItemTemplate>
                  </asp:TemplateField>
                  <asp:TemplateField HeaderText="Adres" SortExpression="adres_org">
                      <EditItemTemplate>
                          <asp:TextBox ID="TextBox3" runat="server" Text='<%# Bind("adres_org") %>'></asp:TextBox>
                      </EditItemTemplate>
                      <ItemTemplate>
                          <asp:Label ID="Label3" runat="server" Text='<%# Bind("adres_org") %>'></asp:Label>
                          <br />
                          <asp:Label ID="Label9" runat="server" Text='<%# Eval("adres_mod") %>'></asp:Label>
                      </ItemTemplate>
                  </asp:TemplateField>
                  <asp:TemplateField HeaderText="Kod pocztowy" SortExpression="kod_poczt_org">
                      <EditItemTemplate>
                          <asp:TextBox ID="TextBox4" runat="server" Text='<%# Bind("kod_poczt_org") %>'></asp:TextBox>
                      </EditItemTemplate>
                      <ItemTemplate>
                          <asp:Label ID="Label4" runat="server" Text='<%# Bind("kod_poczt_org") %>'></asp:Label>
                          <br />
                          <asp:Label ID="Label10" runat="server" Text='<%# Eval("kod_poczt_mod") %>'></asp:Label>
                      </ItemTemplate>
                  </asp:TemplateField>
                  <asp:TemplateField HeaderText="Miejscowość" SortExpression="miejscowosc_org">
                      <EditItemTemplate>
                          <asp:TextBox ID="TextBox5" runat="server" Text='<%# Bind("miejscowosc_org") %>'></asp:TextBox>
                      </EditItemTemplate>
                      <ItemTemplate>
                          <asp:Label ID="Label5" runat="server" Text='<%# Bind("miejscowosc_org") %>'></asp:Label>
                          <br />
                          <asp:Label ID="Label11" runat="server" Text='<%# Eval("miejscowosc_mod") %>'></asp:Label>
                      </ItemTemplate>
                  </asp:TemplateField>
                  <asp:TemplateField HeaderText="Początek" SortExpression="data_pocz_org">
                      <EditItemTemplate>
                          <asp:TextBox ID="TextBox6" runat="server" Text='<%# Bind("data_pocz_org") %>'></asp:TextBox>
                      </EditItemTemplate>
                      <ItemTemplate>
                          <asp:Label ID="Label6" runat="server" 
                              Text='<%# Bind("data_pocz_org", "{0:dd.MM.yyyy}") %>'></asp:Label>
                          <br />
                          <asp:Label ID="Label12" runat="server" 
                              Text='<%# Eval("data_pocz_mod", "{0:dd.MM.yyyy}") %>'></asp:Label>
                      </ItemTemplate>
                  </asp:TemplateField>
                  <asp:TemplateField HeaderText="Koniec" SortExpression="data_konc_org">
                      <EditItemTemplate>
                          <asp:TextBox ID="TextBox7" runat="server" Text='<%# Bind("data_konc_org") %>'></asp:TextBox>
                      </EditItemTemplate>
                      <ItemTemplate>
                          <asp:Label ID="Label7" runat="server" 
                              Text='<%# Bind("data_konc_org", "{0:dd.MM.yyyy}") %>'></asp:Label>
                          <br />
                          <asp:Label ID="Label13" runat="server" 
                              Text='<%# Eval("data_konc_mod", "{0:dd.MM.yyyy}") %>'></asp:Label>
                      </ItemTemplate>
                  </asp:TemplateField>
                  <asp:BoundField DataField="user_" HeaderText="Użytkownik" 
                      SortExpression="user_" />
                  <asp:BoundField DataField="Expr1" HeaderText="Expr1" SortExpression="Expr1" />
                  <asp:BoundField DataField="Expr2" HeaderText="Expr2" SortExpression="Expr2" />
              </Columns>
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
          <asp:SqlDataSource ID="modyfikacje_osob" runat="server" 
              ConnectionString="<%$ ConnectionStrings:wap %>" SelectCommand="SELECT     tbl_modyfikacje.id_, tbl_modyfikacje.data_modyfikacji, tbl_modyfikacje.imie_org, tbl_modyfikacje.imie_mod, tbl_modyfikacje.nazwisko_org, tbl_modyfikacje.nazwisko_mod, 
                      tbl_modyfikacje.adres_org, tbl_modyfikacje.adres_mod, tbl_modyfikacje.kod_poczt_org, tbl_modyfikacje.kod_poczt_mod, tbl_modyfikacje.miejscowosc_org, tbl_modyfikacje.miejscowosc_mod, 
                      tbl_modyfikacje.uwagi, tbl_modyfikacje.data_pocz_org, tbl_modyfikacje.data_pocz_mod, tbl_modyfikacje.data_konc_org, tbl_modyfikacje.data_konc_mod, tbl_osoby.imie, tbl_osoby.nazwisko, 
                      tbl_users_.user_, tbl_users_.imie AS Expr1, tbl_users_.nazwisko AS Expr2
FROM         tbl_modyfikacje LEFT OUTER JOIN
                      tbl_users_ ON tbl_modyfikacje.id_modyfikujacego = tbl_users_.id_ LEFT OUTER JOIN
                      tbl_osoby ON tbl_modyfikacje.id_osoby = tbl_osoby.ident">
          </asp:SqlDataSource>
       
    </asp:Panel>
        <asp:Panel ID="Panel7" runat="server" Height="420px" 
        Visible="False" CssClass="admin_panel">
            <h4>
                <asp:Label ID="Label14" runat="server" Text="Specjalizacje"></asp:Label>
            </h4>
            &nbsp;<table width="100%">
                <tr>
                    <td class="auto-style5">
                        <asp:SqlDataSource ID="specjalizacje" runat="server" 
                            ConnectionString="<%$ ConnectionStrings:wap %>" 
                            
                            
                            SelectCommand="SELECT glo_specjalizacje.id_, rtrim(glo_specjalizacje.nazwa) as nazwa
, glo_specjalizacje.grupa, rtrim(glo_grupy_specjalizacji.Nazwa) AS nazwa_grupy FROM glo_specjalizacje LEFT OUTER JOIN glo_grupy_specjalizacji ON glo_specjalizacje.grupa = glo_grupy_specjalizacji.ident ORDER BY glo_specjalizacje.nazwa">
                        </asp:SqlDataSource>
                        <asp:Panel ID="Panel12" runat="server" Height="350px" ScrollBars="Vertical" Width="800px">
                            <asp:GridView ID="GridView5" runat="server" AutoGenerateColumns="False" 
                                CellPadding="3" DataKeyNames="id_,nazwa,grupa" DataSourceID="specjalizacje" 
                                ForeColor="#333333" GridLines="None" 
                                onselectedindexchanged="GridView5_SelectedIndexChanged" Width="97%" Height="350px">
                                <AlternatingRowStyle BackColor="White" />
                                <Columns>
                                    <asp:CommandField ButtonType="Image" SelectImageUrl="~/img/details.jpg" 
                                        ShowSelectButton="True" >
                                    <ItemStyle Width="25px" />
                                    </asp:CommandField>
                                    <asp:BoundField DataField="nazwa" HeaderText="nazwa" ReadOnly="True" 
                                        SortExpression="nazwa" >
                                    <ItemStyle Width="360px" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="nazwa_grupy" HeaderText="Grupa" 
                                        SortExpression="nazwa_grupy" >
                                    <ItemStyle CssClass="col_left " Width="300px" />
                                    </asp:BoundField>
                                </Columns>
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
                        </asp:Panel>
                    </td>
                    <td style="top: 0px; vertical-align: top;" 
                        width="30%">
                        <table style="width:95%;">
                            <tr valign="top">
                                <td>
                                    Nazwa:</td>
                                <td>
                                    <asp:TextBox ID="TextBox8" runat="server" Width="80%"></asp:TextBox>
                                    <asp:RequiredFieldValidator ID="RequiredFieldValidator1" runat="server" 
                                        ControlToValidate="TextBox8" ErrorMessage="Nazwa specyfikacji musi być podana"></asp:RequiredFieldValidator>
                                </td>
                            </tr>
                            <tr>
                                <td>
                                    Grupa</td>
                                <td>
                                    <asp:SqlDataSource ID="Grupy_specjalizacji" runat="server" 
                                        ConnectionString="<%$ ConnectionStrings:wap %>" 
                                        SelectCommand="SELECT DISTINCT [ident], rtrim([Nazwa]) as nazwa FROM [glo_grupy_specjalizacji] ORDER BY [Nazwa]">
                                    </asp:SqlDataSource>
                                    <asp:DropDownList ID="DropDownList3" runat="server" 
                                        DataSourceID="Grupy_specjalizacji" DataTextField="Nazwa" DataValueField="ident">
                                    </asp:DropDownList>
                                </td>
                            </tr>
                            <tr>
                                <td colspan="2">
                                    <table style="width:100%;">
                                        <tr>
                                            <td style="text-align: left" align="left">
                                                <asp:LinkButton ID="LinkButton9" runat="server" CssClass="button_" onclick="LinkButton9_Click" Width="80px">Zapisz</asp:LinkButton>
                                                <asp:LinkButton ID="LinkButton10" runat="server" CausesValidation="False" CssClass="button_" onclick="LinkButton10_Click" Width="80px">Anuluj</asp:LinkButton>
                                                <asp:LinkButton ID="LinkButton11" runat="server" CausesValidation="False" CssClass="button_" onclick="LinkButton11_Click" Width="80px">Usuń</asp:LinkButton>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td style="text-align: left">
                                                <asp:LinkButton ID="LinkButton8" runat="server" CssClass="button_" Height="0px" 
                                                    onclick="LinkButton8_Click" Visible="False" Width="253px">Dodaj specjalizację</asp:LinkButton>
                                            </td>
                                        </tr>
                                    </table>
                                </td>
                            </tr>
                        </table>
                    </td>
                </tr>
            </table>
           
       
      
       
    </asp:Panel>
   <asp:Panel ID="Panel8" runat="server" Height="450px" 
        Visible="False" CssClass="admin_panel">
        <h4>
            <asp:Label ID="Label15" runat="server" Text="Bazy danych"></asp:Label>
        </h4>
       

        <p>
        </p>
        <asp:Panel ID="Panel16" runat="server" BackColor="#9999FF" Visible="False">
            <strong>Dodanie nowej bazy danych</strong><br />
            <asp:Panel ID="Panel17" runat="server" BackColor="#CCCCFF">
                <table style="width:100%;">
                    <tr>
                        <td class="style1">Nazwa bazy (opisowa)</td>
                        <td class="style1">
                            <asp:TextBox ID="dbOpis" runat="server" required="required"></asp:TextBox>
                        </td>
                        <td class="style1"></td>
                    </tr>
                    <tr>
                        <td>Nazwa Servera</td>
                        <td>
                            <asp:TextBox ID="dbServer" runat="server" required="required"></asp:TextBox>
                        </td>
                        <td>&nbsp;</td>
                    </tr>
                    <tr>
                        <td>Nazwa bazy danych</td>
                        <td>
                            <asp:TextBox ID="dbNazwa" runat="server" required="required"></asp:TextBox>
                        </td>
                        <td>&nbsp;</td>
                    </tr>
                    <tr>
                        <td>Nazwa użytkownika</td>
                        <td>
                            <asp:TextBox ID="dbUser" runat="server" required="required"></asp:TextBox>
                        </td>
                        <td>&nbsp;</td>
                    </tr>
                    <tr>
                        <td>Hasło</td>
                        <td>
                            <asp:TextBox ID="dbPasswd" runat="server" required="required"></asp:TextBox>
                        </td>
                        <td>&nbsp;</td>
                    </tr>
                    <tr>
                        <td>Kwerenda odczytująca</td>
                        <td>
                            <asp:DropDownList ID="DropDownList6" runat="server" DataSourceID="Kwerendy0" DataTextField="Nazwa" DataValueField="id_">
                            </asp:DropDownList>
                        </td>
                        <td>&nbsp;</td>
                    </tr>
                </table>
                <br />
                <table style="width:100%;">
                    <tr>
                        <td>
                            <asp:Button ID="Button1" runat="server" Text="Dodaj bazę danych" CssClass="auto-style6" EnableTheming="False" Height="21px" Width="300px" OnClick="Button1_Click" />
                        </td>
                        <td>
                            <asp:LinkButton ID="LinkButton31" runat="server" CssClass="button_" Height="21px" OnClick="LinkButton31_Click" Width="300px" CausesValidation="False">Anuluj</asp:LinkButton>
                        </td>
                    </tr>
                </table>
                <br />
                <br />
            </asp:Panel>
            <br />
        </asp:Panel>
        <asp:Panel ID="Panel14" runat="server">
            <table cellpadding="0" cellspacing="0" style="width: 100%; height: 400px;">
                <tr>
                    <td class="auto-style4">
                        <asp:Panel ID="Panel15" runat="server" Height="400px">
                            <asp:GridView ID="GridView6" runat="server" AutoGenerateColumns="False" CellPadding="4" DataKeyNames="id_,nazwa,db_name,db_catalog,db_user,db_paswd,id_kwerendy" DataSourceID="bazy_danych0" ForeColor="#333333" GridLines="None" onselectedindexchanged="GridView6_SelectedIndexChanged" Width="95%">
                                <AlternatingRowStyle BackColor="White" />
                                <Columns>
                                    <asp:CommandField ButtonType="Image" SelectImageUrl="~/img/details.jpg" ShowSelectButton="True">
                                    <ItemStyle Width="25px" />
                                    </asp:CommandField>
                                    <asp:BoundField DataField="nazwa" HeaderText="nazwa" ReadOnly="True" SortExpression="nazwa" />
                                </Columns>
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
                            <asp:SqlDataSource ID="bazy_danych0" runat="server" ConnectionString="<%$ ConnectionStrings:wap %>" SelectCommand="SELECT DISTINCT id_, nazwa, db_name, db_user, db_paswd, db_catalog, id_kwerendy FROM tbl_bazy_danych ORDER BY nazwa"></asp:SqlDataSource>
                        </asp:Panel>
                    </td>
                    <td valign="top">
                        <table style="width:100%;" valign="top">
                            <tr>
                                <td colspan="2">
                                    <asp:LinkButton ID="LinkButton26" runat="server" CausesValidation="False" CssClass="button_" onclick="LinkButton16_Click" Width="99%">Dodaj nową bazę danych</asp:LinkButton>
                                </td>
                            </tr>
                            <tr>
                                <td>Nazwa bazy (opisowa)</td>
                                <td>
                                    <asp:TextBox ID="TextBox9" runat="server"></asp:TextBox>
                                    <asp:RequiredFieldValidator ID="RequiredFieldValidator11" runat="server" ControlToValidate="TextBox9" ErrorMessage="Pole musi być wypełnione"></asp:RequiredFieldValidator>
                                </td>
                            </tr>
                            <tr>
                                <td class="style1">Nazwa serwera</td>
                                <td class="style1">
                                    <asp:TextBox ID="TextBox10" runat="server"></asp:TextBox>
                                    <asp:RequiredFieldValidator ID="RequiredFieldValidator12" runat="server" ControlToValidate="TextBox10" ErrorMessage="Pole musi być wypełnione"></asp:RequiredFieldValidator>
                                </td>
                            </tr>
                            <tr>
                                <td>Nazwa bazy danych</td>
                                <td>
                                    <asp:TextBox ID="TextBox11" runat="server"></asp:TextBox>
                                    <asp:RequiredFieldValidator ID="RequiredFieldValidator13" runat="server" ControlToValidate="TextBox11" ErrorMessage="Pole musi być wypełnione"></asp:RequiredFieldValidator>
                                </td>
                            </tr>
                            <tr>
                                <td>Nazwa użytkownika</td>
                                <td>
                                    <asp:TextBox ID="TextBox12" runat="server"></asp:TextBox>
                                    <asp:RequiredFieldValidator ID="RequiredFieldValidator14" runat="server" ControlToValidate="TextBox12" ErrorMessage="Pole musi być wypełnione"></asp:RequiredFieldValidator>
                                </td>
                            </tr>
                            <tr>
                                <td>Hasło</td>
                                <td>
                                    <asp:TextBox ID="TextBox13" runat="server"></asp:TextBox>
                                    <asp:RequiredFieldValidator ID="RequiredFieldValidator15" runat="server" ControlToValidate="TextBox13" ErrorMessage="Pole musi być wypełnione"></asp:RequiredFieldValidator>
                                </td>
                            </tr>
                            <tr>
                                <td>Kwerenda importująca</td>
                                <td>
                                    <asp:DropDownList ID="DropDownList4" runat="server" DataSourceID="Kwerendy0" DataTextField="Nazwa" DataValueField="id_">
                                    </asp:DropDownList>
                                    <asp:SqlDataSource ID="Kwerendy0" runat="server" ConnectionString="<%$ ConnectionStrings:wap %>" SelectCommand="SELECT DISTINCT [id_], [Nazwa] FROM [glo_kwerenda] ORDER BY [Nazwa]"></asp:SqlDataSource>
                                </td>
                            </tr>
                            <tr>
                                <td colspan="2" valign="top">
                                    <table width="95%">
                                        <tr>
                                            <td>
                                                <asp:LinkButton ID="LinkButton27" runat="server" CssClass="button_" onclick="LinkButton13_Click">Zapisz</asp:LinkButton>
                                            </td>
                                            <td>
                                                <asp:LinkButton ID="LinkButton28" runat="server" CausesValidation="False" CssClass="button_">Anuluj</asp:LinkButton>
                                            </td>
                                            <td>
                                                <asp:LinkButton ID="LinkButton29" runat="server" CausesValidation="False" CssClass="button_" onclick="LinkButton15_Click">Usuń</asp:LinkButton>
                                            </td>
                                            <td>
                                                <asp:LinkButton ID="LinkButton30" runat="server" CausesValidation="False" CssClass="button_" onclick="LinkButton17_Click">Testuj</asp:LinkButton>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td colspan="4">
                                                <asp:Label ID="Label16" runat="server"></asp:Label>
                                                &nbsp;</td>
                                        </tr>
                                    </table>
                                </td>
                            </tr>
                        </table>
                    </td>
                </tr>
            </table>
            <br />
        </asp:Panel>
       
       
      
       
    </asp:Panel>

    <asp:Panel ID="Panel3" runat="server"  Visible="False" CssClass="admin_panel">
        <h4>
            <asp:Label ID="Label18" runat="server" Text="Grupy specjalizacji"></asp:Label>
        </h4>
       

       <table>
           <tr>
               <td style="vertical-align: top" class="auto-style1">
                   <asp:Panel ID="Panel13" runat="server" Height="400px" ScrollBars="Vertical">
 <asp:GridView ID="GridView7" runat="server" AutoGenerateColumns="False" 
                       CellPadding="3" DataSourceID="Grupy_specjalizacji" 
                       ForeColor="#333333" GridLines="None" Width="95%" 
                       onselectedindexchanged="GridView7_SelectedIndexChanged" 
                           DataKeyNames="ident,Nazwa">
                       <AlternatingRowStyle BackColor="White" />
                       <Columns>
                           <asp:CommandField ButtonType="Image" SelectImageUrl="~/img/details.jpg" 
                               ShowSelectButton="True" >
                           <ItemStyle Width="25px" />
                           </asp:CommandField>
                           <asp:BoundField DataField="Nazwa" HeaderText="Nazwa" ReadOnly="True" 
                               SortExpression="Nazwa" />
                       </Columns>
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
                   <asp:SqlDataSource ID="SqlDataSource1" runat="server" 
                       ConnectionString="<%$ ConnectionStrings:wap %>" 
                       SelectCommand="SELECT DISTINCT [id_], [nazwa], [db_name], [db_user], [db_paswd], [db_catalog] FROM [tbl_bazy_danych] ORDER BY [nazwa]">
                   </asp:SqlDataSource>
                   </asp:Panel>
                  
               </td>
               <td style="vertical-align: top;" class="auto-style2">
                   <table cellpadding="0" cellspacing="0">
                       <tr>
                           <td>
                               &nbsp;</td>
                       </tr>
                       <tr>
                           <td valign="top">
                               Nazwa grupy specjalizacji&nbsp;&nbsp;
                               <asp:TextBox ID="TextBox19" runat="server" Width="255px"></asp:TextBox>
                           </td>
                       </tr>
                       <tr>
                           <td>
                               <asp:RequiredFieldValidator ID="RequiredFieldValidator10" runat="server" ControlToValidate="TextBox19" ErrorMessage="Pole musi być wypełnione"></asp:RequiredFieldValidator>
                           </td>
                       </tr>
                       <tr>
                           <td>
                               <table>
                                   <tr>
                                       <td>
                                           <asp:LinkButton ID="LinkButton21" runat="server" CausesValidation="False" CssClass="button_" onclick="LinkButton21_Click">Dodaj</asp:LinkButton>
                                           <asp:LinkButton ID="LinkButton22" runat="server" CssClass="button_" onclick="LinkButton22_Click">Zapisz</asp:LinkButton>
                                           <asp:LinkButton ID="LinkButton23" runat="server" CausesValidation="False" CssClass="button_" onclick="LinkButton23_Click">Anuluj</asp:LinkButton>
                                           <asp:LinkButton ID="LinkButton24" runat="server" CausesValidation="False" CssClass="button_" onclick="LinkButton24_Click">Usuń</asp:LinkButton>
                                       </td>
                                   </tr>
                                   <tr>
                                       <td>
                                           <asp:Label ID="Label19" runat="server" Text=""></asp:Label>
                                           &nbsp;</td>
                                   </tr>
                               </table>
                           </td>
                       </tr>
                   </table>
               </td>
           </tr>
       </table>
       
       
      
       
    </asp:Panel>
    <asp:Panel ID="Panel10" runat="server" Visible="False" CssClass="Panel_2_2">
        <asp:Panel ID="Panel11" runat="server" CssClass="Panel_1_1">
        
        <table style="border-style: 0; border-width: 0px; width:100%;" cellpadding="0" 
                cellspacing="0">
            <tr>
                <td style="background-color: #0066FF">
                    <asp:Label ID="Label17" runat="server" Text="Użytkownicy" Font-Bold="True" 
                        ForeColor="#FFCC00"></asp:Label>
                </td>
                <td style="background-color: #0066FF" class="style2">
                    &nbsp;</td>
                <td style="text-align: right; background-color: #0066FF">
                    <asp:ImageButton ID="ImageButton1" runat="server" 
                        ImageUrl="~/img/zamknij.jpg" CausesValidation="False" 
                        onclick="ImageButton1_Click" />
                </td>
            </tr>
            <tr>
                <td>
                    Nazwa użytkownka</td>
                <td class="style2">
                    <asp:TextBox ID="TextBox14" runat="server" Width="90%"></asp:TextBox>
                </td>
                <td>
                    <asp:RequiredFieldValidator ID="RequiredFieldValidator7" runat="server" 
                        ControlToValidate="TextBox14" ErrorMessage="Proszę wypełnic pole"></asp:RequiredFieldValidator>
                </td>
            </tr>
            <tr>
                <td>
                    Hasło</td>
                <td class="style2">
                    <asp:TextBox ID="TextBox15" runat="server" TextMode="Password" Width="90%"></asp:TextBox>
                </td>
                <td>
                    <asp:RequiredFieldValidator ID="RequiredFieldValidator8" runat="server" 
                        ControlToValidate="TextBox15" ErrorMessage="Proszę wypełnic pole"></asp:RequiredFieldValidator>
                </td>
            </tr>
            <tr>
                <td>
                    Powtórz hasło</td>
                <td class="style2">
                    <asp:TextBox ID="TextBox16" runat="server" TextMode="Password" Width="90%"></asp:TextBox>
                </td>
                <td>
                    <asp:RequiredFieldValidator ID="RequiredFieldValidator9" runat="server" 
                        ControlToValidate="TextBox16" ErrorMessage="Proszę wypełnic pole"></asp:RequiredFieldValidator>
                    <asp:CompareValidator ID="CompareValidator1" runat="server" 
                        ControlToCompare="TextBox15" ControlToValidate="TextBox16" 
                        ErrorMessage="Hasło i jego powtórzenie muszą być identyczne"></asp:CompareValidator>
                </td>
            </tr>
            <tr>
                <td>
                    Imię</td>
                <td class="style2">
                    <asp:TextBox ID="TextBox17" runat="server" Width="90%"></asp:TextBox>
                </td>
                <td>
                    &nbsp;</td>
            </tr>
            <tr>
                <td class="style1">
                    Nazwisko</td>
                <td class="style3">
                    <asp:TextBox ID="TextBox18" runat="server" Width="90%"></asp:TextBox>
                </td>
                <td class="style1">
                </td>
            </tr>
            <tr>
                <td>
                    Rola</td>
                <td class="style2">
                    <asp:DropDownList ID="DropDownList1" runat="server" DataSourceID="role" 
                        DataTextField="nazwa" DataValueField="id_">
                    </asp:DropDownList>
                    <asp:SqlDataSource ID="role" runat="server" 
                        ConnectionString="<%$ ConnectionStrings:wap %>" 
                        SelectCommand="SELECT DISTINCT [id_], [nazwa] FROM [glo_role_userow] ORDER BY [nazwa]">
                    </asp:SqlDataSource>
                </td>
                <td>
                    &nbsp;</td>
            </tr>
            <tr>
                <td>
                    
                </td>
                <td class="style2">
                    
                </td>
                <td>
                    &nbsp;</td>
            </tr>
            <tr>
                <td colspan="3">
                    <table style="width:100%;">
                        <tr>
                            <td>
                               
                                <asp:LinkButton ID="LinkButton18" runat="server" CssClass="button_" 
                        onclick="LinkButton18_Click">Zapisz</asp:LinkButton>
                                </td>
                            <td>
<asp:LinkButton ID="LinkButton19" runat="server" CssClass="button_" 
                        CausesValidation="False">Anuluj</asp:LinkButton>


                               </td>
                            <td>

                            <asp:LinkButton ID="LinkButton20" runat="server" CssClass="button_" 
                        CausesValidation="False" onclick="LinkButton20_Click" Visible="False">Usuń</asp:LinkButton>
</td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
        <br />
        <br />


</asp:Panel>
    </asp:Panel>
       </div>
        <div>
    </div>


    aass
        </asp:Content>


