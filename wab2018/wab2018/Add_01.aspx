<%@ Page Title="" Language="C#" MasterPageFile="~/Master.Master" AutoEventWireup="true" CodeBehind="Add_01.aspx.cs" Inherits="wab2018.Add_01" %>

<%@ Register assembly="DevExpress.Web.v17.1, Version=17.1.17.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.Web" tagprefix="dx" %>

<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="server">
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    w<style type="text/css">
        .style1
        {
            height: 26px;
        }
        
        .style2
        {
            width: 123px;
        }
        .style3
        {
            height: 26px;
            width: 123px;
        }
        
        .auto-style2 {
            width: 33%;
        }
        .auto-style3 {
            height: 46px;
        }
        .auto-style4 {
            width: 123px;
            height: 46px;
        }
       
      
        
    </style><br />

   
      <div id ="mainWindow" class="newPage">   



    <asp:Panel ID="Panel6" runat="server" BorderStyle="Solid" BorderWidth="1px">
 <h2>
        <asp:Label ID="Label1" runat="server" Text="Nowy Biegły"></asp:Label>
      
</h2>
  <hr>
        <table style="width:100%; " cellpadding="0" cellspacing="0" >
            <tr>
                <td style="vertical-align: top;" class="auto-style2">
                      <table style="width:100%;">
            <tr>
                <td >
                    <asp:Label ID="Label3" runat="server" Text="Tytuł"></asp:Label>
                </td>
                <td class="style2">
                    <asp:TextBox ID="TextBox8" runat="server"></asp:TextBox>
                </td>
                <td style="width: 40%">
                    &nbsp;</td>
            </tr>
            <tr>
                <td >
                    Imie</td>
                <td class="style2">
                    <asp:TextBox ID="TextBox1" runat="server"></asp:TextBox>
                </td>
                <td style="width: 40%">
                    <asp:RequiredFieldValidator ID="RequiredFieldValidator1" runat="server" 
                        ControlToValidate="TextBox1" ErrorMessage="&lt;- Proszę uzupełnić pole"></asp:RequiredFieldValidator>
                </td>
            </tr>
            <tr>
                <td>
                    Nazwisko</td>
                <td class="style2">
                    <asp:TextBox ID="TextBox2" runat="server"></asp:TextBox>
                </td>
                <td style="width: 40%">
                    <asp:RequiredFieldValidator ID="RequiredFieldValidator2" runat="server" 
                        ControlToValidate="TextBox2" ErrorMessage="&lt;- Proszę uzupełnić pole"></asp:RequiredFieldValidator>
                </td>
            </tr>
            <tr>
                <td class="style1">
                    Adres</td>
                <td class="style3">
                    <asp:TextBox ID="TextBox3" runat="server"></asp:TextBox>
                </td>
                <td class="style1" style="width: 40%">
                    <asp:RequiredFieldValidator ID="RequiredFieldValidator3" runat="server" 
                        ControlToValidate="TextBox3" ErrorMessage="&lt;- Proszę uzupełnić pole"></asp:RequiredFieldValidator>
                </td>
            </tr>
            <tr>
                <td>
                    Kod pocztowy</td>
                <td class="style2">
                    <asp:TextBox ID="TextBox4" runat="server"></asp:TextBox>
                </td>
                <td style="width: 40%">
                    <asp:RequiredFieldValidator ID="RequiredFieldValidator4" runat="server" 
                        ControlToValidate="TextBox4" ErrorMessage="&lt;- Proszę uzupełnić pole"></asp:RequiredFieldValidator>
                    <asp:RegularExpressionValidator ID="RegularExpressionValidator1" runat="server" 
                        ControlToValidate="TextBox4" 
                        ErrorMessage="Proszę wpisać poprawny kod pocztowy XX-XXX" 
                        ValidationExpression="^[0-9]{2}-[0-9]{2}[0-9]$"></asp:RegularExpressionValidator>
                </td>
            </tr>
            <tr>
                <td>
                    Miejscowość</td>
                <td class="style2">
                    <asp:TextBox ID="TextBox5" runat="server"></asp:TextBox>
                </td>
                <td style="width: 40%">
                    <asp:RequiredFieldValidator ID="RequiredFieldValidator5" runat="server" 
                        ControlToValidate="TextBox5" ErrorMessage="&lt;- Proszę uzupełnić pole"></asp:RequiredFieldValidator>
                </td>
            </tr>
            <tr>
                <td class="style1">
                    Powołanie od:</td>
                <td class="style3">
                    <dx:ASPxDateEdit ID="txDataPoczatku" runat="server" DisplayFormatString="dd.MM.yyyy">
                    </dx:ASPxDateEdit>
                  
                </td>
                <td class="style1" style="width: 40%">
                    &nbsp;</td>
            </tr>
            <tr>
                <td>
                    Powołanie do:</td>
                <td class="style2">
                    <dx:ASPxDateEdit ID="txDataKonca" runat="server" DisplayFormatString="dd.MM.yyyy" EditFormat="Custom" EditFormatString="dd.MM.yyyy">
                    </dx:ASPxDateEdit>
                    <asp:TextBox ID="TextBox7" runat="server"></asp:TextBox>
                </td>
                <td style="width: 40%">
                    &nbsp;</td>
            </tr>
            <tr>
                <td abbr="0" align="center" class="auto-style3">
                    Nr PESEL</td>
                <td class="auto-style4" abbr="0" align="center">
                    <asp:TextBox ID="TextBox9" runat="server"></asp:TextBox>
                </td>
                <td abbr="0" align="center" class="auto-style5">
                    <asp:RegularExpressionValidator ID="RegularExpressionValidator2" runat="server" 
                        ControlToValidate="TextBox9" ErrorMessage="Proszę podać poprawny nr PESEL" 
                        ValidationExpression="^[0-9]{11}$"></asp:RegularExpressionValidator>
                </td>
            </tr>
        </table>
  
   
      
                    </td>
                <td class="auto-style2" style="vertical-align: top;">Dane kontaktowe
                    <table style="width:100%;">
        <tr>
            <td>nr telefonu 1</td>
           
            <td>
                <asp:TextBox ID="tel1" runat="server"></asp:TextBox></td>
        </tr>
        <tr>
            <td>nr telefonu 2</td>
            
            <td>
                <asp:TextBox ID="tel2" runat="server"></asp:TextBox>
            </td>
        </tr>
                         <tr>
            <td>email</td>
            
            <td>
                <asp:TextBox ID="email" runat="server"></asp:TextBox>
            </td>
        </tr>
      
    </table>

                    Adres korespondencyjny
                    <table style="width: 100%;">
                        <tr>
                            <td>adres</td>
                          
                            <td>
                                 <asp:TextBox ID="adrKor" runat="server"></asp:TextBox>
                                </td>
                        </tr>
                         <tr>
                            <td>Kod pocztowy</td>
                          
                            <td>
                                 <asp:TextBox ID="kodKor" runat="server"></asp:TextBox>
                            </td>
                        </tr>
                        <tr>
                            <td>miejscowość</td>
                            
                            <td> <asp:TextBox ID="cityKor" runat="server"></asp:TextBox></td>
                        </tr>
                       
                    </table>

                </td>
                <td>
                    <table style="width: 100%;">
                        <tr>
                            <td style="vertical-align: top" colspan="2">
                                <strong>Specjalizacje</strong></td>
                        </tr>
                        <tr>
                            <td style="vertical-align: top; width:50%" >
                                <asp:GridView ID="GridView2" runat="server" AutoGenerateColumns="False" 
                                    DataKeyNames="id_,id_spec" DataSourceID="Specjalizacje_temp" CellPadding="4" 
                                    ForeColor="#333333" GridLines="None" 
                                    onselectedindexchanged="GridView2_SelectedIndexChanged" Width="100%">
                                    <AlternatingRowStyle BackColor="White" />
                                    <Columns>
                                        <asp:CommandField ShowSelectButton="True" SelectText="Usuń" ButtonType="Image" 
                                            SelectImageUrl="~/img/minus.jpg" >
                                        </asp:CommandField>
                                        <asp:BoundField DataField="nazwa" HeaderText="przyznane" ReadOnly="True" 
                                            SortExpression="nazwa" />
                                    </Columns>
                                    <EditRowStyle BackColor="#2461BF" />
                                    <FooterStyle BackColor="#507CD1" Font-Bold="True" ForeColor="White" />
                                    <HeaderStyle BackColor="#507CD1" Font-Bold="True" ForeColor="White" />
                                    <PagerStyle BackColor="#2461BF" ForeColor="White" HorizontalAlign="Center" />
                                    <RowStyle BackColor="#EFF3FB" Font-Size="X-Small" />
                                    <SelectedRowStyle BackColor="#D1DDF1" Font-Bold="True" ForeColor="#333333" />
                                    <SortedAscendingCellStyle BackColor="#F5F7FB" />
                                    <SortedAscendingHeaderStyle BackColor="#6D95E1" />
                                    <SortedDescendingCellStyle BackColor="#E9EBEF" />
                                    <SortedDescendingHeaderStyle BackColor="#4870BE" />
                                </asp:GridView>
                                <asp:SqlDataSource ID="Specjalizacje_temp" runat="server" 
                                    ConnectionString="<%$ ConnectionStrings:wap %>" SelectCommand="SELECT     tbl_sprcjalizacje_temp.id_, glo_specjalizacje.nazwa, tbl_sprcjalizacje_temp.id_spec
FROM         tbl_sprcjalizacje_temp RIGHT OUTER JOIN
                      glo_specjalizacje ON tbl_sprcjalizacje_temp.id_spec = glo_specjalizacje.id_
WHERE     (tbl_sprcjalizacje_temp.nr_sesji = @sesja)
order by 1">
                                    <SelectParameters>
                                        <asp:SessionParameter Name="sesja" SessionField="sesja" />
                                    </SelectParameters>
                                </asp:SqlDataSource>
                            </td>
                            <td style="width:50%;">
                                <asp:Panel ID="Panel1" runat="server" Height="400px" ScrollBars="Vertical">
                                    <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False" 
                                    DataKeyNames="id_" DataSourceID="specjalizacje" 
                                    
    onselectedindexchanged="GridView1_SelectedIndexChanged" CellPadding="3" 
                                    ForeColor="#333333" GridLines="Vertical" Width="90%">
                                        <AlternatingRowStyle BackColor="White" />
                                        <Columns>
                                            <asp:CommandField ShowSelectButton="True" SelectText="Dodaj" ButtonType="Image" 
                                                SelectImageUrl="~/img/plus-icon.png" >
                                            <ItemStyle Width="5px" />
                                            </asp:CommandField>
                                            <asp:BoundField DataField="nazwa" HeaderText="dostępne" ReadOnly="True" 
                                            SortExpression="nazwa" />
                                        </Columns>
                                        <EditRowStyle BackColor="#2461BF" />
                                        <FooterStyle BackColor="#507CD1" Font-Bold="True" ForeColor="White" />
                                        <HeaderStyle BackColor="#507CD1" Font-Bold="True" ForeColor="White" />
                                        <PagerStyle BackColor="#2461BF" ForeColor="White" HorizontalAlign="Center" />
                                        <RowStyle BackColor="#EFF3FB" Font-Size="X-Small" />
                                        <SelectedRowStyle BackColor="#D1DDF1" Font-Bold="True" ForeColor="#333333" />
                                        <SortedAscendingCellStyle BackColor="#F5F7FB" />
                                        <SortedAscendingHeaderStyle BackColor="#6D95E1" />
                                        <SortedDescendingCellStyle BackColor="#E9EBEF" />
                                        <SortedDescendingHeaderStyle BackColor="#4870BE" />
                                    </asp:GridView>
                                </asp:Panel>
                                <asp:SqlDataSource ID="specjalizacje" runat="server" 
                                    ConnectionString="<%$ ConnectionStrings:wap %>" 
                                    SelectCommand="SELECT DISTINCT * FROM [glo_specjalizacje] ORDER BY [nazwa]">
                                </asp:SqlDataSource>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
          
        </table>
    
   <table style=" width: 100%;">
                <tr>
                    <td style=" width: 100%; text-align:center">
                        &nbsp;<asp:LinkButton ID="LinkButton1" runat="server" CssClass="button_" 
                            onclick="LinkButton1_Click">Zapisz</asp:LinkButton>
                    </td>
                    <td style=" width: 100%; text-align :center">
                       
                        <asp:LinkButton ID="LinkButton2" runat="server" CssClass="button_" 
                            onclick="LinkButton2_Click" CausesValidation="False">Anuluj</asp:LinkButton>
                    </td>
                </tr>
                </table>
    

      
   

          </asp:Panel>


          </div>

    




</asp:Content>
