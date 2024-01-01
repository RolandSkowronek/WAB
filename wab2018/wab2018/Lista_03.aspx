<%@ Page Title="" Language="C#" MasterPageFile="~/Master.Master" AutoEventWireup="true" CodeBehind="Lista_03.aspx.cs" Inherits="wab2018.Lista_03" %>
<%@ Register src="cal1.ascx" tagname="cal1" tagprefix="uc1" %>
<%@ Register src="Lista_osob .ascx" tagname="Lista_osob" tagprefix="uc2" %>
<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="server">
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
      <div id ="mainWindow" class="newPage"> 
       <asp:Panel ID="lista_obciazen" runat="server" Width="100%">
        <h2>
            <asp:Label ID="Label1" runat="server" Text="Lista obciązeń"></asp:Label>
        </h2>

        
        <table style="width:100%;" cellpadding="0" cellspacing="0">
            <tr>
                <td>
                    <asp:LinkButton ID="LinkButton1" runat="server" CssClass="button_" 
                        Width="231px" onclick="LinkButton1_Click1">Dodaj nowe obciążenie</asp:LinkButton>
                </td>
                <td>
                    &nbsp;</td>
                <td>
                    &nbsp;</td>
            </tr>
        </table>
        <asp:Panel ID="Panel1" runat="server" Height="400px" ScrollBars="Auto" 
            Width="100%">
            <asp:GridView ID="GridView1" runat="server" AllowSorting="True" 
                AutoGenerateColumns="False" CellPadding="3" DataSourceID="lista" 
                ForeColor="#333333" GridLines="None" Width="90%" 
                onselectedindexchanged="GridView1_SelectedIndexChanged" DataKeyNames="id_,instytucja,sygnatura,data_wprowadzenia,data_otrzymania,termin,id_osoby,imie,nazwisko,data_zwrotu">
                <AlternatingRowStyle BackColor="White" />
                <Columns>
                    <asp:CommandField ShowSelectButton="True" ButtonType="Image" SelectImageUrl="~/img/edit.jpg" />
                    <asp:BoundField DataField="instytucja" HeaderText="Instytucja" 
                        SortExpression="instytucja" />
                    <asp:BoundField DataField="sygnatura" 
                        HeaderText="Sygnatura" SortExpression="sygnatura" />
                    <asp:BoundField DataField="data_wprowadzenia" DataFormatString="{0:d}" 
                        HeaderText="Data wprowadzenia" SortExpression="data_wprowadzenia" >
                    <ItemStyle HorizontalAlign="Center" />
                    </asp:BoundField>
                    <asp:BoundField DataField="data_otrzymania" DataFormatString="{0:d}" 
                        HeaderText="Data otrzymania" SortExpression="data_otrzymania" >
                    <ItemStyle HorizontalAlign="Center" />
                    </asp:BoundField>
                    <asp:BoundField DataField="data_zwrotu" DataFormatString="{0:d}" HeaderText="Data zwrotu" SortExpression="data_zwrotu" >
                    <HeaderStyle HorizontalAlign="Center" />
                    </asp:BoundField>
                    <asp:BoundField DataField="termin" DataFormatString="{0:d}" HeaderText="Termin" SortExpression="termin" >
                    <ItemStyle HorizontalAlign="Center" />
                    </asp:BoundField>
                    <asp:BoundField DataField="imie" HeaderText="Biegły" ReadOnly="True" SortExpression="imie" >
                    <ItemStyle HorizontalAlign="Center" />
                    </asp:BoundField>
                </Columns>
                <EditRowStyle BackColor="#2461BF" />
                <FooterStyle BackColor="#507CD1" Font-Bold="True" ForeColor="White" />
                <HeaderStyle BackColor="#4B6C9E" Font-Bold="True" ForeColor="White" />
                <PagerStyle BackColor="#2461BF" ForeColor="White" HorizontalAlign="Center" />
                <RowStyle BackColor="#EFF3FB" />
                <SelectedRowStyle BackColor="#D1DDF1" Font-Bold="True" ForeColor="#333333" />
                <SortedAscendingCellStyle BackColor="#F5F7FB" />
                <SortedAscendingHeaderStyle BackColor="#6D95E1" />
                <SortedDescendingCellStyle BackColor="#E9EBEF" />
                <SortedDescendingHeaderStyle BackColor="#4870BE" />
            </asp:GridView>
        </asp:Panel>
        <br />
        <asp:SqlDataSource ID="lista" runat="server" 
            ConnectionString="<%$ ConnectionStrings:wap %>" SelectCommand="SELECT DISTINCT tbl_obciazenia.id_, tbl_obciazenia.instytucja, tbl_obciazenia.sygnatura, tbl_obciazenia.data_wprowadzenia, tbl_obciazenia.data_otrzymania, tbl_obciazenia.termin, tbl_obciazenia.id_osoby, tbl_osoby.imie + ' ' + tbl_osoby.nazwisko AS imie, tbl_obciazenia.data_zwrotu, tbl_osoby.imie AS Expr1, tbl_osoby.nazwisko FROM tbl_obciazenia LEFT OUTER JOIN tbl_osoby ON tbl_obciazenia.id_osoby = tbl_osoby.ident ORDER BY tbl_obciazenia.data_wprowadzenia"></asp:SqlDataSource>
    </asp:Panel>



      <asp:Panel ID="Panel3" runat="server" CssClass="Panel_2_2" Visible="False">
        
        <asp:Panel ID="Panel4" runat="server" CssClass="Panel_1_1">
            
            <div style="background-color: #3399FF; width: 100%; top:0px;">
            <b>Edycja danych </b><br/>
            </div>
             <table style="width:100%;">
            <tr>
                <td style="vertical-align: top;">
                      <table style="width:100%;">
             <tr style="height:40px;">
                <td>
                    Instytucja</td>
                <td style="width: 33%">
                    <asp:TextBox ID="TextBox1" runat="server" Width="100%"></asp:TextBox>
                </td>
                <td colspan="2">
                    <asp:RequiredFieldValidator ID="RequiredFieldValidator1" runat="server" 
                        ControlToValidate="TextBox1" ErrorMessage="&lt;- Proszę uzupełnić pole"></asp:RequiredFieldValidator>
                </td>
            </tr>
             <tr style="height:40px;">
                <td>
                    Sygnatura</td>
                <td style="width: 33%">
                    <asp:TextBox ID="TextBox2" runat="server" Width="100%"></asp:TextBox>
                </td>
                <td>
                    <asp:RequiredFieldValidator ID="RequiredFieldValidator2" runat="server" 
                        ControlToValidate="TextBox2" ErrorMessage="&lt;- Proszę uzupełnić pole"></asp:RequiredFieldValidator>
                </td>
                 <td>&nbsp;</td>
            </tr>
         
            <tr style="height:40px;">
                <td>
                    Data wprowadzenia</td>
                <td style="width: 25%">
                    
                    <uc1:cal1 ID="cal13" runat="server" />
                    
                </td>
                <td>
                    Data Otrzymania</td>
                <td>
                    <uc1:cal1 ID="cal11" runat="server" />
                </td>
            </tr>
             <tr>
                   <td class="style1">
                    <asp:Label ID="Label5" runat="server">Termin:</asp:Label>
                </td>
                 <td class="style2"> 
                   <uc1:cal1 ID="cal1" runat="server" />

                 </td>

                <td class="style1">

                   Data zwrotu opinii

                </td>
                <td class="style2">
                  
                    
                     <uc1:cal1 ID="cal12" runat="server" />
                  
                </td>
              
            </tr>
                           <tr style="height:40px;">
                              <td>
                                  Osoba:</td>
                              <td colspan="3">
                                 <uc2:Lista_osob ID="Lista_osob1" runat="server" />
                                                             </td>
                          </tr>
        </table>
                    
                    </td>
            </tr>
          
        </table>
            <asp:Label ID="Label6" runat="server"></asp:Label>
            <br/>
  <table style="width: 100%;">
                <tr>
                    <td>
                        &nbsp;
                        <asp:LinkButton ID="LinkButton2" runat="server" CssClass="button_" 
                            onclick="LinkButton1_Click">Zapisz</asp:LinkButton>
                    </td>
                    <td>
                        &nbsp;
                        <asp:LinkButton ID="LinkButton3" runat="server" CssClass="button_" 
                            onclick="LinkButton2_Click" CausesValidation="False">Anuluj</asp:LinkButton>
                    </td>
                    <td>
                        <asp:LinkButton ID="LinkButton4" runat="server" CausesValidation="False" 
                            CssClass="button_" onclick="LinkButton3_Click">Usuń </asp:LinkButton>
                    </td>
                </tr>
                </table>

            <br />
        </asp:Panel>
        <br />
    </asp:Panel>
          </div>
    </asp:Content>

