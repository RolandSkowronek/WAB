<%@ Page Title="" Language="C#" MasterPageFile="~/Master.Master" AutoEventWireup="true" CodeBehind="wykazBieglych.aspx.cs" Inherits="wab2018.wykazBieglych" EnableEventValidation="false"  %>

<%@ Register assembly="DevExpress.Web.v17.1, Version=17.1.17.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.Web" tagprefix="dx" %>
<%@ Register assembly="DevExpress.Web.Bootstrap.v17.1, Version=17.1.17.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.Web.Bootstrap" tagprefix="dx" %>
<%@ Register src="userControls/daneStatystyczne.ascx" tagname="daneStatystyczne" tagprefix="uc1" %>
<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="server">
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
     <style type="text/css">
        .InfoTable td
        {
            padding: 0 4px;
            vertical-align: top;
        }
         .auto-style1 {
             width: 229px;
         }
                
         .auto-style2 {
             width: 258px;
         }
         .auto-style3 {
             width: 172px;
         }
                
         </style>
    <script type="text/javascript">
        var keyValue;
        function OnMoreInfoClick(element, key) {
            callbackPanel.SetContentHtml("");
            popup.ShowAtElement(element);
            
            keyValue = key;
        }
        function popup_Shown(s, e) {
            callbackPanel.PerformCallback(keyValue);
        }
    

        var keySpec;
        function OnAddClick(element, key) {
            callbackPanel.SetContentHtml("");
            alert(key);
            //popup.ShowAtElement(element);
            
            keySpec = key;
        }
        function popup_Shown(s, e) {
            callbackPanel.PerformCallback(keyValue);
        }
    </script>
  
        <script type="text/javascript">
        function UpdateDetailGrid(s, e) {
            detailGridView.PerformCallback(masterGridView.GetFocusedRowIndex());
        }
    </script>

<script src="https://code.jquery.com/jquery-1.12.4.min.js"></script>

     <div id ="mainWindow" class="newPage">  <h2>         Wykaz biegłych sądowych</h2>
      &nbsp;<table style="margin-left:auto;margin-right:auto;" class="dxflInternalEditorTable_Aqua">
        <tr>
            <td style="width: 300px">
              
                    
                <asp:CheckBox ID="CheckBox1" runat="server" Text="Wyświetl po specjalizacji" 
                    AutoPostBack="True" oncheckedchanged="uruchomFiltrowaniePoSpecjalizacji" TabIndex="1" meta:resourcekey="CheckBox1Resource1" />
            &nbsp;&nbsp;&nbsp;
                <asp:SqlDataSource ID="specjalizacje" runat="server" ConnectionString="<%$ ConnectionStrings:wap %>" SelectCommand="SELECT [id_], [nazwa] FROM [glo_specjalizacje] ORDER BY [nazwa]"></asp:SqlDataSource>
            </td>
            <td class="auto-style3">
              
                    
<asp:DropDownList ID="DropDownList1" runat="server" AutoPostBack="True" 
                    DataSourceID="specjalizacje" DataTextField="nazwa" DataValueField="id_" 
                     onselectedindexchanged="wyswietlPoSpecjalizacji" Enabled="False" meta:resourcekey="DropDownList1Resource1">
                </asp:DropDownList>
            </td>
            <td class="auto-style2">
              
                    
                <asp:CheckBox ID="CheckBox4" runat="server" Text="Archiwum" AutoPostBack="True" OnCheckedChanged="obsługaArchiwum" meta:resourcekey="CheckBox4Resource1" />
            </td>
           <td align="right" style="width: 410px" >
               <asp:LinkButton ID="LinkButton4" runat="server" Text="Drukuj" CssClass="button_" OnClick="print_" meta:resourcekey="LinkButton4Resource1"></asp:LinkButton>
               <asp:LinkButton ID="LinkButton14" runat="server" CssClass="button_" OnClick="twórzZestawienie" meta:resourcekey="LinkButton14Resource1" Text="zestawienie"></asp:LinkButton>
               <asp:LinkButton ID="LinkButton6" runat="server" Text="Excel" CssClass="button_" OnClick="makeExcell" meta:resourcekey="LinkButton6Resource1" ></asp:LinkButton>
               <asp:LinkButton ID="LinkButton15" runat="server" Text="Excel" CssClass="button_" OnClick="makeExcellforBIP" meta:resourcekey="LinkButton6Resource1" ></asp:LinkButton>
           </td>
         </tr>   

        </table>


         <br />
         <asp:Panel ID="normalny" runat="server">
             <dx:ASPxGridView ID="listaBieglych" runat="server" AutoGenerateColumns="False" DataSourceID="daneBieglych" Width="100%" OnBeforePerformDataSelect="listaBieglych_BeforePerformDataSelect">
                 <SettingsPager PageSize="50">
                 </SettingsPager>
                 <Settings ShowFilterRow="True" ShowFilterRowMenu="True" />
                 <SettingsBehavior AllowFocusedRow="True" />
                 <Columns>
                       <dx:GridViewDataTextColumn VisibleIndex="0" meta:resourcekey="GridViewDataTextColumnResource1" ShowInCustomizationForm="True">
                           <Settings AllowAutoFilter="True" />
                           <SettingsHeaderFilter>
                               <DateRangeCalendarSettings TodayButtonText="Dzisiaj" />
                           </SettingsHeaderFilter>
                           <DataItemTemplate>
                                <a href="javascript:void(0);" onclick="OnMoreInfoClick(this, '<%# Eval("ident") %>')">
                               <asp:Image ID="Image1" runat="server" ImageUrl="~/img/button_edycja.png" meta:resourcekey="Image1Resource1" />
                              </a>
                               </DataItemTemplate>
                    </dx:GridViewDataTextColumn>

                    
                     <dx:GridViewDataTextColumn Caption="Tytuł" FieldName="tytul" ShowInCustomizationForm="True" VisibleIndex="1">
                     </dx:GridViewDataTextColumn>
                     <dx:GridViewDataTextColumn Caption="Imie" FieldName="imie" ShowInCustomizationForm="True" VisibleIndex="2">
                     </dx:GridViewDataTextColumn>
                     <dx:GridViewDataTextColumn Caption="Nazwisko" FieldName="nazwisko" ShowInCustomizationForm="True" VisibleIndex="3">
                     </dx:GridViewDataTextColumn>
                     <dx:GridViewDataTextColumn Caption="Adres" FieldName="adres" ReadOnly="True" ShowInCustomizationForm="True" VisibleIndex="6">
                     </dx:GridViewDataTextColumn>
                     <dx:GridViewDataDateColumn Caption="Powołanie od" FieldName="data_poczatkowa" ShowInCustomizationForm="True" VisibleIndex="4">
                     </dx:GridViewDataDateColumn>
                     <dx:GridViewDataDateColumn Caption="Powołanie do" FieldName="data_koncowa" ShowInCustomizationForm="True" VisibleIndex="5">
                     </dx:GridViewDataDateColumn>
                     <dx:GridViewDataCheckColumn Caption="Z" FieldName="zawieszony" VisibleIndex="6" meta:resourcekey="GridViewDataCheckColumnResource1" ShowInCustomizationForm="True">
                         </dx:GridViewDataCheckColumn>
                     <dx:GridViewDataTextColumn Caption="Specjalizacje" FieldName="specjalizacjeWidok" ShowInCustomizationForm="True" VisibleIndex="16">
                     </dx:GridViewDataTextColumn>
                     <dx:GridViewDataTextColumn Caption="Uwagi" FieldName="uwagi" ShowInCustomizationForm="True" VisibleIndex="15">
                     </dx:GridViewDataTextColumn>
                     <dx:GridViewDataTextColumn Caption="Telefon" FieldName="telefon" ShowInCustomizationForm="True" VisibleIndex="7">
                     </dx:GridViewDataTextColumn>
                 </Columns>
             </dx:ASPxGridView>
             <br />
             <asp:SqlDataSource ID="daneBieglych" runat="server" ConnectionString="<%$ ConnectionStrings:wap %>" SelectCommand="SELECT DISTINCT View_listaBieglych.tytul, View_listaBieglych.imie, View_listaBieglych.uwagi, View_listaBieglych.nazwisko, View_listaBieglych.adres, View_listaBieglych.tel1 AS telefon, View_listaBieglych.specjalizacja_opis AS opis, View_listaBieglych.data_poczatkowa, View_listaBieglych.data_koncowa, View_listaBieglych.ident, View_listaBieglych.adres2, View_listaBieglych.zawieszony, View_listaBieglych.specjalizacjeWidok FROM View_listaBieglych LEFT OUTER JOIN tbl_specjalizacje_osob ON View_listaBieglych.ident = tbl_specjalizacje_osob.id_osoby WHERE (View_listaBieglych.data_koncowa &gt;= GETDATE()) AND (View_listaBieglych.typ &lt; 2) ORDER BY View_listaBieglych.nazwisko"></asp:SqlDataSource>
             <dx:ASPxGridViewExporter ID="ASPxGridViewExporter1" runat="server" GridViewID="listaBieglych" OnRenderBrick="ASPxGridViewExporter1_RenderBrick">
             </dx:ASPxGridViewExporter>
             <br />
             <br />
         </asp:Panel>
         <br />
         <asp:Panel ID="archiwum" runat="server">

             &nbsp;<dx:ASPxGridView ID="listaBieglych0" runat="server" AutoGenerateColumns="False" DataSourceID="daneBieglychArchiwum" Width="100%">
                 <SettingsPager PageSize="50">
                 </SettingsPager>
                 <Settings ShowFilterRow="True" ShowFilterRowMenu="True" />
                 <Columns>
                     <dx:GridViewDataTextColumn meta:resourcekey="GridViewDataTextColumnResource1" ShowInCustomizationForm="True" VisibleIndex="0">
                         <Settings AllowAutoFilter="True" />
                         <SettingsHeaderFilter>
                             <DateRangeCalendarSettings TodayButtonText="Dzisiaj" />
                         </SettingsHeaderFilter>
                         <DataItemTemplate>
                             <a href="javascript:void(0);" onclick='OnMoreInfoClick(this, &#039;<%# Eval("ident") %>&#039;)'>
                             <asp:Image ID="Image2" runat="server" ImageUrl="~/img/button_edycja.png" meta:resourcekey="Image1Resource1" />
                             </a>
                         </DataItemTemplate>
                     </dx:GridViewDataTextColumn>
                     <dx:GridViewDataTextColumn Caption="Tytuł" FieldName="tytul" ShowInCustomizationForm="True" VisibleIndex="1">
                     </dx:GridViewDataTextColumn>
                     <dx:GridViewDataTextColumn Caption="Imie" FieldName="imie" ShowInCustomizationForm="True" VisibleIndex="2">
                     </dx:GridViewDataTextColumn>
                     <dx:GridViewDataTextColumn Caption="Nazwisko" FieldName="nazwisko" ShowInCustomizationForm="True" VisibleIndex="3">
                     </dx:GridViewDataTextColumn>
                     <dx:GridViewDataTextColumn Caption="Adres" FieldName="adres" ReadOnly="True" ShowInCustomizationForm="True" VisibleIndex="6">
                     </dx:GridViewDataTextColumn>
                     <dx:GridViewDataDateColumn Caption="Powołanie od" FieldName="data_poczatkowa" ShowInCustomizationForm="True" VisibleIndex="4">
                     </dx:GridViewDataDateColumn>
                     <dx:GridViewDataDateColumn Caption="Powołanie do" FieldName="data_koncowa" ShowInCustomizationForm="True" VisibleIndex="5">
                     </dx:GridViewDataDateColumn>
                     <dx:GridViewDataCheckColumn Caption="Z" FieldName="zawieszony" meta:resourcekey="GridViewDataCheckColumnResource1" ShowInCustomizationForm="True" VisibleIndex="6">
                     </dx:GridViewDataCheckColumn>
                     <dx:GridViewDataTextColumn Caption="Specjalizacje" FieldName="specjalizacjeWidok" ShowInCustomizationForm="True" VisibleIndex="16">
                     </dx:GridViewDataTextColumn>
                     <dx:GridViewDataTextColumn Caption="Uwagi" FieldName="uwagi" ShowInCustomizationForm="True" VisibleIndex="15">
                     </dx:GridViewDataTextColumn>
                     <dx:GridViewDataTextColumn Caption="Telefon" FieldName="telefon" ShowInCustomizationForm="True" VisibleIndex="7">
                     </dx:GridViewDataTextColumn>
                 </Columns>
             </dx:ASPxGridView>
             <br />
             <asp:SqlDataSource ID="daneBieglychArchiwum" runat="server" ConnectionString="<%$ ConnectionStrings:wap %>" SelectCommand="SELECT DISTINCT View_listaBieglych.tytul as tytul, View_listaBieglych.imie as imie,  View_listaBieglych.uwagi as uwagi, View_listaBieglych.nazwisko as nazwisko, View_listaBieglych.adres as adres,View_listaBieglych.tel1 as telefon, View_listaBieglych.specjalizacja_opis as opis , View_listaBieglych.data_poczatkowa as data_poczatkowa, View_listaBieglych.data_koncowa as data_koncowa,View_listaBieglych.ident as ident, View_listaBieglych.adres2 as adres2, View_listaBieglych.zawieszony as zawieszony, View_listaBieglych.specjalizacjeWidok as specjalizacjeWidok FROM  View_listaBieglych LEFT OUTER JOIN tbl_specjalizacje_osob ON View_listaBieglych.ident = tbl_specjalizacje_osob.id_osoby where  View_listaBieglych.data_koncowa &lt; getdate() order by nazwisko"></asp:SqlDataSource>
         </asp:Panel>


        <dx:ASPxPopupControl ID="popup" ClientInstanceName="popup" runat="server" AllowDragging="True"
        PopupHorizontalAlign="OutsideRight" HeaderText="Dane biegłego" Theme="Aqua" MinHeight="300px" MinWidth="550px" OnWindowCallback ="popup_WindowCallback" CloseAction="CloseButton" OnPopupElementResolve="popup_PopupElementResolve" OnUnload="popup_Unload" meta:resourcekey="popupResource1">
        <ContentCollection>

             
            <dx:PopupControlContentControl runat="server" meta:resourcekey="PopupControlContentControlResource1" >
                
 <dx:ASPxCallbackPanel ID="callbackPanel" ClientInstanceName="callbackPanel" runat="server"
                    Width="450px" Height="120px" OnCallback="callbackPanel_Callback" RenderMode="Table" meta:resourcekey="callbackPanelResource1">
                    <PanelCollection>
                        <dx:PanelContent runat="server" meta:resourcekey="PanelContentResource1">
                            <asp:Panel ID="Panel1" runat="server" Visible="False" meta:resourcekey="Panel1Resource1">

                                  <dx:ASPxPageControl ID="carTabPage" runat="server" ActiveTabIndex="0" EnableHierarchyRecreation="True" Theme="Office2010Blue" meta:resourcekey="carTabPageResource1" >
            <TabPages >
                  <dx:TabPage Text="Dane osobowe" meta:resourcekey="TabPageResource1">
                       <ContentCollection>
                        <dx:ContentControl ID="ContentControl6" runat="server" meta:resourcekey="ContentControl6Resource1">



                            <asp:SqlDataSource ID="daneBieglego" runat="server" ConnectionString="<%$ ConnectionStrings:wap %>" SelectCommand="SELECT DISTINCT ident, imie, nazwisko, pesel, tytul, data_poczatkowa, data_koncowa, czy_zaw,specjalizacja_opis FROM tbl_osoby WHERE (ident = @idBieglego)">
                                <SelectParameters>
                                    <asp:SessionParameter Name="idBieglego" SessionField="id_osoby" />
                                </SelectParameters>
                            </asp:SqlDataSource>




                            <dx:ASPxVerticalGrid ID="ASPxVerticalGrid1" runat="server" AutoGenerateRows="False" DataSourceID="daneBieglego" EnableTheming="True" KeyFieldName="ident" Theme="Office2010Blue" Width="700px" meta:resourcekey="ASPxVerticalGrid1Resource1">
                                <Rows>
                                    <dx:VerticalGridTextRow Caption="Imię" FieldName="imie" VisibleIndex="1" meta:resourcekey="VerticalGridTextRowResource1">
                                    </dx:VerticalGridTextRow>
                                    <dx:VerticalGridTextRow Caption="Nazwisko" FieldName="nazwisko" VisibleIndex="2" meta:resourcekey="VerticalGridTextRowResource2">
                                    </dx:VerticalGridTextRow>
                                    <dx:VerticalGridTextRow Caption="PESEL" FieldName="pesel" ReadOnly="True" VisibleIndex="3" meta:resourcekey="VerticalGridTextRowResource3">
                                    </dx:VerticalGridTextRow>
                                    <dx:VerticalGridTextRow Caption="Tytuł" FieldName="tytul" VisibleIndex="0" meta:resourcekey="VerticalGridTextRowResource4">
                                    </dx:VerticalGridTextRow>
                                    <dx:VerticalGridDateRow Caption="Powołanie od:" FieldName="data_poczatkowa" UnboundType="DateTime" VisibleIndex="4" meta:resourcekey="VerticalGridDateRowResource1">
                                    </dx:VerticalGridDateRow>
                                    <dx:VerticalGridDateRow Caption="Powołanie do:" FieldName="data_koncowa" UnboundType="DateTime" VisibleIndex="5" meta:resourcekey="VerticalGridDateRowResource2">
                                    </dx:VerticalGridDateRow>
                                    <dx:VerticalGridCheckRow Caption="Zawieszenie" FieldName="czy_zaw" VisibleIndex="6" meta:resourcekey="VerticalGridCheckRowResource1">
                                    </dx:VerticalGridCheckRow>
                                    <dx:VerticalGridTextRow Caption="Specjalizacja opis" FieldName="specjalizacja_opis" Height="2" VisibleIndex="7" meta:resourcekey="VerticalGridTextRowResource5">
                                    </dx:VerticalGridTextRow>
                                </Rows>
                                <Settings ShowCategoryIndents="False" />
                                <SettingsPager Mode="ShowPager" Visible="False">
                                </SettingsPager>
                                <SettingsDataSecurity AllowDelete="False" AllowEdit="False" AllowInsert="False" />
                            </dx:ASPxVerticalGrid>

                          </dx:ContentControl>
                            </ContentCollection>
                      </dx:TabPage>

                <dx:TabPage Text="Dane kontaktowe" meta:resourcekey="TabPageResource2">
                    <ContentCollection>
                        <dx:ContentControl ID="ContentControl1" runat="server" meta:resourcekey="ContentControl1Resource1">


                                   <asp:SqlDataSource ID="DaneKontaktowe" runat="server" ConnectionString="<%$ ConnectionStrings:wap %>" SelectCommand="SELECT distinct ulica as Adres, kod_poczt as [kod pocztowy], miejscowosc as [Miejscowość], tel1, tel2, email, adr_kores, kod_poczt_kor, miejscowosc_kor, ident FROM tbl_osoby WHERE (ident = @idBieglego)">
                                       <SelectParameters>
                                           <asp:SessionParameter Name="idBieglego" SessionField="id_osoby" />
                                       </SelectParameters>
                                   </asp:SqlDataSource>
                                <dx:ASPxVerticalGrid ID="ASPxVerticalGrid2" runat="server" AutoGenerateRows="False" DataSourceID="DaneKontaktowe" EnableTheming="True" Theme="Office2010Blue" Width="700px" meta:resourcekey="ASPxVerticalGrid2Resource1">
                                    <Rows>
                                        <dx:VerticalGridTextRow FieldName="Adres" VisibleIndex="0" meta:resourcekey="VerticalGridTextRowResource6">
                                        </dx:VerticalGridTextRow>
                                        <dx:VerticalGridTextRow Caption="Kod pocztowy" FieldName="kod pocztowy" VisibleIndex="1" meta:resourcekey="VerticalGridTextRowResource7">
                                        </dx:VerticalGridTextRow>
                                        <dx:VerticalGridTextRow FieldName="Miejscowość" VisibleIndex="2" meta:resourcekey="VerticalGridTextRowResource8">
                                        </dx:VerticalGridTextRow>
                                        <dx:VerticalGridTextRow Caption="Telefon" FieldName="tel1" VisibleIndex="3" meta:resourcekey="VerticalGridTextRowResource9">
                                        </dx:VerticalGridTextRow>
                                        <dx:VerticalGridTextRow Caption="Telefon" FieldName="tel2" VisibleIndex="4" meta:resourcekey="VerticalGridTextRowResource10">
                                        </dx:VerticalGridTextRow>
                                        <dx:VerticalGridTextRow Caption="Email" FieldName="email" VisibleIndex="5" meta:resourcekey="VerticalGridTextRowResource11">
                                        </dx:VerticalGridTextRow>
                                        <dx:VerticalGridTextRow Caption="Adres korespondecyjny" FieldName="adr_kores" VisibleIndex="7" meta:resourcekey="VerticalGridTextRowResource12">
                                        </dx:VerticalGridTextRow>
                                        <dx:VerticalGridTextRow Caption="Kod pocztowy" FieldName="kod_poczt_kor" VisibleIndex="8" meta:resourcekey="VerticalGridTextRowResource13">
                                        </dx:VerticalGridTextRow>
                                        <dx:VerticalGridTextRow Caption="Miejscowość" FieldName="miejscowosc_kor" VisibleIndex="9" meta:resourcekey="VerticalGridTextRowResource14">
                                        </dx:VerticalGridTextRow>
                                        <dx:VerticalGridTextRow Caption="Dane korespondencyjne" VisibleIndex="6" meta:resourcekey="VerticalGridTextRowResource15">
                                            <PropertiesTextEdit>
                                                <Style Font-Bold="True">
                                                </Style>
                                            </PropertiesTextEdit>
                                            <HeaderStyle Font-Bold="True" />
                                        </dx:VerticalGridTextRow>
                                    </Rows>
                                    <Settings ShowCategoryIndents="False" />
                                    <SettingsPager Mode="ShowPager" Visible="False">
                                    </SettingsPager>
                                    <SettingsDataSecurity AllowDelete="False" AllowEdit="False" AllowInsert="False" />
                                </dx:ASPxVerticalGrid>
                        </dx:ContentControl>
                    </ContentCollection>
                </dx:TabPage>
                <dx:TabPage Text="Specjalizacje" meta:resourcekey="TabPageResource3">
                    <ContentCollection>
                        <dx:ContentControl ID="ContentControl2" runat="server" meta:resourcekey="ContentControl2Resource1">
                     <asp:SqlDataSource ID="SqlDataSource2" runat="server" ConnectionString="<%$ ConnectionStrings:wap %>" SelectCommand="SELECT nazwa FROM View_pozycje_specjalizacji WHERE (id_osoby = @id_biegłego) ORDER BY nazwa">
                         <SelectParameters>
                             <asp:SessionParameter Name="id_biegłego" SessionField="id_osoby" />
                         </SelectParameters>
                     </asp:SqlDataSource>
                              <dx:ASPxGridView ID="ASPxGridView3" runat="server" AutoGenerateColumns="False" DataSourceID="SqlDataSource2" Theme="Office2010Blue" Width="100%" meta:resourcekey="ASPxGridView3Resource1">
                                  <SettingsCommandButton>
                                      <ShowAdaptiveDetailButton ButtonType="Image">
                                      </ShowAdaptiveDetailButton>
                                      <HideAdaptiveDetailButton ButtonType="Image">
                                      </HideAdaptiveDetailButton>
                                  </SettingsCommandButton>
                                  <SettingsDataSecurity AllowDelete="False" AllowEdit="False" AllowInsert="False" />
                                  <Columns>
                                      <dx:GridViewDataTextColumn Caption="Specjalizacje biegłego" FieldName="nazwa" ShowInCustomizationForm="True" VisibleIndex="0" meta:resourcekey="GridViewDataTextColumnResource13">
                                          <PropertiesTextEdit NullText="Brak danych">
                                          </PropertiesTextEdit>
                                      </dx:GridViewDataTextColumn>
                                  </Columns>
                              </dx:ASPxGridView>
                        </dx:ContentControl>
                    </ContentCollection>
                </dx:TabPage>
                <dx:TabPage Text="uwagi" meta:resourcekey="TabPageResource4">
                    <ContentCollection>
                        <dx:ContentControl ID="ContentControl3" runat="server" meta:resourcekey="ContentControl3Resource1">
                            <div style="min-height:275px;"><table style="width: 100%;"><tr><td>Uwagi</td></tr><tr><td><asp:TextBox ID="TextBox9" runat="server" TextMode="MultiLine" Width="100%" Height="100%" Rows="12" meta:resourcekey="TextBox9Resource1"></asp:TextBox></td></tr></table></div>
                        </dx:ContentControl>
                    </ContentCollection>
                </dx:TabPage>
                <dx:TabPage Text="Dane statystyczne" meta:resourcekey="TabPageResource5">
                    <ContentCollection>
                        <dx:ContentControl ID="ContentControl4" runat="server" meta:resourcekey="ContentControl4Resource1">
                            <div style="min-height:275px;"><asp:DropDownList ID="DropDownList3" runat="server" AutoPostBack="True" OnSelectedIndexChanged="changeQuerry" DataSourceID="kwerendyStatystyczne0" DataTextField="Nazwa" DataValueField="kwerenda" meta:resourcekey="DropDownList3Resource1"></asp:DropDownList>
                    <asp:GridView ID="GridView1" runat="server" Css meta:resourcekey="GridView1Resource1"></asp:GridView>
                                <asp:SqlDataSource ID="kwerendyStatystyczne0" runat="server" ConnectionString="<%$ ConnectionStrings:wap %>" SelectCommand="SELECT Nazwa, kwerenda, id_ FROM dane_statystyczne WHERE (czy_us &lt;&gt; 1)  ORDER BY Nazwa">
                                </asp:SqlDataSource>
                            </div>
                        </dx:ContentControl>
                    </ContentCollection>
                </dx:TabPage>
                  <dx:TabPage Text="Historia powołań" meta:resourcekey="TabPageResource6">
                    <ContentCollection>
                        <dx:ContentControl ID="ContentControl5" runat="server" meta:resourcekey="ContentControl5Resource1">
                              <div style="min-height:275px;">
                                  <table style="width: 100%;">
                                      <tr>
                                          <td style="width: 60%" valign="top"><asp:Panel ID="Panel8" runat="server" Height="200px" ScrollBars="Vertical" meta:resourcekey="Panel8Resource1">
                                              <asp:SqlDataSource ID="SqlDataSource7" runat="server" ConnectionString="<%$ ConnectionStrings:wap %>"  SelectCommand="SELECT DISTINCT View_powolania.ident, View_powolania.id_powolania, tbl_powolania.data_od, tbl_powolania.data_do, tbl_powolania.czyus, tbl_powolania.id_bieglego FROM            View_powolania INNER JOIN      tbl_powolania ON View_powolania.ident = tbl_powolania.ident GROUP BY View_powolania.ident, View_powolania.id_powolania, tbl_powolania.data_od, tbl_powolania.data_do, tbl_powolania.czyus, tbl_powolania.id_bieglego HAVING        (tbl_powolania.czyus = 0) AND (tbl_powolania.id_bieglego = @id_bieglego)">
                                                  <SelectParameters>
                                                      <asp:SessionParameter Name="id_bieglego" SessionField="id_osoby" Type="Int32" />
                                                  </SelectParameters>
                                              </asp:SqlDataSource>
                                              <dx:ASPxGridView ID="ASPxGridView4" runat="server" AutoGenerateColumns="False" DataSourceID="SqlDataSource7" EnableTheming="True" KeyFieldName="data_od" Theme="Office2010Blue" Width="100%" meta:resourcekey="ASPxGridView4Resource1">
                                                  <SettingsCommandButton>
                                                      <ShowAdaptiveDetailButton ButtonType="Image">
                                                      </ShowAdaptiveDetailButton>
                                                      <HideAdaptiveDetailButton ButtonType="Image">
                                                      </HideAdaptiveDetailButton>
                                                  </SettingsCommandButton>
                                                  <Columns>
                                                      <dx:GridViewDataDateColumn Caption="Powołanie od:" FieldName="data_od" ReadOnly="True" ShowInCustomizationForm="True" VisibleIndex="2" meta:resourcekey="GridViewDataDateColumnResource5">
                                                      </dx:GridViewDataDateColumn>
                                                      <dx:GridViewDataDateColumn Caption="Powołanie do:" FieldName="data_do" ReadOnly="True" ShowInCustomizationForm="True" UnboundType="DateTime" VisibleIndex="3" meta:resourcekey="GridViewDataDateColumnResource6">
                                                      </dx:GridViewDataDateColumn>
                                                  </Columns>
                                              </dx:ASPxGridView>
                                              </asp:Panel></td></tr></table></div>
                        </dx:ContentControl>
                    </ContentCollection>
                </dx:TabPage>
                  <dx:TabPage Text="Skargi" meta:resourcekey="TabPageResource7">
                    <ContentCollection>
                        <dx:ContentControl ID="ContentControl13" runat="server" meta:resourcekey="ContentControl13Resource1">
                              <div style="min-height:275px;">
                                  <table style="width: 100%;">
                                      <tr>
                                          <td style="width: 60%" valign="top"><asp:Panel ID="Panel5" runat="server" Height="200px" ScrollBars="Vertical" meta:resourcekey="Panel5Resource1">
                                              <dx:ASPxGridView ID="ASPxGridView2" runat="server" AutoGenerateColumns="False" DataSourceID="SqlDataSkargi" EnableTheming="True" Theme="Office2010Blue" Width="100%" meta:resourcekey="ASPxGridView1Resource1">
                                                  <SettingsCommandButton>
                                                      <ShowAdaptiveDetailButton ButtonType="Image">
                                                      </ShowAdaptiveDetailButton>
                                                      <HideAdaptiveDetailButton ButtonType="Image">
                                                      </HideAdaptiveDetailButton>
                                                  </SettingsCommandButton>
                                                  <Columns>
                                                      <dx:GridViewDataTextColumn Caption="Numer" FieldName="numer" ShowInCustomizationForm="True" VisibleIndex="0" meta:resourcekey="GridViewDataTextColumnResource14">
                                                      </dx:GridViewDataTextColumn>
                                                      <dx:GridViewDataTextColumn Caption="Rok" FieldName="rok" ShowInCustomizationForm="True" VisibleIndex="1" meta:resourcekey="GridViewDataTextColumnResource15">
                                                      </dx:GridViewDataTextColumn>
                                                      <dx:GridViewDataDateColumn Caption="Data wpływu" FieldName="dataWplywu" ShowInCustomizationForm="True" VisibleIndex="2" meta:resourcekey="GridViewDataDateColumnResource7">
                                                      </dx:GridViewDataDateColumn>
                                                      <dx:GridViewDataDateColumn Caption="Data pisma" FieldName="dataPisma" ShowInCustomizationForm="True" VisibleIndex="3" meta:resourcekey="GridViewDataDateColumnResource8">
                                                      </dx:GridViewDataDateColumn>
                                                      <dx:GridViewDataTextColumn Caption="Sygnatura" FieldName="Sygnatura" ShowInCustomizationForm="True" VisibleIndex="4" meta:resourcekey="GridViewDataTextColumnResource16">
                                                      </dx:GridViewDataTextColumn>
                                                      <dx:GridViewDataTextColumn Caption="Wizytator" FieldName="wizytator" ShowInCustomizationForm="True" VisibleIndex="5" meta:resourcekey="GridViewDataTextColumnResource17">
                                                      </dx:GridViewDataTextColumn>
                                                      <dx:GridViewDataCheckColumn Caption="Zakreślono" FieldName="zakreslono" ShowInCustomizationForm="True" VisibleIndex="6" meta:resourcekey="GridViewDataCheckColumnResource3">
                                                      </dx:GridViewDataCheckColumn>
                                                      <dx:GridViewDataTextColumn Caption="Uwagi" FieldName="uwagi" ShowInCustomizationForm="True" VisibleIndex="8" meta:resourcekey="GridViewDataTextColumnResource18">
                                                      </dx:GridViewDataTextColumn>
                                                  </Columns>
                                              </dx:ASPxGridView>
                                              <asp:SqlDataSource ID="SqlDataSkargi" runat="server" ConnectionString="<%$ ConnectionStrings:wap %>"  SelectCommand="SELECT ident, numer, rok, dataWplywu, dataPisma, Sygnatura, wizytator, zakreslono, dataZakreslenia, uwagi FROM tbl_skargi WHERE (idBieglego = @idBieglego) AND (czyUs = 0)">
                                                  <SelectParameters>
                                                      <asp:SessionParameter Name="idBieglego" SessionField="id_osoby" DefaultValue="" />
                                                  </SelectParameters>
                                              </asp:SqlDataSource>
                                              </asp:Panel></td></tr></table></div>
                        </dx:ContentControl>
                    </ContentCollection>
                </dx:TabPage>
            </TabPages>
        </dx:ASPxPageControl>
</dx:ASPxTabControl>
                            </asp:Panel>
                         
                              <asp:Panel ID="Panel2" runat="server" Visible="False" meta:resourcekey="Panel2Resource1">

      <dx:ASPxPageControl ID="ASPxPageControl1" runat="server" ActiveTabIndex="0" EnableHierarchyRecreation="True" Theme="Office2010Blue"  meta:resourcekey="ASPxPageControl1Resource1" TabIndex="1">
            <TabPages >
                      <dx:TabPage Text="Dane osobowe" meta:resourcekey="TabPageResource8">
                       <ContentCollection>
                        <dx:ContentControl ID="ContentControl7" runat="server" meta:resourcekey="ContentControl7Resource1">




                            <table style="width:100%;">
                                <tr>
                                    <td>Tytuł</td>
                                    <td>
                                        <asp:TextBox ID="TxBTytul" runat="server" meta:resourcekey="TxBTytulResource1" ></asp:TextBox>
                                    </td>
                                    <td style="width: 50%">Specjalizacja&nbsp; opis (doprecyzowanie) </td>
                                </tr>
                                <tr>
                                    <td>Imie</td>
                                    <td>
                                        <asp:TextBox ID="TxbImie" runat="server" meta:resourcekey="TxbImieResource1"></asp:TextBox>
                                    </td>
                                    <td align="center" rowspan="4" valign="top">
                                        <asp:TextBox ID="txspecjalizacja_opis" runat="server" Height="99%" TextMode="MultiLine" Width="99%" meta:resourcekey="txspecjalizacja_opisResource1"></asp:TextBox>
                                    </td>
                                </tr>
                                <tr>
                                    <td>Nazwisko</td>
                                    <td>
                                        <asp:TextBox ID="TxBNazwisko" runat="server" meta:resourcekey="TxBNazwiskoResource1"></asp:TextBox>
                                    </td>
                                </tr>
                                <tr>
                                    <td>PESEL</td>
                                    <td>
                                        <asp:TextBox ID="TxbPesel" runat="server" meta:resourcekey="TxbPeselResource1"></asp:TextBox>
                                    </td>
                                </tr>
                                <tr>
                                    <td>Data powołania od</td>
                                    <td>
                                        <dx:ASPxDateEdit ID="ASPxDateEdit3" runat="server"  meta:resourcekey="ASPxDateEdit3Resource1">
                                        </dx:ASPxDateEdit>
                                    </td>
                                </tr>
                                <tr>
                                    <td>Data powołania do</td>
                                    <td>
                                        <dx:ASPxDateEdit ID="ASPxDateEdit4" runat="server" meta:resourcekey="ASPxDateEdit4Resource1">
                                        </dx:ASPxDateEdit>
                                    </td>
                                    <td style="width: 50%">&nbsp;</td>
                                </tr>
                                <tr>
                                    <td>
                                        <asp:DropDownList ID="ddlZawiszenie" runat="server" AutoPostBack="True" OnSelectedIndexChanged="zmianaZawieszenia" meta:resourcekey="DropDownList4Resource1">
                                            <asp:ListItem Value="0" meta:resourcekey="ListItemResource1">Brak zawieszenia</asp:ListItem>
                                            <asp:ListItem Value="1" meta:resourcekey="ListItemResource2">Zawieszenie</asp:ListItem>
                                        </asp:DropDownList>
                                    </td>
                                    <td>
                                        <asp:Label ID="lblPoczatekZawieszenia" runat="server" Text="Początek zawieszenia" Visible="False"></asp:Label>
                                    </td>
                                    <td style="width: 50%">
                                        <dx:ASPxDateEdit ID="poczatekZawieszeniaData" runat="server" meta:resourceKey="zawieszenieDataResource1" OnDateChanged="ASPxDateEdit7_DateChanged">
                                        </dx:ASPxDateEdit>
                                    </td>
                                </tr>
                                <tr>
                                    <td>&nbsp;</td>
                                    <td>
                                        <asp:Label ID="lblKoniecZawieszenia" runat="server" Text="Koniec zawieszenia" Visible="False"></asp:Label>
                                    </td>
                                    <td style="width: 50%">
                                        <dx:ASPxDateEdit ID="koniecZawieszeniaData" runat="server" meta:resourceKey="zawieszenieDataResource1" OnDateChanged="ASPxDateEdit7_DateChanged" Visible="False">
                                        </dx:ASPxDateEdit>
                                    </td>
                                </tr>
                            </table>
                            <br />
                            <asp:SqlDataSource ID="SqlDataSource3" runat="server" ConnectionString="<%$ ConnectionStrings:wap %>" SelectCommand="SELECT DISTINCT ident, imie, nazwisko, pesel, tytul, data_poczatkowa, data_koncowa FROM tbl_osoby WHERE (ident = @idBieglego)">
                                <SelectParameters>
                                    <asp:SessionParameter Name="idBieglego" SessionField="id_osoby" />
                                </SelectParameters>
                            </asp:SqlDataSource>


                            </dx:ContentControl>
                            </ContentCollection>
                      </dx:TabPage>

                      <dx:TabPage Text="Dane kontaktowe" meta:resourcekey="TabPageResource9">
                    <ContentCollection>
                        <dx:ContentControl ID="ContentControl8" runat="server" meta:resourcekey="ContentControl8Resource1">
                                <table class="dxeBinImgCPnlSys">
                                <tr>
                                    <td class="auto-style1">Adres</td>
                                    <td>
                                        <asp:TextBox runat="server" ID="TxAdres1" meta:resourcekey="TxAdres1Resource1"></asp:TextBox>

                                    </td>
                                    <td>&nbsp;</td>
                                </tr>
                                <tr>
                                    <td class="auto-style1">Kod pocztowy</td>
                                    <td>
                                        <asp:TextBox ID="Txkod1" runat="server" meta:resourcekey="Txkod1Resource1"></asp:TextBox>
                                    </td>
                                    <td>&nbsp;</td>
                                </tr>
                                <tr>
                                    <td class="auto-style1">Miejscowość</td>
                                    <td>
                                        <asp:TextBox ID="TxMiejscowosc1" runat="server" meta:resourcekey="TxMiejscowosc1Resource1"></asp:TextBox>
                                    </td>
                                    <td>&nbsp;</td>
                                </tr>
                                    <tr>
                                        <td >Telefon</td>
                                        <td >
                                            <asp:TextBox ID="TxTelefon1" runat="server" meta:resourcekey="TxTelefon1Resource1"></asp:TextBox>
                                        </td>
                                        <td >&nbsp;</td>
                                    </tr>
                                <tr>
                                    <td >Telefon</td>
                                    <td >
                                        <asp:TextBox ID="TxTelefon2" runat="server" meta:resourcekey="TxTelefon2Resource1"></asp:TextBox>

                                    </td>
                                    <td >&nbsp;</td>
                                </tr>
                                <tr>
                                    <td class="auto-style1">email</td>
                                    <td>
                                        <asp:TextBox ID="TxEmail" runat="server" meta:resourcekey="TxEmailResource1"></asp:TextBox>
                                    </td>
                                    <td>&nbsp;</td>
                                    <td class="auto-style1">&nbsp;</td>
                                    <td>
                                        &nbsp;</td>
                                    <td>&nbsp;</td>
                                    <td >&nbsp;</td>
                                    <td >
                                        &nbsp;</td>
                                    <td ></td>
                                    <td class="auto-style1">&nbsp;</td>
                                    <td>
                                        &nbsp;</td>
                                    <td>&nbsp;</td>
                                </tr>
                                    <tr>
                                        <td >Dane korespondencyjne</td>
                                        <td>
                                            &nbsp;</td>
                                        <td>&nbsp;</td>
                                    </tr>
                                    <tr>
                                        <td >Adres korespondencyjny</td>
                                        <td>
                                            <asp:TextBox ID="TxAdres2" runat="server" meta:resourcekey="TxAdres2Resource1"></asp:TextBox>
                                        </td>
                                        <td>&nbsp;</td>
                                    </tr>
                                    <tr>
                                        <td class="auto-style1">Kod pocztowy</td>
                                        <td>
                                            <asp:TextBox ID="TxKod2" runat="server" meta:resourcekey="TxKod2Resource1"></asp:TextBox>
                                        </td>
                                        <td>&nbsp;</td>
                                    </tr>
                                    <tr>
                                        <td class="auto-style1">Miejscowość</td>
                                        <td>
                                            <asp:TextBox ID="Txmiejscowosc2" runat="server" meta:resourcekey="Txmiejscowosc2Resource1"></asp:TextBox>
                                        </td>
                                        <td>&nbsp;</td>
                                    </tr>
                            </table>
                       

                                   <br />
                        </dx:ContentControl>
                    </ContentCollection>
                </dx:TabPage>
                      <dx:TabPage Text="Specjalizacje" meta:resourcekey="TabPageResource10">
                    <ContentCollection>
                        <dx:ContentControl ID="ContentControl9" runat="server" meta:resourcekey="ContentControl9Resource1">
                             <table style="width: 100%;">
                                 <tr>
                                     <td style="vertical-align: top;width:50%" width="100%">
                                         Specjalizacje biegłego</td><td style="vertical-align: top;width:50%">
                                             Dostępne specjalizacje</td></tr>
                                 <tr>
                                     <td style="vertical-align: top;width:50%" width="100%">
                                         <asp:Panel ID="Panel4" runat="server" Height="240px" ScrollBars="Vertical" meta:resourcekey="Panel4Resource1">

                                             <dx:ASPxGridView ID="ASPxGridView6" runat="server" AutoGenerateColumns="False" DataSourceID="Specjalizacje_temp" OnCustomButtonCallback="usunSpecjalizacje" Width="100%" EnableCallBacks="False" meta:resourcekey="ASPxGridView6Resource1" >

                                                 <ClientSideEvents FocusedRowChanged="UpdateDetailGrid" />
                                                  <SettingsPager PageSize="8">
                                                 </SettingsPager>
                                                 <Settings ShowFilterRow="True" />
                                                  <SettingsCommandButton>
                                                     <ShowAdaptiveDetailButton ButtonType="Image">
                                                     </ShowAdaptiveDetailButton>
                                                     <HideAdaptiveDetailButton ButtonType="Image">
                                                     </HideAdaptiveDetailButton>
                                                 </SettingsCommandButton>
                                                 <SettingsDataSecurity AllowDelete="False" AllowEdit="False" AllowInsert="False" />
                                                 <Columns>
                                                     <dx:GridViewDataTextColumn FieldName="nazwa" ShowInCustomizationForm="True" VisibleIndex="2" meta:resourcekey="GridViewDataTextColumnResource19">
                                                     </dx:GridViewDataTextColumn>
                                                        <dx:GridViewCommandColumn ShowNewButton="true" VisibleIndex="0" ButtonRenderMode="Image" meta:resourcekey="GridViewCommandColumnResource1">
                                                          <CustomButtons>
                    <dx:GridViewCommandColumnCustomButton ID="Clone2" meta:resourcekey="GridViewCommandColumnCustomButtonResource1">
                        <Image ToolTip="Clone Record" Url="img/minus.jpg" />

                    </dx:GridViewCommandColumnCustomButton>
                </CustomButtons>
                                                               </dx:GridViewCommandColumn>
                                                 </Columns>
                                             </dx:ASPxGridView>
                                             
                                             
                                             <asp:SqlDataSource ID="Specjalizacje_temp" runat="server" ConnectionString="<%$ ConnectionStrings:wap %>" SelectCommand="SELECT tbl_sprcjalizacje_temp.nr_sesji, glo_specjalizacje.nazwa, glo_specjalizacje.id_, tbl_sprcjalizacje_temp.id_ AS id_temp FROM tbl_sprcjalizacje_temp LEFT OUTER JOIN glo_specjalizacje ON tbl_sprcjalizacje_temp.id_spec = glo_specjalizacje.id_ WHERE (tbl_sprcjalizacje_temp.nr_sesji = @id_sesji)">
                                                 <SelectParameters>
                                                     <asp:SessionParameter Name="id_sesji" SessionField="sesja" />
                                                 </SelectParameters>
                                             </asp:SqlDataSource>
                                         </asp:Panel>
                                     </td>
                                     <td style="vertical-align: top;width:50%">
                                             <asp:Panel ID="Panel9" runat="server" ScrollBars="Vertical" Height="250px" meta:resourcekey="Panel9Resource1">
                                                 <dx:ASPxGridView ID="ASPxGridView5" runat="server" AutoGenerateColumns="False" DataSourceID="SqlDataSource1f" KeyFieldName="nazwa" Width="100%" OnCustomButtonCallback="dodajSpecjalizacje" EnableCallbackAnimation="True" EnableCallBacks="False" meta:resourcekey="ASPxGridView5Resource1">
                                                     <SettingsPager PageSize="10000" Visible="False">
                                                     </SettingsPager>
                                                     <Settings ShowFilterRow="True" />
                                                     <SettingsBehavior AllowSelectByRowClick="True" ProcessSelectionChangedOnServer="True" />
                                                     <SettingsCommandButton>
                                                         <ShowAdaptiveDetailButton ButtonType="Image">
                                                         </ShowAdaptiveDetailButton>
                                                         <HideAdaptiveDetailButton ButtonType="Image">
                                                         </HideAdaptiveDetailButton>
                                                     </SettingsCommandButton>
                                                     <SettingsDataSecurity AllowDelete="False" AllowInsert="False" />
                                                     <Columns>
                                                        
                                                          
                                                         <dx:GridViewDataTextColumn FieldName="nazwa" ReadOnly="True" ShowInCustomizationForm="True" VisibleIndex="1" meta:resourcekey="GridViewDataTextColumnResource20">
                                                         </dx:GridViewDataTextColumn>

                                                          <dx:GridViewCommandColumn ShowNewButton="true" ShowEditButton="true" VisibleIndex="0" ButtonRenderMode="Image" meta:resourcekey="GridViewCommandColumnResource2">
                                                          <CustomButtons>
                    <dx:GridViewCommandColumnCustomButton ID="Clone" meta:resourcekey="GridViewCommandColumnCustomButtonResource2">
                        <Image ToolTip="Clone Record" Url="img/plus-icon.png" />

                    </dx:GridViewCommandColumnCustomButton>
                </CustomButtons>
                                                               </dx:GridViewCommandColumn>
                                                     </Columns>
                                                 </dx:ASPxGridView>
                                             </asp:Panel>
                                         
                                     </td>
                                 </tr>
                             </table>                            <br />
                         <asp:SqlDataSource ID="SqlDataSource1f" runat="server" 
                                    ConnectionString="<%$ ConnectionStrings:wap %>" 
                                    
                                    SelectCommand="SELECT DISTINCT id_,nazwa FROM [glo_specjalizacje] ORDER BY [nazwa]">
                                </asp:SqlDataSource>
                        </dx:ContentControl>
                    </ContentCollection>
                </dx:TabPage>
                      <dx:TabPage Text="uwagi" meta:resourcekey="TabPageResource11">
                    <ContentCollection>
                        <dx:ContentControl ID="ContentControl10" runat="server" meta:resourcekey="ContentControl10Resource1">
                            <div style="min-height:275px;"><table style="width: 100%;"><tr><td>Uwagi</td></tr><tr><td><asp:TextBox ID="TxUwagi" runat="server" TextMode="MultiLine" Width="100%" Height="100%" Rows="12" meta:resourcekey="TxUwagiResource1"></asp:TextBox></td></tr></table></div>
                        </dx:ContentControl>
                    </ContentCollection>
                </dx:TabPage>
                      <dx:TabPage Text="Dane statystyczne" meta:resourcekey="TabPageResource12">
                    <ContentCollection>
                        <dx:ContentControl ID="ContentControl11" runat="server" meta:resourcekey="ContentControl11Resource1">
                            <asp:DropDownList ID="DropDownList4" runat="server" AutoPostBack="True" DataSourceID="kwerendyStatystyczne1" DataTextField="Nazwa" DataValueField="kwerenda" meta:resourceKey="DropDownList3Resource1" OnSelectedIndexChanged="DropDownList4_SelectedIndexChanged" AppendDataBoundItems="True" Height="16px" ViewStateMode="Disabled" Width="139px">
                            </asp:DropDownList>
                            <asp:SqlDataSource ID="kwerendyStatystyczne1" runat="server" ConnectionString="<%$ ConnectionStrings:wap %>" SelectCommand="SELECT Nazwa, kwerenda FROM dane_statystyczne WHERE (czy_us &lt;&gt; 1) ORDER BY Nazwa"></asp:SqlDataSource>
                            <br />
                            <asp:GridView ID="GridView26" runat="server" BackColor="White" BorderColor="#E7E7FF" BorderStyle="None" BorderWidth="1px" CellPadding="3" Css="" GridLines="Horizontal" meta:resourceKey="GridView1Resource1">
                                <AlternatingRowStyle BackColor="#F7F7F7" />
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
                            <div style="min-height:275px;">
                                <br />
                               
                                <br />
                                <br />
                            </div>
                        </dx:ContentControl>
                    </ContentCollection>
                </dx:TabPage>
                      <dx:TabPage Text="Historia powołań" meta:resourcekey="TabPageResource13">
                    <ContentCollection>
                        <dx:ContentControl ID="ContentControl12" runat="server" meta:resourcekey="ContentControl12Resource1">
                              <div style="min-height:275px;">
                                  <table style="width: 100%;">
                                      <tr>
                                          <td style="width: 160px" valign="top">
                                              <table style="width:100%;"><tr><td>od:</td><td>
                                                  <dx:ASPxDateEdit ID="ASPxDateEdit1" runat="server" meta:resourcekey="ASPxDateEdit1Resource1">
                                                  </dx:ASPxDateEdit>
                                                  </td></tr><tr><td>do </td><td>
                                                      <dx:ASPxDateEdit ID="ASPxDateEdit2" runat="server" meta:resourcekey="ASPxDateEdit2Resource1" AllowNull="False">
                                                      </dx:ASPxDateEdit>
                                                      </td></tr><tr><td  colspan="2">
                                                      <asp:LinkButton ID="LinkButton5" runat="server" CausesValidation="False" CssClass="button_" OnClick="dodajPowolanie" Width="100%" meta:resourcekey="LinkButton5Resource1">Dodaj</asp:LinkButton>
                                                      </td></tr><tr><td  colspan="2">
                                                      <asp:LinkButton ID="LinkButton8" runat="server" CausesValidation="False" CssClass="button_" OnClick="zmienPowolanie" Width="100%" meta:resourcekey="LinkButton8Resource1">Zmień</asp:LinkButton>
                                                  </td></tr><tr><td  colspan="2">
                                                      <asp:LinkButton ID="LinkButton9" runat="server" CausesValidation="False" CssClass="button_" OnClick="usunPowolanie" Width="100%" meta:resourcekey="LinkButton9Resource1">Usuń</asp:LinkButton>
                                                      </td></tr></table><br />
                                              <asp:Label ID="Label2" runat="server" meta:resourcekey="Label2Resource1"></asp:Label>
                                              <br /><br /></td><td style="width: 60%" valign="top"><asp:Panel ID="Panel3" runat="server" Height="200px" ScrollBars="Vertical" meta:resourcekey="Panel3Resource1">
                                              <asp:SqlDataSource ID="powolania" runat="server" ConnectionString="<%$ ConnectionStrings:wap %>" SelectCommand="SELECT ident, data_od, data_do, czyus, id_bieglego FROM tbl_powolania WHERE (czyus = 0) AND (id_bieglego = @id_bieglego)">
                                                  <SelectParameters>
                                                      <asp:SessionParameter Name="id_bieglego" SessionField="id_osoby" Type="Int32" />
                                                  </SelectParameters>
                                              </asp:SqlDataSource>
                                              <asp:GridView ID="GridView25" runat="server" AutoGenerateColumns="False" CellPadding="0" DataKeyNames="ident,data_od,data_do" DataSourceID="powolania" ForeColor="#333333" GridLines="None" OnSelectedIndexChanged="wybiezPowolanie" Width="100%" meta:resourcekey="GridView25Resource1">
                                                  <AlternatingRowStyle BackColor="White" />
                                                  <Columns>
                                                      <asp:CommandField ShowSelectButton="True" meta:resourcekey="CommandFieldResource1" />
                                                      <asp:BoundField DataField="data_od" DataFormatString="{0:d}" HeaderText="od" SortExpression="data_od" ReadOnly="True" meta:resourcekey="BoundFieldResource1" />
                                                      <asp:BoundField DataField="data_do" DataFormatString="{0:d}" HeaderText="do" SortExpression="data_do" ReadOnly="True" meta:resourcekey="BoundFieldResource2" />
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
                                              </asp:Panel></td></tr></table></div>
                        </dx:ContentControl>
                    </ContentCollection>
                </dx:TabPage>
                      <dx:TabPage Text="Skargi" meta:resourcekey="TabPageResource14">
                    <ContentCollection>
                        <dx:ContentControl ID="ContentControl14" runat="server" meta:resourcekey="ContentControl14Resource1">
                              <div style="min-height:275px;">
                                       <asp:LinkButton ID="LinkButton10" runat="server" CssClass="button_thin" OnClick="nowaSkarga" meta:resourcekey="LinkButton10Resource1">Nowa skarga</asp:LinkButton>
                                  
                                              <asp:Panel ID="Panel6" runat="server" Height="100%" ScrollBars="Vertical" meta:resourcekey="Panel6Resource1">
                                              
                                              <asp:Panel ID="Panel11" runat="server" Visible="False" Height="300px" Width="100%" HorizontalAlign="Center" meta:resourcekey="Panel11Resource1">
                                                  <table style="border: thin solid #000000; width:99%; height: 250px;" cellpadding="0" cellspacing="0">
                                                      <tr>
                                                          <td colspan="3">
                                                              <strong>
                                                              <asp:Label ID="Label6" runat="server" meta:resourcekey="Label6Resource1"></asp:Label>
                                                              <br />
                                                              </strong>
                                                          </td>
                                                          <td>&nbsp;</td>
                                                          <td>&nbsp;</td>
                                                      </tr>
                                                      <tr>
                                                          <td  style="background-color: #E1F0FF; width: 20%;">Numer</td>
                                                          <td  style="background-color: #E1F0FF">
                                                              <asp:TextBox ID="txNumer" runat="server" meta:resourcekey="txNumerResource1"></asp:TextBox>
                                                          </td>
                                                         
                                                          <td  style="background-color: #E1F0FF; width: 20%;">Rok</td>
                                                          <td  style="background-color: #E1F0FF">
                                                              <asp:DropDownList ID="txRok" runat="server" AutoPostBack="True" OnSelectedIndexChanged="zmienRokSkargi">
                                                              </asp:DropDownList>
                                                          </td>
                                                         
                                                      </tr>
                                                      <tr>
                                                          <td style="width: 20%">Data wpływu</td>
                                                          <td style="width: 30%">
                                                              <dx:ASPxDateEdit ID="ASPxDateWplyw" runat="server" meta:resourcekey="ASPxDateWplywResource1">
                                                              </dx:ASPxDateEdit>
                                                          </td>
                                                          <td style="width: 20%">Data pisma</td>
                                                          <td style="width: 30%">
                                                              <dx:ASPxDateEdit ID="ASPxDatePismo" runat="server" meta:resourcekey="ASPxDatePismoResource1">
                                                              </dx:ASPxDateEdit>
                                                          </td>
                                                      </tr>
                                                        <tr>
                                                          <td style="background-color: #E1F0FF">Sygnatura</td>
                                                          <td style="background-color: #E1F0FF">
                                                               <asp:TextBox ID="txSygnatura" runat="server" meta:resourcekey="txSygnaturaResource1"></asp:TextBox>
                                                            </td>
                                                            <td style="background-color: #E1F0FF">Wizytator</td>
                                                            <td style="background-color: #E1F0FF">
                                                                <asp:TextBox ID="txWizytator" runat="server" meta:resourcekey="txWizytatorResource1"></asp:TextBox>
                                                            </td>
                                                      </tr>
                                                       <tr>
                                                          <td style="width: 20%">Zakreślono</td>
                                                          <td style="width: 30%">
                                                              <asp:CheckBox ID="CheckBox3" runat="server" AutoPostBack="True" OnCheckedChanged="pokazDateZ" meta:resourcekey="CheckBox3Resource1" />
                                                           </td>
                                                           <td style="width: 20%">
                                                               <asp:Label ID="Label7" runat="server" Text="Data zakreślenia" meta:resourcekey="Label7Resource1"></asp:Label>
                                                           </td>
                                                           <td style="width: 30%">
                                                               <dx:ASPxDateEdit ID="ASPxDateZakreslenie" runat="server" meta:resourcekey="ASPxDateZakreslenieResource1">
                                                               </dx:ASPxDateEdit>
                                                           </td>
                                                      </tr>
                                                       <tr>
                                                          <td style="background-color: #E1F0FF">Uwagi</td>
                                                          <td colspan="3" style="background-color: #E1F0FF">  <asp:TextBox ID="uwagi" runat="server" Width="100%" TextMode="MultiLine" meta:resourcekey="uwagiResource1"></asp:TextBox></td>
                                                      </tr>
                                               
                                                  </table>

                                                  <table style="width: 100%;">
                                                      <tr>
                                                          <td align="center" style="width: 33%">
                                                                
                                                                <asp:LinkButton ID="LinkButton1" runat="server" CssClass="button_thin" OnClick="zapiszSkarge" meta:resourcekey="LinkButton1Resource1">Zapisz</asp:LinkButton>

                                                          </td>
                                                          <td align="center">
                                                             
                                                                 <asp:LinkButton ID="LinkButton2" runat="server" CssClass="button_thin" OnClick="anulowanieSkargi" meta:resourcekey="LinkButton2Resource1">Anuluj</asp:LinkButton>
                                                         
                                                          </td>
                                                          <td align="center">
                                                                          <asp:LinkButton ID="LinkButton13" runat="server" CssClass="button_thin" OnClick="usunSkarge" style="height: 12px" meta:resourcekey="LinkButton13Resource1">Usuń</asp:LinkButton>
                                                  

                                                          </td>
                                                      </tr>
                                                     
                                                  </table>
                                                 
                                              </asp:Panel>
                                              <dx:ASPxGridView ID="ListaSkarg" runat="server" AutoGenerateColumns="False" DataSourceID="SqlDataSkargi2" EnableTheming="True" Theme="Office2010Blue" Width="100%" OnCustomButtonCallback="ASPxGridView7_CustomButtonCallback" EnableCallbackAnimation="True" EnableCallBacks="False" EnablePagingCallbackAnimation="True" KeyFieldName="ident" ViewStateMode="Enabled" meta:resourcekey="ASPxGridView7Resource1">
                                                  <SettingsBehavior SortMode="Value" AllowFocusedRow="True" AllowSelectByRowClick="True" AllowSelectSingleRowOnly="True" />
                                                  <SettingsCommandButton>
                                                      <ShowAdaptiveDetailButton ButtonType="Image">
                                                      </ShowAdaptiveDetailButton>
                                                      <HideAdaptiveDetailButton ButtonType="Image">
                                                      </HideAdaptiveDetailButton>
                                                      <SelectButton Text="select">
                                                          <Image Url="~/img/expand.jpg">
                                                          </Image>
                                                      </SelectButton>
                                                  </SettingsCommandButton>
                                                  <SettingsDataSecurity AllowDelete="False" AllowEdit="False" AllowInsert="False" />
                                                  <Columns>
                                                      <dx:GridViewDataTextColumn Caption="Numer" FieldName="numer" ShowInCustomizationForm="True" VisibleIndex="1" meta:resourcekey="GridViewDataTextColumnResource21">
                                                      </dx:GridViewDataTextColumn>
                                                      <dx:GridViewDataTextColumn Caption="Rok" FieldName="rok" ShowInCustomizationForm="True" VisibleIndex="2" meta:resourcekey="GridViewDataTextColumnResource22">
                                                      </dx:GridViewDataTextColumn>
                                                      <dx:GridViewDataDateColumn Caption="Data wpływu" FieldName="dataWplywu" ShowInCustomizationForm="True" VisibleIndex="3" meta:resourcekey="GridViewDataDateColumnResource9">
                                                      </dx:GridViewDataDateColumn>
                                                      <dx:GridViewDataDateColumn Caption="Data pisma" FieldName="dataPisma" ShowInCustomizationForm="True" VisibleIndex="4" meta:resourcekey="GridViewDataDateColumnResource10">
                                                      </dx:GridViewDataDateColumn>
                                                      <dx:GridViewDataTextColumn Caption="Sygnatura" FieldName="Sygnatura" ShowInCustomizationForm="True" VisibleIndex="5" meta:resourcekey="GridViewDataTextColumnResource23">
                                                      </dx:GridViewDataTextColumn>
                                                      <dx:GridViewDataTextColumn Caption="Wizytator" FieldName="wizytator" ShowInCustomizationForm="True" VisibleIndex="6" meta:resourcekey="GridViewDataTextColumnResource24">
                                                      </dx:GridViewDataTextColumn>
                                                      <dx:GridViewDataCheckColumn Caption="Zakreślono" FieldName="zakreslono" ShowInCustomizationForm="True" VisibleIndex="7" meta:resourcekey="GridViewDataCheckColumnResource4">
                                                      </dx:GridViewDataCheckColumn>
                                                      <dx:GridViewDataTextColumn Caption="Uwagi" FieldName="uwagi" ShowInCustomizationForm="True" VisibleIndex="8" meta:resourcekey="GridViewDataTextColumnResource25">
                                                      </dx:GridViewDataTextColumn>
                                                             <dx:GridViewCommandColumn ShowNewButton="true" VisibleIndex="0" ButtonRenderMode="Image" Caption=" " meta:resourcekey="GridViewCommandColumnResource3" ShowNewButtonInHeader="True">
                                                          <CustomButtons>
                    <dx:GridViewCommandColumnCustomButton ID="wybierz" meta:resourcekey="GridViewCommandColumnCustomButtonResource3">
                        <Image ToolTip="Edytuj/usuń skagę" Url="img/button_edycja.png" />

                    </dx:GridViewCommandColumnCustomButton>
                </CustomButtons>
                                                               </dx:GridViewCommandColumn>
                                                  </Columns>
                                              </dx:ASPxGridView>
                                              <asp:SqlDataSource ID="SqlDataSkargi2" runat="server" ConnectionString="<%$ ConnectionStrings:wap %>"  SelectCommand="SELECT numer, rok, dataWplywu, dataPisma, Sygnatura, wizytator, zakreslono, dataZakreslenia, uwagi, ident FROM tbl_skargi WHERE (idBieglego = @idBieglego) AND (czyus = 0) ORDER BY numer, rok" UpdateCommand="UPDATE tbl_skargi SET numer = @numer, rok = @rok, dataWplywu = @dataWplywu, dataPisma = @dataPisma, Sygnatura = @sygnatura, wizytator = wizytator, zakreslono = zakreslono, dataZakreslenia = dataZakreslenia, uwagi = uwagi WHERE (ident = @ident)">
                                                  <SelectParameters>
                                                      <asp:SessionParameter Name="idBieglego" SessionField="id_osoby" DefaultValue="" />
                                                  </SelectParameters>
                                                  <UpdateParameters>
                                                      <asp:Parameter Name="numer" />
                                                      <asp:Parameter Name="rok" />
                                                      <asp:Parameter Name="dataWplywu" />
                                                      <asp:Parameter Name="dataPisma" />
                                                      <asp:Parameter Name="sygnatura" />
                                                      <asp:Parameter Name="ident" />
                                                  </UpdateParameters>
                                              </asp:SqlDataSource>
                                                  <br />
                                              </asp:Panel>

                                          </div>
                        </dx:ContentControl>
                    </ContentCollection>
                </dx:TabPage>
            </TabPages>
        </dx:ASPxPageControl>

                                  <table style="width: 100%;">
                                      <tr>
                                          <td>
                                              <asp:Label ID="Label5" runat="server" meta:resourcekey="Label5Resource1"></asp:Label>
                                              
                                          </td>
                                          <td>
                                              &nbsp;</td>
                                          <td>&nbsp;</td>
                                      </tr>
                                     
                                      <tr>
                                          <td>
                                              <asp:Button ID="Button1" runat="server" OnClick="zapiszDaneBieglego" Text="Zapisz" CssClass="button_" meta:resourcekey="Button1Resource2" />
                                          </td>
                                          <td>&nbsp;</td>
                                          <td>
                                              <asp:Button ID="Button3" runat="server" Text="Usuń" CssClass="button_" OnClick="usunBieglego" meta:resourcekey="Button3Resource1" />
                                          </td>
                                      </tr>
                                     
                                  </table>

                            </asp:Panel>
                            <table class="InfoTable">
                                <tr>
                                    <td>
                                        &nbsp;</td>
                                    <td>
                                        &nbsp;</td>
                                </tr>
                            </table>
                        </dx:PanelContent>
                    </PanelCollection>
                </dx:ASPxCallbackPanel>
            </dx:PopupControlContentControl>
              
              
               
        </ContentCollection>
        <ClientSideEvents Shown="popup_Shown" CloseUp="function(s, e) { s.PerformCallback();}" />
    </dx:ASPxPopupControl>

     </div>
</asp:Content>
