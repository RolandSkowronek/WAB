<%@ Page Title="" Language="C#" MasterPageFile="~/Master.Master" AutoEventWireup="true" CodeBehind="wykazSkarg.aspx.cs" Inherits="wab2018.wykazSkarg" %>
<%@ Register assembly="DevExpress.Web.v17.1, Version=17.1.17.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.Web" tagprefix="dx" %>
<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="server">
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <script type="text/javascript">
        var keyValue;
        function OnMoreInfoClick(element, key) {
           
            document.open("biegliLista.aspx?skarga="+key, "", "", true);
          
        }
        function popup_Shown(s, e) {
            callbackPanel.PerformCallback(keyValue);
        }
    </script>
       <div id ="mainWindow" style="background-color:white;" min-height:800 px;" >

        <dx:ASPxGridView ID="ASPxGridView1" runat="server" DataSourceID="skargiSQL" AutoGenerateColumns="False" EnableTheming="True" Theme="Moderno" Width="100%" KeyFieldName="idSkargi">
            <Settings ShowFilterRow="True" />
                                                            
            <SettingsDataSecurity AllowInsert="False" />

            <Columns>
                 <dx:GridViewDataColumn VisibleIndex="1" Width="30px">
                <DataItemTemplate>
                    <a href="javascript:void(0);" onclick="OnMoreInfoClick(this, '<%# Container.KeyValue %>')">B</a>
                </DataItemTemplate>
            </dx:GridViewDataColumn>
                <dx:GridViewDataTextColumn Caption="Numer" FieldName="numer" ShowInCustomizationForm="True" VisibleIndex="2" Width="5%">
                </dx:GridViewDataTextColumn>
                <dx:GridViewDataTextColumn Caption="Rok" FieldName="rok" ShowInCustomizationForm="True" VisibleIndex="3" Width="5%">
                </dx:GridViewDataTextColumn>
                <dx:GridViewDataDateColumn Caption="Data wpływu" FieldName="dataWplywu" ShowInCustomizationForm="True" VisibleIndex="5" Width="7%">
                </dx:GridViewDataDateColumn>
                <dx:GridViewDataDateColumn Caption="Data pisma" FieldName="dataPisma" ShowInCustomizationForm="True" VisibleIndex="6" Width="7%">
                </dx:GridViewDataDateColumn>
                <dx:GridViewDataTextColumn Caption="Sygnatura" FieldName="Sygnatura" ShowInCustomizationForm="True" VisibleIndex="7" Width="10%">
                </dx:GridViewDataTextColumn>
                <dx:GridViewDataTextColumn Caption="Biegły" FieldName="Biegly" ReadOnly="True" ShowInCustomizationForm="True" VisibleIndex="4" Width="15%">
                   <Settings AutoFilterCondition="Contains" />  
                    <EditFormSettings Visible="False" />
                </dx:GridViewDataTextColumn>
                <dx:GridViewDataTextColumn Caption="Wizytator" FieldName="wizytator" ShowInCustomizationForm="True" VisibleIndex="8" Width="15%">
                     <Settings AutoFilterCondition="Contains" />  
                </dx:GridViewDataTextColumn>
                    <dx:GridViewDataCheckColumn Caption="Zakreślono" FieldName="zakreslono" VisibleIndex="9" meta:resourcekey="GridViewDataCheckColumnResource1" ShowInCustomizationForm="True" Width="5%">
                         </dx:GridViewDataCheckColumn>
                 
                <dx:GridViewDataTextColumn Caption="Uwagi" FieldName="uwagi" VisibleIndex="12" Width="15%">
                </dx:GridViewDataTextColumn>
                <dx:GridViewCommandColumn VisibleIndex="0" Caption="  " Width="0px">
                </dx:GridViewCommandColumn>
                 <dx:GridViewDataDateColumn FieldName="dataZakreslenia" VisibleIndex="10" Width="7%">
                 </dx:GridViewDataDateColumn>
            </Columns>
               <SettingsPager Visible="False">
            </SettingsPager>
               <SettingsEditing Mode="Batch" />
            
        </dx:ASPxGridView>
     
      
        <asp:SqlDataSource ID="skargiSQL" runat="server" ConnectionString="<%$ ConnectionStrings:wap %>" SelectCommand="SELECT CAST(RTRIM(tbl_skargi.numer) AS bigint) AS numer, tbl_skargi.rok, tbl_skargi.dataWplywu, tbl_skargi.dataPisma, tbl_skargi.Sygnatura, tbl_osoby.imie + ' ' + RTRIM(tbl_osoby.nazwisko) AS Biegly, tbl_skargi.wizytator, COALESCE (tbl_skargi.zakreslono, NULL, 0) AS zakreslono, tbl_skargi.zakreslono AS Expr1, tbl_osoby.ident, tbl_skargi.uwagi, tbl_skargi.ident AS idSkargi, tbl_skargi.dataZakreslenia FROM tbl_skargi LEFT OUTER JOIN tbl_osoby ON tbl_skargi.idBieglego = tbl_osoby.ident WHERE (tbl_skargi.czyUs = 0) ORDER BY tbl_skargi.rok, numer" UpdateCommand="UPDATE tbl_skargi SET numer = @numer, rok = @rok, dataWplywu = @dataWplywu, dataPisma = @dataPisma, Sygnatura = @Sygnatura, wizytator = @wizytator, zakreslono = @zakreslono, dataZakreslenia = @dataZakreslenia, uwagi = @uwagi WHERE (ident = @idSkargi)" DeleteCommand="UPDATE tbl_skargi SET czyus = 1 WHERE (ident = @idSkargi)" InsertCommand="INSERT INTO tbl_skargi(numer, rok, dataWplywu, dataPisma, Sygnatura, wizytator, dataZakreslenia, uwagi, czyus, idBieglego) VALUES (@numer, @rok, @dataWplywu, @dataPisma, @Sygnatura, @wizytator, @dataZakreslenia, @uwagi, 0, @idBieglego)">
            <DeleteParameters>
                <asp:Parameter Name="idSkargi" />
            </DeleteParameters>
            <InsertParameters>
                <asp:Parameter Name="numer" />
                <asp:Parameter Name="rok" />
                <asp:Parameter Name="dataWplywu" />
                <asp:Parameter Name="dataPisma" />
                <asp:Parameter Name="Sygnatura" />
                <asp:Parameter Name="wizytator" />
                <asp:Parameter Name="dataZakreslenia" />
                <asp:Parameter Name="uwagi" />
                <asp:Parameter Name="idBieglego" />
            </InsertParameters>
            <UpdateParameters>
                <asp:Parameter Name="numer" />
                <asp:Parameter Name="rok" />
                <asp:Parameter Name="dataWplywu" />
                <asp:Parameter Name="dataPisma" />
                <asp:Parameter Name="Sygnatura" />
                <asp:Parameter Name="wizytator" />
                <asp:Parameter Name="zakreslono" />
                <asp:Parameter Name="dataZakreslenia" />
                <asp:Parameter Name="uwagi" />
                <asp:Parameter Name="idSkargi" />
            </UpdateParameters>
        </asp:SqlDataSource>
     
      </div>
   
</asp:Content>
