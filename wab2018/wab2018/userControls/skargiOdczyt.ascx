<%@ Control Language="C#" AutoEventWireup="true" CodeBehind="skargiOdczyt.ascx.cs"  Inherits="wab2018.skargiOdczyt" %>
<%@ Register assembly="DevExpress.Web.v17.1, Version=17.1.17.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.Web" tagprefix="dx" %>


<div style="z-index:99999;">



<asp:SqlDataSource ID="ListaPowolanMediatorow" runat="server" ConnectionString="<%$ ConnectionStrings:wap %>" DeleteCommand="UPDATE tbl_skargi SET czyus =1 WHERE (ident = @ident)" InsertCommand="INSERT INTO tbl_skargi(numer, rok, dataWplywu, dataPisma, Sygnatura, wizytator, zakreslono, dataZakreslenia, uwagi, czyus, idBieglego) VALUES (@numer, @rok, @dataWplywu, @dataPisma, @Sygnatura, @wizytator, @zakreslono, @dataZakreslenia, @uwagi, 0, @idBieglego)" SelectCommand="SELECT ident, numer, rok, dataWplywu, dataPisma, Sygnatura, wizytator, COALESCE ( zakreslono,null,'0') as zakreslono, dataZakreslenia, uwagi, czyus, idBieglego FROM tbl_skargi WHERE (czyus = 0) AND (idBieglego = @id_bieglego) ORDER BY rok, numer" UpdateCommand="UPDATE tbl_skargi SET numer = @numer, rok = @rok, dataWplywu = @dataWplywu, dataPisma = @dataPisma, Sygnatura = @Sygnatura, wizytator = @wizytator, zakreslono = @zakreslono, dataZakreslenia = @dataZakreslenia, uwagi = @uwagi WHERE (ident = @ident)">
    <DeleteParameters>
        <asp:Parameter Name="ident" />

    </DeleteParameters>
    <InsertParameters>
        <asp:Parameter Name="numer" />
        <asp:Parameter Name="rok" />
        <asp:Parameter Name="dataWplywu" />
       
        <asp:Parameter Name="dataPisma" />
        <asp:Parameter Name="Sygnatura" />
        <asp:Parameter Name="wizytator" />
        <asp:Parameter Name="zakreslono" />
        <asp:Parameter Name="dataZakreslenia" />
        <asp:Parameter Name="uwagi" />
        <asp:Parameter Name="idBieglego" />
    </InsertParameters>
    <SelectParameters>
        <asp:SessionParameter Name="id_bieglego" SessionField="idMediatora" Type="Int32" />
    </SelectParameters>
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
        <asp:Parameter Name="ident" />
    </UpdateParameters>
</asp:SqlDataSource>



<dx:ASPxGridView ID="ASPxGridView1" runat="server" AutoGenerateColumns="False" DataSourceID="ListaPowolanMediatorow" KeyFieldName="ident" OnCustomCallback="ASPxGridView1_CustomCallback" OnInitNewRow="ASPxGridView1_InitNewRow" OnRowInserted="ASPxGridView1_RowInserted" OnCellEditorInitialize="ASPxGridView1_CellEditorInitialize" OnRowUpdating="ASPxGridView1_RowUpdating" OnRowInserting="ASPxGridView1_RowInserting" Theme="Moderno" Width="100%">
    <Columns>
        <dx:GridViewCommandColumn VisibleIndex="0">
        </dx:GridViewCommandColumn>
        <dx:GridViewDataTextColumn FieldName="ident" Visible="False" VisibleIndex="4">
        </dx:GridViewDataTextColumn>
        <dx:GridViewDataTextColumn Caption="Numer" FieldName="numer" VisibleIndex="1">
        </dx:GridViewDataTextColumn>
        <dx:GridViewDataTextColumn Caption="Rok" FieldName="rok" VisibleIndex="2">
        </dx:GridViewDataTextColumn>
        <dx:GridViewDataTextColumn Caption="Sygnatura" FieldName="Sygnatura" VisibleIndex="3">
        </dx:GridViewDataTextColumn>
        <dx:GridViewDataTextColumn FieldName="wizytator" Name="Wizytator" VisibleIndex="9" Caption="Sędzia wizytator">
        </dx:GridViewDataTextColumn>
        <dx:GridViewDataDateColumn AllowTextTruncationInAdaptiveMode="True" Caption="Data wpływu" FieldName="dataWplywu" VisibleIndex="5">
            <PropertiesDateEdit DisplayFormatString="yyyy-MM-dd">
            </PropertiesDateEdit>
        </dx:GridViewDataDateColumn>
        <dx:GridViewDataDateColumn FieldName="dataPisma" Name="Data pisma" VisibleIndex="6" Caption="Data wpływu pisma">
            <PropertiesDateEdit DisplayFormatString="yyyy-MM-dd">
            </PropertiesDateEdit>
        </dx:GridViewDataDateColumn>
        <dx:GridViewDataDateColumn FieldName="dataZakreslenia" Name="Data zakreślenia" VisibleIndex="7" Caption="Data Zakreślenia">
            <PropertiesDateEdit DisplayFormatString="yyyy-MM-dd">
            </PropertiesDateEdit>
        </dx:GridViewDataDateColumn>
        <dx:GridViewDataCheckColumn FieldName="zakreslono" Name="Zakreślono" VisibleIndex="8" Caption="Zakreślono">
        </dx:GridViewDataCheckColumn>
        <dx:GridViewDataTextColumn Caption="Uwagi" FieldName="uwagi" VisibleIndex="10">
        </dx:GridViewDataTextColumn>
    </Columns>
     <SettingsEditing Mode="Inline">
    </SettingsEditing>
     <SettingsDataSecurity AllowDelete="False" AllowEdit="False" AllowInsert="False" />
     <EditFormLayoutProperties>
            <SettingsAdaptivity AdaptivityMode="SingleColumnWindowLimit" SwitchToSingleColumnAtWindowInnerWidth="700" />
        </EditFormLayoutProperties>

        <SettingsPopup>
         
        </SettingsPopup>
</dx:ASPxGridView>
    <br />
    <br />
    </div>

