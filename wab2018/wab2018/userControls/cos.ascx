<%@ Control Language="C#" AutoEventWireup="true" CodeBehind="cos.ascx.cs"  Inherits="wab2018.cos" %>
<%@ Register assembly="DevExpress.Web.v17.1, Version=17.1.17.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.Web" tagprefix="dx" %>


<div style="z-index:99999;">



<asp:SqlDataSource ID="ListaPowolanMediatorow" runat="server" ConnectionString="<%$ ConnectionStrings:wap %>" DeleteCommand="UPDATE tbl_skargi SET czyus =1 WHERE (ident = @ident)" InsertCommand="INSERT INTO tbl_skargi(numer, rok, dataWplywu, dataPisma, Sygnatura, wizytator, zakreslono, dataZakreslenia, uwagi, czyus, idBieglego) VALUES (@numer, @rok, @dataWplywu, @dataPisma, @Sygnatura, @wizytator, @zakreslono, @dataZakreslenia, @uwagi, 0, @idBieglego)" SelectCommand="SELECT ident, CAST(RTRIM(numer) AS bigint) AS numer, rok, dataWplywu, dataPisma, Sygnatura, wizytator, COALESCE (zakreslono, NULL, '0') AS zakreslono, dataZakreslenia, uwagi, czyus, idBieglego FROM tbl_skargi WHERE (czyus = 0) AND (idBieglego = @id_bieglego) ORDER BY rok, numer" UpdateCommand="UPDATE tbl_skargi SET numer = @numer, rok = @rok, dataWplywu = @dataWplywu, dataPisma = @dataPisma, Sygnatura = @Sygnatura, wizytator = @wizytator, zakreslono = @zakreslono, dataZakreslenia = @dataZakreslenia, uwagi = @uwagi WHERE (ident = @ident)">
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
        <dx:GridViewCommandColumn ShowDeleteButton="True" ShowEditButton="True" ShowNewButtonInHeader="True" VisibleIndex="0">
        </dx:GridViewCommandColumn>
        <dx:GridViewDataTextColumn FieldName="ident" Visible="False" VisibleIndex="4">
        </dx:GridViewDataTextColumn>
        <dx:GridViewDataTextColumn Caption="Numer" FieldName="numer" VisibleIndex="1" Width="50px">
        </dx:GridViewDataTextColumn>
        <dx:GridViewDataTextColumn Caption="Rok" FieldName="rok" VisibleIndex="2" Width="50px">
        </dx:GridViewDataTextColumn>
        <dx:GridViewDataTextColumn Caption="Sygnatura" FieldName="Sygnatura" VisibleIndex="3" Width="200px">
        </dx:GridViewDataTextColumn>
        <dx:GridViewDataTextColumn FieldName="wizytator" Name="Wizytator" VisibleIndex="9" Width="300px">
        </dx:GridViewDataTextColumn>
        <dx:GridViewDataDateColumn AllowTextTruncationInAdaptiveMode="True" Caption="Data wpływu" FieldName="dataWplywu" VisibleIndex="5" Width="100px">
            <PropertiesDateEdit DisplayFormatString="yyyy-MM-dd">
            </PropertiesDateEdit>
        </dx:GridViewDataDateColumn>
        <dx:GridViewDataDateColumn FieldName="dataPisma" Name="Data pisma" VisibleIndex="6" Width="100px">
            <PropertiesDateEdit DisplayFormatString="yyyy-MM-dd">
            </PropertiesDateEdit>
        </dx:GridViewDataDateColumn>
        <dx:GridViewDataDateColumn FieldName="dataZakreslenia" Name="Data zakreślenia" VisibleIndex="7" Width="100px">
            <PropertiesDateEdit DisplayFormatString="yyyy-MM-dd">
            </PropertiesDateEdit>
        </dx:GridViewDataDateColumn>
        <dx:GridViewDataCheckColumn FieldName="zakreslono" Name="Zakreślono" VisibleIndex="8" Width="50px">
        </dx:GridViewDataCheckColumn>
        <dx:GridViewDataTextColumn Caption="Uwagi" FieldName="uwagi" VisibleIndex="10" Width="300px">
        </dx:GridViewDataTextColumn>
    </Columns>
     <SettingsEditing Mode="Inline">
    </SettingsEditing>
     <EditFormLayoutProperties>
            <SettingsAdaptivity AdaptivityMode="SingleColumnWindowLimit" SwitchToSingleColumnAtWindowInnerWidth="700" />
        </EditFormLayoutProperties>

        <SettingsPopup>
         
        </SettingsPopup>
</dx:ASPxGridView>
    <br />
    <br />
    </div>

