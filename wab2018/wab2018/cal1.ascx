<%@ Control Language="C#" AutoEventWireup="true" CodeBehind="cal1.ascx.cs" Inherits="wab.cal1" %>
  <style type="text/css">
    body, p {
        font-family: Tahoma, Arial, Sans-serif;
        font-size: 10pt;
    }
    h1 {
        margin: 0px;
        padding: 0px;
    }
    .Panel_1
    {
        position:absolute;
        top:103px;
        left:243px;
        width:300px;
        height:170px;
         box-shadow: 10px 10px 5px #888;
         background-color:#99CCFF;
        }
        .Panel_2
    {
        position:absolute;
        top:0px;
        left:0px;
        width:100%;
        height:100%;
        background-color:rgba(0, 0, 0, 0.22);
         
        }
        
    </style>


<asp:Panel ID="Panel1" runat="server" CssClass="Panel_2" Visible="False">
    <asp:Panel ID="Panel2" runat="server" CssClass="Panel_1">
        <div style="float: right; background-color: #99CCFF;">
       
        <asp:ImageButton ID="ImageButton2" runat="server" ImageUrl="~/img/cancel.png" 
            onclick="ImageButton2_Click" />
    </div> 
    <asp:Calendar ID="Calendar1" runat="server" BackColor="White" 
        BorderColor="#3366CC" BorderWidth="1px" CellPadding="1" 
        DayNameFormat="Shortest" Font-Names="Verdana" Font-Size="8pt" 
        ForeColor="#003399" Height="119px" Width="100%" 
        onselectionchanged="Calendar1_SelectionChanged">
        <DayHeaderStyle BackColor="#99CCCC" ForeColor="#336666" Height="1px" />
        <NextPrevStyle Font-Size="8pt" ForeColor="#CCCCFF" />
        <OtherMonthDayStyle ForeColor="#999999" />
        <SelectedDayStyle BackColor="#009999" Font-Bold="True" ForeColor="#CCFF99" />
        <SelectorStyle BackColor="#99CCCC" ForeColor="#336666" />
        <TitleStyle BackColor="#003399" BorderColor="#3366CC" BorderWidth="1px" 
            Font-Bold="True" Font-Size="10pt" ForeColor="#CCCCFF" Height="25px" />
        <TodayDayStyle BackColor="#99CCCC" ForeColor="White" />
        <WeekendDayStyle BackColor="#CCCCFF" />
    </asp:Calendar>
    </asp:Panel>
</asp:Panel>

<table  width="140px">
     <tr style="height:21px;">
        <td width="120px">
            <asp:TextBox ID="TextBox1" runat="server" ReadOnly="True" Width="120px"></asp:TextBox>
        </td>
        <td>
            <asp:ImageButton ID="ImageButton1" runat="server" ImageUrl="~/img/cal_1.jpg" 
                onclick="ImageButton1_Click" Height="25px" Width="25px" />
        </td>
    </tr>
    </table>

<asp:HiddenField ID="HiddenField1" runat="server" />


