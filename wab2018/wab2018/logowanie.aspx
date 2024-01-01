<%@ Page Title="" Language="C#" MasterPageFile="~/Master.Master" AutoEventWireup="true" CodeBehind="logowanie.aspx.cs" Inherits="wab2018.logowanie" %>
<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="server">
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
     
    <div id ="mainWindow" style="background-color:white;" min-height:800 px;" >
        <br />
        <br />
        <br />
                <div class="accountInfo">
                    <div style=" color: #000;background: 0;padding-bottom: 14px;font-size: 24px;border-bottom: 1px solid #e6e6e6;margin-bottom: 15px;height: 18px;">
                        Zaloguj się
                   </div>
                        
                        <br />
                        <div style="padding:24px 24px 15px 15px">
                        <div style=" height:28px;">
                            <div style="width: 140px;float:left; ">
                                <asp:Label ID="Label4" runat="server" Text=" Użytkownik"></asp:Label>
                            </div>
                              <div style="width: 160px;float:right">
                                  <asp:TextBox ID="TextBox1" runat="server" style="margin-top: 0px"></asp:TextBox>
                            </div>
                            <br />
                        </div>
                           <div style=" height:28px;">
                            <div style="width: 140px;float:left; ">
                                <asp:Label ID="Label6" runat="server" Text=" Hasło"></asp:Label>
                            </div>
                              <div style="width: 160px;float:right">
                                  <asp:TextBox ID="TextBox2" runat="server" TextMode="Password" 
                                      ontextchanged="TextBox2_TextChanged" AutoPostBack="True"></asp:TextBox>
                                  <br />
                            </div>
                            <br />
                               <br />
                        </div>
                           </div>

                            <div Class ="OK_login">
                              
                                <table style="width:100%;">
                                    <tr>
                                        <td  align="center">
                              
                               <asp:LinkButton ID="LinkButton1" runat="server" Width="200px"  
                                    CssClass="button_" onclick="LinkButton1_Click">Zaloguj!</asp:LinkButton>
                                        </td>
                                       
                                    </tr>
                                   
                                   
                                </table>
                                
      
                            </div>
                        
                    </div>
                <div>
                    <br />
                    <asp:Label ID="Label3" runat="server"></asp:Label><br/>
                     <asp:RequiredFieldValidator ID="RequiredFieldValidator1" runat="server" 
                            ControlToValidate="TextBox1" ErrorMessage="Proszę wpisać użytkownika"></asp:RequiredFieldValidator>
                     <asp:RequiredFieldValidator ID="RequiredFieldValidator2" runat="server" 
                            ControlToValidate="TextBox2" ErrorMessage="Hasło jest wymagane"></asp:RequiredFieldValidator>
                </div>
        </div>       
      <br />
        <br />
        <br />
</asp:Content>
