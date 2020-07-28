<%@ Page Title="Home Page" Language="C#" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeBehind="Default.aspx.cs" Inherits="ImportSample._Default" %>

<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">

   <h3>Import / Export database data from/to Excel.</h3>
<div>
    <table>
        <tr>
            <td>Select File&nbsp; : </td>
            <td>
                <asp:FileUpload ID="FileUpload1" runat="server" />
                </td>
            <td>
                <asp:Button ID="btnImport" runat="server" Text="Import Data" OnClick="btnImport_Click"  />
            </td>
           
        </tr>
    </table>
    <div>
        <br />
        <asp:Label ID="lblMessage" runat="server"  Font-Bold="true" />
        <br />
        <asp:GridView ID="gvData" runat="server" AutoGenerateColumns="false">
            <EmptyDataTemplate>
                <div style="padding:10px">
                    Data not found!
                </div>
            </EmptyDataTemplate>
            <Columns>
                <asp:BoundField HeaderText=" Employee ID" DataField="EmployeeID" />
                <asp:BoundField HeaderText=" Company Name" DataField="CompanyName" />
                <asp:BoundField HeaderText=" Contact Name" DataField="ContactName" />
                <asp:BoundField HeaderText=" Contact Title" DataField="ContactTitle" />
                <asp:BoundField HeaderText=" Address" DataField="EmployeeAddress" />
                <asp:BoundField HeaderText=" Postal Code" DataField="PostalCode" />
            </Columns>
        </asp:GridView>
       
        <asp:Button ID="btnEdit" runat="server" OnClick="btnUpdate_Click" Text="Edit Data" />
        <br />
    </div>
</div>

</asp:Content>
