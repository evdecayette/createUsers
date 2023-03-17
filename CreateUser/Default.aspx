<%@ Page Title="" Language="C#" MasterPageFile="~/UserMaster.Master" AutoEventWireup="true" CodeBehind="Default.aspx.cs" Inherits="CreateUser.Default" %>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <div class="container">
        <div class="row m-3">
            <div>
                <asp:Label ID="lblMessage" runat="server"></asp:Label>
            </div>
            <div class="col-lg-6 pl-3">
                <div class="card">
                  <div class="card-header">
                    Preview List
                  </div>
                  <div class="card-body">
                      <div>
                          <asp:Button CssClass="btn btn-primary pb-2" runat="server" Text="Start Created" ID="btnCreate" OnClick="btnCreatClick" />
                      </div>
                      <asp:ListView ID="ListView1" runat="server">
                           <ItemTemplate>
                                    <p><%# Eval("NomComplet") %><p>
                           </ItemTemplate>
                      </asp:ListView>
                  </div>
                </div>
            </div>
            <div class="col-lg-6 pl-3">
                <div class="card">
                  <div class="card-header">
                    Created List
                  </div>
                  <div class="card-body">
                      <asp:Literal ID="displayUsers" runat="server"></asp:Literal>
                  </div>
                </div>
            </div>
        </div>
        <div class="row">
        <div class="col-lg-2">

        </div>
        <div class="col-lg-8">
            <div class="progress" role="progressbar" aria-label="Example with label" aria-valuenow="25" aria-valuemin="0" aria-valuemax="100">
              <div class="progress-bar" style="width: 25%">25%</div>
            </div>
        </div>
        <div class="col-lg-2">

        </div>
       </div>
    </div>
</asp:Content>
