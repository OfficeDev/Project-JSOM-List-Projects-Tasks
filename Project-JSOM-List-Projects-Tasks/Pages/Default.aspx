<%-- 
   Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 See LICENSE in the project root for license information.
 --%>
<%-- The following 4 lines are ASP.NET directives needed when using SharePoint components --%>

<%@ Page Inherits="Microsoft.SharePoint.WebPartPages.WebPartPage, Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" MasterPageFile="~masterurl/default.master" Language="C#" %>

<%@ Register TagPrefix="Utilities" Namespace="Microsoft.SharePoint.Utilities" Assembly="Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>
<%@ Register TagPrefix="WebPartPages" Namespace="Microsoft.SharePoint.WebPartPages" Assembly="Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>
<%@ Register TagPrefix="SharePoint" Namespace="Microsoft.SharePoint.WebControls" Assembly="Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>

<%-- The markup and script in the following Content element will be placed in the <head> of the page --%>
<asp:Content ContentPlaceHolderID="PlaceHolderAdditionalPageHead" runat="server">
       <script type="text/javascript" src="https://ajax.aspnetcdn.com/ajax/jQuery/jquery-1.9.1.min.js"></script>
    <SharePoint:ScriptLink name="sp.js" runat="server" OnDemand="true" LoadAfterUI="true" Localizable="false" />
    <SharePoint:ScriptLink name="PS.js" runat="server" OnDemand="false" LoadAfterUI="true" Localizable="false" />
   
    <meta name="WebPartPageExpansion" content="full" />

    <!-- Add your CSS styles to the following file -->
    <link rel="Stylesheet" type="text/css" href="../Content/App.css" />

    <!-- Add your JavaScript to the following file -->
    <script type="text/javascript" src="../Scripts/App.js"></script>
</asp:Content>

<%-- The markup in the following Content element will be placed in the TitleArea of the page --%>
<asp:Content ContentPlaceHolderID="PlaceHolderPageTitleInTitleArea" runat="server">
    List Projects and Tasks using JSOM
</asp:Content>

<%-- The markup and script in the following Content element will be placed in the <body> of the page --%>
<asp:Content ContentPlaceHolderID="PlaceHolderMain" runat="server">

    <div>
        <!-- Removing for runtime. If restoring, fix the embedded end-of-comment marker 
        <p id="message">
            <!-- The following content will be replaced with the user name when you run the app - see App.js -- >
            initializing...
        </p>
        -->

        <p>
           List the tasks for any project by clicking the button to the left edge of the project name. <br />
           Remove all tasks in the listing by clicking the following button.
        </p>
        <p>
           <button height='20px' name='btnRemoveTasks' onclick='btnRemoveTasks()'>Remove tasks</button>
           <button height='20px' id='nosey' name='' hidden=true />

        </p>
    </div> 
    <div>
        <table id="ptable">
            <thead>
                <th align="left"></th>
                <th align="left" width="300px">Project Name</th>
                <th align="left">Project Id</th>
            </thead>

            <tbody id="pbody">
            </tbody>
        </table>
 
        <p id="msg">Messages go here.</p>
    </div>

</asp:Content>
