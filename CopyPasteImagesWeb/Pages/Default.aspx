<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Default.aspx.cs" Inherits="Contoso.Core.CopyPasteImagesWeb.Default" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>Contoso Copy Paste Images</title>
    <script type="text/javascript" src="../Scripts/jquery-1.9.1.min.js"></script>
    <script type="text/javascript" src="../Scripts/app.js"></script>
</head>
<body style="display: none">
    <form id="form1" runat="server">
        <asp:ScriptManager ID="ScriptManager1" runat="server" EnableCdn="True" />
        <div id="divSPChrome"></div>

        <div style="padding-left: 15px; padding-right: 15px; overflow-y: scroll; position: absolute; top: 132px; bottom: 0px; right: 0px; left: 0px;">
            <div>                
                <h1 style="padding-left: 10px;">Step 1: Understanding the App</h1>
                <div style="padding-left: 10px;">
                    This example shows how you can 
                    add copy/paste functionality for images in SharePoint Body field in Pages, using JavaScript SCOM. The example uses the Publishing Images library to save any image that a user try to paste on the Body field of their pages. The languages of the solution supported now are english and portuguese, depending on the current users profile settings. The script could be incorporated in master pages to work over site collections. To setup the solution on the host web follow the steps below. 
                </div>
                <div style="padding: 10px;">
                    Your current display language settings:
                    <asp:Label ID="currentLanguages" Font-Bold="true" runat="server" />
                </div>
            </div>

            <div>
                <h1 style="padding-left: 10px;">Step 2: Add Scripts</h1>
                <div style="padding: 10px;">
                    Please, click the button below to upload jQuery Library and Copy Paste Images JavaScript file to the Site Assets library in the host web. 
                    <br />
                    Also, this step will register script links on the host web.
                </div>
                <div style="padding: 10px;">
                    <asp:Button ID="btnAddAssets" Text="Add scripts, images and links" OnClick="AddAssets_Click" runat="server" />
                </div>
            </div>

            <div style="padding: 10px;">
                <h1 style="padding-left: 10px;">Limitations</h1>
                <div>
                    <ul>
                        <li>The script does not work currently with Minimal Download Strategy.</li>
                        <li>This is no production ready code, so there is no caching when accessing User Profile.</li>
                    </ul>
                </div>
            </div>
            <div style="padding: 10px;">
                <h1 style="padding-left: 10px;">Removal</h1>
                <div>
                    Click the button below to remove the script links from the host web.
                </div>
                <div style="margin-top: 10px;">
                    <asp:Button ID="btnRemoveScripts" Text="Remove script links" OnClick="RemoveScripts_Click" runat="server" />
                </div>
            </div>
        </div>
    </form>
</body>
</html>
