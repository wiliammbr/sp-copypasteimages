using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.UserProfiles;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace Contoso.Core.CopyPasteImagesWeb
{
    public partial class Default : System.Web.UI.Page
    {
        protected void Page_PreInit(object sender, EventArgs e)
        {
            Uri redirectUrl;
            switch (SharePointContextProvider.CheckRedirectionStatus(Context, out redirectUrl))
            {
                case RedirectionStatus.Ok:
                    return;
                case RedirectionStatus.ShouldRedirect:
                    Response.Redirect(redirectUrl.AbsoluteUri, endResponse: true);
                    break;
                case RedirectionStatus.CanNotRedirect:
                    Response.Write("An error occurred while processing your request.");
                    Response.End();
                    break;
            }
        }

        protected void Page_Load(object sender, EventArgs e)
        {
            RegisterChromeControlScript();
            
            var spContext = SharePointContextProvider.Current.GetSharePointContext(Context);

            using (var clientContext = spContext.CreateUserClientContextForSPHost())
            {
                var user = clientContext.Web.CurrentUser;
                clientContext.Load(user);
                clientContext.ExecuteQuery();

                var peopleManager = new PeopleManager(clientContext);
                var userProperties = peopleManager.GetUserProfilePropertyFor(user.LoginName, "SPS-MUILanguages");               
                clientContext.ExecuteQuery();

                currentLanguages.Text = userProperties.Value;
            }
        }

        private void RegisterChromeControlScript()
        {
            string script = @"
            function chromeLoaded() {
                $('body').show();
            }

            //function callback to render chrome after SP.UI.Controls.js loads
            function renderSPChrome() {
                //Set the chrome options for launching Help, Account, and Contact pages
                var options = {
                    'appTitle': document.title,
                    'onCssLoaded': 'chromeLoaded()'
                };

                //Load the Chrome Control in the divSPChrome element of the page
                var chromeNavigation = new SP.UI.Controls.Navigation('divSPChrome', options);
                chromeNavigation.setVisible(true);
            }";

            Page.ClientScript.RegisterClientScriptBlock(typeof(Default), "BasePageScript", script, true);
        }

        protected void AddAssets_Click(object sender, EventArgs e)
        {
            var spContext = SharePointContextProvider.Current.GetSharePointContext(Context);

            using (var clientContext = spContext.CreateUserClientContextForSPHost())
            {
                AddAssetsToHostWeb(clientContext);
                AddScriptLinksToHostWeb(clientContext);
            }
        }

        protected void RemoveScripts_Click(object sender, EventArgs e)
        {
            var spContext = SharePointContextProvider.Current.GetSharePointContext(Context);

            using (var clientContext = spContext.CreateUserClientContextForSPHost())
            {
                var existingActions = clientContext.Web.UserCustomActions;
                clientContext.Load(existingActions);
                clientContext.ExecuteQuery();

                RemoveScriptLinksFromHostWeb(clientContext, existingActions);
            }
        }

        private void AddAssetsToHostWeb(ClientContext clientContext)
        {
            var web = clientContext.Web;
            var library = web.Lists.GetByTitle("Site Assets");
            clientContext.Load(library, l => l.RootFolder);

            UploadScript(clientContext, library, "jquery-1.9.1.min.js");
            UploadScript(clientContext, library, "copypasteimages.js");
            UploadScript(clientContext, library, "copypasteimages_resources.en-US.js");
            UploadScript(clientContext, library, "copypasteimages_resources.pt-br.js");
            UploadImage(clientContext, library, "loading.gif");
        }

        private static void UploadScript(ClientContext clientContext, List library, string fileName)
        {
            UploadAsset(clientContext, library, fileName, "Scripts");
        }

        private static void UploadImage(ClientContext clientContext, List library, string fileName)
        {
            UploadAsset(clientContext, library, fileName, "Images");
        }

        private static void UploadAsset(ClientContext clientContext, List library, string fileName, string folder)
        {
            var filePath = System.Web.Hosting.HostingEnvironment.MapPath(string.Format("~/{0}/{1}", folder, fileName));
            var newFile = new FileCreationInformation();
            newFile.Content = System.IO.File.ReadAllBytes(filePath);
            newFile.Url = fileName;
            newFile.Overwrite = true;
            var uploadFile = library.RootFolder.Files.Add(newFile);
            clientContext.Load(uploadFile);
            clientContext.ExecuteQuery();
        }

        private static void AddScriptLinksToHostWeb(ClientContext clientContext)
        {
            var existingActions = clientContext.Web.UserCustomActions;
            clientContext.Load(existingActions);
            clientContext.ExecuteQuery();

            RemoveScriptLinksFromHostWeb(clientContext, existingActions);
            
            var customActionJQuery = existingActions.Add();
            customActionJQuery.Description = "copyPasteImagesJQuery";
            customActionJQuery.Location = "ScriptLink";
            customActionJQuery.ScriptSrc = "~site/SiteAssets/jquery-1.9.1.min.js";
            customActionJQuery.Sequence = 1000;
            customActionJQuery.Update();

            var copyPasteImages = existingActions.Add();
            copyPasteImages.Description = "copyPasteImagesScript";
            copyPasteImages.Location = "ScriptLink";
            copyPasteImages.ScriptSrc = "~site/SiteAssets/copypasteimages.js";
            copyPasteImages.Sequence = 1010;
            copyPasteImages.Update();
            
            clientContext.ExecuteQuery();
        }

        private static void RemoveScriptLinksFromHostWeb(ClientContext clientContext, UserCustomActionCollection existingActions)
        {
            var actions = existingActions.ToArray();
            foreach (var action in actions)
            {
                if (action.Location.Equals("ScriptLink") &&
                    (action.Description.Equals("copyPasteImagesJQuery") || action.Description.Equals("copyPasteImagesScript")))
                {
                    action.DeleteObject();
                }
            }

            clientContext.ExecuteQuery();
        }
    }
}