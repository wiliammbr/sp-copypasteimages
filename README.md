### Summary ###
This sample shows how to add the functionality of uploading images to SharePoint pages. 

To set up this sample, we used a provider-hosted add-in using .NET CSOM that adds some files and script links to the Host Web.

### Applies to ###
-  Office 365 Multi Tenant (MT)
-  Office 365 Dedicated (D)
-  SharePoint 2013 on-premises
-  SharePoint 2016 on-premises

### Prerequisites ###
It's important that the SharePoint site has the Publishing Feature enabled.

### Solution ###
Solution | Author(s)
---------|----------
CopyPasteImages | Wiliam Rocha

### Version history ###
Version  | Date | Comments
---------| -----| --------
1.0  | May 01st 2017 | Initial release

### Disclaimer ###
**THIS CODE IS PROVIDED *AS IS* WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**


----------

# Step 1: Set up a Publishing Images library #
The first step is to set up a Publishing Images library where the images will be saved. It must be placed on the same web of the Pages library. 

# Step 2: Add scripts #
To demonstrate the functionality, you must add the ScriptLinks to the Host Web. This is all done using .NET CSOM. If building a custom branding solution with master page the links could of course be added directly to the master page.

```javascript
var copyPasteImages = existingActions.Add();
copyPasteImages.Description = "copyPasteImagesScript";
copyPasteImages.Location = "ScriptLink";
copyPasteImages.ScriptSrc = "~site/SiteAssets/copypasteimages.js";
copyPasteImages.Sequence = 1010;
copyPasteImages.Update();
            
```

# Copy Paste Functionality #
The functionality will be added to the Page Body's fields on your page. So you can just copy an image from your photo viewer or editor.

![Screenshot of navigation](http://i.imgur.com/760PG9b.png "Screenshot of an image being copied")

And then, just paste it on the body of your Page:

![Screenshot of navigation](http://i.imgur.com/DGVN6h7.png "Screenshot of an image being pasted")

The end result should look like this:

![Screenshot of navigation](http://i.imgur.com/wxOQ8rF.png "Screenshot of a pasted image")
