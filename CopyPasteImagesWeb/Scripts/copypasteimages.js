"use strict";

(function ($) {
    // Copy Paste info
    var editMode = false;
    var editingContainer = null;
    var ieClipboardDiv = null;
    var pasteRequests = [];
    var pasteDeferred = null;
    // SharePoint
    var context = null;
    var currentSiteUrl = '';
    // Utils
    var isIe = (navigator.userAgent.toLowerCase().indexOf("msie") !== -1 || navigator.userAgent.toLowerCase().indexOf("trident") !== -1);
    var isSafari = !isIe && navigator.appVersion.search("Safari") !== -1 && navigator.appVersion.search("Chrome") == -1 && navigator.appVersion.search("CrMo") == -1 && navigator.appVersion.search("CriOS") == -1;
    var imageMimeRegex = /^image\/(p?jpeg|gif|png)$/i;
    // Default Messages
    var loadingImage = "Loading image...";
    var waitASecond = "Please, just wait a few moments...";
    var requestFailed = "Request failed. ";
    var imagesListInternalName = "PublishingImages";
    var imagesListTitle = "Images";

    // Ready function after jQuery loads
    var copyPasteReady = function () {
        getEditingContainer()
            .then(getUserLanguage, onError)
            .then(defineLocalePreferences, onError)
            .then(enableCopyPaste, onError);
    };

    // Gets the editing container where the user will insert the images
    // It also loads a lot of other javascript files
    var getEditingContainer = function () {
        var deferred = new $.Deferred();
        editingContainer = $("div[id$=RichHtmlField_displayContent]");
        if (editingContainer.length > 0) {
            editMode = true;
            SP.SOD.executeOrDelayUntilScriptLoaded(function () {
                SP.SOD.executeFunc("SP.js", "SP.ClientContext");
                SP.SOD.executeOrDelayUntilScriptLoaded(function () {
                    SP.SOD.registerSod("sp.userprofiles.js", SP.Utilities.Utility.getLayoutsPageUrl("sp.userprofiles.js"));
                    SP.SOD.executeFunc("sp.userprofiles.js", false, function () {
                        deferred.resolve();
                    });
                }, "sp.js");
            }, "core.js");
        } else {
            editMode = false;
            deferred.resolve();
        }
        return deferred.promise();
    };

    // Get the user language preferences
    var getUserLanguage = function () {
        var deferred = new $.Deferred();
        if (editMode) {
            var targetUser = "i:0#.f|membership|" + _spPageContextInfo.userLoginName;
            context = new SP.ClientContext.get_current();
            var peopleManager = new SP.UserProfiles.PeopleManager(context);
            var userProperty = peopleManager.getUserProfilePropertyFor(targetUser, "SPS-MUILanguages");
            context.executeQueryAsync(function () {
                var language = "en-US";
                var userProfileLanguage = userProperty.m_value.split(",")[0];
                if (userProfileLanguage) {
                    language = userProfileLanguage;
                }
                deferred.resolve(language);
            }, function (sender, args) {
                deferred.reject(sender, args);
            });
        } else {
            deferred.resolve();
        }
        return deferred.promise();
    };

    // Loads the resources file based on the language
    var defineLocalePreferences = function (language) {
        var deferred = $.Deferred();

        if (editMode) {
            var scriptUrl = "";
            var scriptRevision = "";
            $("script").each(function (i, el) {
                if (el.src.toLowerCase().indexOf("copypasteimages.js") > -1) {
                    scriptUrl = el.src;
                    scriptRevision = scriptUrl.substring(scriptUrl.indexOf(".js") + 3);
                    scriptUrl = scriptUrl.substring(0, scriptUrl.indexOf(".js"));
                }
            });

            // Load translation files
            var resourcesFile = scriptUrl + "_resources." + language + ".js";
            if (scriptRevision.length > 0) {
                resourcesFile += scriptRevision;
            }

            loadScript(resourcesFile, function () {
                deferred.resolve();
            });
        } else {
            deferred.resolve();
        }

        return deferred;
    };

    // Attaches the event of pasting images
    var enableCopyPaste = function () {
        if (editMode) {
            currentSiteUrl = location.protocol + "//" + location.host + L_Menu_BaseUrl;
            $(document).on('paste', function (e) {
                if (isIe) {
                    ieClipboardEvent(e);
                } else {
                    standardClipboardEvent(e);
                }
            });

            if (isIe) {
                var contentDiv = jQuery("#DeltaPlaceHolderMain");
                if (contentDiv.length == 0)
                    contentDiv = jQuery("#ctl00_MSO_ContentDiv");
                contentDiv.append("<div id='ie-clipboard-contenteditable' class='hidden' contenteditable='true'></div > ");
                ieClipboardDiv = $("#ie-clipboard-contenteditable");
            }
        }
    };

    var ieClipboardEvent = function (e) {
        var nid = SP.UI.Notify.addNotification("<img src='/_layouts/15/images/loadingcirclests16.gif?rev=23' style='vertical-align:bottom; display:inline-block; margin-" + (document.documentElement.dir == "rtl" ? "left" : "right") + ":2px;' />&nbsp;<span style='vertical-align:top;'>" + loadingImage + "</span>", true);

        if (clipboardData) {
            var clipboardText = clipboardData.getData("Text");

            if (!clipboardText) {
                focusIeClipboardDiv();
                ieClipboardDiv.empty();

                setTimeout(function () {
                    var content = ieClipboardDiv.html();
                    var type = content.substring(content.indexOf(":") + 1, content.indexOf(";"));

                    if (imageMimeRegex.test(type)) {
                        publishToImages(type, content, nid);

                        //console.log("Clipboard Plain Text: " + clipboardText);
                        //console.log("Clipboard HTML: " + ieClipboardDiv.html());
                    } else {
                        setTimeout(function () {
                            var div = document.createElement("div");
                            div.innerHTML = content || clipboardText;
                            var curRange = window.getSelection().getRangeAt(0);
                            if (curRange) {
                                curRange.deleteContents();
                                curRange.insertNode(div);
                            }
                            SP.UI.Notify.removeNotification(nid);
                        }, 0);
                    }
                    ieClipboardDiv.empty();
                }, 0);
            } else {
                setTimeout(function () {
                    var div = document.createElement("div");
                    div.innerHTML = content || clipboardText;
                    var range = window.getSelection().getRangeAt(0);
                    if (range) {
                        range.deleteContents();
                        range.insertNode(div);
                    }
                    SP.UI.Notify.removeNotification(nid);
                }, 0);
            }
        }
    };

    var standardClipboardEvent = function (e) {
        if (e && e.originalEvent.clipboardData) {
            pasteDeferred = $.Deferred();
            var items = e.originalEvent.clipboardData.items;
            var obj = null;
            for (var i = 0; i < items.length; i++) {
                obj = items[i];
                if (imageMimeRegex.test(obj.type)) {
                    loadImage(obj);
                    return;
                }
            }
        }
    };

    var publishToImages = function (fileType, content, nid) {
        var imageData = getImageDataFromBase64(content);
        var image = new Image();
        image.onload = function () {
            setTimeout(function () {
                var div = document.createElement("div");
                div.id = "pasteContainer";
                div.style.width = image.width + "px";
                div.style.height = image.height + "px";
                div.style.position = "relative";

                div.innerHTML = "<div id='loading' class='boxLoading'> \
                                            <div class='boxImageLoading' style='position: absolute;'> \
                                                <img style='width: 50%; margin: auto; padding: 5px;' src='/SiteAssets/loading.gif' alt='" + waitASecond + "'> \
                                            </div> \
                                            <div class='boxFormFade'> \
                                            </div> \
                                        </div>";
                var range = window.getSelection().getRangeAt(0);
                if (range) {
                    range.deleteContents();
                    range.insertNode(div);
                }
                SP.UI.Notify.removeNotification(nid);
            }, 0);
        };
        image.src = imageData;

        instantiateTargetFolder(imagesListTitle, document.title)
            .then(uploadImage.bind(null, fileType, content), onError);
    };

    var uploadImage = function (fileType, content, parentFolder) {
        var guid = newGuid();
        var fileName = document.title + "_" + guid + "." + fileType.split("/")[1];
        var destinationUrl = currentSiteUrl + "/" + imagesListInternalName + "/" + document.title + "/" + fileName;

        uploadFile(content, fileName, destinationUrl, function (xData, status) {
            pasteRequests.push({
                time: new Date(),
                data: xData,
                statusResult: status
            });

            if (status == "success") {
                setTimeout(function () {
                    var img = document.createElement("img");
                    img.src = destinationUrl;
                    var range = window.getSelection().getRangeAt(0);
                    if (range) {
                        var pasteContainer = document.getElementById("pasteContainer");
                        pasteContainer.focus();
                        var newRange = document.createRange();
                        newRange.selectNodeContents(pasteContainer);
                        newRange.deleteContents();
                        newRange.insertNode(img);

                        var selection = window.getSelection();
                        selection.removeAllRanges();
                        selection.addRange(newRange);

                        if (pasteContainer && pasteContainer.parentElement) {
                            var spanImage = document.createElement("span");
                            spanImage.innerHTML = pasteContainer.innerHTML;
                            pasteContainer.parentNode.insertBefore(spanImage, pasteContainer.nextSibling);
                            pasteContainer.parentElement.removeChild(pasteContainer);
                        }
                    }
                }, 0);
            }
            if (pasteDeferred) {
                pasteDeferred.resolve();
            }
        });
    };

    var loadScript = function (url, callback) {
        var head = document.getElementsByTagName("head")[0];
        var script = document.createElement("script");
        script.src = url;

        // Attach handlers for all browsers
        var done = false;
        script.onload = function () {
            if (!done && (!this.readyState
                || this.readyState === "loaded"
                || this.readyState === "complete")) {
                done = true;

                // Continue your code
                callback();

                // Handle memory leak in IE
                script.onload = null;
                script.onreadystatechange = null;
                head.removeChild(script);
            }
        };
        script.onreadystatechange = script.onload;

        head.appendChild(script);
    };

    // Chrome paste handler
    var loadImage = function (fileObject) {
        var nid = SP.UI.Notify.addNotification("<img src='/_layouts/15/images/loadingcirclests16.gif?rev=23' style='vertical-align:bottom; display:inline-block; margin-" + (document.documentElement.dir == "rtl" ? "left" : "right") + ":2px;' />&nbsp;<span style='vertical-align:top;'>" + loadingImage + "</span>", true);
        var reader = new FileReader();
        var file = fileObject.getAsFile();
        reader.onloadend = function () {
            publishToImages(file.type, reader.result, nid);
        };
        reader.readAsDataURL(file);
    };

    var focusIeClipboardDiv = function () {
        ieClipboardDiv.focus();
        var range = document.createRange();
        range.selectNodeContents(ieClipboardDiv.get(0));
        var selection = window.getSelection();
        selection.removeAllRanges();
        selection.addRange(range);
    };

    var getSelectedNode = function () {
        var selectedNode = null;
        if (document.selection) {
            selectedNode = document.selection.createRange().parentElement();
        } else {
            var selection = window.getSelection();
            if (selection && selection.rangeCount > 0) {
                selectedNode = selection.getRangeAt(0).startContainer.parentNode;
            }
        }
        return selectedNode;
    };

    var instantiateTargetFolder = function (listTitle, folderUrl) {
        var deferred = $.Deferred();
        var currentContext = SP.ClientContext.get_current();
        var list = currentContext.get_web().get_lists().getByTitle(listTitle);
        var rootFolder = list.get_rootFolder();
        var folderContext = rootFolder.get_context();
        var folderNames = folderUrl.split("/");
        var folderName = folderNames[0];
        var currentFolder = rootFolder.get_folders().add(folderName);
        folderContext.load(currentFolder);
        folderContext.executeQueryAsync(
            function () {
                if (folderNames.length > 1) {
                    var subFolderUrl = folderNames.slice(1, folderNames.length).join("/");
                    createFolderInternal(currentFolder, subFolderUrl, success, error);
                }
                deferred.resolve();
            },
            function (sender, args) {
                deferred.reject(sender, args);
            }
        );
        return deferred.promise();
    };

    var uploadFile = function (contentFile, fileName, destinationUrl, processResult) {
        var soapEnv =
            "<soap:Envelope xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' xmlns:xsd='http://www.w3.org/2001/XMLSchema' xmlns:soap='http://schemas.xmlsoap.org/soap/envelope/'> \
                <soap:Body>\
                    <CopyIntoItems xmlns='http://schemas.microsoft.com/sharepoint/soap/'>\
                        <SourceUrl>test.txt</SourceUrl>\
                        <DestinationUrls>\
                            <string>" + destinationUrl + "</string>\
                        </DestinationUrls>\
                        <Fields>\
                            <FieldInformation Type='Text' DisplayName='Title' InternalName='Title' Value='" + fileName + "' />\
                        </Fields>\
                        <Stream>" + getContentFromBase64(contentFile) + "</Stream>\
                    </CopyIntoItems>\
                </soap:Body>\
            </soap:Envelope>";

        $.ajax({
            url: currentSiteUrl + "/_vti_bin/copy.asmx",
            beforeSend: function (xhr) { xhr.setRequestHeader("SOAPAction", "http://schemas.microsoft.com/sharepoint/soap/CopyIntoItems"); },
            type: "POST",
            dataType: "xml",
            data: soapEnv,
            complete: processResult,
            contentType: "text/xml; charset=\"utf-8\""
        });
    };

    var getContentFromBase64 = function (content) {
        if (content.indexOf(",") > -1) {
            content = content.substring(content.indexOf(",") + 1);
        }
        if (content.indexOf("\"") > -1) {
            content = content.substring(0, content.indexOf("\""));
        }
        return content;
    };

    var getImageDataFromBase64 = function (content) {
        if (content.indexOf("src=\"") > -1) {
            content = content.substring(content.indexOf("\"") + 1);
        }
        if (content.indexOf("\"") > -1) {
            content = content.substring(0, content.indexOf("\""));
        }
        return content;
    };

    var newGuid = function () {
        function s4() { return Math.floor((1 + Math.random()) * 0x10000).toString(16).substring(1); }
        return s4() + s4() + "-" + s4() + "-" + s4() + "-" + s4() + "-" + s4() + s4() + s4();
    };

    var onError = function (sender, args) {
        alert(requestFailed + args.get_message() + "\n" + args.get_stackTrace());
    };

    $(document).on("ready", function (e) {
        copyPasteReady();
    });
})(jQuery);