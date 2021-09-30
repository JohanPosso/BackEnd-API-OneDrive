
var markdownEditor = {
    applicationId: "498c4f17-92a6-4cba-818c-2eb8249c46d5",
    defaultFileName: "MD file.md",
    microsoftGraphApiRoot: "https://graph.microsoft.com/v1.0/",

 
    openFile: function () {
        var editor = this;

        var options = {
           
            clientId: editor.applicationId,
            action: "download",
            multiSelect: false,
            advanced: {
               
                filter: ".md,.mdown,.txt",
              
                queryParameters: "select=*,name,size",
                
                scopes: ["Files.ReadWrite openid User.Read"]
            },
            success: function (files) {
                
                editor.accessToken = files.accessToken;
                
                
                var selectedFile = files.value[0];
                editor.openItemInEditor(selectedFile);
            },
            error: function (e) { editor.showError("Error occurred while picking a file: " + e, e); },
        };
        OneDrive.open(options);
    },

    openItemInEditor: function (fileItem) {
        var editor = this;
        editor.lastSelectedFile = fileItem;
        var downloadLink = fileItem["@microsoft.graph.downloadUrl"];
        
        $.ajax(downloadLink, {
            success: function (data, status, xhr) {
                
                editor.setEditorBody(xhr.responseText);
                editor.setFilename(fileItem.name);
                $("#canvas").attr("disabled", false);
                editor.openFileID = fileItem.id;
            }, 
            error: function(xhr, status, err) {
                editor.showError(err);
            }
        });
    },

    saveAsFile: function () {
        var editor = this;
        var filename = "";

        if (editor.lastSelectedFile) {
            filename = editor.lastSelectedFile.name;
        }
        if (!filename || filename == "") {
            filename = editor.defaultFileName;
        }

        var options = {
            clientId: editor.applicationId,
            action: "query",
            advanced: {
                queryParameters: "select=id,name,parentReference"
            },
            success: function (selection) {
                var folder = selection.value[0]; 
                
                
                if (!editor.lastSelectedFile) {
                    editor.lastSelectedFile = { 
                        id: null,
                        name: filename,
                        parentReference: {
                            driveId: folder.parentReference.driveId,
                            id: folder.id
                        }
                    }
                } else {
                    editor.lastSelectedFile.parentReference.driveId = folder.parentReference.driveId;
                    editor.lastSelectedFile.parentReference.id = folder.id
                }
                
                editor.accessToken = selection.accessToken;

                editor.saveFileWithAPI( { uploadIntoParentFolder: true });
            },
            error: function (e) { editor.showError("An error occurred while saving the file: " + e, e); }
        };

        OneDrive.save(options);
    },


    saveFile: function () {
        var editor = this;
        if (editor.openFileID == "") {
            editor.saveAsFile();
            return;
        }

        editor.saveFileWithAPI();
    },


    saveFileWithAPI: function (state) {
        var editor = this;
        if (editor.accessToken == null) {
            editor.showError("Unable to save file due to an authentication error. Try using Save As instead.");
            return;
        }

        
        var url = editor.generateGraphUrl(editor.lastSelectedFile, (state && state.uploadIntoParentFolder) ? true : false, "/content");

        var bodyContent = $("#canvas").val().replace(/\r\n|\r|\n/g, "\r\n");

        $.ajax(url, {
            method: "PUT",
            contentType: "application/octet-stream",
            data: bodyContent,
            processData: false,
            headers: { Authorization: "Bearer" + editor.accessToken },
            success: function(data, status, xhr) {
                if (data && data.name && data.parentReference) {
                    editor.lastSelectedFile = data;
                    editor.showSuccessMessage("File was saved.");
                }
            },
            error: function(xhr, status, err) {
                editor.showError(err);
            }
        });
    },
   

   
    renameFile: function () {
        var editor = this;

        var oldFilename = (editor.lastSelectedFile && editor.lastSelectedFile.name) ? editor.lastSelectedFile.name : editor.defaultFileName;
        var newFilename = window.prompt("Rename file", oldFilename);
        if (!newFilename) return;
        
        editor.setFilename(newFilename);

        if (editor.lastSelectedFile && editor.lastSelectedFile.id) {
            editor.lastSelectedFile.name = newFilename;
            editor.patchDriveItemWithAPI({ propertyList: [ "name" ]} );
        } else {
        }
    },

    patchDriveItemWithAPI: function(state) {
        var editor = this;
        if (editor.accessToken == null) {
            editor.showError("Unable to save file due to an authentication error. Try using Save As instead.");
            return;
        }
        
        if (state == null) {
            editor.showError("The state parameter is required for this method.");
            return;
        }

        var item = editor.lastSelectedFile;
        var propList = state.propertyList;
       
        var url = editor.generateGraphUrl(item, false, null);

        var patchData = { };
        for(var i=0, len = propList.length; i < len; i++)
        {
            patchData[propList[i]] = item[propList[i]];
        }

        $.ajax(url, {
            method: "PATCH",
            contentType: "application/json; charset=UTF-8",
            data: JSON.stringify(patchData),
            headers: { Authorization: "Bearer" + editor.accessToken },
            success: function(data, status, xhr) {
                if (data && data.name && data.parentReference) {
                    editor.showSuccessMessage("File was updated successfully.");
                    editor.lastSelectedFile = data;
                }
            },
            error: function(xhr, status, err) {
                editor.showError("Unable to patch file metadata: " + err);
            }
        });
    },
    
    shareFile: function () {
        var editor = this;
        if (!editor.lastSelectedFile || !editor.lastSelectedFile.id)
        {
            editor.showError("You need to save the file first before you can share it.");
            return;
        }

        editor.getSharingLinkWithAPI();
    },

    getSharingLinkWithAPI: function() {
        var editor = this;
        var driveItem = editor.lastSelectedFile;

        var url = editor.generateGraphUrl(driveItem, false, "/createLink");
        var requestBody = { type: "view" };

        $.ajax(url, {
            method: "POST",
            contentType: "application/json; charset=UTF-8",
            data: JSON.stringify(requestBody),
            headers: { Authorization: "Bearer" + editor.accessToken },
            success: function(data, status, xhr) {
                if (data && data.link && data.link.webUrl) {
                    window.prompt("View-only sharing link", data.link.webUrl);
                } else {
                    editor.showError("Unable to retrieve a sharing link for this file.");
                }
            },
            error: function(xhr, status, err) {
                editor.showError("Unable to retrieve a sharing link for this file.");
            }
        });
    },


    generateGraphUrl: function(driveItem, targetParentFolder, itemRelativeApiPath) {
        var url = this.microsoftGraphApiRoot;
        if (targetParentFolder)
        {
            url += "drives/" + driveItem.parentReference.driveId + "/items/" +driveItem.parentReference.id + "/children/" + driveItem.name;
        } else {
            url += "drives/" + driveItem.parentReference.driveId + "/items/" + driveItem.id;
        }

        if (itemRelativeApiPath) {
            url += itemRelativeApiPath;
        }
        return url;
    },


    createNewFile: function () {
        this.lastSelectedFile = null;
		this.openFileID = "";
        this.setFilename(this.defaultFileName);
        $("#canvas").attr("disabled", false);
        this.setEditorBody("");
    },


    setFilename: function (filename) {
        var btnRename = this.buttons["rename"];
        if (btnRename) {
            $(btnRename).text(filename);
        }
    },

    setEditorBody: function (text) {
        $("#canvas").val(text);
    },

    buttons: {},
    wireUpCommandButton: function(element, cmd)
    {
        this.buttons[cmd] = element;
        switch(cmd) {
            case "new": 
                element.onclick = function () { markdownEditor.createNewFile(); return false; }
                break;
            case "open": 
                element.onclick = function () { markdownEditor.openFile(); return false; }
                break;
            case "save":
                element.onclick = function () { markdownEditor.saveFile(); return false; }
                break;
            case "saveAs":
                element.onclick = function () { markdownEditor.saveAsFile(); return false; }
                break;
            case "rename":
                element.onclick = function () { markdownEditor.renameFile(); return false; }
                break;
            case "share":
                element.onclick = function () { markdownEditor.shareFile(); return false; }
                break;
        }
    },

    showError: function (msg, e) {
        window.alert(msg);
    },

    showSuccessMessage: function(msg) {
        window.alert(msg);
    },

    lastSelectedFile: null,

    accessToken: null,

    user: {
        id: "nouser@contoso.com",
        domain: "organizations"
    }
}

$("#canvas").attr("disabled", true);
