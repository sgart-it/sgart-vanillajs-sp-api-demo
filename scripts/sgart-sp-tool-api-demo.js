(function () {
    /* 
        SharePoint Tool Api Demo (Sgart.it)
        javascript:(function(){var s=document.createElement('script');s.src='/SiteAssets/ToolApiDemo/sgart-sp-tool-api-demo.js?t='+(new Date()).getTime();document.head.appendChild(s);})();
     */
    const VERSION = "1.2025-11-10";

    const LOG_SOURCE = "Sgart.it:SharePoint API Demo:";

    const HTML_ID_WRAPPER = "sgart-content-wrapper";

    const HTML_ID_PUPUP = "sgart-popup";

    const HTML_ID_BTN_EXIT = "sgart-btn-exit";
    const HTML_ID_BTN_CLEAR_OUTPUT = "sgart-btn-clear-output";
    const HTML_ID_BTN_COPY_OUTPUT = "sgart-btn-copy-output";
    const HTML_ID_BTN_EXAMPLES = "sgart-btn-examples";
    const HTML_ID_BTN_HISTORY = "sgart-btn-history";
    const HTML_ID_LBL_COUNT = "sgart-label-count";
    const HTML_ID_HTTP_STATUS = "sgart-http-status";

    const HTML_ID_OUTPUT_RAW = "sgart-output-raw";
    const HTML_ID_OUTPUT_SIMPLE = "sgart-output-simple";
    const HTML_ID_OUTPUT_TABLE = "sgart-output-table";

    const HTML_ID_BTN_EXECUTE = "sgart-btn-execute";
    const HTML_ID_TXT_INPUT = "sgart-txt-input";
    const HTML_ID_SELECT_ODATA = "sgart-select-odata";

    const HTML_ID_TAB_RAW = "sgart-tab-response-raw";
    const HTML_ID_TAB_SIMPLE = "sgart-tab-response-simple";
    const HTML_ID_TAB_TABLE = "sgart-tab-response-table";

    const TAB_KEY_RAW = 'raw';
    const TAB_KEY_SIMPLE = 'symple';
    const TAB_KEY_TABLE = 'table';
    let currentTab = TAB_KEY_SIMPLE;
    let serverRelativeUrlPrefix = "/";

    const EXAMPLES = {
        groups: [
            {
                id: "site",
                title: "Site",
                actions: [
                    {
                        id: "getSite",
                        title: "Get site",
                        url: "site"
                    },
                    {
                        id: "getSiteId",
                        title: "Get site id",
                        url: "site/id"
                    }
                ]
            },
            {
                id: "web",
                title: "Web",
                actions: [
                    {
                        id: "getWeb",
                        title: "Get web",
                        url: "web"
                    },
                    {
                        id: "getWebById",
                        title: "Get sub webs",
                        url: "web/webs"
                    },
                    {
                        id: "getWebSiteUsers",
                        title: "Get site users",
                        url: "web/siteusers"
                    },
                    {
                        id: "getWebSiteGrous",
                        title: "Get site groups",
                        url: "web/sitegroups"
                    },
                    {
                        id: "getWebRoleDefinitions",
                        title: "Get site role definitions",
                        url: "web/roledefinitions"
                    }
                ]
            },
            {
                id: "user",
                title: "User",
                actions: [
                    {
                        id: "getCurrentUser",
                        title: "Get current user",
                        url: "web/CurrentUser"
                    },
                    {
                        id: "getUsers",
                        title: "Get users",
                        url: "web/SiteUsers"
                    },
                    {
                        id: "getUserById",
                        title: "Get user by id",
                        url: "web/GetUserById(1)"
                    },
                    {
                        id: "getGroups",
                        title: "Get groups",
                        url: "web/sitegroups"
                    },
                    {
                        id: "getGroupsMembers",
                        title: "Get group members",
                        url: "web/sitegroups(1)/users"
                    },
                ]
            },
            {
                id: "list",
                title: "List",
                actions: [
                    {
                        id: "getLists",
                        title: "Get lists",
                        url: "web/lists?$select=Id,Title,BaseType,ItemCount,EntityTypeName,Hidden,LastItemUserModifiedDate&$top=100&$orderBy=Title"
                    },
                    {
                        id: "getListByGuid",
                        title: "Get list by guid",
                        url: "web/lists(guid'00000000-0000-0000-0000-000000000000')"
                    },
                    {
                        id: "getListByTitle",
                        title: "Get list by title",
                        url: "web/lists/getbytitle('Documents')"
                    },
                    {
                        id: "getListByTitleRootFolder",
                        title: "Get list by title RootFolder",
                        url: "web/lists/getbytitle('Documents')/RootFolder"
                    },
                    {
                        id: "getFields",
                        title: "Get fields",
                        url: "web/lists/getbytitle('Documents')/fields"
                    },
                    {
                        id: "getViews",
                        title: "Get views",
                        url: "web/lists/getbytitle('Documents')/views?$top=100&$orderBy=Title"
                    },
                    {
                        id: "getViewsHidden",
                        title: "Get views hidden",
                        url: "web/lists/getbytitle('Documents')/views?$select=Id,Title,ServerRelativeUrl,ViewQuery&$top=100&$orderBy=Title&$filter=Hidden eq true"
                    },
                    {
                        id: "getContenttypes",
                        title: "Get content types",
                        url: "web/lists/getbytitle('Documents')/contenttypes?$select=*&$orderBy=Name"
                    },
                ]
            },
            {
                id: "item",
                title: "Item",
                actions: [
                    {
                        id: "getItems",
                        title: "Get items",
                        url: "web/lists/getbytitle('Documents')/items?$top=10&$orderBy=Id desc"
                    },
                    {
                        id: "getItemById",
                        title: "Get item by id",
                        url: "web/lists/getbytitle('Documents')/items(1)"
                    },
                    {
                        id: "getItemByIdWithSelect",
                        title: "Get item by id with select and expand",
                        url: "web/lists/getbytitle('Documents')/items(1)?$select=Title,Id,Created,Modified,Author/Title,Editor/Title&$expand=Author,Editor"
                    },
                    {
                        id: "getItemAttachments",
                        title: "Get item attachments",
                        url: "web/lists/getbytitle('Documents')/items(1)/AttachmentFiles",
                    }
                ]
            },
            {
                id: "file",
                title: "File",
                actions: [
                    {
                        id: "getFileById",
                        title: "Get file by id",
                        url: "web/getfilebyid('00000000-0000-0000-0000-000000000000')"
                    },
                    {
                        id: "getFileByServerRelativeUrl",
                        title: "Get file by server relative url",
                        url: "web/getfilebyserverrelativeurl('/sites/someSite/Shared Documents/file.txt')"
                    },
                    {
                        id: "getFileContent",
                        title: "Get file content",
                        url: "web/getfilebyserverrelativeurl('/sites/someSite/Shared Documents/file.txt')/$value"
                    }
                ]
            },
            {
                id: "folder",
                title: "Folder",
                actions: [
                    {
                        id: "getFolderByServerRelativeUrl",
                        title: "Get folder by server relative url",
                        url: "web/getfolderbyserverrelativeurl('Shared Documents')"
                    },
                    {
                        id: "getFolderById",
                        title: "Get folder by id",
                        url: "web/getfolderbyid('00000000-0000-0000-0000-000000000000')"
                    },
                    {
                        id: "getFolderFiles",
                        title: "Get folder files",
                        url: "web/getfolderbyserverrelativeurl('Shared Documents')/files"
                    },
                    {
                        id: "getFolderFileContent",
                        title: "Get folder files",
                        url: "web/getfolderbyserverrelativeurl('/sites/someSite/Shared Documents/file name')/$value"
                    }
                ]
            },
            {
                id: "search",
                title: "Search",
                actions: [
                    {
                        id: "searchSites",
                        title: "Search sites",
                        url: "search/query",
                        query: {
                            mode: "search",
                            filter: "sharepoint (contentclass:STS_Site) Path:\"https://sgart.sharepoint.com/*\"",
                            select: "Title,Path,Description,SiteLogo,WebTemplate,WebId,SiteId,Created,LastModifiedTime"
                        }
                    }
                ]
            },
            {
                id: "userProfile",
                title: "User Profile",
                actions: [
                    {
                        id: "getPMInstance",
                        title: "PeopleManager instance",
                        url: "SP.UserProfiles.PeopleManager"
                    },
                    {
                        id: "getPMFollowedByMe",
                        title: "Followed by ME",
                        url: "SP.UserProfiles.PeopleManager/getpeoplefollowedbyme",
                        query: {
                            select: "*"
                        }
                    },
                    {
                        id: "getPMFollowedBy",
                        title: "Followed by ...",
                        url: "SP.UserProfiles.PeopleManager/getpeoplefollowedby(@v)?@v='i%3A0%23.f%7Cmembership%7Cuser%40domain.onme",
                        query: {
                            select: "*"
                        }
                    }


                ]
            },
            {
                id: "taxonomy",
                title: "Taxonomy",
                actions: [
                    {
                        id: "getTermStoreGroups",
                        title: "Get groups",
                        url: "v2.1/termStore/groups"
                    },
                    {
                        id: "getTermStoreTermGroups",
                        title: "Get term groups",
                        url: "v2.1/termStore/termGroups"
                    },
                    {
                        id: "getTermStoreSets",
                        title: "Get term sets",
                        url: "v2.1/termStore/groups/{groupid}/sets"
                    },
                    {
                        id: "getTermStoreSetById",
                        title: "Get terms by set id",
                        url: "v2.1/termStore/groups/{groupid}/sets/{termSetId}/terms"
                    },
                    {
                        id: "getTermById",
                        title: "Get terms by set id next level",
                        url: "v2.1/termStore/groups/{groupid}/sets/{termSetId}/terms/{termId}/terms"
                    }
                ]
            },
            {
                id: "webhooks",
                title: "Webhooks",
                actions: [
                    {
                        id: "getWebhooks",
                        title: "Get webhooks",
                        url: "web/lists/getbytitle('Documents')/subscriptions"
                    }
                ]
            },
            {
                id: "tenant",
                title: "Tenant",
                actions: [
                    {
                        id: "getTenantAppCatalog",
                        title: "Get tenant app catalog",
                        url: "SP_TenantSettings_Current"
                    }
                ]
            }
        ]
    };


    // encode dei caratteri in html
    String.prototype.htmlEncode = function () {
        const node = document.createTextNode(this);
        return document.createElement("a").appendChild(node).parentNode.innerHTML.replace(/'/g, "&#39;").replace(/"/g, "&#34;");
    };

    function copyToClipboard(text) {
        navigator.clipboard.writeText(text)
            .then(() => {
                console.log('Text copied to clipboard:', text);
            })
            .catch(err => {
                console.error('Failed to copy text: ', err);
            });
    }

    const simplifyObjectOrArray = (response) => {
        if (!response) {
            console.debug(LOG_SOURCE, 'response is undefined or null');
            return {};
        }

        if (response.value && Array.isArray(response.value)) {
            // console.debug(LOG_SOURCE, 'response.value is an array', response);
            return response.value
        }

        if (response.d) {
            if (response.d.results && Array.isArray(response.d.results)) {
                // console.debug(LOG_SOURCE, 'response.d.results is an array', response);
                return response.d.results;
            } else {
                // console.debug(LOG_SOURCE, 'response.d is a single object', response);
                return response.d;
            }
        }

        if (Array.isArray(response)) {
            // console.debug(LOG_SOURCE, 'response is an array', response);
            return response;
        }
        // console.debug(LOG_SOURCE, 'response is object', response);
        return response;
    };


    const htmlTableFromJson = (function () {
        const buildTableItem = (item) => {
            const table = {
                columns: [
                    {
                        key: 'internalName',
                        name: 'InternalName',
                        fieldName: 'internalName',
                        minWidth: 250,
                        isRowHeader: true,
                        isResizable: true
                    },
                    {
                        key: 'value',
                        name: 'Value',
                        fieldName: 'value',
                        minWidth: 450,
                        isResizable: true
                    }
                ],
                items: []
            };

            if (item) {
                table.items = Object.keys(item).map(key => {
                    const value = item[key];
                    if (typeof value === 'object') {
                        return { internalName: key, value: JSON.stringify(value, null, 2) };
                    }
                    return { internalName: key, value };
                });
            }
            return table;
        };

        const buildTableItems = (items) => {
            const table = {
                columns: [],
                items: []
            };

            if (!items || items.length === 0) {
                return table;
            }

            const item = items[0];
            if (item) {
                table.columns = Object.keys(item).map(key => ({
                    key,
                    name: key,
                    fieldName: key,
                    minWidth: 50,
                    isResizable: true
                }));

                const newItems = [];
                items.forEach(item => {
                    const ni = {};
                    Object.keys(item).forEach(key => {
                        const value = item[key];
                        ni[key] = typeof value === 'object'
                            ? JSON.stringify(value, null, 2)
                            : value;
                    });
                    newItems.push(ni);
                });
                table.items = newItems;

            }
            return table;
        };

        const buildTable = (items) => {
            const data = simplifyObjectOrArray(items);
            // console.debug(LOG_SOURCE, 'buildTable: items is object', data);
            if (!items) {
                // console.debug(LOG_SOURCE, 'buildTable: items is undefined or null');
                return { columns: [], items: [] };
            }
            if (Array.isArray(data)) {
                return buildTableItems(data);
            }
            return buildTableItem(data);
        };

        const renderTable = (table) => {
            let html = '<table border="1" cellpadding="5" cellspacing="0"><thead><tr>';
            table.columns.forEach(col => {
                html += `<th>${col.name}</th>`;
            });
            html += '</tr></thead><tbody>';
            table.items.forEach(item => {
                html += '<tr>';
                table.columns.forEach(col => {
                    html += `<td>${item[col.fieldName]}</td>`;
                });
                html += '</tr>';
            });
            html += '</tbody></table>';
            return html;
        };

        const buid = (json) => {
            const tableItems = buildTable(json);
            const html = renderTable(tableItems);
            return html;
        };

        return {
            buid
        };
    })();

    function getQueryParam(query) {
        if (!query) return "";
        const q = "?"
            + Object.keys(query)
                //.map(k => "$" + encodeURIComponent(k) + "=" + encodeURIComponent(query[k]))
                .map(k => "$" + k + "=" + query[k])
                .join("&");
        return q;
    }

    function injectStyle() {
        const css = `
            :root{
            --sgart-primary-color: #8a2e11;
            --sgart-primary-color-light: rgba(138,46,17,0.45);
            --sgart-primary-color-hover: #6f250e;
            --sgart-primary-color-dark: #4c1609;
            --sgart-secondary-color: #0a0a0a;
            --sgart-secondary-color-dark: #060606;
            --sgart-secondary-color-white: #ffffff;
            --sgart-secondary-color-gray-light: #e6e6e6;
            --sgart-btn-color-execute: #097BED;
            }
            .sgart-content-wrapper {
            font-family: Arial, sans-serif;
            border: 0;
            display: flex;
            flex-direction: column;
            position: fixed;
            top: 0;
            left: 0;
            right: 0;
            bottom: 0;
            background-color: var(--sgart-secondary-color-white);
            margin: 0;
            padding: 0;
            z-index: 10000;
            box-sizing: border-box;
            }   
            .sgart-content-wrapper input, .sgart-content-wrapper textarea, .sgart-content-wrapper select, .sgart-content-wrapper .sgart-button {
            font-family: Arial, sans-serif;
            font-size: 14px;
            height: 32px;
            padding: 0 10px;
            border: 1px solid var(--sgart-primary-color);
            background-color: var(--sgart-secondary-color-white);
            box-sizing: border-box;
            }
            .sgart-content-wrapper select {
            width: 180px;
            }
            .sgart-content-wrapper #sgart-api-demo {
            width: 200px;
            }
            .sgart-content-wrapper .sgart-button  {
            background-color: var(--sgart-primary-color);
            color: var(--sgart-secondary-color-white);
            padding: 0 10px;
            cursor: pointer;
            width: 120px;
            overflow: hidden;
            white-space: nowrap;
            display: flex;
            align-items: center;
            justify-content: center;
            gap: 5px;
            }
            .sgart-content-wrapper .sgart-button svg {
            color: var(--sgart-secondary-color-white);
            fill: var(--sgart-secondary-color-white);
            height: 20px;
            width: 20px;
            }
            .sgart-content-wrapper .sgart-button.sgart-button-tab {
            background-color: var(--sgart-primary-color-light);
            color: var(--sgart-secondary-color);
            }
            .sgart-content-wrapper .sgart-button.selected, .sgart-content-wrapper .sgart-button:hover, .sgart-content-wrapper .sgart-button.sgart-button-tab.selected, .sgart-content-wrapper .sgart-button.sgart-button-tab:hover {
            background-color: var(--sgart-primary-color-hover);
            color: var(--sgart-secondary-color-white);
            font-weight: bold;
            }
            #${HTML_ID_BTN_EXECUTE} {
            background-color: var(--sgart-btn-color-execute);
            }
            .sgart-button.sgart-button-tab:hover 
            {
            border-color: var(--sgart-secondary-color);
            }
            .sgart-content-wrapper .sgart-separator{
            margin: 0;
            }
            .sgart-header {
            background-color: var(--sgart-secondary-color);  
            color: white;
            padding: 5px 10px;
            border-bottom: 1px solid var(--sgart-secondary-color-gray-light);
            height: 40px;
            display: flex;
            flex-direction: row;
            align-items: center;
            justify-content: space-between;
            }       
            .sgart-header .sgart-button {
            background-color: var(--sgart-secondary-color);
            color: var(--sgart-secondary-color-white);
            padding: 0px 20px;
            }
            .sgart-header .logo {
            height: 33px;
            margin-right: 10px;
            }
                .sgart-toolbar {
                    display:flex;
                    flex-direction: row;
                    align-items: center;
                    justify-content: space-between;
                }
                .sgart-toolbar-left {
            display: flex;
            gap: 10px;
            justify-content: left;
            align-items: center;
            flex-wrap: wrap;
                }
                .sgart-toolbar-right{ 
                    justify-content: right;
                }
            .sgart-body {
            display: flex;
            flex-direction: column;
            flex-grow: 1;
            padding: 10px;
            gap: 10px;
            }   
            .sgart-input-area {
            display: flex;
            gap: 10px;
            align-items: center;
            justify-content: space-between;
            }
            .sgart-input {
            flex-grow: 1;   
            }
            .sgart-output-area {
            flex-grow: 1;   
            display: flex;
            overflow: hidden;
            position: relative;
            }
            .sgart-output-area > div {
            position: absolute;
            top: 0;
            left: 0;
            right: 0;
            bottom: 0;    
            overflow: auto;
            flex-grow: 1;   
            display: flex;
            box-sizing: border-box;
            border: 1px solid var(--sgart-primary-color);
            background-color: var(--sgart-secondary-color-white);
            }
            .sgart-output-area table {
            border-collapse: collapse;
            width: 100%;
            }
            .sgart-content-wrapper .sgart-output-txt, .sgart-content-wrapper .sgart-output-table {
            width: 100%;    
            height: 100%;
            flex-grow: 1;
            gap: 10px;
            font-family: monospace;
            resize: none;
            box-sizing: border-box;
            border: none;
            }
            .sgart-content-wrapper table th {
            background-color: var(--sgart-primary-color);
            color: var(--sgart-secondary-color-white);
            text-align: left;
            position: sticky;
            top: 0;
            z-index: 1000;
            padding: 5px;
            }
            .sgart-content-wrapper .sgart-http-status {
            border: 1px solid var(--sgart-secondary-color);
            padding: 5px 10px;
            background-color: var(--sgart-secondary-color-gray-light);
            color: var(--sgart-secondary-color);
            }
            .sgart-content-wrapper .sgart-http-status-100 { background-color: #e7f3fe; color: #31708f; border-color: #bce8f1; }  
            .sgart-content-wrapper .sgart-http-status-200 { background-color: #dff0d8; color: #3c763d; border-color: #d6e9c6; }
            .sgart-content-wrapper .sgart-http-status-300 { background-color: #fcf8e3; color: #8a6d3b; border-color: #faebcc; }
            .sgart-content-wrapper .sgart-http-status-400 { background-color: #f2dede; color: #a94442; border-color: #ebccd1; }
            .sgart-content-wrapper .sgart-http-status-500 { background-color: #f2dede; color: #a94442; border-color: #ebccd1; }   
            .sgart-popup {
            position: fixed;
            display: none;   /*flex;*/
            top: 0;
            left: 0;
            right: 0;
            bottom: 0;
            backdrop-filter: blur(5px);
            z-index: 10001;
            padding: 40px 20px 20px 20px;
            }
            .sgart-popup .sgart-popup-wrapper {
            display: flex;
            flex-direction: column;
            width: 100%;
            background-color: var(--sgart-secondary-color-white);
            border: 2px solid var(--sgart-primary-color);
            box-shadow: 0 4px 8px rgba(0, 0, 0, 0.2);
            z-index: 10002;
            }
            .sgart-popup .sgart-pupup-header {
            display: flex;
            flex-direction: row;
            justify-content: space-between;
            align-items: center;
            padding: 10px;
            height: 40px;
            border-bottom: 1px solid var(--sgart-primary-color);
            background-color: var(--sgart-primary-color);
            color: var(--sgart-secondary-color-white);
            }
            .sgart-popup .sgart-popup-body {
            display: flex;
            flex-direction: column;
            padding: 10px;
            height: 100%;
            overflow-x: hidden;
            overflow-y: auto;
            }
            .sgart-popup .sgart-popup-group {
            display: flex;
            flex-direction: column;
            padding: 10px;
            height: 100%;
            }
            .sgart-popup .sgart-popup-group > div {
            display: flex;
            flex-direction: row;
            justify-content: space-between;
            padding: 10px;
            height: 100%;
            flex-wrap: wrap;
            }
            .sgart-popup .sgart-popup-action {  
            border: 1px solid var(--sgart-primary-color);
            padding: 10px;
            margin: 5px;
            cursor: pointer;
            width: 45%;
            overflow: hidden;
            text-align: left;
            background-color: var(--sgart-secondary-color-white);
            }
            .sgart-popup .sgart-popup-action h4 {
            margin: 10px 0;
            font-size: 16px;
            }
            .sgart-popup .sgart-popup-action p {
            word-wrap: break-word;
            margin: 10px 0;
            }
            .sgart-popup .sgart-popup-history li {
            display: flex;
            flex-direction: row;
            align-items: center;
            margin: 5px 0;
            gap: 10px;
            justify-content: space-between;
            }
            .sgart-popup .sgart-popup-history button {
            flex: auto;
            }
        `;
        const stylePrev = document.head.getElementsByClassName('sgart-inject-style')[0];
        if (stylePrev) {
            document.head.removeChild(stylePrev);
        }
        const style = document.createElement('style');
        style.className = 'sgart-inject-style';
        //style.type = 'text/css';
        style.appendChild(document.createTextNode(css));
        document.head.appendChild(style);
    }

    function showInterface() {
        const interfaceDivPrev = document.getElementById(HTML_ID_WRAPPER);
        if (interfaceDivPrev) {
            document.body.removeChild(interfaceDivPrev);
        }
        const interfaceDiv = document.createElement('div');
        interfaceDiv.id = HTML_ID_WRAPPER;
        interfaceDiv.className = 'sgart-content-wrapper';
        interfaceDiv.innerHTML = `
            <div class="sgart-header">
                <a href="https://www.sgart.it/IT/informatica/tool-sharepoint-api-demo-vanilla-js/post" target="_blank"><img alt="Logo Sgart.it" class="logo" src="data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAJEAAAAhCAYAAADZEklWAAAABGdBTUEAALGPC/xhBQAAAAlwSFlzAAAOwAAADsABataJCQAAABh0RVh0U29mdHdhcmUAcGFpbnQubmV0IDQuMC42/Ixj3wAAAvtJREFUeF7tki2WFTEQhUvh2AAOxwZwODbAClCsAMcCUDjUbACFQ7EBHAqFYwMoHKpI9ZycE+rdVCXpSvebmYhv0qn7U336DfFf4kNglj+LewjxHzmxGMYROxanQfxbTiyGccSOxWkQ/5ITi5vWA+oQLG1x5yH+KScWTU0T1bOIR74/AnkHIP4hJxZNTRPVs4gHff/A34T4u5xYNDWNeC1QZvE/6LsJyNsD6vB6y/0lwEv8DQsblraIB33viN9AOhDIm0F6JUP8FQsblraIB33vs36Djnch/oKFDUtbxIO+91m/Qce7EH/GwoZoUazeW7zeUq95BO2zQHnB85R6zZMg/pQetDkavSNq50Pt7c0hf+C7jId7iHphzUPt7c0hf+C7EN/IsxKi0TvQTpl5oIx1zzMPlLHueeaBMvquKfWaR9A+C+TXs7I7U+o1T4L4Y3qIQi/NaA15rbzQkpE7ovRokK5nLR7NUZkWJr/LxWCYD3KAuaA15LXywkimhZbekd2iI5A3g3Qv08JIb0fmYjDMeznAXNAa8lp5YSTTQkvvrN2aWXtGejsyF4Nh3skB5oLWkNfKCy0ZuY9QduQe616b7WXWnpHejgzx23RagBDE8moNeb1dLRmvo5WW3qhdJbP2jPR2ZIjfYGHD0jQ9Pcjr7WrJ1DweKGPda7O9zNoz0tuRIX6NhQ1L0/T0IK+3qyVzTb0jzNoz0tuRsctE6wF1CFpDXisvtGSuqXeEWXtGejsyxK+wEIregXbKzANlrHueeaCMda/N9lLbg/A8Wi/vtVlJLQMgfpketDkavSNq55m9UbtKWju99/PutVmJp2eSr928B70jaueZvVG7Slo7vffz7rVZiadnko/4RXo+gnIx0lsoO3KPvkfR0lt6IkA7ang5TxdKjwb5K6S/d4TncjTMFocDh1fJMzkaZovDgcPpPE0HAnkzSPcyi0OAw+k8kQPMPFBGzxaHA4fTeSwHmPcS1bPYBRxOZ/0T3SvgcDqP0hEF6l8cCNM/xi1s5uHihBcAAAAASUVORK5CYII="></a>
                <h3>Tool SharePoint API Demo (Vanilla JS)</h3>
                <button id="${HTML_ID_BTN_EXIT}" class="sgart-button">Exit</button>
            </div>
            <div class="sgart-body">
                <div class="sgart-input-area">
                    <label for="${HTML_ID_TXT_INPUT}">API url:</label>
                    <input type="text" id="${HTML_ID_TXT_INPUT}" class="sgart-input" value="web/lists">
                    <select id="${HTML_ID_SELECT_ODATA}" title="OData HTTP header 'accept'">
                        <option value="nometadata" selected>Nometadata [accept:application/json; odata=nometadata]</option>
                        <option value="verbose">Verbose [accept:application/json; odata=verbose]</option>
                    </select>
                </div>
                <div class="sgart-toolbar">
					<div class="sgart-toolbar-left">
						<button id="${HTML_ID_BTN_EXECUTE}" class="sgart-button" title="Execute api call"><svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 2048 2048"><path d="M1792 1024L512 1920V128l1280 896zM640 1674l929-650-929-650v1300z"></path></svg><span>Execute</span></button>
						<span class="sgart-separator">|</span>
						<button id="${HTML_ID_BTN_CLEAR_OUTPUT}" class="sgart-button" title="Clear all outputs"><svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 2048 2048"><path d="M1115 1024l914 915-90 90-915-914-915 914-90-90 914-915L19 109l90-90 915 914 915-914 90 90-914 915z"></path></svg><span>Clear</span></button>
						<button id="${HTML_ID_BTN_COPY_OUTPUT}" class="sgart-button" title="Copy current response"><svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 2048 2048"><path d="M1920 805v1243H640v-384H128V0h859l384 384h128l421 421zm-384-37h165l-165-165v165zM640 384h549L933 128H256v1408h384V384zm1152 512h-384V512H768v1408h1024V896z"></path></svg><span>Copy</span></button>
						<span class="sgart-separator">|</span>
                        <span>Output:</span>
						<button id="${HTML_ID_TAB_RAW}" class="sgart-button sgart-button-tab" data-tab="${TAB_KEY_RAW}" data-tab-control-id="${HTML_ID_OUTPUT_RAW}" title="API Response">RAW</button>
						<button id="${HTML_ID_TAB_SIMPLE}" class="sgart-button sgart-button-tab" data-tab="${TAB_KEY_SIMPLE}" data-tab-control-id="${HTML_ID_OUTPUT_SIMPLE}" title="Response with 'value' or 'd' property removed">Simple</button>
						<button id="${HTML_ID_TAB_TABLE}" class="sgart-button sgart-button-tab" data-tab="${TAB_KEY_TABLE}" data-tab-control-id="${HTML_ID_OUTPUT_TABLE}" title="Response formatted as table (beta)">Table</button>
                        <span class="sgart-separator">|</span>
                        <button id="${HTML_ID_BTN_EXAMPLES}" class="sgart-button" title="Show popup with examples"><svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 2048 2048"><path d="M1792 0v1792H256V0h1536zm-128 128H384v1536h1280V128zM640 896H512V768h128v128zm896 0H768V768h768v128zm-896 384H512v-128h128v128zm896 0H768v-128h768v128zM640 512H512V384h128v128zm896 0H768V384h768v128z"></path></svg><span>Examples</span></button>
                        <button id="${HTML_ID_BTN_HISTORY}" class="sgart-button" title="Show popup with histories"><svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 2048 2048"><path d="M1024 512v549l365 366-90 90-403-402V512h128zm944 113q80 192 80 399t-80 399q-78 183-220 325t-325 220q-192 80-399 80-174 0-336-57-158-55-289-156-130-101-223-238-47-69-81-144t-57-156l123-34q40 145 123 266t198 208 253 135 289 48q123 0 237-32t214-90 182-141 140-181 91-214 32-238q0-123-32-237t-90-214-141-182-181-140-214-91-238-32q-130 0-252 36T545 268 355 429 215 640h297v128H0V256h128v274q17-32 37-62t42-60q94-125 220-216Q559 98 710 49t314-49q207 0 399 80 183 78 325 220t220 325z"></path></svg><span>History</span></button>
                        <span class="sgart-separator">|</span>
                        <span id="${HTML_ID_HTTP_STATUS}" class="sgart-http-status" title="HTTP response status"></span>
                        <span title="Response items count">Count: <span id="${HTML_ID_LBL_COUNT}"></span></span>

					</div>
					<div class="sgart-toolbar-right"><small>v. ${VERSION}</small></div>
                </div>

                <div class="sgart-output-area">
                    <div>
                        <textarea id="${HTML_ID_OUTPUT_RAW}" class="sgart-output-txt"></textarea>
                        <textarea id="${HTML_ID_OUTPUT_SIMPLE}" class="sgart-output-txt"></textarea>
                        <div id="${HTML_ID_OUTPUT_TABLE}" class="sgart-output-table"></div>
                    </div>
                </div>
            </div>
            <div id="${HTML_ID_PUPUP}" class="sgart-popup"></div>            
        `;
        document.body.appendChild(interfaceDiv);
    }

    const fetchGetJson = async (url, odataVerbose) => {
        /* to paste into the 'Browser Developer Console' in another Tenant */

        const ct = "application/json; odata=" + (odataVerbose ? "verbose" : "nometadata");
        const response = await fetch(url, { method: "GET", headers: { "Accept": ct, "Content-Type": ct } });
        const data = await response.json();

        return {
            status: response.status,
            data: data ?? {}
        };

    };

    /* History */

    const history = (function () {
        const LOCAL_STORAGE_KEY_HISTORY = "sgart_it_sp_api_demo_history_v1";
        const MAX_HISTORY_ITEMS = 99;
        const historyList = [];

        const loadFromStorage = () => {
            try {
                const historyJson = localStorage.getItem(LOCAL_STORAGE_KEY_HISTORY);
                if (historyJson) {
                    const historyArray = JSON.parse(historyJson);
                    if (Array.isArray(historyArray)) {
                        historyList.length = 0;
                        historyArray.forEach(item => historyList.push(item));
                    }
                }
            } catch (error) {
                console.error(LOG_SOURCE, "Error loading history from local storage:", error);
            }
        };

        const saveToStorage = () => {
            try {
                const historyJson = JSON.stringify(historyList);
                localStorage.setItem(LOCAL_STORAGE_KEY_HISTORY, historyJson);
            } catch (error) {
                console.error(LOG_SOURCE, "Error saving history to local storage:", error);
            }
        };

        const clear = () => {
            historyList.length = 0;
            saveToStorage();
        };

        const getList = () => {
            if (historyList.length === 0) {
                loadFromStorage();
            }
            return historyList;
        };

        const add = (url, odataVerbose) => {
            if (historyList.length > 0 && historyList[0].url.toLocaleLowerCase() === url.toLocaleLowerCase()) {
                // do nothing, same as last     
            } else {
                if (historyList.length >= MAX_HISTORY_ITEMS) {
                    historyList.pop();
                }
                const historyItem = {
                    url: url,
                    odataVerbose: odataVerbose,
                    timestamp: new Date().toISOString(),
                    //response: response
                };
                historyList.unshift(historyItem);
                saveToStorage();
            }
        };

        return {
            init: loadFromStorage,
            clear,
            getList,
            add
        };
    })();

    /* END History */


    /* POPUP */

    const EVENT_CLOSE = "popup-close";
    const EVENT_SET_URL = "set-url";

    const popup = (function () {
        let fnHandleClick = undefined;

        const show = (title, bodyContentHtml, fnHandle) => {
            const elmPopup = document.getElementById(HTML_ID_PUPUP);
            elmPopup.innerHTML = `
                <div class="sgart-popup-wrapper">
                    <div class="sgart-pupup-header">
                        <h2>${title.htmlEncode()}</h2>
                        <button class="sgart-button sgart-popup-event" data-event="${EVENT_CLOSE}" title="close popup">Close</button>
                    </div>
                    <div class="sgart-popup-body">${bodyContentHtml}</div>
                </div>`;
            elmPopup.style.display = 'flex';
            elmPopup.addEventListener("click", fnHandle);
        };

        const hide = () => {
            const elmPopup = document.getElementById(HTML_ID_PUPUP);
            elmPopup.style.display = 'none';
            elmPopup.innerHTML = '';
            if (fnHandleClick && typeof fnHandleClick === 'function') {
                elmPopup.removeEventListener("click", fnHandleClick);
                fnHandleClick = undefined;
            }
        };

        return {
            show,
            hide
        };
    })();

    function handlePopupClickEvent(event) {
        const target = event.target;
        const actionElem = target.closest('.sgart-popup-event');
        if (actionElem) {
            const poupEvent = actionElem.getAttribute('data-event')
            if (poupEvent === EVENT_CLOSE) {
                popup.hide();
            } else if (poupEvent === EVENT_SET_URL) {
                const url = actionElem.getAttribute('data-url');
                document.getElementById(HTML_ID_TXT_INPUT).value = url;
                popup.hide();
                handleExecuteClickEvent();
            } else {
                console.error(LOG_SOURCE, "Unknown popup event:", poupEvent);
            }
        }
    }

    function popupShowExamples() {
        let html = "";
        EXAMPLES.groups.forEach(group => {
            html += "<div class='sgart-popup-group'><h3>" + group.title.htmlEncode() + "</h3><div>";
            group.actions.forEach(action => {
                const relativeUrl = '_api/' + action.url + getQueryParam(action.query);
                const url = (serverRelativeUrlPrefix + relativeUrl).htmlEncode();
                const title = action.title.htmlEncode();
                html += "<button class='sgart-popup-action sgart-popup-event'"
                    + " data-event=\"" + EVENT_SET_URL + "\""
                    + " data-url=\"" + url + "\""
                    + " data-group=\"" + group.id + "\""
                    + " data-action=\"" + action.id + "\""
                    + " title=\"" + url + "\""
                    + "><h4>" + title + "</h4><p>" + relativeUrl + "</p>"
                    + "</button>";
            });
            html += "</div></div>";
        });

        popup.show("Examples and usage", html, handlePopupClickEvent);
    }

    function popupShowHistory() {
        let html = "";
        const historyList = history.getList();
        if (historyList.length === 0) {
            html = "<p>No history available.</p>";
        } else {
            html += "<div class='sgart-popup-history'><ol>";
            historyList.forEach((historyItem, index) => {
                const url = historyItem.url.htmlEncode();
                const date = `${new Date(historyItem.timestamp).toLocaleString()}`.htmlEncode();
                html += "<li><span>" + (index + 1) + "</span><span>" + date + "</span><button class='sgart-popup-action sgart-popup-event'"
                    + " data-event=\"" + EVENT_SET_URL + "\""
                    + " data-url=\"" + url + "\""
                    + "><strong>" + url + "</strong>"
                    + "</button></li>";
            });
            html += "</ol></div>";
        }
        popup.show("History", html, handlePopupClickEvent);
    }

    /* END POPUP */

    function handleExecuteKeydownEvent(event) {
        if (event.keyCode === 13) {
            handleExecuteClickEvent();
        }
    }

    function handleExecuteClickEvent() {
        const input = document.getElementById(HTML_ID_TXT_INPUT).value;

        const outputRaw = document.getElementById(HTML_ID_OUTPUT_RAW);
        const outputPretty = document.getElementById(HTML_ID_OUTPUT_SIMPLE);
        const outputTable = document.getElementById(HTML_ID_OUTPUT_TABLE);
        const waitTxt = "Executing...";
        outputRaw.value = waitTxt;
        outputPretty.value = waitTxt;
        outputTable.innerHTML = waitTxt;

        const modeVerbose = document.getElementById(HTML_ID_SELECT_ODATA).value === 'verbose';

        fetchGetJson(input, modeVerbose).then(response => {
            const statusTxt = response.status;
            let statusGroup = statusTxt >= 500
                ? "500"
                : statusTxt >= 400
                    ? "400"
                    : statusTxt >= 300
                        ? "300"
                        : statusTxt >= 200
                            ? "200"
                            : statusTxt >= 100
                                ? "100"
                                : "000";
            const elmStatus = document.getElementById(HTML_ID_HTTP_STATUS);
            elmStatus.innerText = statusTxt;
            elmStatus.className = `sgart-http-status sgart-http-status-${statusGroup}`;

            const data = response.data;
            outputRaw.value = JSON.stringify(data, null, 2);

            const simplified = simplifyObjectOrArray(data);
            outputPretty.value = JSON.stringify(simplified, null, 2);

            document.getElementById(HTML_ID_LBL_COUNT).innerText = Array.isArray(simplified) ? simplified.length : "1";

            const tableHtml = htmlTableFromJson.buid(data);
            outputTable.innerHTML = tableHtml;

            // add to history
            history.add(input, modeVerbose);
        }).catch(error => {
            console.error(LOG_SOURCE, "Error executing API request:", error);
            const msg = "Error: " + error.message;
            outputRaw.value = msg;
            outputPretty.value = msg;
            outputArea.value = msg;
        });
    }

    function handleSwitchTabEvent(event) {
        currentTab = event.currentTarget.getAttribute('data-tab');
        const tabs = document.getElementsByClassName('sgart-button-tab');
        Array.from(tabs).forEach(btn => {
            btn.classList.remove('selected');
            const dataTab = btn.getAttribute('data-tab');
            const controlId = btn.getAttribute('data-tab-control-id');
            const controlElem = document.getElementById(controlId);
            if (dataTab === currentTab) {
                btn.classList.add('selected');
                controlElem.style.display = 'flex';
            } else {
                btn.classList.remove('selected');
                controlElem.style.display = 'none';
            }
        });
    }

    function handleExitClickEvent() {
        const interfaceDiv = document.getElementById(HTML_ID_WRAPPER);
        document.body.removeChild(interfaceDiv);
        const style = document.head.getElementsByClassName('sgart-inject-style')[0];
        if (style) {
            document.head.removeChild(style);
        }
        console.log(LOG_SOURCE, "Interface closed");
    }


    function addEvents() {
        const btnExecute = document.getElementById(HTML_ID_BTN_EXECUTE);
        btnExecute.addEventListener("click", handleExecuteClickEvent);

        const txtInput = document.getElementById(HTML_ID_TXT_INPUT);
        txtInput.addEventListener("keydown", handleExecuteKeydownEvent);

        document.getElementById(HTML_ID_BTN_EXIT).addEventListener("click", handleExitClickEvent);
        document.getElementById(HTML_ID_BTN_EXAMPLES).addEventListener("click", popupShowExamples);
        document.getElementById(HTML_ID_BTN_HISTORY).addEventListener("click", popupShowHistory);
        document.getElementById(HTML_ID_BTN_CLEAR_OUTPUT).addEventListener("click", () => {
            document.getElementById(HTML_ID_OUTPUT_RAW).value = "";
            document.getElementById(HTML_ID_OUTPUT_SIMPLE).value = "";
            document.getElementById(HTML_ID_OUTPUT_TABLE).innerHTML = "";
        });
        document.getElementById(HTML_ID_BTN_COPY_OUTPUT).addEventListener("click", () => {
            if (currentTab === TAB_KEY_TABLE) {
                copyToClipboard(document.getElementById(HTML_ID_OUTPUT_TABLE).innerHTML);
            } else if (currentTab === TAB_KEY_SIMPLE) {
                copyToClipboard(document.getElementById(HTML_ID_OUTPUT_SIMPLE).innerHTML);
            } else {
                copyToClipboard(document.getElementById(HTML_ID_OUTPUT_SIMPLE).value);
            }
        });

        const tabs = document.getElementsByClassName('sgart-button-tab');
        Array.from(tabs).forEach(btn => {
            btn.onclick = handleSwitchTabEvent;
        });
        tabs[0].click();
    }

    function init() {
        console.log(LOG_SOURCE, `v.${VERSION} - https://www.sgart.it/IT/informatica/tool-sharepoint-api-demo-vanilla-js/post`);

        const i = window.location.pathname.toLocaleLowerCase().indexOf('/sites/');
        if (i >= 0) {
            const j = window.location.pathname.indexOf('/', i + 7);
            if (j >= 0) {
                serverRelativeUrlPrefix = window.location.pathname.substring(i, j + 1);
            } else {
                serverRelativeUrlPrefix = window.location.pathname.substring(i) + "/";
            }
        } else {
            serverRelativeUrlPrefix = "/";
        }
        console.log(LOG_SOURCE, "Site detected in URL", serverRelativeUrlPrefix);

        injectStyle();
        showInterface();
        addEvents();

        history.init();

        // set default
        const elmTxt = document.getElementById(HTML_ID_TXT_INPUT);
        elmTxt.value = serverRelativeUrlPrefix + "_api/web";
        elmTxt.focus();
        handleExecuteClickEvent();
    }

    init();
})();