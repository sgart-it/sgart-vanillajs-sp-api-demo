(function () {
	/* 
        SharePoint Tool Api Demo (Sgart.it)
		javascript:(function(){var s=document.createElement('script');s.src='/SiteAssets/ToolApiDemo/sgart-sp-tool-api-demo.js?t='+(new Date()).getTime();document.head.appendChild(s);})();
	 */
    let serverRelativeUrlPrefix = "/";
	const VERSION = "1.1.2025-10-27";

    function injectStyle() {
        const css = `
            .sgart-content {
                font-family: Arial, sans-serif;
                border: 0;
                display: flex;
                flex-direction: column;
                position: fixed;
                top: 0;
                left: 0;
                right: 0;
                bottom: 0;
                background-color: white;
                margin: 0;
                padding: 0;
                z-index: 10000;
            }   
                .sgart-content input, .sgart-content textarea, .sgart-content select, .sgart-content .sgart-button {
                    font-family: Arial, sans-serif;
                    font-size: 14px;
                    height: 32px;
                    padding: 0 10px;
                    border: 1px solid rgb(149, 60, 15);
                    background-color: white;
                }
                .sgart-content select {
                    width: 110px;
                }
                .sgart-content #sgart-api-demo {
                    width: 200px;
                }
                .sgart-content .sgart-button  {
                    background-color: rgb(149, 60, 15);
                    color: white;
                    padding: 0px 20px;
                    cursor: pointer;
                }
                .sgart-content .sgart-separator{
                    margin: 0px 10px;
                }

            .sgart-header {
                background-color: #080808;  
                color: white;
                padding: 10px;
                border-bottom: 1px solid #ccc;
                height: 40px;
                display: flex;
                flwex-direction: row;
                align-items: center;
                justify-content: space-between;
            }       
                .sgart-header .sgart-button  {
                    background-color: #080808;
                    color: white;
                    padding: 0px 20px;
                }
                .sgart-header .logo {
                    height: 33px;
                    margin-right: 10px;
                }
			.sgart-toolbar{
				display:flex;
				flex-direction: row;
				align-items: center;
				justify-content: space-between;
			}
			.sgart-toolbar-left{
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
                jsutify-content: space-between;
            }
            #sgart-input {
                flex-grow: 1;   
            }
            .sgart-output-area {
                flex-grow: 1;   
                display: flex;
            }
            #sgart-output {
                width: 100%;    
                height: 100%;
                flex-grow: 1;
                gap: 10px;
                font-family: monospace;
                resize: none;
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

    function copyToClipboard(text) {
        navigator.clipboard.writeText(text)
            .then(() => {
                //console.log('Text copied to clipboard:', text);
            })
            .catch(err => {
                console.error('Failed to copy text: ', err);
            });
    }

    function getQueryParam(query) {
        if (!query) return "";
        const q = "?"
            + Object.keys(query)
                //.map(k => "$" + encodeURIComponent(k) + "=" + encodeURIComponent(query[k]))
                .map(k => "$" + k + "=" + query[k])
                .join("&");
        return q;
    }

    function loadAPiUrlOptions() {
        const configs = {
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
                            url: "web/lists"
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
                            url: "web/lists/getbytitle('Documents')/items"
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

        const select = document.getElementById('sgart-api-demo');
        const optionEmpty = document.createElement('option');
        optionEmpty.value = "";
        optionEmpty.text = "API URL";
        select.appendChild(optionEmpty);

        configs.groups.forEach(group => {
            const optGroup = document.createElement('optgroup');
            optGroup.label = group.title;
            group.actions.forEach(action => {
                const relativeUrl = '_api/' + action.url + getQueryParam(action.query);
                const option = document.createElement('option');
                option.value = serverRelativeUrlPrefix + relativeUrl;
                option.text = action.title + ": [ " + relativeUrl + " ]";
                option.setAttribute('data-group', group.id);
                option.setAttribute('data-action', action.id);
                optGroup.appendChild(option);
            });
            select.appendChild(optGroup);
        });
        select.onchange = function () {
            document.getElementById('sgart-input').value = this.value;
            document.getElementById('sgart-output').value = "";
        };

        document.querySelector('#sgart-api-demo [data-action=getWeb]').selected = true;
        select.onchange();
        executeApiCall();
    }

    function showInterface() {
        const interfaceDivPrev = document.getElementById('sgart-content');
        if (interfaceDivPrev) {
            document.body.removeChild(interfaceDivPrev);
        }
        const interfaceDiv = document.createElement('div');
        interfaceDiv.id = 'sgart-content';
        interfaceDiv.className = 'sgart-content';
        interfaceDiv.innerHTML = `
            <div class="sgart-header">
                <a href="https://www.sgart.it" target="_blank"><img alt="Logo Sgart.it" class="logo" src="data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAJEAAAAhCAYAAADZEklWAAAABGdBTUEAALGPC/xhBQAAAAlwSFlzAAAOwAAADsABataJCQAAABh0RVh0U29mdHdhcmUAcGFpbnQubmV0IDQuMC42/Ixj3wAAAvtJREFUeF7tki2WFTEQhUvh2AAOxwZwODbAClCsAMcCUDjUbACFQ7EBHAqFYwMoHKpI9ZycE+rdVCXpSvebmYhv0qn7U336DfFf4kNglj+LewjxHzmxGMYROxanQfxbTiyGccSOxWkQ/5ITi5vWA+oQLG1x5yH+KScWTU0T1bOIR74/AnkHIP4hJxZNTRPVs4gHff/A34T4u5xYNDWNeC1QZvE/6LsJyNsD6vB6y/0lwEv8DQsblraIB33viN9AOhDIm0F6JUP8FQsblraIB33vs36Djnch/oKFDUtbxIO+91m/Qce7EH/GwoZoUazeW7zeUq95BO2zQHnB85R6zZMg/pQetDkavSNq50Pt7c0hf+C7jId7iHphzUPt7c0hf+C7EN/IsxKi0TvQTpl5oIx1zzMPlLHueeaBMvquKfWaR9A+C+TXs7I7U+o1T4L4Y3qIQi/NaA15rbzQkpE7ovRokK5nLR7NUZkWJr/LxWCYD3KAuaA15LXywkimhZbekd2iI5A3g3Qv08JIb0fmYjDMeznAXNAa8lp5YSTTQkvvrN2aWXtGejsyF4Nh3skB5oLWkNfKCy0ZuY9QduQe616b7WXWnpHejgzx23RagBDE8moNeb1dLRmvo5WW3qhdJbP2jPR2ZIjfYGHD0jQ9Pcjr7WrJ1DweKGPda7O9zNoz0tuRIX6NhQ1L0/T0IK+3qyVzTb0jzNoz0tuRsctE6wF1CFpDXisvtGSuqXeEWXtGejsyxK+wEIregXbKzANlrHueeaCMda/N9lLbg/A8Wi/vtVlJLQMgfpketDkavSNq55m9UbtKWju99/PutVmJp2eSr928B70jaueZvVG7Slo7vffz7rVZiadnko/4RXo+gnIx0lsoO3KPvkfR0lt6IkA7ang5TxdKjwb5K6S/d4TncjTMFocDh1fJMzkaZovDgcPpPE0HAnkzSPcyi0OAw+k8kQPMPFBGzxaHA4fTeSwHmPcS1bPYBRxOZ/0T3SvgcDqP0hEF6l8cCNM/xi1s5uHihBcAAAAASUVORK5CYII="></a>
                <h3>Tool SharePoint API Demo (Vanilla JS)</h3>
                <button id="sgart-close" class="sgart-button">Close</button>
            </div>
            <div class="sgart-body">
                <div class="sgart-input-area">
                    <label for="sgart-input">API url:</label>
                    <input type="text" id="sgart-input" value="web/lists">
                    <select id="sgart-api-demo" title="API Demo URLs">
                    </select>
                    <select id="sgart-odata-mode" title="OData http header accept">
                        <option value="nometadata" selected>No Metadata [accept:application/json; odata=nometadata]</option>
                        <option value="verbose">OData Verbose [accept:application/json; odata=verbose]</option>
                    </select>
                </div>
                <div class="sgart-toolbar">
					<div class="sgart-toolbar-left">
						<button id="sgart-execute" class="sgart-button">Execute</button>
						<span class="sgart-separator">|</span>
						<button id="sgart-clear-output" class="sgart-button">Clear</button>
						<button id="sgart-copy-output" class="sgart-button">Copy</button>
						<!--
						<span class="sgart-separator">|</span>
						<button id="sgart-tab-response" class="sgart-button">Response</button>
						<button id="sgart-tab-table" class="sgart-button">Table</button>
						-->
					</div>
					<div class="sgart-toolbar-right">v. ${VERSION}</div>
                </div>
                <div class="sgart-output-area">
                    <textarea id="sgart-output"></textarea>
                </div>
            </div>
        `;
        interfaceDiv.querySelector('#sgart-close').onclick = function () {
            document.body.removeChild(interfaceDiv);
            const style = document.head.getElementsByClassName('sgart-inject-style')[0];
            if (style) {
                document.head.removeChild(style);
            }
        };

        interfaceDiv.querySelector('#sgart-clear-output').onclick = function () {
            document.getElementById('sgart-output').value = "";
        };

        interfaceDiv.querySelector('#sgart-copy-output').onclick = function () {
            copyToClipboard(document.getElementById('sgart-output').value);
        };

        document.body.appendChild(interfaceDiv);

        loadAPiUrlOptions();
    }

    const fetchGetJson = async (url, odataVerbose, outputNormal) => {
        /* to paste into the 'Browser Developer Console' in another Tenant */
        const ct = "application/json; odata=" + (odataVerbose ? "verbose" : "nometadata");
        const response = await fetch(url, { method: "GET", headers: { "Accept": ct, "Content-Type": ct } });
        if (!response.ok) {
            const txt = await response.json();
            throw new Error(`Response status: ${response.status}, ${response.statusText}, ${txt}`);
        }
        const data = await response.json();
        if (outputNormal)
            return data;
        if (odataVerbose) {
            if (data.d && data.d.results) return data.d.results;
            else if (data.d) return data.d;
            return data;
        }
        return data.value ?? data;
    };

    function executeApiCall() {
        const input = document.getElementById('sgart-input').value;
        const outputArea = document.getElementById('sgart-output');
        const modeVerbose = document.getElementById('sgart-odata-mode').value === 'verbose';
        outputArea.value = "Executing...";
        fetchGetJson(input, modeVerbose).then(data => {
            outputArea.value = JSON.stringify(data, null, 2);
        }).catch(error => {
            outputArea.value = "Error: " + error.message;
        });
    }

    function init() {
        const i = window.location.pathname.toLocaleLowerCase().indexOf('/sites/');
        if (i >= 0) {
            serverRelativeUrlPrefix = window.location.pathname.substring(0, window.location.pathname.indexOf('/', i + 7)) + "/";
        } else {
            serverRelativeUrlPrefix = "/";
        }

        injectStyle();
        showInterface();
        document.getElementById('sgart-execute').onclick = executeApiCall;

        console.log("Sgart.it SharePoint API Test Interface initialized v." + VERSION);
        //console.log(fetchGetJson.toString());
    }

    init();
})();