(function () {
    /*
        Power Automate Tool Info (Sgart.it)
        https://www.sgart.it/IT/informatica/???/post

        javascript:(function(){var s=document.createElement('script');s.src='/SiteAssets/ToolApiDemo/sgart-pa-tool-info.js?t='+(new Date()).getTime();document.head.appendChild(s);})();
    */
    const VERSION = "1.2025-11-29";

    const LOG_SOURCE = "Sgart.it:PowerAutomate:ToolInfo:";
    const LOG_COLOR_SOURCE = "%c" + LOG_SOURCE;
    const LOG_COLOR_LOG = "color: #000; background: #5cb85c; padding: 1px 4px;";
    const LOG_COLOR_DEBUG = "color: #000; background: #5bc0de; padding: 1px 4px;";
    const LOG_COLOR_INFO = "color: #000; background: #5cb85c; padding: 1px 4px;";
    const LOG_COLOR_WARN = "color: #000; background: #f0ad4e; padding: 1px 4px";
    const LOG_COLOR_ERROR = "color: #fff; background: #d9534f; padding: 1px 4px";

    const HTML_ID_WRAPPER = "sgart-content-wrapper";

    const HTML_ID_PUPUP = "sgart-popup";

    const HTML_ID_BTN_EXIT = "sgart-btn-exit";
    const HTML_ID_BTN_CLEAR_OUTPUT = "sgart-btn-clear-output";
    const HTML_ID_BTN_COPY_OUTPUT = "sgart-btn-copy-output";
    const HTML_ID_BTN_HISTORY = "sgart-btn-history";
    const HTML_ID_HTTP_STATUS = "sgart-http-status";
    const HTML_ID_HTTP_EXECUTION_TIME = "sgart-http-execution-time";

    const HTML_ID_BTN_EXECUTE = "sgart-btn-execute";
    const HTML_ID_TXT_INPUT = "sgart-txt-input";

    const TAB_OUTPUT = '-output';
    const TAB_KEY_RAW = 'sgart-tab-raw';
    const TAB_KEY_TREE = 'sgart-tab-tree';
    const TAB_KEY_TABLE_CONN = 'sgart-tab-table-connection';
    const globalTabs = [
        { key: TAB_KEY_RAW, controlId: TAB_KEY_RAW + TAB_OUTPUT, type: "txt", title: "API response", text: "RAW", iconSVG: `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 2048 2048"><path d="M896 512H0V384h896v128zM384 768h896v128H384V768zm1024 0h640v128h-640V768zm640-384v128H1024V384h1024zM384 1152h1280v128H384v-128zM0 1536h1280v128H0v-128z"></path></svg>` },
        { key: TAB_KEY_TREE, controlId: TAB_KEY_TREE + TAB_OUTPUT, type: "tree", title: "API response formatted as tree", text: "Tree", iconSVG: `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 2048 2048" class="svg_dd790ee3" focusable="false"><path d="M512 384h1536v128H512V384zm512 640V896h1024v128H1024zm0 512v-128h1024v128H1024zM0 640V256h384v384H0zm128-256v128h128V384H128zm384 768V768h384v384H512zm128-256v128h128V896H640zm-128 768v-384h384v384H512zm128-256v128h128v-128H640z"></path></svg>` },
        { key: TAB_KEY_TABLE_CONN, controlId: TAB_KEY_TABLE_CONN + TAB_OUTPUT, type: "table", title: "Connection", text: "Connection", iconSVG: `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 2048 2048"><path d="M128 256h1664v1536H128V256zm640 768v256h384v-256H768zm384-128V640H768v256h384zm-512 0V640H256v256h384zm-384 128v256h384v-256H256zm384 640v-256H256v256h384zm512 0v-256H768v256h384zm512 0v-256h-384v256h384zm0-384v-256h-384v256h384zm0-384V640h-384v256h384zM256 512h1408V384H256v128z"></path></svg>` }
    ];
    let currentTabKey = globalTabs[0].key;

    
    const console = {
        log: (msg, value) => {
            if (value)
                window.console.log(LOG_COLOR_SOURCE, LOG_COLOR_LOG, msg, value);
            else
                window.console.log(LOG_COLOR_SOURCE, LOG_COLOR_LOG, msg);
        },
        debug: (msg, value) => {
            if (value)
                window.console.debug(LOG_COLOR_SOURCE, LOG_COLOR_DEBUG, msg, value);
            else
                window.console.debug(LOG_COLOR_SOURCE, LOG_COLOR_DEBUG, msg);
        },
        info: (msg, value) => {
            if (value)
                window.console.info(LOG_COLOR_SOURCE, LOG_COLOR_INFO, msg, value);
            else
                window.console.info(LOG_COLOR_SOURCE, LOG_COLOR_INFO, msg);
        },
        warn: (msg, value) => {
            if (value)
                window.console.warn(LOG_COLOR_SOURCE, LOG_COLOR_WARN, msg, value);
            else
                window.console.warn(LOG_COLOR_SOURCE, LOG_COLOR_WARN, msg);
        },
        error: (msg, value) => {
            if (value)
                window.console.error(LOG_COLOR_SOURCE, LOG_COLOR_ERROR, msg, value);
            else
                window.console.error(LOG_COLOR_SOURCE, LOG_COLOR_ERROR, msg);
        }
    };
    

    const getElmById = (id) => { const elm = document.getElementById(id); return elm };
    const setElmByIdHtml = (id, html) => { const elm = document.getElementById(id); elm.innerHTML = html; return elm };
    const setElmByIdValue = (id, value) => { const elm = document.getElementById(id); elm.value = value; return elm };

    // encode dei caratteri in html
    String.prototype.htmlEncode = function () {
        const node = document.createTextNode(this);
        return document.createElement("a").appendChild(node).parentNode.innerHTML.replace(/'/g, "&#39;").replace(/"/g, "&#34;");
    };

    const copyToClipboard = (text) => navigator.clipboard.writeText(text)
        .then(() => console.debug('Text copied to clipboard:', text))
        .catch(err => console.error('Failed to copy text: ', err));

    const formatObjectAsHtmlTree = (objJson, idContainer, options) => {
        const BASE = "sgart-it-format-json-to-tree";
        const SVG_ADD = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 2048 2048"><path d="M2048 960v128h-960v960H960v-960H0V960h960V0h128v960h960z"></path></svg>`;
        const SVG_SUB = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 2048 2048"><path d="M0 960h2048v128H0V960z"></path></svg>`

        const getSequence = () => "id" + Math.random().toString(16).slice(2);
        const htmlEscape = (str) => (str ?? "").toString().replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/'/g, "&#039;");

        const injectStyle = () => {
            const color = options || {};
            const css = `
        .${BASE} { 
            --${BASE}-prop: ${color.cProp ?? "#0451a5"};
            --${BASE}-sep: ${color.cSep ?? "#444"};
            --${BASE}-string: ${color.cString ?? "#a31515"};
            --${BASE}-number: ${color.cNumber ?? "#098658"};
            --${BASE}-boolean: ${color.cBoolean ?? "#0000ff"};
            --${BASE}-type: ${color.cType ?? "#666"};
            --${BASE}-btn: ${color.cBtn ?? "#222"};
        }
        .${BASE}, .${BASE} * { font-family: consolas, menlo, monaco, "Ubuntu Mono", source-code-pro, monospace; font-size: .9rem; }
        .${BASE} var, .${BASE} i, .${BASE} em { font-style: italic; text-decoration: none; font-weight: normal; color: var(--${BASE}-type); }
        .${BASE} i { padding: 0 5px 0 0; font-style: normal;  color: var(--${BASE}-sep);}
        .${BASE} label { display: inline; font-style: normal; text-decoration: none; font-weight: bold; padding: 0; }
        .${BASE} .button { display: inline-flex; justify-content: center; align-items: center; width: 24px; height: 24px; padding: 0; margin: 0 5px 0 0; border-radius: 0; border: 1px solid var(--${BASE}-btn); color: var(--${BASE}-btn); background-color: transparent; overflow: hidden; font-size: 1rem; cursor: pointer;}
        .${BASE} .button svg { width: 11px; height: 11px; pointer-events: none; fill: var(--${BASE}-btn);}
        .${BASE} ul { list-style: none; }
        .${BASE} ul li { min-height: 30px; line-height: 30px; vertical-align: middle; }
        .${BASE} label { color: var(--${BASE}-prop); }
        .${BASE} .key-value-boolean span, .${BASE} .key-value-null span, .${BASE} .key-value-undefined span { color: var(--${BASE}-boolean); }
        .${BASE} .key-value-string span { color: var(--${BASE}-string); }
        .${BASE} .key-value-number span { color: var(--${BASE}-number); }            
        `;
            const className = `${BASE}-inject-styles`;
            const stylePrev = document.head.getElementsByClassName(className)[0];
            if (stylePrev) {
                document.head.removeChild(stylePrev);
            }
            const style = document.createElement('style');
            style.className = className;
            style.appendChild(document.createTextNode(css));
            document.head.appendChild(style);
        };

        const getType = (value) => value === null ? "null" : Array.isArray(value) ? "array" : typeof value;

        const formatObject = (obj, level) => {
            const objectName = Array.isArray(obj) ? "array" : "object";
            const items = Object.entries(obj);
            const s = items.reduce((accumulator, current) => {
                const [key, value] = current;
                const type = getType(value);
                if (type === "array" || type === "object")
                    return accumulator + `<li class="key-value-${type}" title="${type}"><label>${key}</label><i>:</i>${formatObject(value, level + 1)}</li>`;

                const str = htmlEscape(value);
                const strTitle = `${key} : ${type} = ${type === "string" ? `&quot;${str}&quot; {${str.length}}` : str}`;
                return accumulator + `<li class="key-value-${type}" title="${strTitle}"><label>${key}</label><i>:</i><span>${str}</span></li>`;
            }, "");
            const id = `${BASE}-${level}-${getSequence()}`;
            return `<button class="button" type="button" aria-expanded="true" aria-controls="${id}">${SVG_SUB}</button><em>${objectName}</em> <var>{${items.length}}</var><ul id="${id}">${s}</ul>`;
        };

        const format = (obj) => obj === null ? "null" : typeof obj === 'object' ? formatObject(obj, 0) : "Unsupported data type";

        const handleClick = (event) => {
            const btn = event.target;
            const ctrlId = btn.getAttribute("aria-controls");
            const control = getElmById(ctrlId);
            const isShow = control.style.display === "";
            control.style.display = isShow ? "none" : "";
            btn.setAttribute("aria-expanded", !isShow);
            btn.innerHTML = isShow ? SVG_ADD : SVG_SUB;
        };

        injectStyle();

        const str = `<section id="${BASE}" class="${BASE}">${format(objJson)}</section>`;
        if (idContainer) {
            const htmlContainer = getElmById(idContainer);
            htmlContainer.innerHTML = str;
            const htmlContaner = getElmById(BASE);
            htmlContaner.addEventListener("click", handleClick);
        }
        return str;
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
            console.warn('buildTable', "TOTO");
            const data = {};    // simplifyObjectOrArray(items);
            // console.debug('buildTable: items is object', data);
            if (!items) {
                // console.debug('buildTable: items is undefined or null');
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

    const BASE = "sgart-it-pa-tool-info-wrapper-1sdfy23";
    const injectStyle = () => {
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
            --sgart-btn-color-execute: #f0ad4e;
            --sgart-btn-color-execute-border: #aa6708;
            }
            .${BASE} {
                font-family: Arial, sans-serif;
                font-size: 14px;
                font-weight: normal;
                line-height: 1.6;
                border: 0;
                display: flex;
                flex-direction: column;
                position: fixed;
                top: 0;
                left: 0;
                right: 0;
                bottom: 0;
                background-color: var(--sgart-secondary-color-white);
                color: var(--sgart-secondary-color-dark);
                margin: 0;
                padding: 0;
                z-index: 10000;
                box-sizing: border-box;
            }   
            .${BASE} input, .${BASE} textarea, .${BASE} select, .${BASE} .sgart-button {
                height: 32px;
                padding: 0 10px;
                border: 1px solid var(--sgart-primary-color);
                background-color: var(--sgart-secondary-color-white);
                box-sizing: border-box;
                background-image: none;
                border-radius: 2px;
            }
            .${BASE} select { width: 180px; }
            .${BASE} #sgart-api-demo { width: 200px; }
            .${BASE} .sgart-button  {
                background-color: var(--sgart-primary-color);
                color: var(--sgart-secondary-color-white);
                padding: 0 10px;
                cursor: pointer;
                width: 110px;
                overflow: hidden;
                white-space: nowrap;
                display: flex;
                align-items: center;
                justify-content: center;
                gap: 5px;
            }
            #${HTML_ID_BTN_EXECUTE} { background-color: var(--sgart-btn-color-execute); border: 1px solid var(--sgart-btn-color-execute-border); color: var(--sgart-secondary-color-dark)}
            #${HTML_ID_BTN_EXECUTE} svg { fill: var(--sgart-secondary-color-dark);}
            .${BASE} .sgart-button svg { color: var(--sgart-secondary-color-white); fill: var(--sgart-secondary-color-white); height: 20px; width: 20px; }
            .${BASE} .sgart-button.sgart-button-tab { background-color: var(--sgart-primary-color-light); color: var(--sgart-secondary-color); }
            .${BASE} .sgart-button.sgart-button-tab svg { fill: var(--sgart-secondary-color); }
            .${BASE} .sgart-button.selected, .${BASE} .sgart-button:hover, .${BASE} .sgart-button.sgart-button-tab.selected, .${BASE} .sgart-button.sgart-button-tab:hover { background-color: var(--sgart-primary-color-hover); color: var(--sgart-secondary-color-white); font-weight: bold; }
            .${BASE} .sgart-button.selected svg { fill: var(--sgart-secondary-color-white); }
            .${BASE} .sgart-button.sgart-button-tab { width: 100px }
            .${BASE} .sgart-button.sgart-button-tab:hover { border-color: var(--sgart-secondary-color); }
            .${BASE} .sgart-button.sgart-button-tab:hover svg { fill: var(--sgart-secondary-color-white); }
            .${BASE} .sgart-separator { margin: 0; }
            .${BASE} .sgart-header {
                position: relative;
                background-color: var(--sgart-secondary-color);  
                color: var(--sgart-secondary-color-white);
                padding: 5px 10px;
                border-bottom: 1px solid var(--sgart-secondary-color-gray-light);
                height: 40px;
                display: flex;
                flex-direction: row;
                align-items: center;
                justify-content: space-between;
            }     
            .${BASE} .sgart-header h1 { font-size: 1.2em; font-weight: bold; margin: 0; } 
            .${BASE} .sgart-header a { display: flex }
            .${BASE} .sgart-header .sgart-button { background-color: var(--sgart-secondary-color); border: 1px solid var(--sgart-secondary-color); color: var(--sgart-secondary-color-white); padding: 5px 10px; width: 80px; }
            .${BASE} .sgart-header .sgart-button:hover { border: 1px solid var(--sgart-secondary-color-white); font-weight: normal; }
            .${BASE} .sgart-header .logo { height: 33px; margin-right: 10px; }
            .${BASE} .sgart-toolbar { display:flex; flex-direction: row; align-items: center; justify-content: space-between; }
            .${BASE} .sgart-toolbar-left { display: flex; gap: 10px; justify-content: left; align-items: center; flex-wrap: wrap; }
            .${BASE} .sgart-toolbar-right{ justify-content: right; }
            .${BASE} .sgart-body { display: flex; flex-direction: column; flex-grow: 1; padding: 10px; gap: 10px; font-weight: normal; }
            .${BASE} .sgart-body label { font-weight: normal; padding: 0; text-wrap-mode: nowrap; }   
            .${BASE} .sgart-input-area { display: flex; gap: 10px; align-items: center; justify-content: space-between; }
            .${BASE} .sgart-input { flex-grow: 1; }
            .${BASE} .sgart-output-area { flex-grow: 1; display: flex; overflow: hidden; position: relative; }
            .${BASE} .sgart-output-area > div {
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
            .${BASE} .sgart-output-area .sgart-output-tree { padding: .5em; }
            .${BASE} .sgart-output-area table { border-collapse: collapse; width: 100%; background-color: var(--sgart-secondary-color-white); color: var(--sgart-secondary-color-dark);}
            .${BASE} .sgart-output-txt, .${BASE} .sgart-output-table {
                width: 100%;    
                height: auto;
                flex-grow: 1;
                gap: 10px;
                font-family: consolas, menlo, monaco, "Ubuntu Mono", source-code-pro, monospace; 
                font-size: 14px;
                white-space: pre;
                box-sizing: border-box;
                border: none;
            }
            .${BASE} table th {
                background-color: var(--sgart-primary-color);
                color: var(--sgart-secondary-color-white);
                text-align: left;
                position: sticky;
                top: 0;
                z-index: 1000;
                padding: 5px;
            }
            .${BASE} .sgart-http-status { border: 1px solid var(--sgart-secondary-color); padding: 0; background-color: var(--sgart-secondary-color-gray-light); color: var(--sgart-secondary-color); font-weight: bold; display: inline-flex; width: 50px; height: 32px; align-items: center; justify-content: center; }
            .${BASE} .sgart-http-status-100 { background-color: #e7f3fe; color: #31708f; border-color: #bce8f1; }  
            .${BASE} .sgart-http-status-200 { background-color: #dff0d8; color: #3c763d; border-color: #d6e9c6; }
            .${BASE} .sgart-http-status-300 { background-color: #fcf8e3; color: #8a6d3b; border-color: #faebcc; }
            .${BASE} .sgart-http-status-400 { background-color: #f2dede; color: #a94442; border-color: #ebccd1; }
            .${BASE} .sgart-http-status-500 { background-color: #f2dede; color: #a94442; border-color: #ebccd1; }   
            .${BASE} .sgart-label-count { border: 1px solid var(--sgart-secondary-color-gray-light); padding: 0; background-color: var(--sgart-secondary-color-white); color: var(--sgart-secondary-color); font-weight: bold; display: inline-flex; width: 50px; height: 32px; align-items: center; justify-content: center; }
            .${BASE} .sgart-popup {
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
            .${BASE} .sgart-popup .sgart-popup-wrapper {
                display: flex;
                flex-direction: column;
                width: 100%;
                background-color: var(--sgart-secondary-color-white);
                border: 2px solid var(--sgart-primary-color);
                box-shadow: 0 4px 8px rgba(0, 0, 0, 0.2);
                z-index: 10002;
            }
            .${BASE} .sgart-popup .sgart-pupup-header {
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
            .${BASE} .sgart-popup .sgart-popup-body {
                display: flex;
                flex-direction: column;
                padding: 10px;
                height: 100%;
                overflow-x: hidden;
                overflow-y: auto;
            }
            .${BASE} .sgart-popup h3 { margin: 0; font-size: 18px;}
            .${BASE} .sgart-popup .sgart-popup-group { display: block; padding: 10px; }
            .${BASE} .sgart-popup .sgart-popup-group > div { display: flex; flex-direction: row; justify-content: flex-start; padding: 10px; flex-wrap: wrap; }
            .${BASE} .sgart-popup .sgart-popup-action { border: 1px solid var(--sgart-primary-color); padding: 10px; margin: 5px; cursor: pointer; width: 32%; overflow: hidden; text-align: left; background-color: var(--sgart-secondary-color-white); }
            .${BASE} .sgart-popup .sgart-popup-action h4 { margin: 0 0 8px 0; font-size: 16px;}
            .${BASE} .sgart-popup .sgart-popup-action p { word-wrap: break-word; margin: 8px 0;}
            .${BASE} .sgart-popup .sgart-popup-action > div { word-wrap: break-word; margin: 8px 0;}
            .${BASE} .sgart-popup .sgart-popup-history li { display: flex; flex-direction: row; align-items: center; margin: 5px 0; gap: 10px; justify-content: space-between;}
            .${BASE} .sgart-popup .sgart-popup-history button { flex: auto;}
        `;
        const stylePrev = document.head.getElementsByClassName('sgart-inject-style')[0];
        if (stylePrev) {
            document.head.removeChild(stylePrev);
        }
        const style = document.createElement('style');
        style.className = 'sgart-inject-style';
        style.appendChild(document.createTextNode(css));
        document.head.appendChild(style);
    };

    function showInterface() {
        const interfaceDivPrev = getElmById(HTML_ID_WRAPPER);
        if (interfaceDivPrev) {
            document.body.removeChild(interfaceDivPrev);
        }
        const htmlTabs = globalTabs.reduce((accumulator, current) => {
            const k = current.key;
            const title = current.title.htmlEncode();
            const text = (current.iconSVG ?? "") + current.text.htmlEncode();
            return accumulator + `<button id="${k}" class="sgart-button sgart-button-tab" data-tab-key="${k}" data-tab-control-id="${k + TAB_OUTPUT}" title="${title}">${text}</button>`;
        }, "");
        const htmlOutputs = globalTabs.reduce((accumulator, current) => {
            return accumulator + `<div id="${current.key + TAB_OUTPUT}" class="sgart-output-${current.type}"></div>`;
        }, "");
        const interfaceDiv = document.createElement('div');
        interfaceDiv.id = HTML_ID_WRAPPER;
        interfaceDiv.className = BASE;
        interfaceDiv.innerHTML = `
            <div class="sgart-header">
                <a href="https://www.sgart.it/IT/informatica/???/post" target="_blank"><img alt="Logo Sgart.it" class="logo" src="data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAJEAAAAhCAYAAADZEklWAAAABGdBTUEAALGPC/xhBQAAAAlwSFlzAAAOwAAADsABataJCQAAABh0RVh0U29mdHdhcmUAcGFpbnQubmV0IDQuMC42/Ixj3wAAAvtJREFUeF7tki2WFTEQhUvh2AAOxwZwODbAClCsAMcCUDjUbACFQ7EBHAqFYwMoHKpI9ZycE+rdVCXpSvebmYhv0qn7U336DfFf4kNglj+LewjxHzmxGMYROxanQfxbTiyGccSOxWkQ/5ITi5vWA+oQLG1x5yH+KScWTU0T1bOIR74/AnkHIP4hJxZNTRPVs4gHff/A34T4u5xYNDWNeC1QZvE/6LsJyNsD6vB6y/0lwEv8DQsblraIB33viN9AOhDIm0F6JUP8FQsblraIB33vs36Djnch/oKFDUtbxIO+91m/Qce7EH/GwoZoUazeW7zeUq95BO2zQHnB85R6zZMg/pQetDkavSNq50Pt7c0hf+C7jId7iHphzUPt7c0hf+C7EN/IsxKi0TvQTpl5oIx1zzMPlLHueeaBMvquKfWaR9A+C+TXs7I7U+o1T4L4Y3qIQi/NaA15rbzQkpE7ovRokK5nLR7NUZkWJr/LxWCYD3KAuaA15LXywkimhZbekd2iI5A3g3Qv08JIb0fmYjDMeznAXNAa8lp5YSTTQkvvrN2aWXtGejsyF4Nh3skB5oLWkNfKCy0ZuY9QduQe616b7WXWnpHejgzx23RagBDE8moNeb1dLRmvo5WW3qhdJbP2jPR2ZIjfYGHD0jQ9Pcjr7WrJ1DweKGPda7O9zNoz0tuRIX6NhQ1L0/T0IK+3qyVzTb0jzNoz0tuRsctE6wF1CFpDXisvtGSuqXeEWXtGejsyxK+wEIregXbKzANlrHueeaCMda/N9lLbg/A8Wi/vtVlJLQMgfpketDkavSNq55m9UbtKWju99/PutVmJp2eSr928B70jaueZvVG7Slo7vffz7rVZiadnko/4RXo+gnIx0lsoO3KPvkfR0lt6IkA7ang5TxdKjwb5K6S/d4TncjTMFocDh1fJMzkaZovDgcPpPE0HAnkzSPcyi0OAw+k8kQPMPFBGzxaHA4fTeSwHmPcS1bPYBRxOZ/0T3SvgcDqP0hEF6l8cCNM/xi1s5uHihBcAAAAASUVORK5CYII="></a>
                <h1>Tool Power Automate Info (Vanilla JS)</h1>
                <button id="${HTML_ID_BTN_EXIT}" class="sgart-button"><svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 2048 2048"><path d="M128 128h1792v1792H128V128zm1664 1664V256H256v1536h1536zM621 1517l-90-90 402-403-402-403 90-90 403 402 403-402 90 90-402 403 402 403-90 90-403-402-403 402z"></path></svg>Exit</button>
            </div>
            <div class="sgart-body">
                <div class="sgart-input-area">
                    <label for="${HTML_ID_TXT_INPUT}">API url:</label>
                    <input type="text" id="${HTML_ID_TXT_INPUT}" class="sgart-input" value="web/lists">
                </div>
                <div class="sgart-toolbar">
					<div class="sgart-toolbar-left">
						<button id="${HTML_ID_BTN_EXECUTE}" class="sgart-button" title="Execute api call"><svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 2048 2048"><path d="M1792 1024L512 1920V128l1280 896zM640 1674l929-650-929-650v1300z"></path></svg><span>Execute</span></button>
						<span class="sgart-separator">|</span>
						<button id="${HTML_ID_BTN_CLEAR_OUTPUT}" class="sgart-button" title="Clear all outputs"><svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 2048 2048"><path d="M1115 1792h421v128H453L50 1516q-24-24-37-56t-13-68q0-35 13-67t38-58L1248 69l794 795-927 928zm133-1542L538 960l614 613 709-709-613-614zM933 1792l128-128-613-614-306 307q-14 14-14 35t14 35l364 365h427z"></path></svg><span>Clear</span></button>
						<button id="${HTML_ID_BTN_COPY_OUTPUT}" class="sgart-button" title="Copy current response"><svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 2048 2048"><path d="M1920 805v1243H640v-384H128V0h859l384 384h128l421 421zm-384-37h165l-165-165v165zM640 384h549L933 128H256v1408h384V384zm1152 512h-384V512H768v1408h1024V896z"></path></svg><span>Copy</span></button>
						<span class="sgart-separator">|</span>
                        <label>Output:</label>${htmlTabs}
                        <span class="sgart-separator">|</span>
                        <button id="${HTML_ID_BTN_HISTORY}" class="sgart-button" title="Show popup with histories"><svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 2048 2048"><path d="M1024 512v549l365 366-90 90-403-402V512h128zm944 113q80 192 80 399t-80 399q-78 183-220 325t-325 220q-192 80-399 80-174 0-336-57-158-55-289-156-130-101-223-238-47-69-81-144t-57-156l123-34q40 145 123 266t198 208 253 135 289 48q123 0 237-32t214-90 182-141 140-181 91-214 32-238q0-123-32-237t-90-214-141-182-181-140-214-91-238-32q-130 0-252 36T545 268 355 429 215 640h297v128H0V256h128v274q17-32 37-62t42-60q94-125 220-216Q559 98 710 49t314-49q207 0 399 80 183 78 325 220t220 325z"></path></svg><span>History</span></button>
                        <span class="sgart-separator">|</span>
                        <span><label>Status:</label> <span id="${HTML_ID_HTTP_STATUS}" class="sgart-http-status" title="HTTP response status"></span></span>
                        <span><label>Time:</label> <span id="${HTML_ID_HTTP_EXECUTION_TIME}" class="sgart-http-execution-time" title="HTTP execution time"></span></span>
    				</div>
					<div class="sgart-toolbar-right"><small>v. ${VERSION}</small></div>
                </div>
                <div class="sgart-output-area"><div>${htmlOutputs}</div></div>
            </div>
            <div id="${HTML_ID_PUPUP}" class="sgart-popup"></div>            
        `;
        document.body.appendChild(interfaceDiv);
    }

    /* FETCH JSON POWER AUTOMATE */

    /**
     * verificare se c'è un metodo migliore per ricavare environment
     */
    var getEnvironmentId = () => {
        const m = window.location.pathname.match(/\/environments\/([a-z0-9\-]+)\//i);
        return m.length === 2 ? m[1] : undefined;
    };

    const getFlowId = () => {
        const m = window.location.pathname.match(/\/flows\/([a-z0-9\-]+)/i);
        return m.length === 2 ? m[1] : undefined;
    };

    const getFlowUrl = () => {
        try {
            const env = getEnvironmentId().replace(/-/g, "");
            console.debug('env', env);
            const env1 = env.substring(0, env.length - 2);
            console.debug('env1', env1);
            const env2 = env.substring(env.length - 2);
            console.debug('env2', env2);
            const flowId = getFlowId();
            console.debug('flowId', flowId);
            return `https://${env1}.${env2}.environment.api.powerplatform.com/powerautomate/flows/${flowId}?api-version=1&$expand=properties.connectionreferences.apidefinition,properties.definitionsummary.operations.apioperation,operationDefinition,plan,properties.throttleData,properties.estimatedsuspensiondata,properties.licenseData,properties.billingContext,properties.throttlingBehavior,properties.powerFlowType,properties.protectionStatus,properties.owningUser&draftFlow=true`;
        } catch (error) {
            console.error('getFlowUrl', error);
        }
        return undefined;
    };

    const getToken = () => {
        try {
            for (let i = 0; i < localStorage.length; i++) {
                const k = localStorage.key(i);
                const v = localStorage.getItem(k);
                if (k.indexOf('-login.windows.net-accesstoken-') >= 0 && k.indexOf('-https://service.flow.microsoft.com//user_impersonation https://service.flow.microsoft.com//.default--') >= 0) {
                    const obj = JSON.parse(v);
                    return obj.secret;
                }
            }
        } catch (error) {
            console.error('getToken', error);
        }
        return undefined;
    };

    const fetchGetJson = async (url) => {
        try {
            const token = getToken();
            console.debug('token', token.substring(0, 20) + "..." + token.substring(token.length - 20));
            const ct = "application/json";
            const response = await fetch(url, { method: "GET", headers: { "Authorization": `Bearer ${token}`, "Accept": ct, "Content-Type": ct } });
            const data = await response.json();
            return {
                status: response.status,
                data: data ?? {}
            };
        } catch (error) {
            console.error("fetchGetJson: Error parsing JSON response:", error);
            return {
                status: response.status,
                data: { error: error.message }
            };
        }
    };

    /* END: FETCH JSON POWER AUTOMATE */

    /* History */

    const history = (function () {
        const LOCAL_STORAGE_KEY_HISTORY = "sgart_it_pa_tool_info_history_v1";
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
                console.error("Error loading history from local storage:", error);
            }
        };

        const saveToStorage = () => {
            try {
                const historyJson = JSON.stringify(historyList);
                localStorage.setItem(LOCAL_STORAGE_KEY_HISTORY, historyJson);
            } catch (error) {
                console.error("Error saving history to local storage:", error);
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

        const add = (url) => {
            if (historyList.length > 0 && historyList[0].url.toLocaleLowerCase() === url.toLocaleLowerCase()) {
                // do nothing, same as last     
            } else {
                if (historyList.length >= MAX_HISTORY_ITEMS) {
                    historyList.pop();
                }
                const historyItem = {
                    url: url,
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
            const elmPopup = getElmById(HTML_ID_PUPUP);
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
            const elmPopup = getElmById(HTML_ID_PUPUP);
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
                setElmByIdValue(HTML_ID_TXT_INPUT, url);
                popup.hide();
                handleExecuteClickEvent();
            } else {
                console.error("Unknown popup event:", poupEvent);
            }
        }
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


    function renderContent(data) {
        if (data) {
            globalTabs.forEach(item => {
                const elm = getElmById(item.controlId);
                switch (item.key) {
                    case TAB_KEY_RAW:
                        elm.innerHTML = JSON.stringify(data, null, 2).htmlEncode();
                        break;
                    case TAB_KEY_TREE:
                        formatObjectAsHtmlTree(data, item.controlId);
                        break;
                    case TAB_KEY_TABLE_CONN:
                        const tableHtml = htmlTableFromJson.buid(data);
                        elm.innerHTML = tableHtml;
                        break;
                }
            });
        } else {
            globalTabs.forEach(item => setElmByIdHtml(item.controlId, msg));
        }
    }

    function handleExecuteClickEvent() {
        try {

            const input = getElmById(HTML_ID_TXT_INPUT).value;

            globalTabs.forEach(item => setElmByIdHtml(item.controlId, "Executing..."));

            setElmByIdHtml(HTML_ID_HTTP_STATUS, "...");
            setElmByIdHtml(HTML_ID_HTTP_EXECUTION_TIME, "-");

            const startTime = performance.now();

            fetchGetJson(input).then(response => {
                const endTime = performance.now();
                setElmByIdHtml(HTML_ID_HTTP_EXECUTION_TIME, (Math.round((endTime - startTime) * 10) / 10) + " ms");

                const statusGroup = parseInt(response.status / 100).toString() + "00";
                setElmByIdHtml(HTML_ID_HTTP_STATUS, response.status).className = `sgart-http-status sgart-http-status-${statusGroup}`;

                globalTabs.forEach(item => setElmByIdHtml(item.controlId, ""));

                history.add(input);

                renderContent(response.data);
            }).catch(error => {
                //renderContent();
                console.error("Error executing API request:", error);
                const msg = "Error: " + error.message;
                globalTabs.forEach(item => setElmByIdHtml(item.controlId, msg));
            });
        } catch (error) {
            console.error('handleExecuteClickEvent', error);
        }
    }

    function handleSwitchTabEvent(event) {
        const key = event.currentTarget.getAttribute('data-tab-key');
        globalTabs.forEach(item => {
            const btn = document.getElementById(item.key);
            btn.classList.remove('selected');
            const controlElem = document.getElementById(item.controlId);
            if (item.key === key) {
                btn.classList.add('selected');
                controlElem.style.display = 'flex';
            } else {
                btn.classList.remove('selected');
                controlElem.style.display = 'none';
            }
        });
    }

    function handleExitClickEvent() {
        window.removeEventListener("beforeunload", handleBeforeunloadEvent);
        const interfaceDiv = document.getElementById(HTML_ID_WRAPPER);
        document.body.removeChild(interfaceDiv);
        const style = document.head.getElementsByClassName('sgart-inject-style')[0];
        if (style) {
            document.head.removeChild(style);
        }
        console.log("Interface closed");
    }

    function addEvents() {
        const btnExecute = document.getElementById(HTML_ID_BTN_EXECUTE);
        btnExecute.addEventListener("click", handleExecuteClickEvent);

        const txtInput = document.getElementById(HTML_ID_TXT_INPUT);
        txtInput.addEventListener("keydown", handleExecuteKeydownEvent);

        document.getElementById(HTML_ID_BTN_EXIT).addEventListener("click", handleExitClickEvent);
        document.getElementById(HTML_ID_BTN_HISTORY).addEventListener("click", popupShowHistory);
        document.getElementById(HTML_ID_BTN_CLEAR_OUTPUT).addEventListener("click", () => globalTabs.forEach(item => setElmByIdHtml(item.controlId, "")));
        document.getElementById(HTML_ID_BTN_COPY_OUTPUT).addEventListener("click", () => {
            console.log('copy-top', currentTabKey);
            console.log('copy-tabs', globalTabs);
            const tab = globalTabs[globalTabs.findIndex(item => item.key === currentTabKey)];
            console.log('copy', tab);
            if (tab) {
                copyToClipboard(getElmById(tab.controlId).innerHTML);
            }else{
                console.error('copy', 'element not found');
            }
        });

        const htmlTabs = document.getElementsByClassName('sgart-button-tab');
        Array.from(htmlTabs).forEach(btn => {
            btn.onclick = handleSwitchTabEvent;
        });
        htmlTabs[0].click();
    }

    function handleBeforeunloadEvent(event) {
        event.preventDefault();
        console.log("beforeunload", event);
    }

    function init() {
        console.log(`v.${VERSION} - https://www.sgart.it/IT/informatica/???/post`);

        injectStyle();
        showInterface();
        addEvents();

        history.init();

        // set default
        const elmTxt = document.getElementById(HTML_ID_TXT_INPUT);
        elmTxt.value = getFlowUrl();
        elmTxt.focus();
        handleExecuteClickEvent();

        window.addEventListener("beforeunload", handleBeforeunloadEvent);
    }

    init();
})();