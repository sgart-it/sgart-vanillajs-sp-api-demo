/**
 * https://www.sgart.it
 * Formatta un oggetto JSON in HTML
 * ritorna l'HTML formattato
 * @param {object} objJson 
 * @param {string} idContainer optional, se passato aggiunge gli eventi expand collapse
 * @param {object} options colori {cProp: "#0451a5", cSep: "#444", cString: "#a31515", cNumber: "#098658", cBoolean: "#0000ff", cType: "666", cBtn: "#222"}
 * @returns 
 */
function formatObjectAsHtmlTree(objJson, idContainer, options) {
    const IDBASE = "sgart-it-format-json-to-tree";
    const SIGN_ADD = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 2048 2048"><path d="M2048 960v128h-960v960H960v-960H0V960h960V0h128v960h960z"></path></svg>`;
    const SIGN_SUB = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 2048 2048"><path d="M0 960h2048v128H0V960z"></path></svg>`

    const getSequence = () => "id" + Math.random().toString(16).slice(2);

    const getType = (value) => value === null ? "null" : Array.isArray(value) ? "array" : typeof value;

    const htmlEscape = (str) => (str ?? "").toString().replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/'/g, "&#039;");

    const injectStyle = () => {
        const color = options || {};
        const css = `
        .${IDBASE} { 
            --${IDBASE}-prop: ${color.cProp ?? "#0451a5"};
            --${IDBASE}-sep: ${color.cSep ?? "#444"};
            --${IDBASE}-string: ${color.cString ?? "#a31515"};
            --${IDBASE}-number: ${color.cNumber ?? "#098658"};
            --${IDBASE}-boolean: ${color.cBoolean ?? "#0000ff"};
            --${IDBASE}-type: ${color.cType ?? "#666"};
            --${IDBASE}-btn: ${color.cBtn ?? "#222"};
        }
        .${IDBASE}, .${IDBASE} * { font-family: consolas, menlo, monaco, "Ubuntu Mono", source-code-pro, monospace; font-size: .9rem; }
        .${IDBASE} var, .${IDBASE} i, .${IDBASE} em { font-style: italic; text-decoration: none; font-weight: normal; color: var(--${IDBASE}-type); }
        .${IDBASE} i { padding: 0 5px 0 0; font-style: normal;  color: var(--${IDBASE}-sep);}
        .${IDBASE} label { display: inline; font-style: normal; text-decoration: none; font-weight: bold; padding: 0; }
        .${IDBASE} .button { display: inline-flex; justify-content: center; align-items: center; width: 24px; height: 24px; padding: 0; margin: 0 5px 0 0; border-radius: 0; border: 1px solid var(--${IDBASE}-btn); color: var(--${IDBASE}-btn); background-color: transparent; overflow: hidden; font-size: 1rem;}
        .${IDBASE} .button svg { width: 11px; height: 11px; pointer-events: none; fill: var(--${IDBASE}-btn);}
        .${IDBASE} ul { list-style: none; }
        .${IDBASE} ul li { min-height: 30px; line-height: 30px; vertical-align: middle; }
        .${IDBASE} label { color: var(--${IDBASE}-prop); }
        .${IDBASE} .key-value-boolean span, .${IDBASE} .key-value-null span, .${IDBASE} .key-value-undefined span { color: var(--${IDBASE}-boolean); }
        .${IDBASE} .key-value-string span { color: var(--${IDBASE}-string); }
        .${IDBASE} .key-value-number span { color: var(--${IDBASE}-number); }            
        `;
        const className = `${IDBASE}'-inject-styles`;
        const stylePrev = document.head.getElementsByClassName(className)[0];
        if (stylePrev) {
            document.head.removeChild(stylePrev);
        }
        const style = document.createElement('style');
        style.className = className;
        style.appendChild(document.createTextNode(css));
        document.head.appendChild(style);
    };

    const formatObject = (obj, level) => {
        let s = "";
        let c = 0;
        const objectName = Array.isArray(obj) ? "array" : "object";
        for (const [key, value] of Object.entries(obj)) {
            const type = getType(value);
            if (type === "array" || type === "object")
                s += `<li class="key-value-${type}" title="${type}"><label>${key}</label><i>:</i>${formatObject(value, level + 1)}</li>`;
            else {
                const str = htmlEscape(value);
                const strTitle = type + ": " + (type === "string" ? `&quot;${str}&quot; length ${str.length}` : str);
                s += `<li class="key-value-${type}"><label>${key}</label><i>:</i><span title="${strTitle}">${str}</span></li>`;
            }
            c++;
        }
        const id = `${IDBASE}-${level}-${getSequence()}`;        
        return `<button class="button" role="button" aria-expanded="true" aria-controls="${id}">${SIGN_SUB}</button><em>${objectName}</em> <var>{${c}}</var><ul id="${id}">${s}</ul>`;
    };

    const format = (obj) => obj === null ? "null" : typeof obj === 'object' ? formatObject(obj, 0) : "Unsupported data type";

    const handleClick = (event) => {
        const btn = event.target;
        const ctrlId = btn.getAttribute("aria-controls");
        const control = document.getElementById(ctrlId);
        const isShow = control.style.display === "" || control.style.display === "block" || control.style.display === "flex";
        control.style.display = isShow ? "none" : "";
        btn.setAttribute("aria-expanded", !isShow);
        btn.innerHTML = isShow ? SIGN_ADD : SIGN_SUB;
    };

    injectStyle();

    const s = `<div id="${IDBASE}" class="${IDBASE}">${format(objJson)}</div>`;
    if (idContainer) {
        const htmlContainer = document.getElementById(idContainer);
        htmlContainer.innerHTML = s;
        const htmlContaner = document.getElementById(IDBASE);
        htmlContaner.addEventListener("click", handleClick);
    }
    return s;
}