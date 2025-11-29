/**
 * https://www.sgart.it
 * Formatta un oggetto JSON in HTML
 * ritorna l'HTML formattato
 * @param {object} objJson 
 * @param {string} idContainer optional, id del contanier, se passato aggiunge la gestione degli eventi expand collapse
 * @param {object} options colori {cProp: "#0451a5", cSep: "#444", cString: "#a31515", cNumber: "#098658", cBoolean: "#0000ff", cType: "666", cBtn: "#222"}
 * @returns 
 */
function formatObjectAsHtmlTree(objJson, idContainer, options) {
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
        .${BASE} .button { display: inline-flex; justify-content: center; align-items: center; width: 24px; height: 24px; padding: 0; margin: 0 5px 0 0; border-radius: 0; border: 1px solid var(--${BASE}-btn); color: var(--${BASE}-btn); background-color: transparent; overflow: hidden; font-size: 1rem;}
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
        const control = document.getElementById(ctrlId);
        const isShow = control.style.display === "";
        control.style.display = isShow ? "none" : "";
        btn.setAttribute("aria-expanded", !isShow);
        btn.innerHTML = isShow ? SVG_ADD : SVG_SUB;
    };

    injectStyle();

    const str = `<section id="${BASE}" class="${BASE}">${format(objJson)}</section>`;
    if (idContainer) {
        const htmlContainer = document.getElementById(idContainer);
        htmlContainer.innerHTML = str;
        const htmlContaner = document.getElementById(BASE);
        htmlContaner.addEventListener("click", handleClick);
    }
    return str;
}