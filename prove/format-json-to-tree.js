function formatObjectAsJsonHtml(idContainer, obj) {
    const IDBASE = "sgart-it-format-json-to-tree";
    const SIGN_ADD = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 2048 2048"><path d="M2048 960v128h-960v960H960v-960H0V960h960V0h128v960h960z"></path></svg>`;
    const SIGN_SUB = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 2048 2048"><path d="M0 896h1920v128H0V896z"></path></svg>`
    let n = -1;

    const getType = (value) => value === null ? "null" : Array.isArray(value) ? "array" : typeof value;

    const htmlEscape = (str) => (str ?? "").toString().replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/'/g, "&#039;");

    const formatObject = (obj, level) => {
        let s = "";
        let c = 0;
        n++;
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
        const id = `${IDBASE}-${level}-${n}`;
        return `<button class="button" role="button" aria-expanded="true" aria-controls="${id}">${SIGN_SUB}</button><em>${objectName}</em> <var>{${c}}</var><ul id="${id}">${s}</ul>`;
    };

    const format = (obj) =>  obj === null ? "null" : typeof obj === 'object' ? formatObject(obj, 0) : "Unsupported data type";

    const handleClick = (event) => {
        const htmlButton = event.target;
        const ctrlId = htmlButton.getAttribute("aria-controls");
        const htmlExpand = document.getElementById(ctrlId);
        const isShow = htmlExpand.style.display === "" || htmlExpand.style.display === "block" || htmlExpand.style.display === "flex";
        htmlExpand.style.display = isShow ? "none" : "";
        htmlButton.setAttribute("aria-expanded", !isShow);
        htmlButton.innerHTML = isShow ? SIGN_ADD : SIGN_SUB;
    };

    const s = `<div id="${IDBASE}" class="${IDBASE}">${format(obj)}</div>`;
    const htmlContainer = document.getElementById(idContainer);
    htmlContainer.innerHTML = s;
    const htmlContaner = document.getElementById(IDBASE);
    htmlContaner.addEventListener("click", handleClick);
    return s;
}