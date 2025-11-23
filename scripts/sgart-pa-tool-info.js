(function () {
    /*
        Power Automate Tool Info (Sgart.it)
        https://www.sgart.it/IT/informatica/???/post

        javascript:(function(){var s=document.createElement('script');s.src='/SiteAssets/ToolApiDemo/sgart-pa-tool-info.js?t='+(new Date()).getTime();document.head.appendChild(s);})();
    */
    /*
    https://defaultbxxxxxxxxxxxxxxxxxxxxxxxxxxxxx.xx.environment.api.powerplatform.com/powerautomate/flows/xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx?api-version=1&$expand=properties.connectionreferences.apidefinition,properties.definitionsummary.operations.apioperation,operationDefinition,plan,properties.throttleData,properties.estimatedsuspensiondata,properties.licenseData,properties.billingContext,properties.throttlingBehavior,properties.powerFlowType,properties.protectionStatus,properties.owningUser&draftFlow=true
    localSorage
    xxxxxxxx-79fa-4149-a2cb-ede7e2219573.b32d8140-xxxx-xxxx-xxxx-6237a532dca9-login.windows.net-accesstoken-xxxxxx-4712-4c46-a7d9-3ed63d992682-b32d8140-xxxx-xxxx-xxxx-6237a532dca9-https://service.flow.microsoft.com//user_impersonation https://service.flow.microsoft.com//.default--
    {
        ...
        "secret":"xxx.xxx.xxx",
        ...
        "tokenType":"Bearer"
    }
    */

    const VERSION = "1.2025-11-20";

    const LOG_SOURCE = "Sgart.it:PowerAutomate:ToolInfo:";

    const console = {
        log: (msg1, msg2) => {
            window.console.log("%c" + LOG_SOURCE, 'color: #000; background: #5cb85c; padding: 1px 4px;', msg1, msg2 ?? '');
        },
        debug: (msg1, msg2) => {
            window.console.debug("%c" + LOG_SOURCE, 'color: #000; background: #5bc0de; padding: 1px 4px;', msg1, msg2 ?? '');
        },
        info: (msg1, msg2) => {
            window.console.info("%c" + LOG_SOURCE, 'color: #000; background: #5cb85c; padding: 1px 4px;', msg1, msg2 ?? '');
        },
        warn: (msg1, msg2) => {
            window.console.warn("%c" + LOG_SOURCE, 'color: #000; background: #f0ad4e; padding: 1px 4px', msg1, msg2 ?? '');
        },
        error: (msg1, msg2) => {
            window.console.error("%c" + LOG_SOURCE, 'color: #fff; background: #d9534f; padding: 1px 4px', msg1, msg2 ?? '');
        }
    };

    /**
     * verificare se c'è un metodo migliore
     */
    var getEnvironmentId = () => {

        //"https://make.powerautomate.com/environments/Default-b32d8140-3c2a-469d-9492-6237a532dca9/flows/4cb20f81-9769-4929-9766-83851d0dfdf5?v3=true".match(/\/(Default\-[a-z0-9\-]+)\//i)

        const env = "/environments/";
        const url = window.location.pathname.toLowerCase();
        const i = url.indexOf(env);
        if (i !== -1) {
            const subUrl = url.substring(i + env.length);
            console.log('subUrl', subUrl);
            const f = subUrl.indexOf("/");
            if (i !== -1) {
                return subUrl.substring(0, f);
            }
        }
        return null;
    };

    const getFlowId = () => {
        const flows = "/flows/";
        const url = window.location.pathname.toLowerCase();
        const i = url.indexOf(flows);
        if (i !== -1) {
            return url.substring(i + flows.length);
        }
        return null;
    };

    const getFlowUrl = () => {
        const envTmp = getEnvironmentId().replace(/-/g, "");
        const env1 = envTmp.substring(0, envTmp.length - 2);
        const env2 = envTmp.substring(envTmp.length - 2);
        const flowId = getFlowId();
        return `https://${env1}.${env2}.environment.api.powerplatform.com/powerautomate/flows/${flowId}?api-version=1&$expand=properties.connectionreferences.apidefinition,properties.definitionsummary.operations.apioperation,operationDefinition,plan,properties.throttleData,properties.estimatedsuspensiondata,properties.licenseData,properties.billingContext,properties.throttlingBehavior,properties.powerFlowType,properties.protectionStatus,properties.owningUser&draftFlow=true`;
    };

    const getToken = () => {
        try {
            console.debug('getToken');
            for (let i = 0; i < localStorage.length; i++) {
                const k = localStorage.key(i);
                const v = localStorage.getItem(k);
                if (k.indexOf('-login.windows.net-accesstoken-') >= 0 && k.indexOf('-https://service.flow.microsoft.com//user_impersonation https://service.flow.microsoft.com//.default--') >= 0) {
                    const obj = JSON.parse(v);

                    return obj.secret;
                }
            }
            return undefined;
        } catch (error) {
            console.error('getToken', error);
        }
    };

    const getJson = async (url) => {
        console.debug('getJson');
        try {
            const token = getToken();
            console.log('token', token);
            const response = await fetch(url, { headers: { 'Authorization': `Bearer ${token}`, 'Content-Type': 'application/json' } });
            return await response.json();
        } catch (error) {
            console.error('getJson', error);
        }
    };

    const init = async () => {
        console.log(`v.${VERSION} - https://www.sgart.it/IT/informatica/????`);

        try {
            const url = getFlowUrl();
            console.log('url', url);
            const content = await getJson(url);
            console.log('content', content);
        } catch (error) {
            console.error('init', error);
        }
    };

    console.debug('START',)
    init();
})();
