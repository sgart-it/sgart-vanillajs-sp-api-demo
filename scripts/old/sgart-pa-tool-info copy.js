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

    const VERSION = "1.2025-11-23";

    const LOG_SOURCE = "Sgart.it:PowerAutomate:ToolInfo:";
    const LOG_COLOR_LOG = "color: #000; background: #5cb85c; padding: 1px 4px;";
    const LOG_COLOR_DEBUG = "color: #000; background: #5bc0de; padding: 1px 4px;";
    const LOG_COLOR_INFO = "color: #000; background: #5cb85c; padding: 1px 4px;";
    const LOG_COLOR_WARN = "color: #000; background: #f0ad4e; padding: 1px 4px";
    const LOG_COLOR_ERROR = "color: #fff; background: #d9534f; padding: 1px 4px";

    const console = {
        log: (msg, value) => {
            if (value)
                window.console.log("%c" + LOG_SOURCE, LOG_COLOR_LOG, msg, value);
            else
                window.console.log("%c" + LOG_SOURCE, LOG_COLOR_LOG, msg);
        },
        debug: (msg, value) => {
            if (value)
                window.console.debug("%c" + LOG_SOURCE, LOG_COLOR_DEBUG, msg, value);
            else
                window.console.debug("%c" + LOG_SOURCE, LOG_COLOR_DEBUG, msg);
        },
        info: (msg, value) => {
            if (value)
                window.console.info("%c" + LOG_SOURCE, LOG_COLOR_INFO, msg, value);
            else
                window.console.info("%c" + LOG_SOURCE, LOG_COLOR_INFO, msg);
        },
        warn: (msg, value) => {
            if (value)
                window.console.warn("%c" + LOG_SOURCE, LOG_COLOR_WARN, msg, value);
            else
                window.console.warn("%c" + LOG_SOURCE, LOG_COLOR_WARN, msg);
        },
        error: (msg, value) => {
            if (value)
                window.console.error("%c" + LOG_SOURCE, LOG_COLOR_ERROR, msg, value);
            else
                window.console.error("%c" + LOG_SOURCE, LOG_COLOR_ERROR, msg);
        }
    };

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

    const getJson = async (url) => {
        try {
            const token = getToken();
            console.debug('token', token.substring(0,10) + "...");
            const response = await fetch(url, { headers: { 'Authorization': `Bearer ${token}`, 'Content-Type': 'application/json' } });
            return await response.json();
        } catch (error) {
            console.error('getJson', error);
        }
        return undefined;
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

    console.debug('START');
    init();
})();
