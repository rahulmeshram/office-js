/* Office JavaScript API library - Custom Functions */
/* Version: 16.0.10723.30000 */
/*
	Copyright (c) Microsoft Corporation.  All rights reserved.
*/


/*
    Your use of this file is governed by the Microsoft Services Agreement http://go.microsoft.com/fwlink/?LinkId=266419.

    This file also contains the following Promise implementation (with a few small modifications):
        * @overview es6-promise - a tiny implementation of Promises/A+.
        * @copyright Copyright (c) 2014 Yehuda Katz, Tom Dale, Stefan Penner and contributors (Conversion to ES6 API by Jake Archibald)
        * @license   Licensed under MIT license
        *            See https://raw.githubusercontent.com/jakearchibald/es6-promise/master/LICENSE
        * @version   2.3.0
*/
var OSF = OSF || {};
OSF.ConstantNames = {
    OfficeJS: "office.customfunctions.js",
    OfficeDebugJS: "office.customfunctions.debug.js",
    HostFileScriptSuffix: "core"
};
var OSF = OSF || {};
OSF.HostSpecificFileVersionDefault = "16.00";
OSF.HostSpecificFileVersionMap = {
    "access": {
        "web": "16.00"
    },
    "agavito": {
        "winrt": "16.00"
    },
    "excel": {
        "ios": "16.00",
        "mac": "16.00",
        "web": "16.00",
        "win32": "16.01",
        "winrt": "16.00"
    },
    "onenote": {
        "web": "16.00",
        "win32": "16.00",
        "winrt": "16.00"
    },
    "outlook": {
        "ios": "16.00",
        "mac": "16.00",
        "web": "16.01",
        "win32": "16.02"
    },
    "powerpoint": {
        "ios": "16.00",
        "mac": "16.00",
        "web": "16.00",
        "win32": "16.01",
        "winrt": "16.00"
    },
    "project": {
        "win32": "16.00"
    },
    "sway": {
        "web": "16.00"
    },
    "word": {
        "ios": "16.00",
        "mac": "16.00",
        "web": "16.00",
        "win32": "16.01",
        "winrt": "16.00"
    }
};
OSF.SupportedLocales = {
    "ar-sa": true,
    "bg-bg": true,
    "bn-in": true,
    "ca-es": true,
    "cs-cz": true,
    "da-dk": true,
    "de-de": true,
    "el-gr": true,
    "en-us": true,
    "es-es": true,
    "et-ee": true,
    "eu-es": true,
    "fa-ir": true,
    "fi-fi": true,
    "fr-fr": true,
    "gl-es": true,
    "he-il": true,
    "hi-in": true,
    "hr-hr": true,
    "hu-hu": true,
    "id-id": true,
    "it-it": true,
    "ja-jp": true,
    "kk-kz": true,
    "ko-kr": true,
    "lo-la": true,
    "lt-lt": true,
    "lv-lv": true,
    "ms-my": true,
    "nb-no": true,
    "nl-nl": true,
    "nn-no": true,
    "pl-pl": true,
    "pt-br": true,
    "pt-pt": true,
    "ro-ro": true,
    "ru-ru": true,
    "sk-sk": true,
    "sl-si": true,
    "sr-cyrl-cs": true,
    "sr-cyrl-rs": true,
    "sr-latn-cs": true,
    "sr-latn-rs": true,
    "sv-se": true,
    "th-th": true,
    "tr-tr": true,
    "uk-ua": true,
    "ur-pk": true,
    "vi-vn": true,
    "zh-cn": true,
    "zh-tw": true
};
OSF.AssociatedLocales = {
    ar: "ar-sa",
    bg: "bg-bg",
    bn: "bn-in",
    ca: "ca-es",
    cs: "cs-cz",
    da: "da-dk",
    de: "de-de",
    el: "el-gr",
    en: "en-us",
    es: "es-es",
    et: "et-ee",
    eu: "eu-es",
    fa: "fa-ir",
    fi: "fi-fi",
    fr: "fr-fr",
    gl: "gl-es",
    he: "he-il",
    hi: "hi-in",
    hr: "hr-hr",
    hu: "hu-hu",
    id: "id-id",
    it: "it-it",
    ja: "ja-jp",
    kk: "kk-kz",
    ko: "ko-kr",
    lo: "lo-la",
    lt: "lt-lt",
    lv: "lv-lv",
    ms: "ms-my",
    nb: "nb-no",
    nl: "nl-nl",
    nn: "nn-no",
    pl: "pl-pl",
    pt: "pt-br",
    ro: "ro-ro",
    ru: "ru-ru",
    sk: "sk-sk",
    sl: "sl-si",
    sr: "sr-cyrl-cs",
    sv: "sv-se",
    th: "th-th",
    tr: "tr-tr",
    uk: "uk-ua",
    ur: "ur-pk",
    vi: "vi-vn",
    zh: "zh-cn"
};
OSF.getSupportedLocale = function OSF$getSupportedLocale(locale, defaultLocale) {
    if (defaultLocale === void 0) { defaultLocale = "en-us"; }
    if (!locale) {
        return defaultLocale;
    }
    var supportedLocale;
    locale = locale.toLowerCase();
    if (locale in OSF.SupportedLocales) {
        supportedLocale = locale;
    }
    else {
        var localeParts = locale.split('-', 1);
        if (localeParts && localeParts.length > 0) {
            supportedLocale = OSF.AssociatedLocales[localeParts[0]];
        }
    }
    if (!supportedLocale) {
        supportedLocale = defaultLocale;
    }
    return supportedLocale;
};
var ScriptLoading;
(function (ScriptLoading) {
    var ScriptInfo = (function () {
        function ScriptInfo(url, isReady, hasStarted, timer, pendingCallback) {
            this.url = url;
            this.isReady = isReady;
            this.hasStarted = hasStarted;
            this.timer = timer;
            this.hasError = false;
            this.pendingCallbacks = [];
            this.pendingCallbacks.push(pendingCallback);
        }
        return ScriptInfo;
    })();
    var ScriptTelemetry = (function () {
        function ScriptTelemetry(scriptId, startTime, msResponseTime) {
            this.scriptId = scriptId;
            this.startTime = startTime;
            this.msResponseTime = msResponseTime;
        }
        return ScriptTelemetry;
    })();
    var LoadScriptHelper = (function () {
        function LoadScriptHelper(constantNames) {
            if (constantNames === void 0) { constantNames = {
                OfficeJS: "office.js",
                OfficeDebugJS: "office.debug.js"
            }; }
            this.constantNames = constantNames;
            this.defaultScriptLoadingTimeout = 10000;
            this.loadedScriptByIds = {};
            this.scriptTelemetryBuffer = [];
            this.osfControlAppCorrelationId = "";
            this.basePath = null;
        }
        LoadScriptHelper.prototype.isScriptLoading = function (id) {
            return !!(this.loadedScriptByIds[id] && this.loadedScriptByIds[id].hasStarted);
        };
        LoadScriptHelper.prototype.getOfficeJsBasePath = function () {
            if (this.basePath) {
                return this.basePath;
            }
            else {
                var getScriptBase = function (scriptSrc, scriptNameToCheck) {
                    var scriptBase, indexOfJS, scriptSrcLowerCase;
                    scriptSrcLowerCase = scriptSrc.toLowerCase();
                    indexOfJS = scriptSrcLowerCase.indexOf(scriptNameToCheck);
                    if (indexOfJS >= 0 && indexOfJS === (scriptSrc.length - scriptNameToCheck.length) && (indexOfJS === 0 || scriptSrc.charAt(indexOfJS - 1) === '/' || scriptSrc.charAt(indexOfJS - 1) === '\\')) {
                        scriptBase = scriptSrc.substring(0, indexOfJS);
                    }
                    else if (indexOfJS >= 0
                        && indexOfJS < (scriptSrc.length - scriptNameToCheck.length)
                        && scriptSrc.charAt(indexOfJS + scriptNameToCheck.length) === '?'
                        && (indexOfJS === 0 || scriptSrc.charAt(indexOfJS - 1) === '/' || scriptSrc.charAt(indexOfJS - 1) === '\\')) {
                        scriptBase = scriptSrc.substring(0, indexOfJS);
                    }
                    return scriptBase;
                };
                var scripts = document.getElementsByTagName("script");
                var scriptsCount = scripts.length;
                var officeScripts = [this.constantNames.OfficeJS, this.constantNames.OfficeDebugJS];
                var officeScriptsCount = officeScripts.length;
                var i, j;
                for (i = 0; !this.basePath && i < scriptsCount; i++) {
                    if (scripts[i].src) {
                        for (j = 0; !this.basePath && j < officeScriptsCount; j++) {
                            this.basePath = getScriptBase(scripts[i].src, officeScripts[j]);
                        }
                    }
                }
                return this.basePath;
            }
        };
        LoadScriptHelper.prototype.loadScript = function (url, scriptId, callback, highPriority, timeoutInMs) {
            this.loadScriptInternal(url, scriptId, callback, highPriority, timeoutInMs);
        };
        LoadScriptHelper.prototype.loadScriptParallel = function (url, scriptId, timeoutInMs) {
            this.loadScriptInternal(url, scriptId, null, false, timeoutInMs);
        };
        LoadScriptHelper.prototype.waitForFunction = function (scriptLoadTest, callback, numberOfTries, delay) {
            var attemptsRemaining = numberOfTries;
            var timerId;
            var validateFunction = function () {
                attemptsRemaining--;
                if (scriptLoadTest()) {
                    callback(true);
                    return;
                }
                else if (attemptsRemaining > 0) {
                    timerId = window.setTimeout(validateFunction, delay);
                    attemptsRemaining--;
                }
                else {
                    window.clearTimeout(timerId);
                    callback(false);
                }
            };
            validateFunction();
        };
        LoadScriptHelper.prototype.waitForScripts = function (ids, callback) {
            var _this = this;
            if (this.invokeCallbackIfScriptsReady(ids, callback) == false) {
                for (var i = 0; i < ids.length; i++) {
                    var id = ids[i];
                    var loadedScriptEntry = this.loadedScriptByIds[id];
                    if (loadedScriptEntry) {
                        loadedScriptEntry.pendingCallbacks.push(function () {
                            _this.invokeCallbackIfScriptsReady(ids, callback);
                        });
                    }
                }
            }
        };
        LoadScriptHelper.prototype.logScriptLoading = function (scriptId, startTime, msResponseTime) {
            startTime = Math.floor(startTime);
            if (OSF.AppTelemetry && OSF.AppTelemetry.onScriptDone) {
                if (OSF.AppTelemetry.onScriptDone.length == 3) {
                    OSF.AppTelemetry.onScriptDone(scriptId, startTime, msResponseTime);
                }
                else {
                    OSF.AppTelemetry.onScriptDone(scriptId, startTime, msResponseTime, this.osfControlAppCorrelationId);
                }
            }
            else {
                var scriptTelemetry = new ScriptTelemetry(scriptId, startTime, msResponseTime);
                this.scriptTelemetryBuffer.push(scriptTelemetry);
            }
        };
        LoadScriptHelper.prototype.setAppCorrelationId = function (appCorrelationId) {
            this.osfControlAppCorrelationId = appCorrelationId;
        };
        LoadScriptHelper.prototype.invokeCallbackIfScriptsReady = function (ids, callback) {
            var hasError = false;
            for (var i = 0; i < ids.length; i++) {
                var id = ids[i];
                var loadedScriptEntry = this.loadedScriptByIds[id];
                if (!loadedScriptEntry) {
                    loadedScriptEntry = new ScriptInfo("", false, false, null, null);
                    this.loadedScriptByIds[id] = loadedScriptEntry;
                }
                if (loadedScriptEntry.isReady == false) {
                    return false;
                }
                else if (loadedScriptEntry.hasError) {
                    hasError = true;
                }
            }
            callback(!hasError);
            return true;
        };
        LoadScriptHelper.prototype.getScriptEntryByUrl = function (url) {
            for (var key in this.loadedScriptByIds) {
                var scriptEntry = this.loadedScriptByIds[key];
                if (this.loadedScriptByIds.hasOwnProperty(key) && scriptEntry.url === url) {
                    return scriptEntry;
                }
            }
            return null;
        };
        LoadScriptHelper.prototype.loadScriptInternal = function (url, scriptId, callback, highPriority, timeoutInMs) {
            if (url) {
                var self = this;
                var doc = window.document;
                var loadedScriptEntry = (scriptId && this.loadedScriptByIds[scriptId]) ? this.loadedScriptByIds[scriptId] : this.getScriptEntryByUrl(url);
                if (!loadedScriptEntry || loadedScriptEntry.hasError || loadedScriptEntry.url.toLowerCase() != url.toLowerCase()) {
                    var script = doc.createElement("script");
                    script.type = "text/javascript";
                    if (scriptId) {
                        script.id = scriptId;
                    }
                    if (!loadedScriptEntry) {
                        loadedScriptEntry = new ScriptInfo(url, false, false, null, null);
                        this.loadedScriptByIds[(scriptId ? scriptId : url)] = loadedScriptEntry;
                    }
                    else {
                        loadedScriptEntry.url = url;
                        loadedScriptEntry.hasError = false;
                        loadedScriptEntry.isReady = false;
                    }
                    if (callback) {
                        if (highPriority) {
                            loadedScriptEntry.pendingCallbacks.unshift(callback);
                        }
                        else {
                            loadedScriptEntry.pendingCallbacks.push(callback);
                        }
                    }
                    var timeFromPageInit = -1;
                    if (window.performance && window.performance.now) {
                        timeFromPageInit = window.performance.now();
                    }
                    var startTime = (new Date()).getTime();
                    var logTelemetry = function (succeeded) {
                        if (scriptId) {
                            var totalTime = (new Date()).getTime() - startTime;
                            if (!succeeded) {
                                totalTime = -totalTime;
                            }
                            self.logScriptLoading(scriptId, timeFromPageInit, totalTime);
                        }
                        self.flushTelemetryBuffer();
                    };
                    var onLoadCallback = function OSF_OUtil_loadScript$onLoadCallback() {
                        if (OSF._OfficeAppFactory.getHostInfo().hostType == "onenote" && (typeof OSF.AppTelemetry !== 'undefined') && (typeof OSF.AppTelemetry.enableTelemetry !== 'undefined')) {
                            OSF.AppTelemetry.enableTelemetry = false;
                        }
                        if (!OSF._OfficeAppFactory.getLoggingAllowed() && (typeof OSF.AppTelemetry !== 'undefined')) {
                            OSF.AppTelemetry.enableTelemetry = false;
                        }
                        logTelemetry(true);
                        loadedScriptEntry.isReady = true;
                        if (loadedScriptEntry.timer != null) {
                            clearTimeout(loadedScriptEntry.timer);
                            delete loadedScriptEntry.timer;
                        }
                        var pendingCallbackCount = loadedScriptEntry.pendingCallbacks.length;
                        for (var i = 0; i < pendingCallbackCount; i++) {
                            var currentCallback = loadedScriptEntry.pendingCallbacks.shift();
                            if (currentCallback) {
                                var result = currentCallback(true);
                                if (result === false) {
                                    break;
                                }
                            }
                        }
                    };
                    var onLoadError = function () {
                        logTelemetry(false);
                        loadedScriptEntry.hasError = true;
                        loadedScriptEntry.isReady = true;
                        if (loadedScriptEntry.timer != null) {
                            clearTimeout(loadedScriptEntry.timer);
                            delete loadedScriptEntry.timer;
                        }
                        var pendingCallbackCount = loadedScriptEntry.pendingCallbacks.length;
                        for (var i = 0; i < pendingCallbackCount; i++) {
                            var currentCallback = loadedScriptEntry.pendingCallbacks.shift();
                            if (currentCallback) {
                                var result = currentCallback(false);
                                if (result === false) {
                                    break;
                                }
                            }
                        }
                    };
                    if (script.readyState) {
                        script.onreadystatechange = function () {
                            if (script.readyState == "loaded" || script.readyState == "complete") {
                                script.onreadystatechange = null;
                                onLoadCallback();
                            }
                        };
                    }
                    else {
                        script.onload = onLoadCallback;
                    }
                    script.onerror = onLoadError;
                    timeoutInMs = timeoutInMs || this.defaultScriptLoadingTimeout;
                    loadedScriptEntry.timer = setTimeout(onLoadError, timeoutInMs);
                    loadedScriptEntry.hasStarted = true;
                    script.setAttribute("crossOrigin", "anonymous");
                    script.src = url;
                    doc.getElementsByTagName("head")[0].appendChild(script);
                }
                else if (loadedScriptEntry.isReady) {
                    callback(true);
                }
                else {
                    if (highPriority) {
                        loadedScriptEntry.pendingCallbacks.unshift(callback);
                    }
                    else {
                        loadedScriptEntry.pendingCallbacks.push(callback);
                    }
                }
            }
        };
        LoadScriptHelper.prototype.flushTelemetryBuffer = function () {
            if (OSF.AppTelemetry && OSF.AppTelemetry.onScriptDone) {
                for (var i = 0; i < this.scriptTelemetryBuffer.length; i++) {
                    var scriptTelemetry = this.scriptTelemetryBuffer[i];
                    if (OSF.AppTelemetry.onScriptDone.length == 3) {
                        OSF.AppTelemetry.onScriptDone(scriptTelemetry.scriptId, scriptTelemetry.startTime, scriptTelemetry.msResponseTime);
                    }
                    else {
                        OSF.AppTelemetry.onScriptDone(scriptTelemetry.scriptId, scriptTelemetry.startTime, scriptTelemetry.msResponseTime, this.osfControlAppCorrelationId);
                    }
                }
                this.scriptTelemetryBuffer = [];
            }
        };
        return LoadScriptHelper;
    })();
    ScriptLoading.LoadScriptHelper = LoadScriptHelper;
})(ScriptLoading || (ScriptLoading = {}));
var OfficeExt;
(function (OfficeExt) {
    var HostName;
    (function (HostName) {
        var Host = (function () {
            function Host() {
                this.getDiagnostics = function _getDiagnostics(version) {
                    var diagnostics = {
                        host: this.getHost(),
                        version: (version || this.getDefaultVersion()),
                        platform: this.getPlatform()
                    };
                    return diagnostics;
                };
                this.platformRemappings = {
                    web: Microsoft.Office.WebExtension.PlatformType.OfficeOnline,
                    winrt: Microsoft.Office.WebExtension.PlatformType.Universal,
                    win32: Microsoft.Office.WebExtension.PlatformType.PC,
                    mac: Microsoft.Office.WebExtension.PlatformType.Mac,
                    ios: Microsoft.Office.WebExtension.PlatformType.iOS,
                    android: Microsoft.Office.WebExtension.PlatformType.Android
                };
                this.camelCaseMappings = {
                    powerpoint: Microsoft.Office.WebExtension.HostType.PowerPoint,
                    onenote: Microsoft.Office.WebExtension.HostType.OneNote
                };
                this.hostInfo = OSF._OfficeAppFactory.getHostInfo();
                this.getHost = this.getHost.bind(this);
                this.getPlatform = this.getPlatform.bind(this);
                this.getDiagnostics = this.getDiagnostics.bind(this);
            }
            Host.prototype.capitalizeFirstLetter = function (input) {
                if (input) {
                    return (input[0].toUpperCase() + input.slice(1).toLowerCase());
                }
                return input;
            };
            Host.getInstance = function () {
                if (Host.hostObj === undefined) {
                    Host.hostObj = new Host();
                }
                return Host.hostObj;
            };
            Host.prototype.getPlatform = function (appNumber) {
                if (this.hostInfo.hostPlatform) {
                    var hostPlatform = this.hostInfo.hostPlatform.toLowerCase();
                    if (this.platformRemappings[hostPlatform]) {
                        return this.platformRemappings[hostPlatform];
                    }
                }
                return null;
            };
            Host.prototype.getHost = function (appNumber) {
                if (this.hostInfo.hostType) {
                    var hostType = this.hostInfo.hostType.toLowerCase();
                    if (this.camelCaseMappings[hostType]) {
                        return this.camelCaseMappings[hostType];
                    }
                    hostType = this.capitalizeFirstLetter(this.hostInfo.hostType);
                    if (Microsoft.Office.WebExtension.HostType[hostType]) {
                        return Microsoft.Office.WebExtension.HostType[hostType];
                    }
                }
                return null;
            };
            Host.prototype.getDefaultVersion = function () {
                if (this.getHost()) {
                    return "16.0.0000.0000";
                }
                return null;
            };
            return Host;
        })();
        HostName.Host = Host;
    })(HostName = OfficeExt.HostName || (OfficeExt.HostName = {}));
})(OfficeExt || (OfficeExt = {}));
var Office;
(function (Office) {
    var _Internal;
    (function (_Internal) {
        var PromiseImpl;
        (function (PromiseImpl) {
            function Init() {
                return (function () {
                    "use strict";
                    function lib$es6$promise$utils$$objectOrFunction(x) {
                        return typeof x === 'function' || (typeof x === 'object' && x !== null);
                    }
                    function lib$es6$promise$utils$$isFunction(x) {
                        return typeof x === 'function';
                    }
                    function lib$es6$promise$utils$$isMaybeThenable(x) {
                        return typeof x === 'object' && x !== null;
                    }
                    var lib$es6$promise$utils$$_isArray;
                    if (!Array.isArray) {
                        lib$es6$promise$utils$$_isArray = function (x) {
                            return Object.prototype.toString.call(x) === '[object Array]';
                        };
                    }
                    else {
                        lib$es6$promise$utils$$_isArray = Array.isArray;
                    }
                    var lib$es6$promise$utils$$isArray = lib$es6$promise$utils$$_isArray;
                    var lib$es6$promise$asap$$len = 0;
                    var lib$es6$promise$asap$$toString = {}.toString;
                    var lib$es6$promise$asap$$vertxNext;
                    var lib$es6$promise$asap$$customSchedulerFn;
                    var lib$es6$promise$asap$$asap = function asap(callback, arg) {
                        lib$es6$promise$asap$$queue[lib$es6$promise$asap$$len] = callback;
                        lib$es6$promise$asap$$queue[lib$es6$promise$asap$$len + 1] = arg;
                        lib$es6$promise$asap$$len += 2;
                        if (lib$es6$promise$asap$$len === 2) {
                            if (lib$es6$promise$asap$$customSchedulerFn) {
                                lib$es6$promise$asap$$customSchedulerFn(lib$es6$promise$asap$$flush);
                            }
                            else {
                                lib$es6$promise$asap$$scheduleFlush();
                            }
                        }
                    };
                    function lib$es6$promise$asap$$setScheduler(scheduleFn) {
                        lib$es6$promise$asap$$customSchedulerFn = scheduleFn;
                    }
                    function lib$es6$promise$asap$$setAsap(asapFn) {
                        lib$es6$promise$asap$$asap = asapFn;
                    }
                    var lib$es6$promise$asap$$browserWindow = (typeof window !== 'undefined') ? window : undefined;
                    var lib$es6$promise$asap$$browserGlobal = lib$es6$promise$asap$$browserWindow || {};
                    var lib$es6$promise$asap$$BrowserMutationObserver = lib$es6$promise$asap$$browserGlobal.MutationObserver || lib$es6$promise$asap$$browserGlobal.WebKitMutationObserver;
                    var lib$es6$promise$asap$$isNode = typeof process !== 'undefined' && {}.toString.call(process) === '[object process]';
                    var lib$es6$promise$asap$$isWorker = typeof Uint8ClampedArray !== 'undefined' &&
                        typeof importScripts !== 'undefined' &&
                        typeof MessageChannel !== 'undefined';
                    function lib$es6$promise$asap$$useNextTick() {
                        var nextTick = process.nextTick;
                        var version = process.versions.node.match(/^(?:(\d+)\.)?(?:(\d+)\.)?(\*|\d+)$/);
                        if (Array.isArray(version) && version[1] === '0' && version[2] === '10') {
                            nextTick = setImmediate;
                        }
                        return function () {
                            nextTick(lib$es6$promise$asap$$flush);
                        };
                    }
                    function lib$es6$promise$asap$$useVertxTimer() {
                        return function () {
                            lib$es6$promise$asap$$vertxNext(lib$es6$promise$asap$$flush);
                        };
                    }
                    function lib$es6$promise$asap$$useMutationObserver() {
                        var iterations = 0;
                        var observer = new lib$es6$promise$asap$$BrowserMutationObserver(lib$es6$promise$asap$$flush);
                        var node = document.createTextNode('');
                        observer.observe(node, { characterData: true });
                        return function () {
                            node.data = (iterations = ++iterations % 2);
                        };
                    }
                    function lib$es6$promise$asap$$useMessageChannel() {
                        var channel = new MessageChannel();
                        channel.port1.onmessage = lib$es6$promise$asap$$flush;
                        return function () {
                            channel.port2.postMessage(0);
                        };
                    }
                    function lib$es6$promise$asap$$useSetTimeout() {
                        return function () {
                            setTimeout(lib$es6$promise$asap$$flush, 1);
                        };
                    }
                    var lib$es6$promise$asap$$queue = new Array(1000);
                    function lib$es6$promise$asap$$flush() {
                        for (var i = 0; i < lib$es6$promise$asap$$len; i += 2) {
                            var callback = lib$es6$promise$asap$$queue[i];
                            var arg = lib$es6$promise$asap$$queue[i + 1];
                            callback(arg);
                            lib$es6$promise$asap$$queue[i] = undefined;
                            lib$es6$promise$asap$$queue[i + 1] = undefined;
                        }
                        lib$es6$promise$asap$$len = 0;
                    }
                    var lib$es6$promise$asap$$scheduleFlush;
                    if (lib$es6$promise$asap$$isNode) {
                        lib$es6$promise$asap$$scheduleFlush = lib$es6$promise$asap$$useNextTick();
                    }
                    else if (lib$es6$promise$asap$$BrowserMutationObserver) {
                        lib$es6$promise$asap$$scheduleFlush = lib$es6$promise$asap$$useMutationObserver();
                    }
                    else if (lib$es6$promise$asap$$isWorker) {
                        lib$es6$promise$asap$$scheduleFlush = lib$es6$promise$asap$$useMessageChannel();
                    }
                    else {
                        lib$es6$promise$asap$$scheduleFlush = lib$es6$promise$asap$$useSetTimeout();
                    }
                    function lib$es6$promise$$internal$$noop() { }
                    var lib$es6$promise$$internal$$PENDING = void 0;
                    var lib$es6$promise$$internal$$FULFILLED = 1;
                    var lib$es6$promise$$internal$$REJECTED = 2;
                    var lib$es6$promise$$internal$$GET_THEN_ERROR = new lib$es6$promise$$internal$$ErrorObject();
                    function lib$es6$promise$$internal$$selfFullfillment() {
                        return new TypeError("You cannot resolve a promise with itself");
                    }
                    function lib$es6$promise$$internal$$cannotReturnOwn() {
                        return new TypeError('A promises callback cannot return that same promise.');
                    }
                    function lib$es6$promise$$internal$$getThen(promise) {
                        try {
                            return promise.then;
                        }
                        catch (error) {
                            lib$es6$promise$$internal$$GET_THEN_ERROR.error = error;
                            return lib$es6$promise$$internal$$GET_THEN_ERROR;
                        }
                    }
                    function lib$es6$promise$$internal$$tryThen(then, value, fulfillmentHandler, rejectionHandler) {
                        try {
                            then.call(value, fulfillmentHandler, rejectionHandler);
                        }
                        catch (e) {
                            return e;
                        }
                    }
                    function lib$es6$promise$$internal$$handleForeignThenable(promise, thenable, then) {
                        lib$es6$promise$asap$$asap(function (promise) {
                            var sealed = false;
                            var error = lib$es6$promise$$internal$$tryThen(then, thenable, function (value) {
                                if (sealed) {
                                    return;
                                }
                                sealed = true;
                                if (thenable !== value) {
                                    lib$es6$promise$$internal$$resolve(promise, value);
                                }
                                else {
                                    lib$es6$promise$$internal$$fulfill(promise, value);
                                }
                            }, function (reason) {
                                if (sealed) {
                                    return;
                                }
                                sealed = true;
                                lib$es6$promise$$internal$$reject(promise, reason);
                            }, 'Settle: ' + (promise._label || ' unknown promise'));
                            if (!sealed && error) {
                                sealed = true;
                                lib$es6$promise$$internal$$reject(promise, error);
                            }
                        }, promise);
                    }
                    function lib$es6$promise$$internal$$handleOwnThenable(promise, thenable) {
                        if (thenable._state === lib$es6$promise$$internal$$FULFILLED) {
                            lib$es6$promise$$internal$$fulfill(promise, thenable._result);
                        }
                        else if (thenable._state === lib$es6$promise$$internal$$REJECTED) {
                            lib$es6$promise$$internal$$reject(promise, thenable._result);
                        }
                        else {
                            lib$es6$promise$$internal$$subscribe(thenable, undefined, function (value) {
                                lib$es6$promise$$internal$$resolve(promise, value);
                            }, function (reason) {
                                lib$es6$promise$$internal$$reject(promise, reason);
                            });
                        }
                    }
                    function lib$es6$promise$$internal$$handleMaybeThenable(promise, maybeThenable) {
                        if (maybeThenable.constructor === promise.constructor) {
                            lib$es6$promise$$internal$$handleOwnThenable(promise, maybeThenable);
                        }
                        else {
                            var then = lib$es6$promise$$internal$$getThen(maybeThenable);
                            if (then === lib$es6$promise$$internal$$GET_THEN_ERROR) {
                                lib$es6$promise$$internal$$reject(promise, lib$es6$promise$$internal$$GET_THEN_ERROR.error);
                            }
                            else if (then === undefined) {
                                lib$es6$promise$$internal$$fulfill(promise, maybeThenable);
                            }
                            else if (lib$es6$promise$utils$$isFunction(then)) {
                                lib$es6$promise$$internal$$handleForeignThenable(promise, maybeThenable, then);
                            }
                            else {
                                lib$es6$promise$$internal$$fulfill(promise, maybeThenable);
                            }
                        }
                    }
                    function lib$es6$promise$$internal$$resolve(promise, value) {
                        if (promise === value) {
                            lib$es6$promise$$internal$$reject(promise, lib$es6$promise$$internal$$selfFullfillment());
                        }
                        else if (lib$es6$promise$utils$$objectOrFunction(value)) {
                            lib$es6$promise$$internal$$handleMaybeThenable(promise, value);
                        }
                        else {
                            lib$es6$promise$$internal$$fulfill(promise, value);
                        }
                    }
                    function lib$es6$promise$$internal$$publishRejection(promise) {
                        if (promise._onerror) {
                            promise._onerror(promise._result);
                        }
                        lib$es6$promise$$internal$$publish(promise);
                    }
                    function lib$es6$promise$$internal$$fulfill(promise, value) {
                        if (promise._state !== lib$es6$promise$$internal$$PENDING) {
                            return;
                        }
                        promise._result = value;
                        promise._state = lib$es6$promise$$internal$$FULFILLED;
                        if (promise._subscribers.length !== 0) {
                            lib$es6$promise$asap$$asap(lib$es6$promise$$internal$$publish, promise);
                        }
                    }
                    function lib$es6$promise$$internal$$reject(promise, reason) {
                        if (promise._state !== lib$es6$promise$$internal$$PENDING) {
                            return;
                        }
                        promise._state = lib$es6$promise$$internal$$REJECTED;
                        promise._result = reason;
                        lib$es6$promise$asap$$asap(lib$es6$promise$$internal$$publishRejection, promise);
                    }
                    function lib$es6$promise$$internal$$subscribe(parent, child, onFulfillment, onRejection) {
                        var subscribers = parent._subscribers;
                        var length = subscribers.length;
                        parent._onerror = null;
                        subscribers[length] = child;
                        subscribers[length + lib$es6$promise$$internal$$FULFILLED] = onFulfillment;
                        subscribers[length + lib$es6$promise$$internal$$REJECTED] = onRejection;
                        if (length === 0 && parent._state) {
                            lib$es6$promise$asap$$asap(lib$es6$promise$$internal$$publish, parent);
                        }
                    }
                    function lib$es6$promise$$internal$$publish(promise) {
                        var subscribers = promise._subscribers;
                        var settled = promise._state;
                        if (subscribers.length === 0) {
                            return;
                        }
                        var child, callback, detail = promise._result;
                        for (var i = 0; i < subscribers.length; i += 3) {
                            child = subscribers[i];
                            callback = subscribers[i + settled];
                            if (child) {
                                lib$es6$promise$$internal$$invokeCallback(settled, child, callback, detail);
                            }
                            else {
                                callback(detail);
                            }
                        }
                        promise._subscribers.length = 0;
                    }
                    function lib$es6$promise$$internal$$ErrorObject() {
                        this.error = null;
                    }
                    var lib$es6$promise$$internal$$TRY_CATCH_ERROR = new lib$es6$promise$$internal$$ErrorObject();
                    function lib$es6$promise$$internal$$tryCatch(callback, detail) {
                        try {
                            return callback(detail);
                        }
                        catch (e) {
                            lib$es6$promise$$internal$$TRY_CATCH_ERROR.error = e;
                            return lib$es6$promise$$internal$$TRY_CATCH_ERROR;
                        }
                    }
                    function lib$es6$promise$$internal$$invokeCallback(settled, promise, callback, detail) {
                        var hasCallback = lib$es6$promise$utils$$isFunction(callback), value, error, succeeded, failed;
                        if (hasCallback) {
                            value = lib$es6$promise$$internal$$tryCatch(callback, detail);
                            if (value === lib$es6$promise$$internal$$TRY_CATCH_ERROR) {
                                failed = true;
                                error = value.error;
                                value = null;
                            }
                            else {
                                succeeded = true;
                            }
                            if (promise === value) {
                                lib$es6$promise$$internal$$reject(promise, lib$es6$promise$$internal$$cannotReturnOwn());
                                return;
                            }
                        }
                        else {
                            value = detail;
                            succeeded = true;
                        }
                        if (promise._state !== lib$es6$promise$$internal$$PENDING) {
                        }
                        else if (hasCallback && succeeded) {
                            lib$es6$promise$$internal$$resolve(promise, value);
                        }
                        else if (failed) {
                            lib$es6$promise$$internal$$reject(promise, error);
                        }
                        else if (settled === lib$es6$promise$$internal$$FULFILLED) {
                            lib$es6$promise$$internal$$fulfill(promise, value);
                        }
                        else if (settled === lib$es6$promise$$internal$$REJECTED) {
                            lib$es6$promise$$internal$$reject(promise, value);
                        }
                    }
                    function lib$es6$promise$$internal$$initializePromise(promise, resolver) {
                        try {
                            resolver(function resolvePromise(value) {
                                lib$es6$promise$$internal$$resolve(promise, value);
                            }, function rejectPromise(reason) {
                                lib$es6$promise$$internal$$reject(promise, reason);
                            });
                        }
                        catch (e) {
                            lib$es6$promise$$internal$$reject(promise, e);
                        }
                    }
                    function lib$es6$promise$enumerator$$Enumerator(Constructor, input) {
                        var enumerator = this;
                        enumerator._instanceConstructor = Constructor;
                        enumerator.promise = new Constructor(lib$es6$promise$$internal$$noop);
                        if (enumerator._validateInput(input)) {
                            enumerator._input = input;
                            enumerator.length = input.length;
                            enumerator._remaining = input.length;
                            enumerator._init();
                            if (enumerator.length === 0) {
                                lib$es6$promise$$internal$$fulfill(enumerator.promise, enumerator._result);
                            }
                            else {
                                enumerator.length = enumerator.length || 0;
                                enumerator._enumerate();
                                if (enumerator._remaining === 0) {
                                    lib$es6$promise$$internal$$fulfill(enumerator.promise, enumerator._result);
                                }
                            }
                        }
                        else {
                            lib$es6$promise$$internal$$reject(enumerator.promise, enumerator._validationError());
                        }
                    }
                    lib$es6$promise$enumerator$$Enumerator.prototype._validateInput = function (input) {
                        return lib$es6$promise$utils$$isArray(input);
                    };
                    lib$es6$promise$enumerator$$Enumerator.prototype._validationError = function () {
                        return new Error('Array Methods must be provided an Array');
                    };
                    lib$es6$promise$enumerator$$Enumerator.prototype._init = function () {
                        this._result = new Array(this.length);
                    };
                    var lib$es6$promise$enumerator$$default = lib$es6$promise$enumerator$$Enumerator;
                    lib$es6$promise$enumerator$$Enumerator.prototype._enumerate = function () {
                        var enumerator = this;
                        var length = enumerator.length;
                        var promise = enumerator.promise;
                        var input = enumerator._input;
                        for (var i = 0; promise._state === lib$es6$promise$$internal$$PENDING && i < length; i++) {
                            enumerator._eachEntry(input[i], i);
                        }
                    };
                    lib$es6$promise$enumerator$$Enumerator.prototype._eachEntry = function (entry, i) {
                        var enumerator = this;
                        var c = enumerator._instanceConstructor;
                        if (lib$es6$promise$utils$$isMaybeThenable(entry)) {
                            if (entry.constructor === c && entry._state !== lib$es6$promise$$internal$$PENDING) {
                                entry._onerror = null;
                                enumerator._settledAt(entry._state, i, entry._result);
                            }
                            else {
                                enumerator._willSettleAt(c.resolve(entry), i);
                            }
                        }
                        else {
                            enumerator._remaining--;
                            enumerator._result[i] = entry;
                        }
                    };
                    lib$es6$promise$enumerator$$Enumerator.prototype._settledAt = function (state, i, value) {
                        var enumerator = this;
                        var promise = enumerator.promise;
                        if (promise._state === lib$es6$promise$$internal$$PENDING) {
                            enumerator._remaining--;
                            if (state === lib$es6$promise$$internal$$REJECTED) {
                                lib$es6$promise$$internal$$reject(promise, value);
                            }
                            else {
                                enumerator._result[i] = value;
                            }
                        }
                        if (enumerator._remaining === 0) {
                            lib$es6$promise$$internal$$fulfill(promise, enumerator._result);
                        }
                    };
                    lib$es6$promise$enumerator$$Enumerator.prototype._willSettleAt = function (promise, i) {
                        var enumerator = this;
                        lib$es6$promise$$internal$$subscribe(promise, undefined, function (value) {
                            enumerator._settledAt(lib$es6$promise$$internal$$FULFILLED, i, value);
                        }, function (reason) {
                            enumerator._settledAt(lib$es6$promise$$internal$$REJECTED, i, reason);
                        });
                    };
                    function lib$es6$promise$promise$all$$all(entries) {
                        return new lib$es6$promise$enumerator$$default(this, entries).promise;
                    }
                    var lib$es6$promise$promise$all$$default = lib$es6$promise$promise$all$$all;
                    function lib$es6$promise$promise$race$$race(entries) {
                        var Constructor = this;
                        var promise = new Constructor(lib$es6$promise$$internal$$noop);
                        if (!lib$es6$promise$utils$$isArray(entries)) {
                            lib$es6$promise$$internal$$reject(promise, new TypeError('You must pass an array to race.'));
                            return promise;
                        }
                        var length = entries.length;
                        function onFulfillment(value) {
                            lib$es6$promise$$internal$$resolve(promise, value);
                        }
                        function onRejection(reason) {
                            lib$es6$promise$$internal$$reject(promise, reason);
                        }
                        for (var i = 0; promise._state === lib$es6$promise$$internal$$PENDING && i < length; i++) {
                            lib$es6$promise$$internal$$subscribe(Constructor.resolve(entries[i]), undefined, onFulfillment, onRejection);
                        }
                        return promise;
                    }
                    var lib$es6$promise$promise$race$$default = lib$es6$promise$promise$race$$race;
                    function lib$es6$promise$promise$resolve$$resolve(object) {
                        var Constructor = this;
                        if (object && typeof object === 'object' && object.constructor === Constructor) {
                            return object;
                        }
                        var promise = new Constructor(lib$es6$promise$$internal$$noop);
                        lib$es6$promise$$internal$$resolve(promise, object);
                        return promise;
                    }
                    var lib$es6$promise$promise$resolve$$default = lib$es6$promise$promise$resolve$$resolve;
                    function lib$es6$promise$promise$reject$$reject(reason) {
                        var Constructor = this;
                        var promise = new Constructor(lib$es6$promise$$internal$$noop);
                        lib$es6$promise$$internal$$reject(promise, reason);
                        return promise;
                    }
                    var lib$es6$promise$promise$reject$$default = lib$es6$promise$promise$reject$$reject;
                    var lib$es6$promise$promise$$counter = 0;
                    function lib$es6$promise$promise$$needsResolver() {
                        throw new TypeError('You must pass a resolver function as the first argument to the promise constructor');
                    }
                    function lib$es6$promise$promise$$needsNew() {
                        throw new TypeError("Failed to construct 'Promise': Please use the 'new' operator, this object constructor cannot be called as a function.");
                    }
                    var lib$es6$promise$promise$$default = lib$es6$promise$promise$$Promise;
                    function lib$es6$promise$promise$$Promise(resolver) {
                        this._id = lib$es6$promise$promise$$counter++;
                        this._state = undefined;
                        this._result = undefined;
                        this._subscribers = [];
                        if (lib$es6$promise$$internal$$noop !== resolver) {
                            if (!lib$es6$promise$utils$$isFunction(resolver)) {
                                lib$es6$promise$promise$$needsResolver();
                            }
                            if (!(this instanceof lib$es6$promise$promise$$Promise)) {
                                lib$es6$promise$promise$$needsNew();
                            }
                            lib$es6$promise$$internal$$initializePromise(this, resolver);
                        }
                    }
                    lib$es6$promise$promise$$Promise.all = lib$es6$promise$promise$all$$default;
                    lib$es6$promise$promise$$Promise.race = lib$es6$promise$promise$race$$default;
                    lib$es6$promise$promise$$Promise.resolve = lib$es6$promise$promise$resolve$$default;
                    lib$es6$promise$promise$$Promise.reject = lib$es6$promise$promise$reject$$default;
                    lib$es6$promise$promise$$Promise._setScheduler = lib$es6$promise$asap$$setScheduler;
                    lib$es6$promise$promise$$Promise._setAsap = lib$es6$promise$asap$$setAsap;
                    lib$es6$promise$promise$$Promise._asap = lib$es6$promise$asap$$asap;
                    lib$es6$promise$promise$$Promise.prototype = {
                        constructor: lib$es6$promise$promise$$Promise,
                        then: function (onFulfillment, onRejection) {
                            var parent = this;
                            var state = parent._state;
                            if (state === lib$es6$promise$$internal$$FULFILLED && !onFulfillment || state === lib$es6$promise$$internal$$REJECTED && !onRejection) {
                                return this;
                            }
                            var child = new this.constructor(lib$es6$promise$$internal$$noop);
                            var result = parent._result;
                            if (state) {
                                var callback = arguments[state - 1];
                                lib$es6$promise$asap$$asap(function () {
                                    lib$es6$promise$$internal$$invokeCallback(state, child, callback, result);
                                });
                            }
                            else {
                                lib$es6$promise$$internal$$subscribe(parent, child, onFulfillment, onRejection);
                            }
                            return child;
                        },
                        'catch': function (onRejection) {
                            return this.then(null, onRejection);
                        }
                    };
                    return lib$es6$promise$promise$$default;
                }).call(this);
            }
            PromiseImpl.Init = Init;
        })(PromiseImpl = _Internal.PromiseImpl || (_Internal.PromiseImpl = {}));
    })(_Internal = Office._Internal || (Office._Internal = {}));
    var _Internal;
    (function (_Internal) {
        function isEdgeLessThan14() {
            var userAgent = window.navigator.userAgent;
            var versionIdx = userAgent.indexOf("Edge/");
            if (versionIdx >= 0) {
                userAgent = userAgent.substring(versionIdx + 5, userAgent.length);
                if (userAgent < "14.14393")
                    return true;
                else
                    return false;
            }
            return false;
        }
        function determinePromise() {
            if (typeof (window) === "undefined" && typeof (Promise) === "function") {
                return Promise;
            }
            if (typeof (window) !== "undefined" && window.Promise) {
                if (isEdgeLessThan14()) {
                    return _Internal.PromiseImpl.Init();
                }
                else {
                    return window.Promise;
                }
            }
            else {
                return _Internal.PromiseImpl.Init();
            }
        }
        _Internal.OfficePromise = determinePromise();
    })(_Internal = Office._Internal || (Office._Internal = {}));
    var OfficePromise = _Internal.OfficePromise;
    Office.Promise = OfficePromise;
})(Office || (Office = {}));
(function () {
    var previousConstantNames = OSF.ConstantNames || {};
    OSF.ConstantNames = {
        FileVersion: "16.0.10723.30000",
        OfficeJS: "office.js",
        OfficeDebugJS: "office.debug.js",
        DefaultLocale: "en-us",
        LocaleStringLoadingTimeout: 5000,
        MicrosoftAjaxId: "MSAJAX",
        OfficeStringsId: "OFFICESTRINGS",
        OfficeJsId: "OFFICEJS",
        HostFileId: "HOST",
        O15MappingId: "O15Mapping",
        OfficeStringJS: "office_strings.js",
        O15InitHelper: "o15apptofilemappingtable.js",
        SupportedLocales: OSF.SupportedLocales,
        AssociatedLocales: OSF.AssociatedLocales
    };
    for (var key in previousConstantNames) {
        OSF.ConstantNames[key] = previousConstantNames[key];
    }
})();
OSF.InitializationHelper = function OSF_InitializationHelper(hostInfo, webAppState, context, settings, hostFacade) {
    this._hostInfo = hostInfo;
    this._webAppState = webAppState;
    this._context = context;
    this._settings = settings;
    this._hostFacade = hostFacade;
};
OSF.InitializationHelper.prototype.saveAndSetDialogInfo = function OSF_InitializationHelper$saveAndSetDialogInfo(hostInfoValue) {
};
OSF.InitializationHelper.prototype.getAppContext = function OSF_InitializationHelper$getAppContext(wnd, gotAppContext) {
};
OSF.InitializationHelper.prototype.setAgaveHostCommunication = function OSF_InitializationHelper$setAgaveHostCommunication() {
};
OSF.InitializationHelper.prototype.prepareRightBeforeWebExtensionInitialize = function OSF_InitializationHelper$prepareRightBeforeWebExtensionInitialize(appContext) {
};
OSF.InitializationHelper.prototype.loadAppSpecificScriptAndCreateOM = function OSF_InitializationHelper$loadAppSpecificScriptAndCreateOM(appContext, appReady, basePath) {
};
OSF.InitializationHelper.prototype.prepareRightAfterWebExtensionInitialize = function OSF_InitializationHelper$prepareRightAfterWebExtensionInitialize() {
};
OSF._OfficeAppFactory = (function OSF__OfficeAppFactory() {
    var _setNamespace = function OSF_OUtil$_setNamespace(name, parent) {
        if (parent && name && !parent[name]) {
            parent[name] = {};
        }
    };
    _setNamespace("Office", window);
    _setNamespace("Microsoft", window);
    _setNamespace("Office", Microsoft);
    _setNamespace("WebExtension", Microsoft.Office);
    if (window.Office.Promise) {
        Microsoft.Office.WebExtension.Promise = window.Office.Promise;
    }
    window.Office = Microsoft.Office.WebExtension;
    Microsoft.Office.WebExtension.PlatformType = {
        PC: "PC",
        OfficeOnline: "OfficeOnline",
        Mac: "Mac",
        iOS: "iOS",
        Android: "Android",
        Universal: "Universal"
    };
    Microsoft.Office.WebExtension.HostType = {
        Word: "Word",
        Excel: "Excel",
        PowerPoint: "PowerPoint",
        Outlook: "Outlook",
        OneNote: "OneNote",
        Project: "Project",
        Access: "Access"
    };
    var _context = {};
    var _settings = {};
    var _hostFacade = {};
    var _WebAppState = { id: null, webAppUrl: null, conversationID: null, clientEndPoint: null, wnd: window.parent, focused: false };
    var _hostInfo = { isO15: true, isRichClient: true, hostType: "", hostPlatform: "", hostSpecificFileVersion: "", hostLocale: "", osfControlAppCorrelationId: "", isDialog: false, disableLogging: false };
    var _isLoggingAllowed = true;
    var _initializationHelper = {};
    var _appInstanceId = null;
    var _isOfficeJsLoaded = false;
    var _officeOnReadyPendingResolves = [];
    var _isOfficeOnReadyCalled = false;
    var _loadScriptHelper = new ScriptLoading.LoadScriptHelper({
        OfficeJS: OSF.ConstantNames.OfficeJS,
        OfficeDebugJS: OSF.ConstantNames.OfficeDebugJS
    });
    if (window.performance && window.performance.now) {
        _loadScriptHelper.logScriptLoading(OSF.ConstantNames.OfficeJsId, -1, window.performance.now());
    }
    var _windowLocationHash = window.location.hash;
    var _windowLocationSearch = window.location.search;
    var _windowName = window.name;
    var getHostAndPlatform = function (appNumber) {
        return {
            host: OfficeExt.HostName.Host.getInstance().getHost(appNumber),
            platform: OfficeExt.HostName.Host.getInstance().getPlatform(appNumber)
        };
    };
    var setOfficeJsAsLoadedAndDispatchPendingOnReadyCallbacks = function (_a) {
        var host = _a.host, platform = _a.platform;
        _isOfficeJsLoaded = true;
        while (_officeOnReadyPendingResolves.length > 0) {
            _officeOnReadyPendingResolves.shift()({ host: host, platform: platform });
        }
    };
    Microsoft.Office.WebExtension.onReady = function Microsoft_Office_WebExtension_onReady(callback) {
        _isOfficeOnReadyCalled = true;
        if (_isOfficeJsLoaded) {
            var _a = getHostAndPlatform(1), host = _a.host, platform = _a.platform;
            if (callback) {
                var result = callback({ host: host, platform: platform });
                if (result && result.then && typeof result.then === "function") {
                    return result.then(function () { return Office.Promise.resolve({ host: host, platform: platform }); });
                }
            }
            return Office.Promise.resolve({ host: host, platform: platform });
        }
        if (callback) {
            return new Office.Promise(function (resolve) {
                _officeOnReadyPendingResolves.push(function (receivedHostAndPlatform) {
                    var result = callback(receivedHostAndPlatform);
                    if (result && result.then && typeof result.then === "function") {
                        return result.then(function () { return resolve(receivedHostAndPlatform); });
                    }
                    resolve(receivedHostAndPlatform);
                });
            });
        }
        return new Office.Promise(function (resolve) {
            _officeOnReadyPendingResolves.push(resolve);
        });
    };
    var getQueryStringValue = function OSF__OfficeAppFactory$getQueryStringValue(paramName) {
        var hostInfoValue;
        var searchString = window.location.search;
        if (searchString) {
            var hostInfoParts = searchString.split(paramName + "=");
            if (hostInfoParts.length > 1) {
                var hostInfoValueRestString = hostInfoParts[1];
                var separatorRegex = new RegExp("[&#]", "g");
                var hostInfoValueParts = hostInfoValueRestString.split(separatorRegex);
                if (hostInfoValueParts.length > 0) {
                    hostInfoValue = hostInfoValueParts[0];
                }
            }
        }
        return hostInfoValue;
    };
    var compareVersions = function _compareVersions(version1, version2) {
        var splitVersion1 = version1.split(".");
        var splitVersion2 = version2.split(".");
        var iter;
        for (iter in splitVersion1) {
            if (parseInt(splitVersion1[iter]) < parseInt(splitVersion2[iter])) {
                return false;
            }
            else if (parseInt(splitVersion1[iter]) > parseInt(splitVersion2[iter])) {
                return true;
            }
        }
        return false;
    };
    var shouldLoadOldMacJs = function _shouldLoadOldMacJs() {
        var versionToUseNewJS = "15.30.1128.0";
        var currentHostVersion = window.external.GetContext().GetHostFullVersion();
        return !!compareVersions(versionToUseNewJS, currentHostVersion);
    };
    var _retrieveLoggingAllowed = function OSF__OfficeAppFactory$_retrieveLoggingAllowed() {
        _isLoggingAllowed = true;
        try {
            if (_hostInfo.disableLogging) {
                _isLoggingAllowed = false;
                return;
            }
            window.external = window.external || {};
            if (typeof window.external.GetLoggingAllowed === 'undefined') {
                _isLoggingAllowed = true;
            }
            else {
                _isLoggingAllowed = window.external.GetLoggingAllowed();
            }
        }
        catch (Exception) {
        }
    };
    var _retrieveHostInfo = function OSF__OfficeAppFactory$_retrieveHostInfo() {
        var hostInfoParaName = "_host_Info";
        var hostInfoValue = getQueryStringValue(hostInfoParaName);
        if (!hostInfoValue) {
            try {
                var windowNameObj = JSON.parse(_windowName);
                hostInfoValue = windowNameObj ? windowNameObj["hostInfo"] : null;
            }
            catch (Exception) {
            }
        }
        if (!hostInfoValue) {
            try {
                window.external = window.external || {};
                if (typeof agaveHost !== "undefined" && agaveHost.GetHostInfo) {
                    window.external.GetHostInfo = function () {
                        return agaveHost.GetHostInfo();
                    };
                }
                var fallbackHostInfo = window.external.GetHostInfo();
                if (fallbackHostInfo == "isDialog") {
                    _hostInfo.isO15 = true;
                    _hostInfo.isDialog = true;
                }
                else if (fallbackHostInfo.toLowerCase().indexOf("mac") !== -1 && fallbackHostInfo.toLowerCase().indexOf("outlook") !== -1 && shouldLoadOldMacJs()) {
                    _hostInfo.isO15 = true;
                }
                else {
                    var hostInfoParts = fallbackHostInfo.split(hostInfoParaName + "=");
                    if (hostInfoParts.length > 1) {
                        hostInfoValue = hostInfoParts[1];
                    }
                    else {
                        hostInfoValue = fallbackHostInfo;
                    }
                }
            }
            catch (Exception) {
            }
        }
        var getSessionStorage = function OSF__OfficeAppFactory$_retrieveHostInfo$getSessionStorage() {
            var osfSessionStorage = null;
            try {
                if (window.sessionStorage) {
                    osfSessionStorage = window.sessionStorage;
                }
            }
            catch (ex) {
            }
            return osfSessionStorage;
        };
        var osfSessionStorage = getSessionStorage();
        if (!hostInfoValue && osfSessionStorage && osfSessionStorage.getItem("hostInfoValue")) {
            hostInfoValue = osfSessionStorage.getItem("hostInfoValue");
        }
        if (hostInfoValue) {
            hostInfoValue = decodeURIComponent(hostInfoValue);
            _hostInfo.isO15 = false;
            var items = hostInfoValue.split("$");
            if (typeof items[2] == "undefined") {
                items = hostInfoValue.split("|");
            }
            _hostInfo.hostType = (typeof items[0] == "undefined") ? "" : items[0].toLowerCase();
            _hostInfo.hostPlatform = (typeof items[1] == "undefined") ? "" : items[1].toLowerCase();
            ;
            _hostInfo.hostSpecificFileVersion = (typeof items[2] == "undefined") ? "" : items[2].toLowerCase();
            _hostInfo.hostLocale = (typeof items[3] == "undefined") ? "" : items[3].toLowerCase();
            _hostInfo.osfControlAppCorrelationId = (typeof items[4] == "undefined") ? "" : items[4];
            if (_hostInfo.osfControlAppCorrelationId == "telemetry") {
                _hostInfo.osfControlAppCorrelationId = "";
            }
            _hostInfo.isDialog = (((typeof items[5]) != "undefined") && items[5] == "isDialog") ? true : false;
            _hostInfo.disableLogging = (((typeof items[6]) != "undefined") && items[6] == "disableLogging") ? true : false;
            var hostSpecificFileVersionValue = parseFloat(_hostInfo.hostSpecificFileVersion);
            var fallbackVersion = OSF.HostSpecificFileVersionDefault;
            if (OSF.HostSpecificFileVersionMap[_hostInfo.hostType] && OSF.HostSpecificFileVersionMap[_hostInfo.hostType][_hostInfo.hostPlatform]) {
                fallbackVersion = OSF.HostSpecificFileVersionMap[_hostInfo.hostType][_hostInfo.hostPlatform];
            }
            if (hostSpecificFileVersionValue > parseFloat(fallbackVersion)) {
                _hostInfo.hostSpecificFileVersion = fallbackVersion;
            }
            if (osfSessionStorage) {
                try {
                    osfSessionStorage.setItem("hostInfoValue", hostInfoValue);
                }
                catch (e) {
                }
            }
        }
        else {
            _hostInfo.isO15 = true;
            _hostInfo.hostLocale = getQueryStringValue("locale");
        }
    };
    var getAppContextAsync = function OSF__OfficeAppFactory$getAppContextAsync(wnd, gotAppContext) {
        if (OSF.AppTelemetry && OSF.AppTelemetry.logAppCommonMessage) {
            OSF.AppTelemetry.logAppCommonMessage("getAppContextAsync starts");
        }
        _initializationHelper.getAppContext(wnd, gotAppContext);
    };
    var initialize = function OSF__OfficeAppFactory$initialize() {
        _retrieveHostInfo();
        _retrieveLoggingAllowed();
        if (_hostInfo.hostPlatform == "web" && _hostInfo.isDialog && window == window.top && window.opener == null) {
            window.open('', '_self', '');
            window.close();
        }
        _loadScriptHelper.setAppCorrelationId(_hostInfo.osfControlAppCorrelationId);
        var basePath = _loadScriptHelper.getOfficeJsBasePath();
        var requiresMsAjax = false;
        if (!basePath)
            throw "Office Web Extension script library file name should be " + OSF.ConstantNames.OfficeJS + " or " + OSF.ConstantNames.OfficeDebugJS + ".";
        var isMicrosftAjaxLoaded = function OSF$isMicrosftAjaxLoaded() {
            if ((typeof (Sys) !== 'undefined' && typeof (Type) !== 'undefined' &&
                Sys.StringBuilder && typeof (Sys.StringBuilder) === "function" &&
                Type.registerNamespace && typeof (Type.registerNamespace) === "function" &&
                Type.registerClass && typeof (Type.registerClass) === "function") ||
                (typeof (OfficeExt) !== "undefined" && OfficeExt.MsAjaxError)) {
                return true;
            }
            else {
                return false;
            }
        };
        var officeStrings = null;
        var loadLocaleStrings = function OSF__OfficeAppFactory_initialize$loadLocaleStrings(appLocale) {
            var fallbackLocaleTried = false;
            var loadLocaleStringCallback = function OSF__OfficeAppFactory_initialize$loadLocaleStringCallback() {
                if (typeof Strings == 'undefined' || typeof Strings.OfficeOM == 'undefined') {
                    if (!fallbackLocaleTried) {
                        fallbackLocaleTried = true;
                        var fallbackLocaleStringFile = basePath + OSF.ConstantNames.DefaultLocale + "/" + OSF.ConstantNames.OfficeStringJS;
                        _loadScriptHelper.loadScript(fallbackLocaleStringFile, OSF.ConstantNames.OfficeStringsId, loadLocaleStringCallback, true, OSF.ConstantNames.LocaleStringLoadingTimeout);
                        return false;
                    }
                    else {
                        throw "Neither the locale, " + appLocale.toLowerCase() + ", provided by the host app nor the fallback locale " + OSF.ConstantNames.DefaultLocale + " are supported.";
                    }
                }
                else {
                    fallbackLocaleTried = false;
                    officeStrings = Strings.OfficeOM;
                }
            };
            if (!isMicrosftAjaxLoaded()) {
                window.Type = Function;
                Type.registerNamespace = function (ns) {
                    window[ns] = window[ns] || {};
                };
                Type.prototype.registerClass = function (cls) {
                    cls = {};
                };
            }
            var localeStringFile = basePath + OSF.getSupportedLocale(appLocale, OSF.ConstantNames.DefaultLocale) + "/" + OSF.ConstantNames.OfficeStringJS;
            _loadScriptHelper.loadScript(localeStringFile, OSF.ConstantNames.OfficeStringsId, loadLocaleStringCallback, true, OSF.ConstantNames.LocaleStringLoadingTimeout);
        };
        var onAppCodeAndMSAjaxReady = function OSF__OfficeAppFactory_initialize$onAppCodeAndMSAjaxReady(loadSuccess) {
            if (loadSuccess) {
                _initializationHelper = new OSF.InitializationHelper(_hostInfo, _WebAppState, _context, _settings, _hostFacade);
                if (_hostInfo.hostPlatform == "web" && _initializationHelper.saveAndSetDialogInfo) {
                    _initializationHelper.saveAndSetDialogInfo(getQueryStringValue("_host_Info"));
                }
                _initializationHelper.setAgaveHostCommunication();
                getAppContextAsync(_WebAppState.wnd, function (appContext) {
                    if (OSF.AppTelemetry && OSF.AppTelemetry.logAppCommonMessage) {
                        OSF.AppTelemetry.logAppCommonMessage("getAppContextAsync callback start");
                    }
                    _appInstanceId = appContext._appInstanceId;
                    var updateVersionInfo = function updateVersionInfo() {
                        var hostVersionItems = _hostInfo.hostSpecificFileVersion.split(".");
                        if (appContext.get_appMinorVersion) {
                            var isIOS = _hostInfo.hostPlatform == "ios";
                            if (!isIOS) {
                                if (isNaN(appContext.get_appMinorVersion())) {
                                    appContext._appMinorVersion = parseInt(hostVersionItems[1]);
                                }
                                else if (hostVersionItems.length > 1 && !isNaN(Number(hostVersionItems[1]))) {
                                    appContext._appMinorVersion = parseInt(hostVersionItems[1]);
                                }
                            }
                        }
                        if (_hostInfo.isDialog) {
                            appContext._isDialog = _hostInfo.isDialog;
                        }
                    };
                    updateVersionInfo();
                    var appReady = function appReady() {
                        _initializationHelper.prepareApiSurface && _initializationHelper.prepareApiSurface(appContext);
                        _loadScriptHelper.waitForFunction(function () { return (Microsoft.Office.WebExtension.initialize != undefined || _isOfficeOnReadyCalled); }, function (initializedDeclaredOrOfficeOnReadyCalled) {
                            if (initializedDeclaredOrOfficeOnReadyCalled) {
                                if (_initializationHelper.prepareApiSurface) {
                                    if (Microsoft.Office.WebExtension.initialize) {
                                        Microsoft.Office.WebExtension.initialize(_initializationHelper.getInitializationReason(appContext));
                                    }
                                }
                                else {
                                    _initializationHelper.prepareRightBeforeWebExtensionInitialize(appContext);
                                }
                                _initializationHelper.prepareRightAfterWebExtensionInitialize && _initializationHelper.prepareRightAfterWebExtensionInitialize();
                            }
                            else {
                                throw new Error("Office.js has not fully loaded. Your app must call \"Office.onReady()\" as part of it's loading sequence (or set the \"Office.initialize\" function). If your app has this functionality, try reloading this page.");
                            }
                        }, 400, 50);
                        setOfficeJsAsLoadedAndDispatchPendingOnReadyCallbacks(getHostAndPlatform(appContext.get_appName()));
                    };
                    if (!_loadScriptHelper.isScriptLoading(OSF.ConstantNames.OfficeStringsId)) {
                        loadLocaleStrings(appContext.get_appUILocale());
                    }
                    _loadScriptHelper.waitForScripts([OSF.ConstantNames.OfficeStringsId], function () {
                        if (officeStrings && !Strings.OfficeOM) {
                            Strings.OfficeOM = officeStrings;
                        }
                        _initializationHelper.loadAppSpecificScriptAndCreateOM(appContext, appReady, basePath);
                    });
                });
                if (_hostInfo.isO15) {
                    var wacXdmInfoIsMissing = (OSF.OUtil.parseXdmInfo() == null);
                    if (wacXdmInfoIsMissing) {
                        var isPlainBrowser = true;
                        if (window.external && typeof window.external.GetContext !== 'undefined') {
                            try {
                                window.external.GetContext();
                                isPlainBrowser = false;
                            }
                            catch (e) {
                            }
                        }
                        if (isPlainBrowser) {
                            setOfficeJsAsLoadedAndDispatchPendingOnReadyCallbacks({ host: null, platform: null });
                        }
                    }
                }
            }
            else {
                var errorMsg = "MicrosoftAjax.js is not loaded successfully.";
                if (OSF.AppTelemetry && OSF.AppTelemetry.logAppException) {
                    OSF.AppTelemetry.logAppException(errorMsg);
                }
                throw errorMsg;
            }
        };
        var onAppCodeReady = function OSF__OfficeAppFactory_initialize$onAppCodeReady() {
            if (OSF.AppTelemetry && OSF.AppTelemetry.setOsfControlAppCorrelationId) {
                OSF.AppTelemetry.setOsfControlAppCorrelationId(_hostInfo.osfControlAppCorrelationId);
            }
            if (_loadScriptHelper.isScriptLoading(OSF.ConstantNames.MicrosoftAjaxId)) {
                _loadScriptHelper.waitForScripts([OSF.ConstantNames.MicrosoftAjaxId], onAppCodeAndMSAjaxReady);
            }
            else {
                _loadScriptHelper.waitForFunction(isMicrosftAjaxLoaded, onAppCodeAndMSAjaxReady, 500, 100);
            }
        };
        if (_hostInfo.isO15) {
            _loadScriptHelper.loadScript(basePath + OSF.ConstantNames.O15InitHelper, OSF.ConstantNames.O15MappingId, onAppCodeReady);
        }
        else {
            var hostSpecificFileName = ([
                _hostInfo.hostType,
                _hostInfo.hostPlatform,
                _hostInfo.hostSpecificFileVersion,
                OSF.ConstantNames.HostFileScriptSuffix || null,
            ]
                .filter(function (part) { return part != null; })
                .join("-"))
                +
                    ".js";
            _loadScriptHelper.loadScript(basePath + hostSpecificFileName.toLowerCase(), OSF.ConstantNames.HostFileId, onAppCodeReady);
        }
        if (_hostInfo.hostLocale) {
            loadLocaleStrings(_hostInfo.hostLocale);
        }
        if (requiresMsAjax && !isMicrosftAjaxLoaded()) {
            var msAjaxCDNPath = (window.location.protocol.toLowerCase() === 'https:' ? 'https:' : 'http:') + '//ajax.aspnetcdn.com/ajax/3.5/MicrosoftAjax.js';
            _loadScriptHelper.loadScriptParallel(msAjaxCDNPath, OSF.ConstantNames.MicrosoftAjaxId);
        }
        window.confirm = function OSF__OfficeAppFactory_initialize$confirm(message) {
            throw new Error('Function window.confirm is not supported.');
        };
        window.alert = function OSF__OfficeAppFactory_initialize$alert(message) {
            throw new Error('Function window.alert is not supported.');
        };
        window.prompt = function OSF__OfficeAppFactory_initialize$prompt(message, defaultvalue) {
            throw new Error('Function window.prompt is not supported.');
        };
        var isOutlookAndroid = _hostInfo.hostType == "outlook" && _hostInfo.hostPlatform == "android";
        if (!isOutlookAndroid) {
            window.history.replaceState = null;
            window.history.pushState = null;
        }
    };
    initialize();
    return {
        getId: function OSF__OfficeAppFactory$getId() { return _WebAppState.id; },
        getClientEndPoint: function OSF__OfficeAppFactory$getClientEndPoint() { return _WebAppState.clientEndPoint; },
        getContext: function OSF__OfficeAppFactory$getContext() { return _context; },
        setContext: function OSF__OfficeAppFactory$setContext(context) { _context = context; },
        getHostInfo: function OSF_OfficeAppFactory$getHostInfo() { return _hostInfo; },
        getLoggingAllowed: function OSF_OfficeAppFactory$getLoggingAllowed() { return _isLoggingAllowed; },
        getHostFacade: function OSF__OfficeAppFactory$getHostFacade() { return _hostFacade; },
        setHostFacade: function setHostFacade(hostFacade) { _hostFacade = hostFacade; },
        getInitializationHelper: function OSF__OfficeAppFactory$getInitializationHelper() { return _initializationHelper; },
        getCachedSessionSettingsKey: function OSF__OfficeAppFactory$getCachedSessionSettingsKey() {
            return (_WebAppState.conversationID != null ? _WebAppState.conversationID : _appInstanceId) + "CachedSessionSettings";
        },
        getWebAppState: function OSF__OfficeAppFactory$getWebAppState() { return _WebAppState; },
        getWindowLocationHash: function OSF__OfficeAppFactory$getHash() { return _windowLocationHash; },
        getWindowLocationSearch: function OSF__OfficeAppFactory$getSearch() { return _windowLocationSearch; },
        getLoadScriptHelper: function OSF__OfficeAppFactory$getLoadScriptHelper() { return _loadScriptHelper; },
        getWindowName: function OSF__OfficeAppFactory$getWindowName() { return _windowName; }
    };
})();
var CustomFunctionMappings = {};


/////////////////////////////////////////////////


/******/ (function(modules) { // webpackBootstrap
/******/ 	// The module cache
/******/ 	var installedModules = {};
/******/
/******/ 	// The require function
/******/ 	function __webpack_require__(moduleId) {
/******/
/******/ 		// Check if module is in cache
/******/ 		if(installedModules[moduleId]) {
/******/ 			return installedModules[moduleId].exports;
/******/ 		}
/******/ 		// Create a new module (and put it into the cache)
/******/ 		var module = installedModules[moduleId] = {
/******/ 			i: moduleId,
/******/ 			l: false,
/******/ 			exports: {}
/******/ 		};
/******/
/******/ 		// Execute the module function
/******/ 		modules[moduleId].call(module.exports, module, module.exports, __webpack_require__);
/******/
/******/ 		// Flag the module as loaded
/******/ 		module.l = true;
/******/
/******/ 		// Return the exports of the module
/******/ 		return module.exports;
/******/ 	}
/******/
/******/
/******/ 	// expose the modules object (__webpack_modules__)
/******/ 	__webpack_require__.m = modules;
/******/
/******/ 	// expose the module cache
/******/ 	__webpack_require__.c = installedModules;
/******/
/******/ 	// define getter function for harmony exports
/******/ 	__webpack_require__.d = function(exports, name, getter) {
/******/ 		if(!__webpack_require__.o(exports, name)) {
/******/ 			Object.defineProperty(exports, name, { enumerable: true, get: getter });
/******/ 		}
/******/ 	};
/******/
/******/ 	// define __esModule on exports
/******/ 	__webpack_require__.r = function(exports) {
/******/ 		if(typeof Symbol !== 'undefined' && Symbol.toStringTag) {
/******/ 			Object.defineProperty(exports, Symbol.toStringTag, { value: 'Module' });
/******/ 		}
/******/ 		Object.defineProperty(exports, '__esModule', { value: true });
/******/ 	};
/******/
/******/ 	// create a fake namespace object
/******/ 	// mode & 1: value is a module id, require it
/******/ 	// mode & 2: merge all properties of value into the ns
/******/ 	// mode & 4: return value when already ns object
/******/ 	// mode & 8|1: behave like require
/******/ 	__webpack_require__.t = function(value, mode) {
/******/ 		if(mode & 1) value = __webpack_require__(value);
/******/ 		if(mode & 8) return value;
/******/ 		if((mode & 4) && typeof value === 'object' && value && value.__esModule) return value;
/******/ 		var ns = Object.create(null);
/******/ 		__webpack_require__.r(ns);
/******/ 		Object.defineProperty(ns, 'default', { enumerable: true, value: value });
/******/ 		if(mode & 2 && typeof value != 'string') for(var key in value) __webpack_require__.d(ns, key, function(key) { return value[key]; }.bind(null, key));
/******/ 		return ns;
/******/ 	};
/******/
/******/ 	// getDefaultExport function for compatibility with non-harmony modules
/******/ 	__webpack_require__.n = function(module) {
/******/ 		var getter = module && module.__esModule ?
/******/ 			function getDefault() { return module['default']; } :
/******/ 			function getModuleExports() { return module; };
/******/ 		__webpack_require__.d(getter, 'a', getter);
/******/ 		return getter;
/******/ 	};
/******/
/******/ 	// Object.prototype.hasOwnProperty.call
/******/ 	__webpack_require__.o = function(object, property) { return Object.prototype.hasOwnProperty.call(object, property); };
/******/
/******/ 	// __webpack_public_path__
/******/ 	__webpack_require__.p = "";
/******/
/******/
/******/ 	// Load entry module and return exports
/******/ 	return __webpack_require__(__webpack_require__.s = 3);
/******/ })
/************************************************************************/
/******/ ([
/* 0 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

// core.ts: It contains the richapi message definition to be sent to host and the HttpUtility methods that sends message to host.
// common.ts: It contains Action, ObjectPath and ClientRequestBase
// operational.ts: It contains helpers to build operational APIs
// batch.ts: It contains the batch support for context.sync()
// embedded.ts: It contains WAC embedded support.
var __extends = (this && this.__extends) || (function () {
    var extendStatics = Object.setPrototypeOf ||
        ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
        function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p]; };
    return function (d, b) {
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
Object.defineProperty(exports, "__esModule", { value: true });
var SessionBase = /** @class */ (function () {
    function SessionBase() {
    }
    SessionBase.prototype._resolveRequestUrlAndHeaderInfo = function () {
        return CoreUtility._createPromiseFromResult(null);
    };
    SessionBase.prototype._createRequestExecutorOrNull = function () {
        return null;
    };
    Object.defineProperty(SessionBase.prototype, "eventRegistration", {
        get: function () {
            return null;
        },
        enumerable: true,
        configurable: true
    });
    return SessionBase;
}());
exports.SessionBase = SessionBase;
var HttpUtility = /** @class */ (function () {
    function HttpUtility() {
    }
    HttpUtility.setCustomSendRequestFunc = function (func) {
        HttpUtility.s_customSendRequestFunc = func;
    };
    /**
     * Send request using XMLHttpRequest
     */
    HttpUtility.xhrSendRequestFunc = function (request) {
        return CoreUtility.createPromise(function (resolve, reject) {
            var xhr = new XMLHttpRequest();
            xhr.open(request.method, request.url);
            xhr.onload = function () {
                var resp = {
                    statusCode: xhr.status,
                    headers: CoreUtility._parseHttpResponseHeaders(xhr.getAllResponseHeaders()),
                    body: xhr.responseText
                };
                resolve(resp);
            };
            xhr.onerror = function () {
                reject(new _Internal.RuntimeError({
                    code: CoreErrorCodes.connectionFailure,
                    message: CoreUtility._getResourceString(CoreResourceStrings.connectionFailureWithStatus, xhr.statusText)
                }));
            };
            if (request.headers) {
                for (var key in request.headers) {
                    xhr.setRequestHeader(key, request.headers[key]);
                }
            }
            xhr.send(CoreUtility._getRequestBodyText(request));
        });
    };
    /**
     * Send request
     */
    HttpUtility.sendRequest = function (request) {
        HttpUtility.validateAndNormalizeRequest(request);
        var func = HttpUtility.s_customSendRequestFunc;
        if (!func) {
            func = HttpUtility.xhrSendRequestFunc;
        }
        return func(request);
    };
    HttpUtility.setCustomSendLocalDocumentRequestFunc = function (func) {
        HttpUtility.s_customSendLocalDocumentRequestFunc = func;
    };
    /**
     * Send request to local document
     */
    HttpUtility.sendLocalDocumentRequest = function (request) {
        HttpUtility.validateAndNormalizeRequest(request);
        var func;
        func = HttpUtility.s_customSendLocalDocumentRequestFunc || HttpUtility.officeJsSendLocalDocumentRequestFunc;
        return func(request);
    };
    HttpUtility.officeJsSendLocalDocumentRequestFunc = function (request) {
        request = CoreUtility._validateLocalDocumentRequest(request);
        var requestSafeArray = CoreUtility._buildRequestMessageSafeArray(request);
        return CoreUtility.createPromise(function (resolve, reject) {
            OSF.DDA.RichApi.executeRichApiRequestAsync(requestSafeArray, function (asyncResult) {
                var response;
                if (asyncResult.status == 'succeeded') {
                    response = {
                        statusCode: RichApiMessageUtility.getResponseStatusCode(asyncResult),
                        headers: RichApiMessageUtility.getResponseHeaders(asyncResult),
                        body: RichApiMessageUtility.getResponseBody(asyncResult)
                    };
                }
                else {
                    response = RichApiMessageUtility.buildHttpResponseFromOfficeJsError(asyncResult.error.code, asyncResult.error.message);
                }
                CoreUtility.log(JSON.stringify(response));
                resolve(response);
            });
        });
    };
    HttpUtility.validateAndNormalizeRequest = function (request) {
        if (CoreUtility.isNullOrUndefined(request)) {
            throw _Internal.RuntimeError._createInvalidArgError({
                argumentName: 'request'
            });
        }
        if (CoreUtility.isNullOrEmptyString(request.method)) {
            request.method = 'GET';
        }
        request.method = request.method.toUpperCase();
    };
    HttpUtility.logRequest = function (request) {
        if (CoreUtility._logEnabled) {
            CoreUtility.log('---HTTP Request---');
            CoreUtility.log(request.method + ' ' + request.url);
            if (request.headers) {
                for (var key in request.headers) {
                    CoreUtility.log(key + ': ' + request.headers[key]);
                }
            }
            if (HttpUtility._logBodyEnabled) {
                CoreUtility.log(CoreUtility._getRequestBodyText(request));
            }
        }
    };
    HttpUtility.logResponse = function (response) {
        if (CoreUtility._logEnabled) {
            CoreUtility.log('---HTTP Response---');
            CoreUtility.log('' + response.statusCode);
            if (response.headers) {
                for (var key in response.headers) {
                    CoreUtility.log(key + ': ' + response.headers[key]);
                }
            }
            if (HttpUtility._logBodyEnabled) {
                CoreUtility.log(response.body);
            }
        }
    };
    HttpUtility._logBodyEnabled = false;
    return HttpUtility;
}());
exports.HttpUtility = HttpUtility;
var HostBridge = /** @class */ (function () {
    function HostBridge(m_bridge) {
        var _this = this;
        this.m_bridge = m_bridge;
        this.m_promiseResolver = {};
        this.m_handlers = [];
        this.m_bridge.onMessageFromHost = function (messageText) {
            var message = JSON.parse(messageText);
            _this.dispatchMessage(message);
        };
    }
    HostBridge.init = function (bridge) {
        if (typeof bridge !== 'object' || !bridge) {
            return;
        }
        var instance = new HostBridge(bridge);
        HostBridge.s_instance = instance;
        HttpUtility.setCustomSendLocalDocumentRequestFunc(function (request) {
            request = CoreUtility._validateLocalDocumentRequest(request);
            var requestFlags = 0;
            if (!CoreUtility.isReadonlyRestRequest(request.method)) {
                requestFlags = 1 /* WriteOperation */;
            }
            var index = request.url.indexOf('?');
            if (index >= 0) {
                var query = request.url.substr(index + 1);
                var flagsInQueryString = CoreUtility._parseRequestFlagsFromQueryStringIfAny(query);
                if (flagsInQueryString >= 0) {
                    requestFlags = flagsInQueryString;
                }
            }
            var bridgeMessage = {
                id: HostBridge.nextId(),
                type: 1 /* request */,
                flags: requestFlags,
                message: request
            };
            return instance.sendMessageToHostAndExpectResponse(bridgeMessage).then(function (bridgeResponse) {
                var responseInfo = bridgeResponse.message;
                return responseInfo;
            });
        });
        for (var i = 0; i < HostBridge.s_onInitedHandlers.length; i++) {
            HostBridge.s_onInitedHandlers[i](instance);
        }
    };
    Object.defineProperty(HostBridge, "instance", {
        get: function () {
            return HostBridge.s_instance;
        },
        enumerable: true,
        configurable: true
    });
    HostBridge.prototype.sendMessageToHost = function (message) {
        this.m_bridge.sendMessageToHost(JSON.stringify(message));
    };
    HostBridge.prototype.sendMessageToHostAndExpectResponse = function (message) {
        var _this = this;
        var ret = CoreUtility.createPromise(function (resolve, reject) {
            _this.m_promiseResolver[message.id] = resolve;
        });
        this.m_bridge.sendMessageToHost(JSON.stringify(message));
        return ret;
    };
    HostBridge.prototype.addHostMessageHandler = function (handler) {
        this.m_handlers.push(handler);
    };
    HostBridge.prototype.removeHostMessageHandler = function (handler) {
        var index = this.m_handlers.indexOf(handler);
        if (index >= 0) {
            this.m_handlers.splice(index, 1);
        }
    };
    HostBridge.onInited = function (handler) {
        HostBridge.s_onInitedHandlers.push(handler);
        if (HostBridge.s_instance) {
            // If the instance is already inited, invoke the handler.
            handler(HostBridge.s_instance);
        }
    };
    HostBridge.prototype.dispatchMessage = function (message) {
        if (typeof message.id === 'number') {
            var resolve = this.m_promiseResolver[message.id];
            if (resolve) {
                resolve(message);
                delete this.m_promiseResolver[message.id];
                return;
            }
        }
        for (var i = 0; i < this.m_handlers.length; i++) {
            this.m_handlers[i](message);
        }
    };
    HostBridge.nextId = function () {
        return HostBridge.s_nextId++;
    };
    HostBridge.s_onInitedHandlers = [];
    HostBridge.s_nextId = 1;
    return HostBridge;
}());
exports.HostBridge = HostBridge;
if (typeof _richApiNativeBridge === 'object' && _richApiNativeBridge) {
    HostBridge.init(_richApiNativeBridge);
}
var _Internal;
(function (_Internal) {
    var RuntimeError = /** @class */ (function (_super) {
        __extends(RuntimeError, _super);
        function RuntimeError(error) {
            var _this = _super.call(this, typeof error === 'string' ? error : error.message) || this;
            Object.setPrototypeOf(_this, RuntimeError.prototype);
            _this.name = 'RichApi.Error';
            if (typeof error === 'string') {
                _this.message = error;
            }
            else {
                _this.code = error.code;
                _this.message = error.message;
                _this.traceMessages = error.traceMessages || [];
                _this.innerError = error.innerError || null;
                _this.debugInfo = _this._createDebugInfo(error.debugInfo || {});
            }
            return _this;
        }
        RuntimeError.prototype.toString = function () {
            return this.code + ': ' + this.message;
        };
        RuntimeError.prototype._createDebugInfo = function (partialDebugInfo) {
            var debugInfo = {
                code: this.code,
                message: this.message
            };
            debugInfo.toString = function () {
                return JSON.stringify(this);
            };
            for (var key in partialDebugInfo) {
                debugInfo[key] = partialDebugInfo[key];
            }
            if (this.innerError) {
                if (this.innerError instanceof _Internal.RuntimeError) {
                    debugInfo.innerError = this.innerError.debugInfo;
                }
                else {
                    debugInfo.innerError = this.innerError;
                }
            }
            return debugInfo;
        };
        RuntimeError._createInvalidArgError = function (error) {
            return new _Internal.RuntimeError({
                code: CoreErrorCodes.invalidArgument,
                message: CoreUtility.isNullOrEmptyString(error.argumentName)
                    ? CoreUtility._getResourceString(CoreResourceStrings.invalidArgumentGeneric)
                    : CoreUtility._getResourceString(CoreResourceStrings.invalidArgument, error.argumentName),
                debugInfo: error.errorLocation ? { errorLocation: error.errorLocation } : {},
                innerError: error.innerError
            });
        };
        return RuntimeError;
    }(Error));
    _Internal.RuntimeError = RuntimeError;
})(_Internal = exports._Internal || (exports._Internal = {}));
// Export a publically-visible class so that can do checks like "if (e instanceof Error)"
exports.Error = _Internal.RuntimeError;
var CoreErrorCodes = /** @class */ (function () {
    function CoreErrorCodes() {
    }
    CoreErrorCodes.apiNotFound = 'ApiNotFound';
    CoreErrorCodes.accessDenied = 'AccessDenied';
    CoreErrorCodes.generalException = 'GeneralException';
    CoreErrorCodes.activityLimitReached = 'ActivityLimitReached';
    CoreErrorCodes.invalidArgument = 'InvalidArgument';
    CoreErrorCodes.connectionFailure = 'ConnectionFailure';
    CoreErrorCodes.timeout = 'Timeout';
    CoreErrorCodes.invalidOrTimedOutSession = 'InvalidOrTimedOutSession';
    CoreErrorCodes.invalidObjectPath = 'InvalidObjectPath';
    CoreErrorCodes.invalidRequestContext = 'InvalidRequestContext';
    CoreErrorCodes.valueNotLoaded = 'ValueNotLoaded';
    return CoreErrorCodes;
}());
exports.CoreErrorCodes = CoreErrorCodes;
var CoreResourceStrings = /** @class */ (function () {
    function CoreResourceStrings() {
    }
    // IMPORTANT! Please add the default english resource string value to
    // both ResourceStringValues.ts and
    // %SRCROOT%\osfweb\jscript\office_strings.js.resx
    // Note that in the office_strings.js.resx file, each string will be
    // prefixed with "L_" (e.g., "L_PropertyDoesNotExist")
    /** Message when Api is not available (as determined via the codegen-ed
          "throwIfApiNotSupported" method call).
          {0} is method/prop name, {1} is API Set name, and {2} is the application name (e.g., "Excel") */
    CoreResourceStrings.apiNotFoundDetails = 'ApiNotFoundDetails';
    CoreResourceStrings.connectionFailureWithStatus = 'ConnectionFailureWithStatus';
    CoreResourceStrings.connectionFailureWithDetails = 'ConnectionFailureWithDetails';
    /** An "invalid argument" that specifies the parameter name to be substituted into '{0}'. */
    CoreResourceStrings.invalidArgument = 'InvalidArgument';
    /** A generic "invalid argument" that does NOT specify the parameter name. */
    CoreResourceStrings.invalidArgumentGeneric = 'InvalidArgumentGeneric';
    CoreResourceStrings.timeout = 'Timeout';
    CoreResourceStrings.invalidOrTimedOutSessionMessage = 'InvalidOrTimedOutSessionMessage';
    CoreResourceStrings.invalidObjectPath = 'InvalidObjectPath';
    CoreResourceStrings.invalidRequestContext = 'InvalidRequestContext';
    CoreResourceStrings.valueNotLoaded = 'ValueNotLoaded';
    return CoreResourceStrings;
}());
exports.CoreResourceStrings = CoreResourceStrings;
var CoreConstants = /** @class */ (function () {
    function CoreConstants() {
    }
    CoreConstants.flags = 'flags';
    CoreConstants.sourceLibHeader = 'SdkVersion';
    CoreConstants.processQuery = 'ProcessQuery';
    CoreConstants.localDocument = 'http://document.localhost/';
    CoreConstants.localDocumentApiPrefix = 'http://document.localhost/_api/';
    return CoreConstants;
}());
exports.CoreConstants = CoreConstants;
var RichApiMessageUtility = /** @class */ (function () {
    function RichApiMessageUtility() {
    }
    RichApiMessageUtility.buildMessageArrayForIRequestExecutor = function (customData, requestFlags, requestMessage, sourceLibHeaderValue) {
        var requestMessageText = JSON.stringify(requestMessage.Body);
        CoreUtility.log('Request:');
        CoreUtility.log(requestMessageText);
        var headers = {};
        headers[CoreConstants.sourceLibHeader] = sourceLibHeaderValue;
        var messageSafearray = RichApiMessageUtility.buildRequestMessageSafeArray(customData, requestFlags, 'POST', CoreConstants.processQuery, headers, requestMessageText);
        return messageSafearray;
    };
    RichApiMessageUtility.buildResponseOnSuccess = function (responseBody, responseHeaders) {
        var response = { ErrorCode: '', ErrorMessage: '', Headers: null, Body: null };
        response.Body = JSON.parse(responseBody);
        response.Headers = responseHeaders;
        return response;
    };
    RichApiMessageUtility.buildResponseOnError = function (errorCode, message) {
        var response = { ErrorCode: '', ErrorMessage: '', Headers: null, Body: null };
        response.ErrorCode = CoreErrorCodes.generalException;
        response.ErrorMessage = message;
        if (errorCode == RichApiMessageUtility.OfficeJsErrorCode_ooeNoCapability) {
            response.ErrorCode = CoreErrorCodes.accessDenied;
        }
        else if (errorCode == RichApiMessageUtility.OfficeJsErrorCode_ooeActivityLimitReached) {
            response.ErrorCode = CoreErrorCodes.activityLimitReached;
        }
        else if (errorCode == RichApiMessageUtility.OfficeJsErrorCode_ooeInvalidOrTimedOutSession) {
            response.ErrorCode = CoreErrorCodes.invalidOrTimedOutSession;
            response.ErrorMessage = CoreUtility._getResourceString(CoreResourceStrings.invalidOrTimedOutSessionMessage);
        }
        return response;
    };
    RichApiMessageUtility.buildHttpResponseFromOfficeJsError = function (errorCode, message) {
        var statusCode = 500;
        var errorBody = {};
        errorBody['error'] = {};
        errorBody['error']['code'] = CoreErrorCodes.generalException;
        errorBody['error']['message'] = message;
        if (errorCode === RichApiMessageUtility.OfficeJsErrorCode_ooeNoCapability) {
            statusCode = 403;
            errorBody['error']['code'] = CoreErrorCodes.accessDenied;
        }
        else if (errorCode === RichApiMessageUtility.OfficeJsErrorCode_ooeActivityLimitReached) {
            statusCode = 429;
            errorBody['error']['code'] = CoreErrorCodes.activityLimitReached;
        }
        return { statusCode: statusCode, headers: {}, body: JSON.stringify(errorBody) };
    };
    RichApiMessageUtility.buildRequestMessageSafeArray = function (customData, requestFlags, method, path, headers, body) {
        var headerArray = [];
        if (headers) {
            for (var headerName in headers) {
                headerArray.push(headerName);
                headerArray.push(headers[headerName]);
            }
        }
        // Following fields will be updated by the Agave framework before the
        // message was sent to server.
        var appPermission = 0;
        var solutionId = '';
        var instanceId = '';
        var marketplaceType = '';
        return [
            customData,
            method,
            path,
            headerArray,
            body,
            appPermission,
            requestFlags,
            solutionId,
            instanceId,
            marketplaceType
        ];
    };
    RichApiMessageUtility.getResponseBody = function (result /*OSF.DDA.RichApi.ExecuteRichApiRequestResult*/) {
        return RichApiMessageUtility.getResponseBodyFromSafeArray(result.value.data);
    };
    RichApiMessageUtility.getResponseHeaders = function (result /*OSF.DDA.RichApi.ExecuteRichApiRequestResult*/) {
        return RichApiMessageUtility.getResponseHeadersFromSafeArray(result.value.data);
    };
    RichApiMessageUtility.getResponseBodyFromSafeArray = function (data) {
        var ret = data[2 /* Body */];
        if (typeof ret === 'string') {
            return ret;
        }
        var arr = ret;
        return arr.join('');
    };
    RichApiMessageUtility.getResponseHeadersFromSafeArray = function (data) {
        var arrayHeader = data[1 /* Headers */];
        if (!arrayHeader) {
            return null;
        }
        var headers = {};
        for (var i = 0; i < arrayHeader.length - 1; i += 2) {
            headers[arrayHeader[i]] = arrayHeader[i + 1];
        }
        return headers;
    };
    RichApiMessageUtility.getResponseStatusCode = function (result /*OSF.DDA.RichApi.ExecuteRichApiRequestResult*/) {
        return RichApiMessageUtility.getResponseStatusCodeFromSafeArray(result.value.data);
    };
    RichApiMessageUtility.getResponseStatusCodeFromSafeArray = function (data) {
        return data[0 /* StatusCode */];
    };
    RichApiMessageUtility.OfficeJsErrorCode_ooeInvalidOrTimedOutSession = 5012;
    RichApiMessageUtility.OfficeJsErrorCode_ooeActivityLimitReached = 5102;
    RichApiMessageUtility.OfficeJsErrorCode_ooeNoCapability = 7000;
    return RichApiMessageUtility;
}());
exports.RichApiMessageUtility = RichApiMessageUtility;
(function (_Internal) {
    function getPromiseType() {
        if (typeof Promise !== 'undefined') {
            return Promise;
        }
        if (typeof Office !== 'undefined') {
            if (Office.Promise) {
                return Office.Promise;
            }
        }
        if (typeof OfficeExtension !== 'undefined') {
            if (OfficeExtension.Promise) {
                return OfficeExtension.Promise;
            }
        }
        // This statement should never be hit
        throw new _Internal.Error('No Promise implementation found');
    }
    _Internal.getPromiseType = getPromiseType;
})(_Internal = exports._Internal || (exports._Internal = {}));
var CoreUtility = /** @class */ (function () {
    function CoreUtility() {
    }
    CoreUtility.log = function (message) {
        if (CoreUtility._logEnabled && typeof console !== 'undefined' && console.log) {
            console.log(message);
        }
    };
    CoreUtility.checkArgumentNull = function (value, name) {
        if (CoreUtility.isNullOrUndefined(value)) {
            throw _Internal.RuntimeError._createInvalidArgError({ argumentName: name });
        }
    };
    CoreUtility.isNullOrUndefined = function (value) {
        if (value === null) {
            return true;
        }
        if (typeof value === 'undefined') {
            return true;
        }
        return false;
    };
    CoreUtility.isUndefined = function (value) {
        if (typeof value === 'undefined') {
            return true;
        }
        return false;
    };
    CoreUtility.isNullOrEmptyString = function (value) {
        if (value === null) {
            return true;
        }
        if (typeof value === 'undefined') {
            return true;
        }
        if (value.length == 0) {
            return true;
        }
        return false;
    };
    CoreUtility.isPlainJsonObject = function (value) {
        if (CoreUtility.isNullOrUndefined(value)) {
            return false;
        }
        if (typeof value !== 'object') {
            return false;
        }
        return Object.getPrototypeOf(value) === Object.getPrototypeOf({});
    };
    CoreUtility.trim = function (str) {
        return str.replace(new RegExp('^\\s+|\\s+$', 'g'), '');
    };
    CoreUtility.caseInsensitiveCompareString = function (str1, str2) {
        if (CoreUtility.isNullOrUndefined(str1)) {
            return CoreUtility.isNullOrUndefined(str2);
        }
        else {
            if (CoreUtility.isNullOrUndefined(str2)) {
                return false;
            }
            else {
                return str1.toUpperCase() == str2.toUpperCase();
            }
        }
    };
    CoreUtility.isReadonlyRestRequest = function (method) {
        return CoreUtility.caseInsensitiveCompareString(method, 'GET');
    };
    CoreUtility._getResourceString = function (resourceId, arg) {
        var ret;
        if (typeof window !== 'undefined' && window.Strings && window.Strings.OfficeOM) {
            var stringName = 'L_' + resourceId;
            var stringValue = window.Strings.OfficeOM[stringName];
            if (stringValue) {
                ret = stringValue;
            }
        }
        if (!ret) {
            ret = CoreUtility.s_resourceStringValues[resourceId];
        }
        if (!ret) {
            ret = resourceId;
        }
        if (!CoreUtility.isNullOrUndefined(arg)) {
            if (Array.isArray(arg)) {
                var arrArg = arg;
                ret = CoreUtility._formatString(ret, arrArg);
            }
            else {
                ret = ret.replace('{0}', arg);
            }
        }
        return ret;
    };
    CoreUtility._formatString = function (format, arrArg) {
        return format.replace(/\{\d\}/g, function (v) {
            var position = parseInt(v.substr(1, v.length - 2));
            if (position < arrArg.length) {
                return arrArg[position];
            }
            else {
                throw _Internal.RuntimeError._createInvalidArgError({ argumentName: 'format' });
            }
        });
    };
    Object.defineProperty(CoreUtility, "Promise", {
        get: function () {
            return _Internal.getPromiseType();
        },
        enumerable: true,
        configurable: true
    });
    CoreUtility.createPromise = function (executor) {
        var ret = new CoreUtility.Promise(executor);
        return ret;
    };
    CoreUtility._createPromiseFromResult = function (value) {
        return CoreUtility.createPromise(function (resolve, reject) {
            resolve(value);
        });
    };
    CoreUtility._createPromiseFromException = function (reason) {
        return CoreUtility.createPromise(function (resolve, reject) {
            reject(reason);
        });
    };
    CoreUtility._createTimeoutPromise = function (timeout) {
        return CoreUtility.createPromise(function (resolve, reject) {
            setTimeout(function () {
                resolve(null);
            }, timeout);
        });
    };
    CoreUtility._createInvalidArgError = function (error) {
        return _Internal.RuntimeError._createInvalidArgError(error);
    };
    CoreUtility._isLocalDocumentUrl = function (url) {
        return CoreUtility._getLocalDocumentUrlPrefixLength(url) > 0;
    };
    CoreUtility._getLocalDocumentUrlPrefixLength = function (url) {
        var localDocumentPrefixes = [
            'http://document.localhost',
            'https://document.localhost',
            '//document.localhost'
        ];
        var urlLower = url.toLowerCase().trim();
        for (var i = 0; i < localDocumentPrefixes.length; i++) {
            if (urlLower === localDocumentPrefixes[i]) {
                return localDocumentPrefixes[i].length;
            }
            else if (urlLower.substr(0, localDocumentPrefixes[i].length + 1) === localDocumentPrefixes[i] + '/') {
                return localDocumentPrefixes[i].length + 1;
            }
        }
        return 0;
    };
    CoreUtility._validateLocalDocumentRequest = function (request) {
        var index = CoreUtility._getLocalDocumentUrlPrefixLength(request.url);
        if (index <= 0) {
            throw _Internal.RuntimeError._createInvalidArgError({
                argumentName: 'request'
            });
        }
        var path = request.url.substr(index);
        var pathLower = path.toLowerCase();
        if (pathLower === '_api') {
            path = '';
        }
        else if (pathLower.substr(0, '_api/'.length) === '_api/') {
            path = path.substr('_api/'.length);
        }
        return {
            method: request.method,
            url: path,
            headers: request.headers,
            body: request.body
        };
    };
    CoreUtility._parseRequestFlagsFromQueryStringIfAny = function (queryString) {
        var parts = queryString.split('&');
        for (var i = 0; i < parts.length; i++) {
            var keyvalue = parts[i].split('=');
            if (keyvalue[0].toLowerCase() === CoreConstants.flags) {
                var flags = parseInt(keyvalue[1]);
                // Ensure the requestFlags is not out of range.
                flags = flags & 255 /* MaxMask */;
                return flags;
            }
        }
        return -1;
    };
    CoreUtility._getRequestBodyText = function (request) {
        var body = '';
        if (typeof request.body === 'string') {
            body = request.body;
        }
        else if (request.body && typeof request.body === 'object') {
            body = JSON.stringify(request.body);
        }
        return body;
    };
    CoreUtility._parseResponseBody = function (response) {
        if (typeof response.body === 'string') {
            var bodyText = CoreUtility.trim(response.body);
            return JSON.parse(bodyText);
        }
        else {
            return response.body;
        }
    };
    CoreUtility._buildRequestMessageSafeArray = function (request) {
        var requestFlags = 0 /* None */;
        if (!CoreUtility.isReadonlyRestRequest(request.method)) {
            requestFlags = 1 /* WriteOperation */;
        }
        if (request.url.substr(0, CoreConstants.processQuery.length).toLowerCase() ===
            CoreConstants.processQuery.toLowerCase()) {
            // for ProcessQuery request, check the flags from query string.
            var index = request.url.indexOf('?');
            if (index > 0) {
                var queryString = request.url.substr(index + 1);
                var flagsInQueryString = CoreUtility._parseRequestFlagsFromQueryStringIfAny(queryString);
                if (flagsInQueryString >= 0) {
                    requestFlags = flagsInQueryString;
                }
            }
        }
        return RichApiMessageUtility.buildRequestMessageSafeArray('', requestFlags, request.method, request.url, request.headers, CoreUtility._getRequestBodyText(request));
    };
    CoreUtility._parseHttpResponseHeaders = function (allResponseHeaders) {
        var responseHeaders = {};
        if (!CoreUtility.isNullOrEmptyString(allResponseHeaders)) {
            var regex = new RegExp('\r?\n');
            var entries = allResponseHeaders.split(regex);
            for (var i = 0; i < entries.length; i++) {
                var entry = entries[i];
                if (entry != null) {
                    var index = entry.indexOf(':');
                    if (index > 0) {
                        var key = entry.substr(0, index);
                        var value = entry.substr(index + 1);
                        key = CoreUtility.trim(key);
                        value = CoreUtility.trim(value);
                        responseHeaders[key.toUpperCase()] = value;
                    }
                }
            }
        }
        return responseHeaders;
    };
    CoreUtility._parseErrorResponse = function (responseInfo) {
        var errorObj = null;
        if (CoreUtility.isPlainJsonObject(responseInfo.body)) {
            errorObj = responseInfo.body;
        }
        else if (!CoreUtility.isNullOrEmptyString(responseInfo.body)) {
            var errorResponseBody = CoreUtility.trim(responseInfo.body);
            try {
                errorObj = JSON.parse(errorResponseBody);
            }
            catch (e) {
                // The server may return HTML instead of JSON in case of error.
                // We need to ignore the error
                CoreUtility.log('Error when parse ' + errorResponseBody);
            }
        }
        var errorMessage;
        var errorCode;
        if (!CoreUtility.isNullOrUndefined(errorObj) && typeof errorObj === 'object' && errorObj.error) {
            errorCode = errorObj.error.code;
            errorMessage = CoreUtility._getResourceString(CoreResourceStrings.connectionFailureWithDetails, [
                responseInfo.statusCode.toString(),
                errorObj.error.code,
                errorObj.error.message
            ]);
        }
        else {
            errorMessage = CoreUtility._getResourceString(CoreResourceStrings.connectionFailureWithStatus, responseInfo.statusCode.toString());
        }
        if (CoreUtility.isNullOrEmptyString(errorCode)) {
            errorCode = CoreErrorCodes.connectionFailure;
        }
        return { errorCode: errorCode, errorMessage: errorMessage };
    };
    CoreUtility._copyHeaders = function (src, dest) {
        if (src && dest) {
            for (var key in src) {
                dest[key] = src[key];
            }
        }
    };
    CoreUtility.addResourceStringValues = function (values) {
        for (var key in values) {
            CoreUtility.s_resourceStringValues[key] = values[key];
        }
    };
    CoreUtility._logEnabled = false;
    CoreUtility.s_resourceStringValues = {
        ApiNotFoundDetails: 'The method or property {0} is part of the {1} requirement set, which is not available in your version of {2}.',
        ConnectionFailureWithStatus: 'The request failed with status code of {0}.',
        ConnectionFailureWithDetails: 'The request failed with status code of {0}, error code {1} and the following error message: {2}',
        InvalidArgument: "The argument '{0}' doesn't work for this situation, is missing, or isn't in the right format.",
        InvalidObjectPath: 'The object path \'{0}\' isn\'t working for what you\'re trying to do. If you\'re using the object across multiple "context.sync" calls and outside the sequential execution of a ".run" batch, please use the "context.trackedObjects.add()" and "context.trackedObjects.remove()" methods to manage the object\'s lifetime.',
        InvalidRequestContext: 'Cannot use the object across different request contexts.',
        Timeout: 'The operation has timed out.',
        ValueNotLoaded: 'The value of the result object has not been loaded yet. Before reading the value property, call "context.sync()" on the associated request context.'
    };
    return CoreUtility;
}());
exports.CoreUtility = CoreUtility;


/***/ }),
/* 1 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

var __extends = (this && this.__extends) || (function () {
    var extendStatics = Object.setPrototypeOf ||
        ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
        function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p]; };
    return function (d, b) {
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
function __export(m) {
    for (var p in m) if (!exports.hasOwnProperty(p)) exports[p] = m[p];
}
Object.defineProperty(exports, "__esModule", { value: true });
var Core = __webpack_require__(0);
__export(__webpack_require__(0));
exports._internalConfig = {
    showDisposeInfoInDebugInfo: false,
    showInternalApiInDebugInfo: false,
    enableEarlyDispose: true,
    // Always polyfill the clientObject.update() method.
    alwaysPolyfillClientObjectUpdateMethod: false,
    // Always polyfill the clientObject.retrieve() method.
    alwaysPolyfillClientObjectRetrieveMethod: false,
    enableConcurrentFlag: true,
    enableUndoableFlag: true
};
/**
 * The configuration for office.js.
 */
exports.config = {
    /**
     * Determines whether to have extended error logging on failure.
     *
     * When true, the error object will include a "debugInfo.fullStatements" property that lists out all the actions that were part of the batch request, both before and after the point of failure.
     *
     * Having this feature on will introduce a performance penalty, and will also log possibly-sensitive data (e.g., the contents of the commands being sent to the host).
     * It is recommended that you only have it on during debugging.  Also, if you are logging the error.debugInfo to a database or analytics service,
     * you should strip out the "debugInfo.fullStatements" property before sending it.
     */
    extendedErrorLogging: false
};
var ClientObjectBase = /** @class */ (function () {
    function ClientObjectBase(contextBase, objectPath) {
        this.m_contextBase = contextBase;
        this.m_objectPath = objectPath;
    }
    Object.defineProperty(ClientObjectBase.prototype, "_objectPath", {
        get: function () {
            return this.m_objectPath;
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(ClientObjectBase.prototype, "_context", {
        get: function () {
            return this.m_contextBase;
        },
        enumerable: true,
        configurable: true
    });
    return ClientObjectBase;
}());
exports.ClientObjectBase = ClientObjectBase;
var Action = /** @class */ (function () {
    function Action(actionInfo, operationType, flags) {
        this.m_actionInfo = actionInfo;
        this.m_operationType = operationType;
        this.m_flags = flags;
    }
    Object.defineProperty(Action.prototype, "actionInfo", {
        get: function () {
            return this.m_actionInfo;
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(Action.prototype, "operationType", {
        get: function () {
            return this.m_operationType;
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(Action.prototype, "flags", {
        get: function () {
            return this.m_flags;
        },
        enumerable: true,
        configurable: true
    });
    return Action;
}());
exports.Action = Action;
var ObjectPath = /** @class */ (function () {
    function ObjectPath(objectPathInfo, parentObjectPath, isCollection, isInvalidAfterRequest, operationType, flags) {
        this.m_objectPathInfo = objectPathInfo;
        this.m_parentObjectPath = parentObjectPath;
        this.m_isCollection = isCollection;
        this.m_isInvalidAfterRequest = isInvalidAfterRequest;
        this.m_isValid = true;
        this.m_operationType = operationType;
        this.m_flags = flags;
    }
    Object.defineProperty(ObjectPath.prototype, "objectPathInfo", {
        get: function () {
            return this.m_objectPathInfo;
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(ObjectPath.prototype, "operationType", {
        get: function () {
            return this.m_operationType;
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(ObjectPath.prototype, "flags", {
        get: function () {
            return this.m_flags;
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(ObjectPath.prototype, "isCollection", {
        get: function () {
            return this.m_isCollection;
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(ObjectPath.prototype, "isInvalidAfterRequest", {
        get: function () {
            return this.m_isInvalidAfterRequest;
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(ObjectPath.prototype, "parentObjectPath", {
        get: function () {
            return this.m_parentObjectPath;
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(ObjectPath.prototype, "argumentObjectPaths", {
        get: function () {
            return this.m_argumentObjectPaths;
        },
        set: function (value) {
            this.m_argumentObjectPaths = value;
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(ObjectPath.prototype, "isValid", {
        get: function () {
            return this.m_isValid;
        },
        set: function (value) {
            this.m_isValid = value;
            if (!value &&
                this.m_objectPathInfo.ObjectPathType === 6 /* ReferenceId */ &&
                this.m_savedObjectPathInfo) {
                ObjectPath.copyObjectPathInfo(this.m_savedObjectPathInfo.pathInfo, this.m_objectPathInfo);
                this.m_parentObjectPath = this.m_savedObjectPathInfo.parent;
                this.m_isValid = true;
                this.m_savedObjectPathInfo = null;
            }
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(ObjectPath.prototype, "originalObjectPathInfo", {
        get: function () {
            return this.m_originalObjectPathInfo;
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(ObjectPath.prototype, "getByIdMethodName", {
        get: function () {
            return this.m_getByIdMethodName;
        },
        set: function (value) {
            this.m_getByIdMethodName = value;
        },
        enumerable: true,
        configurable: true
    });
    ObjectPath.prototype._updateAsNullObject = function () {
        this.resetForUpdateUsingObjectData();
        this.m_objectPathInfo.ObjectPathType = 7 /* NullObject */;
        this.m_objectPathInfo.Name = '';
        this.m_parentObjectPath = null;
    };
    ObjectPath.prototype.saveOriginalObjectPathInfo = function () {
        if (exports.config.extendedErrorLogging && !this.m_originalObjectPathInfo) {
            this.m_originalObjectPathInfo = {};
            ObjectPath.copyObjectPathInfo(this.m_objectPathInfo, this.m_originalObjectPathInfo);
        }
    };
    ObjectPath.prototype.updateUsingObjectData = function (value, clientObject) {
        var referenceId = value[CommonConstants.referenceId];
        if (!Core.CoreUtility.isNullOrEmptyString(referenceId)) {
            if (!this.m_savedObjectPathInfo &&
                !this.isInvalidAfterRequest &&
                ObjectPath.isRestorableObjectPath(this.m_objectPathInfo.ObjectPathType)) {
                // The object has referenceId and it's not invalidateAfterRequest.
                // For example, the Font property of Range object (Range.Font).
                // The Font object has ReferenceId and the Range.Font object is not marked as invalidateAfterRequest.
                //
                // Because the object has ReferenceId, it will be invalidated at the end of Word.run().
                // In the above example, the Font object will be invalidated at the end of Word.run().
                //
                // Suppose the Range object is added into the trackedObjects, we should still be able
                // to access the range.Font property and invoke method on the range.Font in the next Word.run().
                // However, because the Font object has ReferenceId, it will be invadated at the end of Word.run().
                //
                // To make the above scenario work, we need to save the original object path info. When the object
                // is invalidated at the end of Word.run(), we will restore the object path to its original object path.
                //
                // var range;
                // await Word.run((context) => {
                //      range = context.document.getRange(...);
                //      range.font.load();
                //      range.track();
                //      await context.sync();
                // });
                //
                // As the range.font object path was changed to ReferenceId and we did not explicitly
                // track the range.font object, the range.font object's object path will be invalidated
                // as Word.run() cleanup code will release all references. To make it work, we will also
                // restore the object path to its m_savedObjectPathInfo.
                //
                // await Word.run(range, (context) => {
                //      range.font.load();
                //      As the range object is tracked and range.font object path was restored, it works now.
                //      await context.sync();
                // });
                //
                // Please note that object has m_savedObjectPathInfo only if the object is not marked with
                // InvalidateAfterRequest.
                var pathInfo = {};
                ObjectPath.copyObjectPathInfo(this.m_objectPathInfo, pathInfo);
                this.m_savedObjectPathInfo = {
                    pathInfo: pathInfo,
                    parent: this.m_parentObjectPath
                };
            }
            this.saveOriginalObjectPathInfo();
            this.resetForUpdateUsingObjectData();
            this.m_objectPathInfo.ObjectPathType = 6 /* ReferenceId */;
            this.m_objectPathInfo.Name = referenceId;
            // Remove parentObjectPathId because the object path is now re-mapped to a reference ID.
            // Its original parent is no longer relevant.
            delete this.m_objectPathInfo.ParentObjectPathId;
            this.m_parentObjectPath = null;
            return;
        }
        if (clientObject) {
            var collectionPropertyPath = clientObject[CommonConstants.collectionPropertyPath];
            if (!Core.CoreUtility.isNullOrEmptyString(collectionPropertyPath) && clientObject.context) {
                var id = CommonUtility.tryGetObjectIdFromLoadOrRetrieveResult(value);
                if (!Core.CoreUtility.isNullOrUndefined(id)) {
                    var propNames = collectionPropertyPath.split('.');
                    var parent_1 = clientObject.context[propNames[0]];
                    for (var i = 1; i < propNames.length; i++) {
                        parent_1 = parent_1[propNames[i]];
                    }
                    this.saveOriginalObjectPathInfo();
                    this.resetForUpdateUsingObjectData();
                    this.m_parentObjectPath = parent_1._objectPath;
                    this.m_objectPathInfo.ParentObjectPathId = this.m_parentObjectPath.objectPathInfo.Id;
                    this.m_objectPathInfo.ObjectPathType = 5 /* Indexer */;
                    this.m_objectPathInfo.Name = '';
                    this.m_objectPathInfo.ArgumentInfo.Arguments = [id];
                    return;
                }
            }
        }
        var parentIsCollection = this.parentObjectPath && this.parentObjectPath.isCollection;
        var getByIdMethodName = this.getByIdMethodName;
        if (parentIsCollection || !Core.CoreUtility.isNullOrEmptyString(getByIdMethodName)) {
            var id = CommonUtility.tryGetObjectIdFromLoadOrRetrieveResult(value);
            if (!Core.CoreUtility.isNullOrUndefined(id)) {
                this.saveOriginalObjectPathInfo();
                this.resetForUpdateUsingObjectData();
                if (!Core.CoreUtility.isNullOrEmptyString(getByIdMethodName)) {
                    this.m_objectPathInfo.ObjectPathType = 3 /* Method */;
                    this.m_objectPathInfo.Name = getByIdMethodName;
                    this.m_getByIdMethodName = null;
                }
                else {
                    // parentIsCollection
                    this.m_objectPathInfo.ObjectPathType = 5 /* Indexer */;
                    this.m_objectPathInfo.Name = '';
                }
                this.m_objectPathInfo.ArgumentInfo.Arguments = [id];
                return;
            }
        }
    };
    ObjectPath.prototype.resetForUpdateUsingObjectData = function () {
        this.m_isInvalidAfterRequest = false;
        this.m_isValid = true;
        this.m_operationType = 1 /* Read */;
        this.m_flags = 4 /* concurrent */;
        this.m_objectPathInfo.ArgumentInfo = {};
        this.m_argumentObjectPaths = null;
    };
    ObjectPath.isRestorableObjectPath = function (objectPathType) {
        return (objectPathType === 1 /* GlobalObject */ ||
            objectPathType === 5 /* Indexer */ ||
            objectPathType === 3 /* Method */ ||
            objectPathType === 4 /* Property */);
    };
    ObjectPath.copyObjectPathInfo = function (src, dest) {
        dest.Id = src.Id;
        dest.ArgumentInfo = src.ArgumentInfo;
        dest.Name = src.Name;
        dest.ObjectPathType = src.ObjectPathType;
        dest.ParentObjectPathId = src.ParentObjectPathId;
    };
    return ObjectPath;
}());
exports.ObjectPath = ObjectPath;
var ClientRequestContextBase = /** @class */ (function () {
    function ClientRequestContextBase() {
        this.m_nextId = 0;
    }
    ClientRequestContextBase.prototype._nextId = function () {
        return ++this.m_nextId;
    };
    ClientRequestContextBase.prototype._addServiceApiAction = function (action, resultHandler, resolve, reject) {
        if (!this.m_serviceApiQueue) {
            this.m_serviceApiQueue = new ServiceApiQueue(this);
        }
        this.m_serviceApiQueue.add(action, resultHandler, resolve, reject);
    };
    return ClientRequestContextBase;
}());
exports.ClientRequestContextBase = ClientRequestContextBase;
var InstantiateActionUpdateObjectPathHandler = /** @class */ (function () {
    function InstantiateActionUpdateObjectPathHandler(m_objectPath) {
        this.m_objectPath = m_objectPath;
    }
    InstantiateActionUpdateObjectPathHandler.prototype._handleResult = function (value) {
        if (Core.CoreUtility.isNullOrUndefined(value)) {
            this.m_objectPath._updateAsNullObject();
        }
        else {
            this.m_objectPath.updateUsingObjectData(value, null);
        }
    };
    return InstantiateActionUpdateObjectPathHandler;
}());
var ClientRequestBase = /** @class */ (function () {
    function ClientRequestBase(context) {
        this.m_contextBase = context;
        this.m_actions = [];
        this.m_actionResultHandler = {};
        this.m_referencedObjectPaths = {};
        this.m_instantiatedObjectPaths = {};
        this.m_preSyncPromises = [];
    }
    ClientRequestBase.prototype.addAction = function (action) {
        this.m_actions.push(action);
        if (action.actionInfo.ActionType == 1 /* Instantiate */) {
            this.m_instantiatedObjectPaths[action.actionInfo.ObjectPathId] = action;
        }
    };
    Object.defineProperty(ClientRequestBase.prototype, "hasActions", {
        get: function () {
            return this.m_actions.length > 0;
        },
        enumerable: true,
        configurable: true
    });
    ClientRequestBase.prototype._getLastAction = function () {
        return this.m_actions[this.m_actions.length - 1];
    };
    ClientRequestBase.prototype.ensureInstantiateObjectPath = function (objectPath) {
        if (objectPath) {
            if (this.m_instantiatedObjectPaths[objectPath.objectPathInfo.Id]) {
                return;
            }
            this.ensureInstantiateObjectPath(objectPath.parentObjectPath);
            this.ensureInstantiateObjectPaths(objectPath.argumentObjectPaths);
            if (!this.m_instantiatedObjectPaths[objectPath.objectPathInfo.Id]) {
                var actionInfo = {
                    Id: this.m_contextBase._nextId(),
                    ActionType: 1 /* Instantiate */,
                    Name: '',
                    ObjectPathId: objectPath.objectPathInfo.Id
                };
                var instantiateAction = new Action(actionInfo, 1 /* Read */, 4 /* concurrent */);
                instantiateAction.referencedObjectPath = objectPath;
                this.addReferencedObjectPath(objectPath);
                this.addAction(instantiateAction);
                var resultHandler = new InstantiateActionUpdateObjectPathHandler(objectPath);
                this.addActionResultHandler(instantiateAction, resultHandler);
            }
        }
    };
    ClientRequestBase.prototype.ensureInstantiateObjectPaths = function (objectPaths) {
        if (objectPaths) {
            for (var i = 0; i < objectPaths.length; i++) {
                this.ensureInstantiateObjectPath(objectPaths[i]);
            }
        }
    };
    ClientRequestBase.prototype.addReferencedObjectPath = function (objectPath) {
        if (this.m_referencedObjectPaths[objectPath.objectPathInfo.Id]) {
            return;
        }
        if (!objectPath.isValid) {
            throw new Core._Internal.RuntimeError({
                code: Core.CoreErrorCodes.invalidObjectPath,
                message: Core.CoreUtility._getResourceString(Core.CoreResourceStrings.invalidObjectPath, CommonUtility.getObjectPathExpression(objectPath)),
                debugInfo: {
                    errorLocation: CommonUtility.getObjectPathExpression(objectPath)
                }
            });
        }
        while (objectPath) {
            this.m_referencedObjectPaths[objectPath.objectPathInfo.Id] = objectPath;
            if (objectPath.objectPathInfo.ObjectPathType == 3 /* Method */) {
                this.addReferencedObjectPaths(objectPath.argumentObjectPaths);
            }
            objectPath = objectPath.parentObjectPath;
        }
    };
    ClientRequestBase.prototype.addReferencedObjectPaths = function (objectPaths) {
        if (objectPaths) {
            for (var i = 0; i < objectPaths.length; i++) {
                this.addReferencedObjectPath(objectPaths[i]);
            }
        }
    };
    ClientRequestBase.prototype.addActionResultHandler = function (action, resultHandler) {
        this.m_actionResultHandler[action.actionInfo.Id] = resultHandler;
    };
    ClientRequestBase.prototype.aggregrateRequestFlags = function (requestFlags, operationType, flags) {
        if (operationType === 0 /* Default */) {
            // set writeoperation
            requestFlags = requestFlags | 1 /* WriteOperation */;
            if ((flags & 2 /* undoable */) === 0) {
                // clear undoable
                requestFlags = requestFlags & ~16 /* Undoable */;
            }
            // clear concurrent
            requestFlags = requestFlags & ~4 /* Concurrent */;
        }
        if (flags & 1 /* restrictedResourceAccess */) {
            // set restricted resource
            requestFlags = requestFlags | 2 /* RestrictedResourceAccess */;
        }
        if ((flags & 4 /* concurrent */) === 0) {
            // clear concurrent
            requestFlags = requestFlags & ~4 /* Concurrent */;
        }
        return requestFlags;
    };
    ClientRequestBase.prototype.finallyNormalizeFlags = function (requestFlags) {
        // Undoable only makes sense for write request.
        if ((requestFlags & 1 /* WriteOperation */) === 0 /* None */) {
            // It's not write operation, clear the Undoable flag.
            requestFlags = requestFlags & ~16 /* Undoable */;
        }
        if (!exports._internalConfig.enableConcurrentFlag) {
            // In case there is issue with the host concurrent implementation, we could have a quick way to turn the feature off.
            requestFlags = requestFlags & ~4 /* Concurrent */;
        }
        if (!exports._internalConfig.enableUndoableFlag) {
            // In case there is issue with the host undoable implementation, we could have a quick way to turn the feature off.
            requestFlags = requestFlags & ~16 /* Undoable */;
        }
        if (!CommonUtility.isSetSupported('RichApiRuntimeFlag', '1.1')) {
            // If the host cannot accept the new flags, we should not send the new flags to the server.
            // For example, the Excel WAC server defines flags as C# enum. If the server code is not updated and the client
            // sends new flags to older Excel WAC server, such as older Excel WAC server on-premise installation, the older
            // Excel WAC server cannot serialize the request flags and will fail.
            requestFlags = requestFlags & ~4 /* Concurrent */;
            requestFlags = requestFlags & ~16 /* Undoable */;
        }
        if (typeof this.m_flagsForTesting === 'number') {
            // We want to be able to simulate the case when the client and server does not agree on the flags by setting another flag.
            requestFlags = this.m_flagsForTesting;
        }
        return requestFlags;
    };
    ClientRequestBase.prototype.buildRequestMessageBodyAndRequestFlags = function () {
        if (exports._internalConfig.enableEarlyDispose) {
            ClientRequestBase._calculateLastUsedObjectPathIds(this.m_actions);
        }
        var requestFlags = 4 /* Concurrent */ | 16 /* Undoable */;
        var objectPaths = {};
        for (var i in this.m_referencedObjectPaths) {
            requestFlags = this.aggregrateRequestFlags(requestFlags, this.m_referencedObjectPaths[i].operationType, this.m_referencedObjectPaths[i].flags);
            objectPaths[i] = this.m_referencedObjectPaths[i].objectPathInfo;
        }
        var actions = [];
        var hasKeepReference = false;
        for (var index = 0; index < this.m_actions.length; index++) {
            var action = this.m_actions[index];
            if (action.actionInfo.ActionType === 3 /* Method */ &&
                action.actionInfo.Name === CommonConstants.keepReference) {
                hasKeepReference = true;
            }
            requestFlags = this.aggregrateRequestFlags(requestFlags, action.operationType, action.flags);
            actions.push(action.actionInfo);
        }
        requestFlags = this.finallyNormalizeFlags(requestFlags);
        var body = {
            AutoKeepReference: this.m_contextBase._autoCleanup && hasKeepReference,
            Actions: actions,
            ObjectPaths: objectPaths
        };
        return {
            body: body,
            flags: requestFlags
        };
    };
    ClientRequestBase.prototype.processResponse = function (actionResults) {
        if (actionResults) {
            for (var i = 0; i < actionResults.length; i++) {
                var actionResult = actionResults[i];
                var handler = this.m_actionResultHandler[actionResult.ActionId];
                if (handler) {
                    handler._handleResult(actionResult.Value);
                }
            }
        }
    };
    ClientRequestBase.prototype.invalidatePendingInvalidObjectPaths = function () {
        for (var i in this.m_referencedObjectPaths) {
            if (this.m_referencedObjectPaths[i].isInvalidAfterRequest) {
                this.m_referencedObjectPaths[i].isValid = false;
            }
        }
    };
    ClientRequestBase.prototype._addPreSyncPromise = function (value) {
        this.m_preSyncPromises.push(value);
    };
    Object.defineProperty(ClientRequestBase.prototype, "_preSyncPromises", {
        get: function () {
            return this.m_preSyncPromises;
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(ClientRequestBase.prototype, "_actions", {
        get: function () {
            return this.m_actions;
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(ClientRequestBase.prototype, "_objectPaths", {
        get: function () {
            return this.m_referencedObjectPaths;
        },
        enumerable: true,
        configurable: true
    });
    ClientRequestBase.prototype._removeKeepReferenceAction = function (objectPathId) {
        for (var i = this.m_actions.length - 1; i >= 0; i--) {
            var actionInfo = this.m_actions[i].actionInfo;
            if (actionInfo.ObjectPathId === objectPathId &&
                actionInfo.ActionType === 3 /* Method */ &&
                actionInfo.Name === CommonConstants.keepReference) {
                this.m_actions.splice(i, 1);
                break;
            }
        }
    };
    ClientRequestBase._updateLastUsedActionIdOfObjectPathId = function (lastUsedActionIdOfObjectPathId, objectPath, actionId) {
        while (objectPath) {
            if (lastUsedActionIdOfObjectPathId[objectPath.objectPathInfo.Id]) {
                return;
            }
            lastUsedActionIdOfObjectPathId[objectPath.objectPathInfo.Id] = actionId;
            var argumentObjectPaths = objectPath.argumentObjectPaths;
            if (argumentObjectPaths) {
                var argumentObjectPathsLength = argumentObjectPaths.length;
                for (var i = 0; i < argumentObjectPathsLength; i++) {
                    ClientRequestBase._updateLastUsedActionIdOfObjectPathId(lastUsedActionIdOfObjectPathId, argumentObjectPaths[i], actionId);
                }
            }
            objectPath = objectPath.parentObjectPath;
        }
    };
    ClientRequestBase._calculateLastUsedObjectPathIds = function (actions) {
        var lastUsedActionIdOfObjectPathId = {};
        // lastUsedActionId is the object's last used action id.
        var actionsLength = actions.length;
        for (var index = actionsLength - 1; index >= 0; --index) {
            var action = actions[index];
            var actionId = action.actionInfo.Id;
            if (action.referencedObjectPath) {
                ClientRequestBase._updateLastUsedActionIdOfObjectPathId(lastUsedActionIdOfObjectPathId, action.referencedObjectPath, actionId);
            }
            var referencedObjectPaths = action.referencedArgumentObjectPaths;
            if (referencedObjectPaths) {
                var referencedObjectPathsLength = referencedObjectPaths.length;
                for (var refIndex = 0; refIndex < referencedObjectPathsLength; refIndex++) {
                    ClientRequestBase._updateLastUsedActionIdOfObjectPathId(lastUsedActionIdOfObjectPathId, referencedObjectPaths[refIndex], actionId);
                }
            }
        }
        var lastUsedObjectPathIdsOfAction = {};
        // lastUsedObjectPathIdsOfAction is the action's last used object paths.
        // After that action, we do not need those objects any more.
        for (var key in lastUsedActionIdOfObjectPathId) {
            var actionId = lastUsedActionIdOfObjectPathId[key];
            var objectPathIds = lastUsedObjectPathIdsOfAction[actionId];
            if (!objectPathIds) {
                objectPathIds = [];
                lastUsedObjectPathIdsOfAction[actionId] = objectPathIds;
            }
            objectPathIds.push(parseInt(key));
        }
        // Set the LastUsedObjectPathIds on the actionInfo
        for (var index = 0; index < actionsLength; index++) {
            var action = actions[index];
            var lastUsedObjectPathIds = lastUsedObjectPathIdsOfAction[action.actionInfo.Id];
            if (lastUsedObjectPathIds && lastUsedObjectPathIds.length > 0) {
                // action.actionInfo.LastUsedObjectPathIds
                action.actionInfo.L = lastUsedObjectPathIds;
            }
            else if (action.actionInfo.L) {
                // action.actionInfo.LastUsedObjectPathIds
                delete action.actionInfo.L;
            }
        }
    };
    return ClientRequestBase;
}());
exports.ClientRequestBase = ClientRequestBase;
var ClientResult = /** @class */ (function () {
    function ClientResult(type) {
        this.m_type = type;
    }
    Object.defineProperty(ClientResult.prototype, "value", {
        get: function () {
            if (!this.m_isLoaded) {
                throw new Core._Internal.RuntimeError({
                    code: Core.CoreErrorCodes.valueNotLoaded,
                    message: Core.CoreUtility._getResourceString(Core.CoreResourceStrings.valueNotLoaded),
                    debugInfo: {
                        errorLocation: 'clientResult.value'
                    }
                });
            }
            return this.m_value;
        },
        enumerable: true,
        configurable: true
    });
    ClientResult.prototype._handleResult = function (value) {
        this.m_isLoaded = true;
        if (typeof value === 'object' && value && value._IsNull) {
            return;
        }
        if (this.m_type === 1 /* date */) {
            this.m_value = CommonUtility.adjustToDateTime(value);
        }
        else {
            this.m_value = value;
        }
    };
    return ClientResult;
}());
exports.ClientResult = ClientResult;
var ServiceApiQueue = /** @class */ (function () {
    function ServiceApiQueue(m_context) {
        this.m_context = m_context;
        this.m_actions = [];
    }
    ServiceApiQueue.prototype.add = function (action, resultHandler, resolve, reject) {
        var _this = this;
        this.m_actions.push({ action: action, resultHandler: resultHandler, resolve: resolve, reject: reject });
        if (this.m_actions.length === 1) {
            setTimeout(function () { return _this.processActions(); }, 0);
        }
    };
    ServiceApiQueue.prototype.processActions = function () {
        var _this = this;
        if (this.m_actions.length === 0) {
            return;
        }
        var actions = this.m_actions;
        this.m_actions = [];
        var request = new ClientRequestBase(this.m_context);
        for (var i = 0; i < actions.length; i++) {
            var action = actions[i];
            request.ensureInstantiateObjectPath(action.action.referencedObjectPath);
            request.ensureInstantiateObjectPaths(action.action.referencedArgumentObjectPaths);
            request.addAction(action.action);
            request.addReferencedObjectPath(action.action.referencedObjectPath);
            request.addReferencedObjectPaths(action.action.referencedArgumentObjectPaths);
        }
        var _a = request.buildRequestMessageBodyAndRequestFlags(), body = _a.body, flags = _a.flags;
        var requestMessage = {
            Url: Core.CoreConstants.localDocumentApiPrefix,
            Headers: null,
            Body: body
        };
        var executor = new HttpRequestExecutor();
        executor
            .executeAsync(this.m_context._customData, flags, requestMessage)
            .then(function (response) {
            _this.processResponse(request, actions, response);
        })
            .catch(function (ex) {
            for (var i = 0; i < actions.length; i++) {
                var action = actions[i];
                action.reject(ex);
            }
        });
    };
    ServiceApiQueue.prototype.processResponse = function (request, actions, response) {
        var error = this.getErrorFromResponse(response);
        var actionResults = null;
        if (response.Body.Results) {
            actionResults = response.Body.Results;
        }
        else if (response.Body.ProcessedResults && response.Body.ProcessedResults.Results) {
            actionResults = response.Body.ProcessedResults.Results;
        }
        if (!actionResults) {
            actionResults = [];
        }
        this.processActionResults(request, actions, actionResults, error);
    };
    ServiceApiQueue.prototype.getErrorFromResponse = function (response) {
        if (!Core.CoreUtility.isNullOrEmptyString(response.ErrorCode)) {
            return new Core._Internal.RuntimeError({
                code: response.ErrorCode,
                message: response.ErrorMessage
            });
        }
        if (response.Body && response.Body.Error) {
            return new Core._Internal.RuntimeError({
                code: response.Body.Error.Code,
                message: response.Body.Error.Message
            });
        }
        return null;
    };
    ServiceApiQueue.prototype.processActionResults = function (request, actions, actionResults, err) {
        request.processResponse(actionResults);
        for (var i = 0; i < actions.length; i++) {
            var action = actions[i];
            var actionId = action.action.actionInfo.Id;
            var hasResult = false;
            for (var j = 0; j < actionResults.length; j++) {
                if (actionId == actionResults[j].ActionId) {
                    var resultValue = actionResults[j].Value;
                    if (action.resultHandler) {
                        action.resultHandler._handleResult(resultValue);
                        resultValue = action.resultHandler.value;
                    }
                    if (action.resolve) {
                        action.resolve(resultValue);
                    }
                    hasResult = true;
                    break;
                }
            }
            if (!hasResult && action.reject) {
                if (err) {
                    action.reject(err);
                }
                else {
                    action.reject('No response for the action.');
                }
            }
        }
    };
    return ServiceApiQueue;
}());
var HttpRequestExecutor = /** @class */ (function () {
    function HttpRequestExecutor() {
    }
    HttpRequestExecutor.prototype.executeAsync = function (customData, requestFlags, requestMessage) {
        var url = requestMessage.Url;
        if (url.charAt(url.length - 1) != '/') {
            url = url + '/';
        }
        url = url + Core.CoreConstants.processQuery;
        url = url + '?' + Core.CoreConstants.flags + '=' + requestFlags.toString();
        var requestInfo = {
            method: 'POST',
            url: url,
            headers: {},
            body: requestMessage.Body
        };
        requestInfo.headers[Core.CoreConstants.sourceLibHeader] = HttpRequestExecutor.SourceLibHeaderValue;
        requestInfo.headers['CONTENT-TYPE'] = 'application/json';
        if (requestMessage.Headers) {
            for (var key in requestMessage.Headers) {
                requestInfo.headers[key] = requestMessage.Headers[key];
            }
        }
        var sendRequestFunc = Core.CoreUtility._isLocalDocumentUrl(requestInfo.url)
            ? Core.HttpUtility.sendLocalDocumentRequest
            : Core.HttpUtility.sendRequest;
        return sendRequestFunc(requestInfo).then(function (responseInfo) {
            var response;
            if (responseInfo.statusCode === 200) {
                response = {
                    ErrorCode: null,
                    ErrorMessage: null,
                    Headers: responseInfo.headers,
                    Body: Core.CoreUtility._parseResponseBody(responseInfo)
                };
            }
            else {
                Core.CoreUtility.log('Error Response:' + responseInfo.body);
                var error = Core.CoreUtility._parseErrorResponse(responseInfo);
                response = {
                    ErrorCode: error.errorCode,
                    ErrorMessage: error.errorMessage,
                    Headers: responseInfo.headers,
                    Body: null
                };
            }
            return response;
        });
    };
    HttpRequestExecutor.SourceLibHeaderValue = 'officejs-rest';
    return HttpRequestExecutor;
}());
exports.HttpRequestExecutor = HttpRequestExecutor;
var CommonConstants = /** @class */ (function (_super) {
    __extends(CommonConstants, _super);
    function CommonConstants() {
        return _super !== null && _super.apply(this, arguments) || this;
    }
    CommonConstants.collectionPropertyPath = '_collectionPropertyPath';
    CommonConstants.id = 'Id';
    CommonConstants.idLowerCase = 'id';
    CommonConstants.idPrivate = '_Id';
    CommonConstants.keepReference = '_KeepReference';
    CommonConstants.objectPathIdPrivate = '_ObjectPathId';
    CommonConstants.referenceId = '_ReferenceId';
    return CommonConstants;
}(Core.CoreConstants));
exports.CommonConstants = CommonConstants;
var CommonUtility = /** @class */ (function (_super) {
    __extends(CommonUtility, _super);
    function CommonUtility() {
        return _super !== null && _super.apply(this, arguments) || this;
    }
    CommonUtility.adjustToDateTime = function (value) {
        if (Core.CoreUtility.isNullOrUndefined(value)) {
            return null;
        }
        if (typeof value === 'string') {
            // It's ISO8601 datetime. For example 2000-01-02T12:34:56Z
            return new Date(value);
        }
        if (Array.isArray(value)) {
            var arr = value;
            for (var i = 0; i < arr.length; i++) {
                arr[i] = CommonUtility.adjustToDateTime(arr[i]);
            }
            return arr;
        }
        throw Core.CoreUtility._createInvalidArgError({ argumentName: 'date' });
    };
    CommonUtility.tryGetObjectIdFromLoadOrRetrieveResult = function (value) {
        var id = value[CommonConstants.id];
        if (Core.CoreUtility.isNullOrUndefined(id)) {
            id = value[CommonConstants.idLowerCase];
        }
        if (Core.CoreUtility.isNullOrUndefined(id)) {
            id = value[CommonConstants.idPrivate];
        }
        return id;
    };
    CommonUtility.getObjectPathExpression = function (objectPath) {
        var ret = '';
        while (objectPath) {
            switch (objectPath.objectPathInfo.ObjectPathType) {
                case 1 /* GlobalObject */:
                    ret = ret;
                    break;
                case 2 /* NewObject */:
                    ret = 'new()' + (ret.length > 0 ? '.' : '') + ret;
                    break;
                case 3 /* Method */:
                    ret = CommonUtility.normalizeName(objectPath.objectPathInfo.Name) + '()' + (ret.length > 0 ? '.' : '') + ret;
                    break;
                case 4 /* Property */:
                    ret = CommonUtility.normalizeName(objectPath.objectPathInfo.Name) + (ret.length > 0 ? '.' : '') + ret;
                    break;
                case 5 /* Indexer */:
                    ret = 'getItem()' + (ret.length > 0 ? '.' : '') + ret;
                    break;
                case 6 /* ReferenceId */:
                    ret = '_reference()' + (ret.length > 0 ? '.' : '') + ret;
                    break;
            }
            objectPath = objectPath.parentObjectPath;
        }
        return ret;
    };
    CommonUtility.setMethodArguments = function (context, argumentInfo, args) {
        if (Core.CoreUtility.isNullOrUndefined(args)) {
            return null;
        }
        var referencedObjectPaths = new Array();
        var referencedObjectPathIds = new Array();
        var hasOne = CommonUtility.collectObjectPathInfos(context, args, referencedObjectPaths, referencedObjectPathIds);
        argumentInfo.Arguments = args;
        if (hasOne) {
            argumentInfo.ReferencedObjectPathIds = referencedObjectPathIds;
        }
        return referencedObjectPaths;
    };
    CommonUtility.validateContext = function (context, obj) {
        if (context && obj && obj._context !== context) {
            throw new Core._Internal.RuntimeError({
                code: Core.CoreErrorCodes.invalidRequestContext,
                message: Core.CoreUtility._getResourceString(Core.CoreResourceStrings.invalidRequestContext)
            });
        }
    };
    CommonUtility.isSetSupported = function (apiSetName, apiSetVersion) {
        if (typeof window !== 'undefined' &&
            window.Office &&
            window.Office.context &&
            window.Office.context.requirements) {
            return window.Office.context.requirements.isSetSupported(apiSetName, apiSetVersion);
        }
        return true;
    };
    CommonUtility.throwIfApiNotSupported = function (apiFullName, apiSetName, apiSetVersion, hostName) {
        if (!CommonUtility._doApiNotSupportedCheck) {
            return;
        }
        if (!CommonUtility.isSetSupported(apiSetName, apiSetVersion)) {
            var message = Core.CoreUtility._getResourceString(Core.CoreResourceStrings.apiNotFoundDetails, [
                apiFullName,
                apiSetName + ' ' + apiSetVersion,
                hostName
            ]);
            throw new Core._Internal.RuntimeError({
                code: Core.CoreErrorCodes.apiNotFound,
                message: message,
                debugInfo: { errorLocation: apiFullName }
            });
        }
    };
    CommonUtility.collectObjectPathInfos = function (context, args, referencedObjectPaths, referencedObjectPathIds) {
        var hasOne = false;
        for (var i = 0; i < args.length; i++) {
            if (args[i] instanceof ClientObjectBase) {
                var clientObject = args[i];
                CommonUtility.validateContext(context, clientObject);
                args[i] = clientObject._objectPath.objectPathInfo.Id;
                referencedObjectPathIds.push(clientObject._objectPath.objectPathInfo.Id);
                referencedObjectPaths.push(clientObject._objectPath);
                hasOne = true;
            }
            else if (Array.isArray(args[i])) {
                var childArrayObjectPathIds = new Array();
                var childArrayHasOne = CommonUtility.collectObjectPathInfos(context, args[i], referencedObjectPaths, childArrayObjectPathIds);
                if (childArrayHasOne) {
                    referencedObjectPathIds.push(childArrayObjectPathIds);
                    hasOne = true;
                }
                else {
                    referencedObjectPathIds.push(0);
                }
            }
            else if (Core.CoreUtility.isPlainJsonObject(args[i])) {
                referencedObjectPathIds.push(0);
                CommonUtility.replaceClientObjectPropertiesWithObjectPathIds(args[i], referencedObjectPaths);
            }
            else {
                referencedObjectPathIds.push(0);
            }
        }
        return hasOne;
    };
    CommonUtility.replaceClientObjectPropertiesWithObjectPathIds = function (value, referencedObjectPaths) {
        for (var key in value) {
            var propValue = value[key];
            if (propValue instanceof ClientObjectBase) {
                referencedObjectPaths.push(propValue._objectPath);
                value[key] = (_a = {}, _a[CommonConstants.objectPathIdPrivate] = propValue._objectPath.objectPathInfo.Id, _a);
            }
            else if (Array.isArray(propValue)) {
                for (var i = 0; i < propValue.length; i++) {
                    if (propValue[i] instanceof ClientObjectBase) {
                        var elem = propValue[i];
                        referencedObjectPaths.push(elem._objectPath);
                        propValue[i] = (_b = {}, _b[CommonConstants.objectPathIdPrivate] = elem._objectPath.objectPathInfo.Id, _b);
                    }
                    else if (Core.CoreUtility.isPlainJsonObject(propValue[i])) {
                        CommonUtility.replaceClientObjectPropertiesWithObjectPathIds(propValue[i], referencedObjectPaths);
                    }
                }
            }
            else if (Core.CoreUtility.isPlainJsonObject(propValue)) {
                CommonUtility.replaceClientObjectPropertiesWithObjectPathIds(propValue, referencedObjectPaths);
            }
            else {
                // Do nothing; leave the value of the key as is.
            }
        }
        var _a, _b;
    };
    CommonUtility.normalizeName = function (name) {
        return name.substr(0, 1).toLowerCase() + name.substr(1);
    };
    CommonUtility._doApiNotSupportedCheck = false;
    return CommonUtility;
}(Core.CoreUtility));
exports.CommonUtility = CommonUtility;


/***/ }),
/* 2 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

Object.defineProperty(exports, "__esModule", { value: true });
var core_1 = __webpack_require__(0);
exports.CoreUtility = core_1.CoreUtility;
exports.Error = core_1.Error;
exports.HttpUtility = core_1.HttpUtility;
exports.SessionBase = core_1.SessionBase;
var common_1 = __webpack_require__(1);
exports.CommonUtility = common_1.CommonUtility;
exports.ClientResult = common_1.ClientResult;
var batch_runtime_1 = __webpack_require__(5);
exports.ClientRequestContext = batch_runtime_1.ClientRequestContext;
exports.ClientObject = batch_runtime_1.ClientObject;
exports.config = batch_runtime_1.config;
exports.Constants = batch_runtime_1.Constants;
exports.ErrorCodes = batch_runtime_1.ErrorCodes;
exports.EventHandlers = batch_runtime_1.EventHandlers;
exports.GenericEventHandlers = batch_runtime_1.GenericEventHandlers;
exports.ResourceStrings = batch_runtime_1.ResourceStrings;
exports.Utility = batch_runtime_1.Utility;
exports._internalConfig = batch_runtime_1._internalConfig;
var BatchApiHelper = /** @class */ (function () {
    function BatchApiHelper() {
    }
    BatchApiHelper.invokeMethod = function (obj, methodName, operationType, args, flags, resultProcessType) {
        var action = batch_runtime_1.ActionFactory.createMethodAction(obj.context, obj, methodName, operationType, args, flags);
        var result = new common_1.ClientResult(resultProcessType);
        batch_runtime_1.Utility._addActionResultHandler(obj, action, result);
        return result;
    };
    BatchApiHelper.invokeEnsureUnchanged = function (obj, objectState) {
        batch_runtime_1.ActionFactory.createEnsureUnchangedAction(obj.context, obj, objectState);
    };
    BatchApiHelper.invokeSetProperty = function (obj, propName, propValue, flags) {
        batch_runtime_1.ActionFactory.createSetPropertyAction(obj.context, obj, propName, propValue, flags);
    };
    BatchApiHelper.createRootServiceObject = function (type, context) {
        var objectPath = batch_runtime_1.ObjectPathFactory.createGlobalObjectObjectPath(context);
        return new type(context, objectPath);
    };
    BatchApiHelper.createObjectFromReferenceId = function (type, context, referenceId) {
        var objectPath = batch_runtime_1.ObjectPathFactory.createReferenceIdObjectPath(context, referenceId);
        return new type(context, objectPath);
    };
    BatchApiHelper.createTopLevelServiceObject = function (type, context, typeName, isCollection, flags) {
        var objectPath = batch_runtime_1.ObjectPathFactory.createNewObjectObjectPath(context, typeName, isCollection, flags);
        return new type(context, objectPath);
    };
    BatchApiHelper.createPropertyObject = function (type, parent, propertyName, isCollection, flags) {
        var objectPath = batch_runtime_1.ObjectPathFactory.createPropertyObjectPath(parent.context, parent, propertyName, isCollection, false, flags);
        return new type(parent.context, objectPath);
    };
    BatchApiHelper.createIndexerObject = function (type, parent, args) {
        var objectPath = batch_runtime_1.ObjectPathFactory.createIndexerObjectPath(parent.context, parent, args);
        return new type(parent.context, objectPath);
    };
    BatchApiHelper.createMethodObject = function (type, parent, methodName, operationType, args, isCollection, isInvalidAfterRequest, getByIdMethodName, flags) {
        var objectPath = batch_runtime_1.ObjectPathFactory.createMethodObjectPath(parent.context, parent, methodName, operationType, args, isCollection, isInvalidAfterRequest, getByIdMethodName, flags);
        return new type(parent.context, objectPath);
    };
    BatchApiHelper.createChildItemObject = function (type, hasIndexerMethod, parent, chileItem, index) {
        var objectPath = batch_runtime_1.ObjectPathFactory.createChildItemObjectPathUsingIndexerOrGetItemAt(hasIndexerMethod, parent.context, parent, chileItem, index);
        return new type(parent.context, objectPath);
    };
    return BatchApiHelper;
}());
exports.BatchApiHelper = BatchApiHelper;


/***/ }),
/* 3 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

Object.defineProperty(exports, "__esModule", { value: true });
var customfunctions_runtime_1 = __webpack_require__(4);
var OfficeExtensionBatch = __webpack_require__(2);
__webpack_require__(6);
// Expose global variables
// CustomFunctionMappings
window.CustomFunctionMappings = CustomFunctionMappings;
// OfficeExtensionBatch
window.OfficeExtensionBatch = OfficeExtensionBatch;
// Promise
if (typeof (Promise) === 'undefined') {
    window.Promise = Office.Promise;
}
// Now initialize the custom functions
Office.onReady(function () {
    function initializeCustomFunctionsOrDelay() {
        if (CustomFunctionMappings && CustomFunctionMappings['__delay__']) {
            setTimeout(initializeCustomFunctionsOrDelay, 50);
        }
        else {
            customfunctions_runtime_1.CustomFunctions.initialize();
        }
    }
    initializeCustomFunctionsOrDelay();
});


/***/ }),
/* 4 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

var __extends = (this && this.__extends) || (function () {
    var extendStatics = Object.setPrototypeOf ||
        ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
        function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p]; };
    return function (d, b) {
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
Object.defineProperty(exports, "__esModule", { value: true });
var OfficeExtension = __webpack_require__(2);
var Core = __webpack_require__(0);
/* End_PlaceHolder_ModuleHeader */
var _hostName = 'Excel';
var _defaultApiSetName = 'ExcelApi';
var _createPropertyObject = OfficeExtension.BatchApiHelper.createPropertyObject;
var _createMethodObject = OfficeExtension.BatchApiHelper.createMethodObject;
var _createIndexerObject = OfficeExtension.BatchApiHelper.createIndexerObject;
var _createRootServiceObject = OfficeExtension.BatchApiHelper.createRootServiceObject;
var _createTopLevelServiceObject = OfficeExtension.BatchApiHelper.createTopLevelServiceObject;
var _createChildItemObject = OfficeExtension.BatchApiHelper.createChildItemObject;
var _invokeMethod = OfficeExtension.BatchApiHelper.invokeMethod;
var _invokeEnsureUnchanged = OfficeExtension.BatchApiHelper.invokeEnsureUnchanged;
var _invokeSetProperty = OfficeExtension.BatchApiHelper.invokeSetProperty;
var _isNullOrUndefined = OfficeExtension.Utility.isNullOrUndefined;
var _isUndefined = OfficeExtension.Utility.isUndefined;
var _throwIfNotLoaded = OfficeExtension.Utility.throwIfNotLoaded;
var _throwIfApiNotSupported = OfficeExtension.Utility.throwIfApiNotSupported;
var _load = OfficeExtension.Utility.load;
var _retrieve = OfficeExtension.Utility.retrieve;
var _toJson = OfficeExtension.Utility.toJson;
var _fixObjectPathIfNecessary = OfficeExtension.Utility.fixObjectPathIfNecessary;
var _handleNavigationPropertyResults = OfficeExtension.Utility._handleNavigationPropertyResults;
var _adjustToDateTime = OfficeExtension.Utility.adjustToDateTime;
var _processRetrieveResult = OfficeExtension.Utility.processRetrieveResult;
var _typeCustomFunctions = 'CustomFunctions';
/* Begin_PlaceHolder_CustomFunctions_BeforeDeclaration */
var CustomFunctionRequestContext = /** @class */ (function (_super) {
    __extends(CustomFunctionRequestContext, _super);
    function CustomFunctionRequestContext(requestInfo) {
        var _this = _super.call(this, requestInfo) || this;
        _this.m_customFunctions = CustomFunctions.newObject(_this);
        _this.m_container = _createRootServiceObject(CustomFunctionsContainer, _this);
        _this._rootObject = _this.m_container;
        _this._rootObjectPropertyName = 'customFunctionsContainer';
        _this._requestFlagModifier = 128 /* UndoPreviewEnabled */;
        return _this;
    }
    Object.defineProperty(CustomFunctionRequestContext.prototype, "customFunctions", {
        get: function () {
            return this.m_customFunctions;
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(CustomFunctionRequestContext.prototype, "customFunctionsContainer", {
        get: function () {
            return this.m_container;
        },
        enumerable: true,
        configurable: true
    });
    CustomFunctionRequestContext.prototype._processOfficeJsErrorResponse = function (officeJsErrorCode, response) {
        var ooeInvalidApiCallInContext = 5004;
        if (officeJsErrorCode === ooeInvalidApiCallInContext) {
            response.ErrorCode = CustomFunctionErrorCode.invalidOperationInCellEditMode;
            response.ErrorMessage = OfficeExtension.Utility._getResourceString(OfficeExtension.ResourceStrings.invalidOperationInCellEditMode);
        }
    };
    return CustomFunctionRequestContext;
}(OfficeExtension.ClientRequestContext));
// Keep in sync with MessageCategory.CustomFunction in %xlshared%\src\api\metadata\current\ExcelApi.cs
var CustomFunctionMessageCategory = 1;
exports.Script = {
    _CustomFunctionMetadata: {}
};
var InvocationContext = /** @class */ (function () {
    function InvocationContext(setResultHandler) {
        this.setResult = setResultHandler;
    }
    Object.defineProperty(InvocationContext.prototype, "onCanceled", {
        get: function () {
            if (!_isNullOrUndefined(this._onCanceled) && typeof this._onCanceled === 'function') {
                return this._onCanceled;
            }
            return undefined;
        },
        set: function (handler) {
            this._onCanceled = handler;
        },
        enumerable: true,
        configurable: true
    });
    return InvocationContext;
}());
exports.InvocationContext = InvocationContext;
/* CustomFunctions Logging classes */
var CustomFunctionLoggingSeverity;
(function (CustomFunctionLoggingSeverity) {
    CustomFunctionLoggingSeverity["Info"] = "Medium";
    CustomFunctionLoggingSeverity["Error"] = "Unexpected";
    CustomFunctionLoggingSeverity["Verbose"] = "Verbose";
})(CustomFunctionLoggingSeverity || (CustomFunctionLoggingSeverity = {}));
var CustomFunctionLog = /** @class */ (function () {
    function CustomFunctionLog(Severity, Message) {
        this.Severity = Severity;
        this.Message = Message;
    }
    return CustomFunctionLog;
}());
var CustomFunctionsLogger = /** @class */ (function () {
    function CustomFunctionsLogger() {
    }
    CustomFunctionsLogger.logEvent = function (log, data) {
        var logMessage = log.Severity + ' ' + log.Message + data;
        OfficeExtension.Utility.log(logMessage);
        if (CustomFunctionsLogger.s_shouldLog) {
            switch (log.Severity) {
                case CustomFunctionLoggingSeverity.Verbose:
                    if (console.log !== null) {
                        console.log(logMessage);
                    }
                    break;
                case CustomFunctionLoggingSeverity.Info:
                    if (console.info !== null) {
                        console.info(logMessage);
                    }
                    break;
                case CustomFunctionLoggingSeverity.Error:
                    if (console.error !== null) {
                        console.error(logMessage);
                    }
                    break;
                default:
                    break;
            }
        }
    };
    CustomFunctionsLogger.shouldLog = function () {
        // Retrieve the logging toggle from the iframe's name.
        try {
            // JSON.parse() may throw exception when window.name[CustomFunctionsLogger.CustomFunctionLoggingFlag] is not valid.
            return (!_isNullOrUndefined(console) &&
                !_isNullOrUndefined(window) &&
                window.name &&
                typeof window.name === 'object' &&
                JSON.parse(window.name)[CustomFunctionsLogger.CustomFunctionLoggingFlag]);
        }
        catch (ex) {
            OfficeExtension.Utility.log(JSON.stringify(ex));
            return false;
        }
    };
    CustomFunctionsLogger.CustomFunctionLoggingFlag = 'CustomFunctionsRuntimeLogging';
    CustomFunctionsLogger.s_shouldLog = CustomFunctionsLogger.shouldLog();
    return CustomFunctionsLogger;
}());
var CustomFunctionProxy = /** @class */ (function () {
    function CustomFunctionProxy() {
        this._whenInit = undefined;
        this._isInit = false;
        this._setResultsDelayMillis = 50;
        this._setResultsLifeMillis = 60 * 1000;
        this._ensureInitRetryDelayMillis = 500;
        this._resultEntryBuffer = [];
        this._isSetResultsTaskScheduled = false;
        this._batchQuotaMillis = 1000;
        this._invocationContextMap = {};
    }
    CustomFunctionProxy.splitName = function (name) {
        // Validate
        var matches = name.match(/[a-z_][a-z_0-9\.]+/gi);
        if (matches === null || matches.length !== 1 || matches[0] !== name) {
            throw OfficeExtension.Utility.createRuntimeError(CustomFunctionErrorCode.invalidOperation, OfficeExtension.Utility._getResourceString(OfficeExtension.ResourceStrings.customFunctionNameContainsBadChars), 'CustomFunctionProxy.splitName');
        }
        // Split
        var splitIndex = name.lastIndexOf('.');
        if (splitIndex < 1 || splitIndex === name.length - 1) {
            // Can't be the first or the last char either.
            throw OfficeExtension.Utility.createRuntimeError(CustomFunctionErrorCode.invalidOperation, OfficeExtension.Utility._getResourceString(OfficeExtension.ResourceStrings.customFunctionNameCannotSplit), 'CustomFunctionProxy.splitName');
        }
        var nameSplit = {
            namespace: name.substring(0, splitIndex),
            name: name.substr(splitIndex + 1)
        };
        return nameSplit;
    };
    CustomFunctionProxy.prototype._initSettings = function () {
        // If internal settings are provided, override the defaults.
        if (typeof exports.Script === 'object' && typeof exports.Script._CustomFunctionSettings === 'object') {
            // setResultsDelayMillis
            if (typeof exports.Script._CustomFunctionSettings.setResultsDelayMillis === 'number') {
                var setResultsDelayMillis = exports.Script._CustomFunctionSettings.setResultsDelayMillis;
                // Make sure we don't end up with a value that is smaller than 0 or bigger than 1000.
                setResultsDelayMillis = Math.max(0, setResultsDelayMillis);
                setResultsDelayMillis = Math.min(1000, setResultsDelayMillis);
                this._setResultsDelayMillis = setResultsDelayMillis;
            }
            if (typeof exports.Script._CustomFunctionSettings.ensureInitRetryDelayMillis === 'number') {
                var ensureInitRetryDelayMillis = exports.Script._CustomFunctionSettings.ensureInitRetryDelayMillis;
                // Make sure we don't end up with a value that is smaller than 0 or bigger than 2 seconds (2 * 1000).
                ensureInitRetryDelayMillis = Math.max(0, ensureInitRetryDelayMillis);
                ensureInitRetryDelayMillis = Math.min(2000, ensureInitRetryDelayMillis);
                this._ensureInitRetryDelayMillis = ensureInitRetryDelayMillis;
            }
            // setResultsLifeMillis
            if (typeof exports.Script._CustomFunctionSettings.setResultsLifeMillis === 'number') {
                var setResultsLifeMillis = exports.Script._CustomFunctionSettings.setResultsLifeMillis;
                // Make sure we don't end up with a value that is smaller than 0 or bigger than 10 * 60 * 1000.
                setResultsLifeMillis = Math.max(0, setResultsLifeMillis);
                setResultsLifeMillis = Math.min(10 * 60 * 1000, setResultsLifeMillis);
                this._setResultsLifeMillis = setResultsLifeMillis;
            }
            // batchQuotaMillis
            if (typeof exports.Script._CustomFunctionSettings.batchQuotaMillis === 'number') {
                var batchQuotaMillis = exports.Script._CustomFunctionSettings.batchQuotaMillis;
                // Make sure we don't end up with a value that is smaller than 0 or bigger than 1000.
                batchQuotaMillis = Math.max(0, batchQuotaMillis);
                batchQuotaMillis = Math.min(1000, batchQuotaMillis);
                this._batchQuotaMillis = batchQuotaMillis;
            }
        }
    };
    CustomFunctionProxy.prototype.ensureInit = function (context) {
        var _this = this;
        this._initSettings();
        // We must hold off context.sync() until the event handler gets registered.
        // Otherwise, server events may get fired but there may not be handler in JavaScirpt,
        // which would lead to the first invocation getting lost.
        if (this._whenInit === undefined) {
            this._whenInit = OfficeExtension.Utility._createPromiseFromResult(null)
                .then(function () {
                if (!_this._isInit) {
                    return context.eventRegistration.register(5 /* RichApiMessageEvent */, '', _this._handleMessage.bind(_this));
                }
            })
                .then(function () {
                _this._isInit = true;
            });
        }
        if (!this._isInit) {
            context._pendingRequest._addPreSyncPromise(this._whenInit);
        }
        return this._whenInit;
    };
    CustomFunctionProxy.prototype._initFromHostBridge = function (hostBridge) {
        var _this = this;
        this._initSettings();
        hostBridge.addHostMessageHandler(function (bridgeMessage) {
            if (bridgeMessage.type === 3 /* genericMessage */) {
                _this._handleMessage(bridgeMessage.message);
            }
        });
        this._isInit = true;
        this._whenInit = OfficeExtension.CoreUtility.Promise.resolve();
    };
    CustomFunctionProxy.prototype._handleMessage = function (args) {
        try {
            OfficeExtension.Utility.log('CustomFunctionProxy._handleMessage');
            OfficeExtension.Utility.checkArgumentNull(args, 'args');
            // Invocation messages and cancellation messages come in with one array. We need to split.
            var entryArray = args.entries;
            var invocationArray = [];
            var cancellationArray = [];
            var metadataArray = [];
            for (var i = 0; i < entryArray.length; i++) {
                if (entryArray[i].messageCategory !== CustomFunctionMessageCategory) {
                    continue;
                }
                if (entryArray[i].messageType === 1000 /* invocationMessage */) {
                    invocationArray.push(entryArray[i]);
                }
                else if (entryArray[i].messageType === 1001 /* cancellationMessage */) {
                    cancellationArray.push(entryArray[i]);
                }
                else if (entryArray[i].messageType === 1002 /* metadataMessage */) {
                    metadataArray.push(entryArray[i]);
                }
                else {
                    throw OfficeExtension.Utility.createRuntimeError(CustomFunctionErrorCode.generalException, 'unexpected message type', 'CustomFunctionProxy._handleMessage');
                }
            }
            // Handle metadata entries before invocation entries!
            if (metadataArray.length > 0) {
                this._handleMetadataEntries(metadataArray);
            }
            if (invocationArray.length > 0) {
                var batchArray = this._batchInvocationEntries(invocationArray);
                this._invokeRemainingBatchEntries(batchArray, 0 /*startIndex*/);
            }
            if (cancellationArray.length > 0) {
                this._handleCancellationEntries(cancellationArray);
            }
        }
        catch (ex) {
            CustomFunctionProxy._tryLog(ex);
            throw ex;
        }
        // The infra doesn't do anything with the returned promise.
        return OfficeExtension.Utility._createPromiseFromResult(null);
    };
    CustomFunctionProxy._tryLog = function (ex) {
        try {
            if (ex.toString) {
                OfficeExtension.Utility.log(ex.toString());
            }
            OfficeExtension.Utility.log(JSON.stringify(ex));
        }
        catch (otherEx) {
            OfficeExtension.Utility.log('Error while logging ex');
        }
    };
    CustomFunctionProxy.prototype._handleMetadataEntries = function (entryArray) {
        for (var i = 0; i < entryArray.length; i++) {
            var messageJson = entryArray[i].message;
            if (OfficeExtension.Utility.isNullOrEmptyString(messageJson)) {
                throw OfficeExtension.Utility.createRuntimeError(CustomFunctionErrorCode.generalException, 'messageJson', 'CustomFunctionProxy._handleMetadataEntries');
            }
            var message = JSON.parse(messageJson);
            if (_isNullOrUndefined(message)) {
                throw OfficeExtension.Utility.createRuntimeError(CustomFunctionErrorCode.generalException, 'message', 'CustomFunctionProxy._handleMetadataEntries');
            }
            exports.Script._CustomFunctionMetadata[message.functionName] = {
                options: {
                    stream: message.isStream,
                    cancelable: message.isCancelable
                }
            };
        } // for (entry)
    };
    CustomFunctionProxy.prototype._handleCancellationEntries = function (entryArray) {
        for (var i = 0; i < entryArray.length; i++) {
            var messageJson = entryArray[i].message;
            if (OfficeExtension.Utility.isNullOrEmptyString(messageJson)) {
                throw OfficeExtension.Utility.createRuntimeError(CustomFunctionErrorCode.generalException, 'messageJson', 'CustomFunctionProxy._handleCancellationEntries');
            }
            var message = JSON.parse(messageJson);
            if (_isNullOrUndefined(message)) {
                throw OfficeExtension.Utility.createRuntimeError(CustomFunctionErrorCode.generalException, 'message', 'CustomFunctionProxy._handleCancellationEntries');
            }
            var invocationId = message.invocationId;
            var invocationContext = this._invocationContextMap[invocationId];
            if (!_isNullOrUndefined(invocationContext)) {
                if (_isNullOrUndefined(invocationContext.onCanceled)) {
                    throw OfficeExtension.Utility.createRuntimeError(CustomFunctionErrorCode.invalidOperation, OfficeExtension.Utility._getResourceString(OfficeExtension.ResourceStrings.customFunctionCancellationHandlerMissing), 'CustomFunctionProxy._handleCancellationEntries');
                }
                invocationContext.onCanceled();
            }
        } // for (entry)
    };
    CustomFunctionProxy.prototype._batchInvocationEntries = function (entryArray) {
        var _this = this;
        var batchArray = [];
        var _loop_1 = function (i) {
            var messageJson = entryArray[i].message;
            if (OfficeExtension.Utility.isNullOrEmptyString(messageJson)) {
                throw OfficeExtension.Utility.createRuntimeError(CustomFunctionErrorCode.generalException, 'messageJson', 'CustomFunctionProxy._batchInvocationEntries');
            }
            var message = JSON.parse(messageJson);
            if (_isNullOrUndefined(message)) {
                throw OfficeExtension.Utility.createRuntimeError(CustomFunctionErrorCode.generalException, 'message', 'CustomFunctionProxy._batchInvocationEntries');
            }
            if (_isNullOrUndefined(message.invocationId) || message.invocationId < 0) {
                throw OfficeExtension.Utility.createRuntimeError(CustomFunctionErrorCode.generalException, 'invocationId', 'CustomFunctionProxy._batchInvocationEntries');
            }
            if (_isNullOrUndefined(message.functionName)) {
                throw OfficeExtension.Utility.createRuntimeError(CustomFunctionErrorCode.generalException, 'functionName', 'CustomFunctionProxy._batchInvocationEntries');
            }
            var batchIndex = -1;
            var call = null;
            var isCancelable = false;
            var isStreaming = false;
            var isBatching = false;
            var metadata = exports.Script._CustomFunctionMetadata[message.functionName];
            if (!_isNullOrUndefined(metadata)) {
                // This branch handles static (new-style) registration.
                call = this_1._getCustomFunctionMappings(message.functionName);
                if (_isNullOrUndefined(call)) {
                    if (_isNullOrUndefined(window)) {
                        throw OfficeExtension.Utility.createRuntimeError(CustomFunctionErrorCode.invalidOperation, OfficeExtension.Utility._getResourceString(OfficeExtension.ResourceStrings.customFunctionWindowMissing), 'CustomFunctionProxy._batchInvocationEntries');
                    }
                    // Verify a nest of parent objects exists.
                    var functionParent = window;
                    var beginOfSegmentIndex = 0;
                    var endOfSegmentIndex = message.functionName.indexOf('.', beginOfSegmentIndex);
                    while (endOfSegmentIndex > beginOfSegmentIndex) {
                        var functionNameSegment = message.functionName.substring(beginOfSegmentIndex, endOfSegmentIndex);
                        if (_isNullOrUndefined(functionParent[functionNameSegment]) ||
                            typeof functionParent[functionNameSegment] !== 'object') {
                            throw OfficeExtension.Utility.createRuntimeError(CustomFunctionErrorCode.invalidOperation, OfficeExtension.Utility._getResourceString(OfficeExtension.ResourceStrings.customFunctionDefintionMissingOnWindow, message.functionName), 'CustomFunctionProxy._batchInvocationEntries');
                        }
                        functionParent = functionParent[functionNameSegment];
                        beginOfSegmentIndex = endOfSegmentIndex + 1; // +1 to skip the '.'
                        endOfSegmentIndex = message.functionName.indexOf('.', beginOfSegmentIndex);
                    }
                    // Verify the function exists.
                    var functionName = message.functionName.substring(beginOfSegmentIndex);
                    if (_isNullOrUndefined(functionParent[functionName])) {
                        throw OfficeExtension.Utility.createRuntimeError(CustomFunctionErrorCode.invalidOperation, OfficeExtension.Utility._getResourceString(OfficeExtension.ResourceStrings.customFunctionDefintionMissingOnWindow, message.functionName), 'CustomFunctionProxy._batchInvocationEntries');
                    }
                    if (typeof functionParent[functionName] != 'function') {
                        throw OfficeExtension.Utility.createRuntimeError(CustomFunctionErrorCode.invalidOperation, OfficeExtension.Utility._getResourceString(OfficeExtension.ResourceStrings.customFunctionInvalidFunction, message.functionName), 'CustomFunctionProxy._batchInvocationEntries');
                    }
                    call = functionParent[functionName];
                }
                isCancelable = metadata.options.cancelable;
                isStreaming = metadata.options.stream;
            }
            // Streaming or Cancelable functions expect an extra parameter called invocationConext.
            // - Streaming functions: invocationContext should contain a member setResult - a callback to set the result. This is defined by us.
            // - Cancelable functions: developers should define their cancellation handler and register it by invocationContext.onCancel()
            //   within their function.
            if (isStreaming || isCancelable) {
                var setResult = undefined;
                if (isStreaming) {
                    setResult = function (result) {
                        _this._setResult(message.invocationId, result);
                    };
                }
                var invocationContext = void 0;
                invocationContext = new InvocationContext(setResult);
                this_1._invocationContextMap[message.invocationId] = invocationContext;
                message.parameterValues.push(invocationContext);
            }
            if (batchIndex >= 0) {
                batchArray[batchIndex].invocationIds.push(message.invocationId);
                batchArray[batchIndex].parameterValueSets.push(message.parameterValues);
            }
            else {
                batchArray.push({
                    call: call,
                    isBatching: isBatching,
                    isStreaming: isStreaming,
                    invocationIds: [message.invocationId],
                    parameterValueSets: [message.parameterValues],
                    functionName: message.functionName
                });
            }
        };
        var this_1 = this;
        for (var i = 0; i < entryArray.length; i++) {
            _loop_1(i);
        } // for (entry)
        return batchArray;
    };
    CustomFunctionProxy.prototype._getCustomFunctionMappings = function (functionName) {
        // Check if CustomFunctionMappings object exists
        if (typeof CustomFunctionMappings === 'object') {
            if (!_isNullOrUndefined(CustomFunctionMappings[functionName])) {
                if (typeof CustomFunctionMappings[functionName] === 'function') {
                    return CustomFunctionMappings[functionName];
                }
                else {
                    throw OfficeExtension.Utility.createRuntimeError(CustomFunctionErrorCode.invalidOperation, OfficeExtension.Utility._getResourceString(OfficeExtension.ResourceStrings.customFunctionInvalidFunctionMapping, functionName), 'CustomFunctionProxy._getCustomFunctionMappings');
                }
            }
        }
        return undefined;
    };
    CustomFunctionProxy.prototype._invokeRemainingBatchEntries = function (batchArray, startIndex) {
        OfficeExtension.Utility.log('CustomFunctionProxy._invokeRemainingBatchEntries');
        var startTimeMillis = Date.now();
        for (var i = startIndex; i < batchArray.length; i++) {
            var currentTimeMillis = Date.now();
            var batchDurationMillis = currentTimeMillis - startTimeMillis;
            if (batchDurationMillis < this._batchQuotaMillis) {
                this._invokeFunctionAndSetResult(batchArray[i]);
            }
            else {
                OfficeExtension.Utility.log('setTimeout(CustomFunctionProxy._invokeRemainingBatchEntries)');
                setTimeout(this._invokeRemainingBatchEntries.bind(this), 0 /*timeout*/, batchArray, i /*startIndex*/);
                break;
            }
        }
    };
    CustomFunctionProxy.prototype._invokeFunctionAndSetResult = function (batch) {
        var _this = this;
        // Invoke.
        var results;
        CustomFunctionsLogger.logEvent(CustomFunctionProxy.CustomFunctionExecutionStartLog, batch.functionName);
        try {
            if (batch.isBatching) {
                // Batching functions: pass an array of parameter value sets, receive an array of results.
                // Use Array.call() to preserve the array of sets, because Array.apply() takes only the first set.
                results = batch.call.call(null, batch.parameterValueSets);
            }
            else {
                // Non-batching functions: pass 1 parameter value set, wrap the received result into an array.
                results = [batch.call.apply(null, batch.parameterValueSets[0])];
            }
        }
        catch (ex) {
            for (var i = 0; i < batch.invocationIds.length; i++) {
                // The function threw an exception.
                this._setError(batch.invocationIds[i], ex);
            }
            CustomFunctionsLogger.logEvent(CustomFunctionProxy.CustomFunctionExecutionFailureLog, batch.functionName);
            return;
        }
        // Set results/errors.
        if (batch.isStreaming) {
            // Do nothing for streaming functions - they will use the callback we passed in.
        }
        else {
            if (results.length === batch.parameterValueSets.length) {
                var _loop_2 = function (i) {
                    if (!_isNullOrUndefined(results[i]) &&
                        typeof results[i] === 'object' &&
                        typeof results[i].then === 'function') {
                        // The function returned a promise.
                        results[i].then(function (value) {
                            CustomFunctionsLogger.logEvent(CustomFunctionProxy.CustomFunctionExecutionFinishLog, batch.functionName);
                            _this._setResult(batch.invocationIds[i], value);
                        }, function (reason) {
                            CustomFunctionsLogger.logEvent(CustomFunctionProxy.CustomFunctionExecutionFailureLog, batch.functionName);
                            _this._setError(batch.invocationIds[i], reason);
                        });
                    }
                    else {
                        // The function returned a value.
                        CustomFunctionsLogger.logEvent(CustomFunctionProxy.CustomFunctionExecutionFinishLog, batch.functionName);
                        this_2._setResult(batch.invocationIds[i], results[i]);
                    }
                };
                var this_2 = this;
                for (var i = 0; i < results.length; i++) {
                    _loop_2(i);
                } // for (result)
            }
            else {
                // The function has screwed up something. We cannot trust the results it has given us.
                // We error out the entire batch.
                CustomFunctionsLogger.logEvent(CustomFunctionProxy.CustomFunctionExecutionFailureLog, batch.functionName);
                for (var i = 0; i < batch.invocationIds.length; i++) {
                    this._setError(batch.invocationIds[i], OfficeExtension.Utility._getResourceString(OfficeExtension.ResourceStrings.customFunctionUnexpectedNumberOfEntriesInResultBatch));
                } // for (invocationId)
            }
        }
    };
    CustomFunctionProxy.prototype._setResult = function (invocationId, result) {
        var invocationResult = { id: invocationId, value: result };
        if (typeof result === 'number' && isNaN(result)) {
            invocationResult.failed = true;
            invocationResult.value = 'NaN';
        }
        this._resultEntryBuffer.push({
            timeCreated: Date.now(),
            result: invocationResult
        });
        this._ensureSetResultsTaskIsScheduled();
    };
    CustomFunctionProxy.prototype._setError = function (invocationId, error) {
        var message;
        if (typeof error === 'object') {
            message = JSON.stringify(error);
        }
        else {
            message = error.toString();
        }
        this._resultEntryBuffer.push({
            timeCreated: Date.now(),
            result: { id: invocationId, failed: true, value: message }
        });
        this._ensureSetResultsTaskIsScheduled();
    };
    CustomFunctionProxy.prototype._ensureSetResultsTaskIsScheduled = function () {
        if (!this._isSetResultsTaskScheduled && this._resultEntryBuffer.length > 0) {
            OfficeExtension.Utility.log('setTimeout(CustomFunctionProxy._executeSetResultsTask)');
            setTimeout(this._executeSetResultsTask.bind(this), this._setResultsDelayMillis);
            this._isSetResultsTaskScheduled = true;
        }
    };
    CustomFunctionProxy.prototype._executeSetResultsTask = function () {
        var _this = this;
        OfficeExtension.Utility.log('CustomFunctionProxy._executeSetResultsTask');
        // Clear this flag first, so that if something happens, we don't block further progress.
        this._isSetResultsTaskScheduled = false;
        // Save result setters locally, so we can restore them if context.sync() fails.
        var resultEntryBufferCopy = [];
        var context = new CustomFunctionRequestContext();
        var invocationResults = [];
        while (this._resultEntryBuffer.length > 0) {
            // Move the next result setter to the local buffer.
            var resultEntry = this._resultEntryBuffer.pop();
            resultEntryBufferCopy.push(resultEntry);
            // Add results to be set
            invocationResults.push(resultEntry.result);
        }
        context.customFunctions.setInvocationResults(invocationResults);
        context.sync().then(function (value) {
            // Results have been successfully sent.
            // Nothing to do.
        }, function (reason) {
            // context.sync() failed.
            // Restore the setters from the local buffer, and schedule a new execution.
            _this._restoreResultEntries(resultEntryBufferCopy);
            _this._ensureSetResultsTaskIsScheduled();
        });
    };
    CustomFunctionProxy.prototype._restoreResultEntries = function (resultEntryBufferCopy) {
        var timeNow = Date.now();
        while (resultEntryBufferCopy.length > 0) {
            var resultSetter = resultEntryBufferCopy.pop();
            if (timeNow - resultSetter.timeCreated <= this._setResultsLifeMillis) {
                this._resultEntryBuffer.push(resultSetter);
            }
        }
    };
    CustomFunctionProxy.CustomFunctionExecutionStartLog = new CustomFunctionLog(CustomFunctionLoggingSeverity.Verbose, 'CustomFunctions [Execution] [Begin] Function=');
    CustomFunctionProxy.CustomFunctionExecutionFailureLog = new CustomFunctionLog(CustomFunctionLoggingSeverity.Error, 'CustomFunctions [Execution] [End] [Failure] Function=');
    CustomFunctionProxy.CustomFunctionExecutionFinishLog = new CustomFunctionLog(CustomFunctionLoggingSeverity.Info, 'CustomFunctions [Execution] [End] [Success] Function=');
    return CustomFunctionProxy;
}());
exports.CustomFunctionProxy = CustomFunctionProxy;
exports.customFunctionProxy = new CustomFunctionProxy();
Core.HostBridge.onInited(function (hostBridge) {
    exports.customFunctionProxy._initFromHostBridge(hostBridge);
});
/* End_PlaceHolder_CustomFunctions_BeforeDeclaration */
/**
 * [Api set: CustomFunctions 1.2]
 */
var CustomFunctions = /** @class */ (function (_super) {
    __extends(CustomFunctions, _super);
    function CustomFunctions() {
        return _super !== null && _super.apply(this, arguments) || this;
    }
    Object.defineProperty(CustomFunctions.prototype, "_className", {
        get: function () {
            return 'CustomFunctions';
        },
        enumerable: true,
        configurable: true
    });
    /* Begin_PlaceHolder_CustomFunctions_Custom_Members */
    CustomFunctions.initialize = function () {
        var context = new CustomFunctionRequestContext();
        return exports.customFunctionProxy.ensureInit(context).then(function () {
            context.customFunctions._SetOsfControlContainerReadyForCustomFunctions();
            return context
                .sync()
                .catch(function (error) {
                return retryOnCellEditMode(error, true /*rethrowOtherError = true, re-throw if it's not cell edit mode*/);
            });
        });
        // Helpers:
        function retryOnCellEditMode(error, rethrowOtherError) {
            var isCellEditModeError = error instanceof OfficeExtension.Error && error.code === CustomFunctionErrorCode.invalidOperationInCellEditMode;
            OfficeExtension.CoreUtility.log('Error on starting custom functions: ' + error);
            if (isCellEditModeError) {
                OfficeExtension.CoreUtility.log('Was in cell-edit mode, will try again');
                var delay_1 = exports.customFunctionProxy._ensureInitRetryDelayMillis;
                return new OfficeExtension.CoreUtility.Promise(function (resolve) { return setTimeout(resolve, delay_1); }).then(function () {
                    return CustomFunctions.initialize();
                });
            }
            // If got here, then it's NOT a cell-edit mode error:
            if (rethrowOtherError) {
                throw error;
            }
        }
    };
    CustomFunctions.prototype.register = function (metadataContent, scriptContent) {
        /* Begin_PlaceHolder_CustomFunctions_Register */
        /* End_PlaceHolder_CustomFunctions_Register */
        _invokeMethod(this, 'Register', 0 /* Default */, [metadataContent, scriptContent], 0 /* none */, 0 /* none */);
    };
    CustomFunctions.prototype.setInvocationResults = function (results) {
        /* Begin_PlaceHolder_CustomFunctions_SetInvocationResults */
        /* End_PlaceHolder_CustomFunctions_SetInvocationResults */
        _invokeMethod(this, 'SetInvocationResults', 0 /* Default */, [results], 2 /* undoable */, 0 /* none */);
    };
    CustomFunctions.prototype._SetInvocationError = function (invocationId, message) {
        /* Begin_PlaceHolder_CustomFunctions__SetInvocationError */
        /* End_PlaceHolder_CustomFunctions__SetInvocationError */
        _invokeMethod(this, '_SetInvocationError', 0 /* Default */, [invocationId, message], 2 /* undoable */, 0 /* none */);
    };
    CustomFunctions.prototype._SetInvocationResult = function (invocationId, result) {
        /* Begin_PlaceHolder_CustomFunctions__SetInvocationResult */
        /* End_PlaceHolder_CustomFunctions__SetInvocationResult */
        _invokeMethod(this, '_SetInvocationResult', 0 /* Default */, [invocationId, result], 2 /* undoable */, 0 /* none */);
    };
    CustomFunctions.prototype._SetOsfControlContainerReadyForCustomFunctions = function () {
        /* Begin_PlaceHolder_CustomFunctions__SetOsfControlContainerReadyForCustomFunctions */
        /* End_PlaceHolder_CustomFunctions__SetOsfControlContainerReadyForCustomFunctions */
        _invokeMethod(this, '_SetOsfControlContainerReadyForCustomFunctions', 0 /* Default */, [], 2 /* undoable */, 0 /* none */);
    };
    /** Handle results returned from the document
     * @private
     */
    CustomFunctions.prototype._handleResult = function (value) {
        _super.prototype._handleResult.call(this, value);
        if (_isNullOrUndefined(value))
            return;
        var obj = value;
        _fixObjectPathIfNecessary(this, obj);
        /* Begin_PlaceHolder_CustomFunctions_HandleResult */
        /* End_PlaceHolder_CustomFunctions_HandleResult */
    };
    /** Handle retrieve results
     * @private
     */
    CustomFunctions.prototype._handleRetrieveResult = function (value, result) {
        _super.prototype._handleRetrieveResult.call(this, value, result);
        /* Begin_PlaceHolder_CustomFunctions_HandleRetrieveResult */
        /* End_PlaceHolder_CustomFunctions_HandleRetrieveResult */
        _processRetrieveResult(this, value, result);
    };
    /**
     * Create a new instance of CustomFunctions object
     */
    CustomFunctions.newObject = function (context) {
        return _createTopLevelServiceObject(CustomFunctions, context, 'Microsoft.ExcelServices.CustomFunctions', false /*isCollection*/, 4 /* concurrent */);
    };
    CustomFunctions.prototype.toJSON = function () {
        return _toJson(this, /* scalarProperties: */ {}, /* navigationProperties: */ {});
    };
    return CustomFunctions;
}(OfficeExtension.ClientObject));
exports.CustomFunctions = CustomFunctions;
var _typeCustomFunctionsContainer = 'CustomFunctionsContainer';
/* Begin_PlaceHolder_CustomFunctionsContainer_BeforeDeclaration */
/* End_PlaceHolder_CustomFunctionsContainer_BeforeDeclaration */
/**
 * [Api set: CustomFunctions 1.1]
 */
var CustomFunctionsContainer = /** @class */ (function (_super) {
    __extends(CustomFunctionsContainer, _super);
    function CustomFunctionsContainer() {
        return _super !== null && _super.apply(this, arguments) || this;
    }
    Object.defineProperty(CustomFunctionsContainer.prototype, "_className", {
        get: function () {
            return 'CustomFunctionsContainer';
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(CustomFunctionsContainer.prototype, "_navigationPropertyNames", {
        get: function () {
            return ['customFunctions'];
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(CustomFunctionsContainer.prototype, "customFunctions", {
        /* Begin_PlaceHolder_CustomFunctionsContainer_Custom_Members */
        /* End_PlaceHolder_CustomFunctionsContainer_Custom_Members */
        get: function () {
            /* Begin_PlaceHolder_CustomFunctionsContainer_CustomFunctions_get */
            /* End_PlaceHolder_CustomFunctionsContainer_CustomFunctions_get */
            _throwIfApiNotSupported('CustomFunctionsContainer.customFunctions', 'CustomFunctions', '1.2', _hostName);
            if (!this._C) {
                this._C = _createPropertyObject(CustomFunctions, this, 'CustomFunctions', false /*isCollection*/, 4 /* concurrent */);
            }
            /* Begin_PlaceHolder_CustomFunctionsContainer_CustomFunctions_get_post */
            /* End_PlaceHolder_CustomFunctionsContainer_CustomFunctions_get_post */
            return this._C;
        },
        enumerable: true,
        configurable: true
    });
    // SET method is absent, because no settable properties on self or children.
    // * "customFunctions" /* NAVIGATION; EXPLICITYLY non-JSON-stringify */
    /** Handle results returned from the document
     * @private
     */
    CustomFunctionsContainer.prototype._handleResult = function (value) {
        _super.prototype._handleResult.call(this, value);
        if (_isNullOrUndefined(value))
            return;
        var obj = value;
        _fixObjectPathIfNecessary(this, obj);
        /* Begin_PlaceHolder_CustomFunctionsContainer_HandleResult */
        /* End_PlaceHolder_CustomFunctionsContainer_HandleResult */
        _handleNavigationPropertyResults(this, obj, ['customFunctions', 'CustomFunctions']);
    };
    CustomFunctionsContainer.prototype.load = function (option) {
        return _load(this, option);
    };
    /** Handle retrieve results
     * @private
     */
    CustomFunctionsContainer.prototype._handleRetrieveResult = function (value, result) {
        _super.prototype._handleRetrieveResult.call(this, value, result);
        /* Begin_PlaceHolder_CustomFunctionsContainer_HandleRetrieveResult */
        /* End_PlaceHolder_CustomFunctionsContainer_HandleRetrieveResult */
        _processRetrieveResult(this, value, result);
    };
    CustomFunctionsContainer.prototype.toJSON = function () {
        return _toJson(this, /* scalarProperties: */ {}, /* navigationProperties: */ {});
    };
    return CustomFunctionsContainer;
}(OfficeExtension.ClientObject));
exports.CustomFunctionsContainer = CustomFunctionsContainer;
/* Begin_PlaceHolder_ErrorCodesTypeName */
var CustomFunctionErrorCode;
(function (CustomFunctionErrorCode) {
    /* End_PlaceHolder_ErrorCodesTypeName */
    CustomFunctionErrorCode["generalException"] = "GeneralException";
    /* Begin_PlaceHolder_ErrorCodesAdditional */
    CustomFunctionErrorCode["invalidOperation"] = "InvalidOperation";
    CustomFunctionErrorCode["invalidOperationInCellEditMode"] = "InvalidOperationInCellEditMode";
    /* End_PlaceHolder_ErrorCodesAdditional */
})(CustomFunctionErrorCode || (CustomFunctionErrorCode = {}));


/***/ }),
/* 5 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

var __extends = (this && this.__extends) || (function () {
    var extendStatics = Object.setPrototypeOf ||
        ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
        function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p]; };
    return function (d, b) {
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
function __export(m) {
    for (var p in m) if (!exports.hasOwnProperty(p)) exports[p] = m[p];
}
Object.defineProperty(exports, "__esModule", { value: true });
var Core = __webpack_require__(0);
var Common = __webpack_require__(1);
__export(__webpack_require__(1));
var ErrorCodes = /** @class */ (function (_super) {
    __extends(ErrorCodes, _super);
    function ErrorCodes() {
        return _super !== null && _super.apply(this, arguments) || this;
    }
    ErrorCodes.propertyNotLoaded = 'PropertyNotLoaded';
    ErrorCodes.runMustReturnPromise = 'RunMustReturnPromise';
    ErrorCodes.cannotRegisterEvent = 'CannotRegisterEvent';
    ErrorCodes.invalidOrTimedOutSession = 'InvalidOrTimedOutSession';
    ErrorCodes.cannotUpdateReadOnlyProperty = 'CannotUpdateReadOnlyProperty';
    return ErrorCodes;
}(Core.CoreErrorCodes));
exports.ErrorCodes = ErrorCodes;
var TraceMarkerActionResultHandler = /** @class */ (function () {
    function TraceMarkerActionResultHandler(callback) {
        this.m_callback = callback;
    }
    TraceMarkerActionResultHandler.prototype._handleResult = function (value) {
        if (this.m_callback) {
            this.m_callback();
        }
    };
    return TraceMarkerActionResultHandler;
}());
var ActionFactory = /** @class */ (function () {
    function ActionFactory() {
    }
    ActionFactory.createSetPropertyAction = function (context, parent, propertyName, value, flags) {
        Utility.validateObjectPath(parent);
        var actionInfo = {
            Id: context._nextId(),
            ActionType: 4 /* SetProperty */,
            Name: propertyName,
            ObjectPathId: parent._objectPath.objectPathInfo.Id,
            ArgumentInfo: {}
        };
        var args = [value];
        var referencedArgumentObjectPaths = Utility.setMethodArguments(context, actionInfo.ArgumentInfo, args);
        Utility.validateReferencedObjectPaths(referencedArgumentObjectPaths);
        context._pendingRequest.ensureInstantiateObjectPath(parent._objectPath);
        context._pendingRequest.ensureInstantiateObjectPaths(referencedArgumentObjectPaths);
        var ret = new Common.Action(actionInfo, 0 /* Default */, flags);
        context._pendingRequest.addAction(ret);
        context._pendingRequest.addReferencedObjectPath(parent._objectPath);
        context._pendingRequest.addReferencedObjectPaths(referencedArgumentObjectPaths);
        ret.referencedObjectPath = parent._objectPath;
        ret.referencedArgumentObjectPaths = referencedArgumentObjectPaths;
        return ret;
    };
    ActionFactory.createMethodAction = function (context, parent, methodName, operationType, args, flags) {
        Utility.validateObjectPath(parent);
        var actionInfo = {
            Id: context._nextId(),
            ActionType: 3 /* Method */,
            Name: methodName,
            ObjectPathId: parent._objectPath.objectPathInfo.Id,
            ArgumentInfo: {}
        };
        var referencedArgumentObjectPaths = Utility.setMethodArguments(context, actionInfo.ArgumentInfo, args);
        Utility.validateReferencedObjectPaths(referencedArgumentObjectPaths);
        context._pendingRequest.ensureInstantiateObjectPath(parent._objectPath);
        context._pendingRequest.ensureInstantiateObjectPaths(referencedArgumentObjectPaths);
        var ret = new Common.Action(actionInfo, operationType, Utility._fixupApiFlags(flags));
        context._pendingRequest.addAction(ret);
        context._pendingRequest.addReferencedObjectPath(parent._objectPath);
        context._pendingRequest.addReferencedObjectPaths(referencedArgumentObjectPaths);
        ret.referencedObjectPath = parent._objectPath;
        ret.referencedArgumentObjectPaths = referencedArgumentObjectPaths;
        return ret;
    };
    ActionFactory.createQueryAction = function (context, parent, queryOption) {
        Utility.validateObjectPath(parent);
        context._pendingRequest.ensureInstantiateObjectPath(parent._objectPath);
        var actionInfo = {
            Id: context._nextId(),
            ActionType: 2 /* Query */,
            Name: '',
            ObjectPathId: parent._objectPath.objectPathInfo.Id
        };
        actionInfo.QueryInfo = queryOption;
        var ret = new Common.Action(actionInfo, 1 /* Read */, 4 /* concurrent */);
        context._pendingRequest.addAction(ret);
        context._pendingRequest.addReferencedObjectPath(parent._objectPath);
        ret.referencedObjectPath = parent._objectPath;
        return ret;
    };
    ActionFactory.createRecursiveQueryAction = function (context, parent, query) {
        Utility.validateObjectPath(parent);
        context._pendingRequest.ensureInstantiateObjectPath(parent._objectPath);
        var actionInfo = {
            Id: context._nextId(),
            ActionType: 6 /* RecursiveQuery */,
            Name: '',
            ObjectPathId: parent._objectPath.objectPathInfo.Id,
            RecursiveQueryInfo: query
        };
        var ret = new Common.Action(actionInfo, 1 /* Read */, 4 /* concurrent */);
        context._pendingRequest.addAction(ret);
        context._pendingRequest.addReferencedObjectPath(parent._objectPath);
        ret.referencedObjectPath = parent._objectPath;
        return ret;
    };
    ActionFactory.createQueryAsJsonAction = function (context, parent, queryOption) {
        Utility.validateObjectPath(parent);
        context._pendingRequest.ensureInstantiateObjectPath(parent._objectPath);
        var actionInfo = {
            Id: context._nextId(),
            ActionType: 7 /* QueryAsJson */,
            Name: '',
            ObjectPathId: parent._objectPath.objectPathInfo.Id
        };
        actionInfo.QueryInfo = queryOption;
        var ret = new Common.Action(actionInfo, 1 /* Read */, 4 /* concurrent */);
        context._pendingRequest.addAction(ret);
        context._pendingRequest.addReferencedObjectPath(parent._objectPath);
        ret.referencedObjectPath = parent._objectPath;
        return ret;
    };
    ActionFactory.createEnsureUnchangedAction = function (context, parent, objectState) {
        Utility.validateObjectPath(parent);
        context._pendingRequest.ensureInstantiateObjectPath(parent._objectPath);
        var actionInfo = {
            Id: context._nextId(),
            ActionType: 8 /* EnsureUnchanged */,
            Name: '',
            ObjectPathId: parent._objectPath.objectPathInfo.Id,
            ObjectState: objectState
        };
        var ret = new Common.Action(actionInfo, 1 /* Read */, 4 /* concurrent */);
        context._pendingRequest.addAction(ret);
        context._pendingRequest.addReferencedObjectPath(parent._objectPath);
        ret.referencedObjectPath = parent._objectPath;
        return ret;
    };
    ActionFactory.createUpdateAction = function (context, parent, objectState) {
        Utility.validateObjectPath(parent);
        context._pendingRequest.ensureInstantiateObjectPath(parent._objectPath);
        var actionInfo = {
            Id: context._nextId(),
            ActionType: 9 /* Update */,
            Name: '',
            ObjectPathId: parent._objectPath.objectPathInfo.Id,
            ObjectState: objectState
        };
        var ret = new Common.Action(actionInfo, 0 /* Default */, 0 /* none */);
        context._pendingRequest.addAction(ret);
        context._pendingRequest.addReferencedObjectPath(parent._objectPath);
        ret.referencedObjectPath = parent._objectPath;
        return ret;
    };
    ActionFactory.createInstantiateAction = function (context, obj) {
        Utility.validateObjectPath(obj);
        context._pendingRequest.ensureInstantiateObjectPath(obj._objectPath.parentObjectPath);
        context._pendingRequest.ensureInstantiateObjectPaths(obj._objectPath.argumentObjectPaths);
        var actionInfo = {
            Id: context._nextId(),
            ActionType: 1 /* Instantiate */,
            Name: '',
            ObjectPathId: obj._objectPath.objectPathInfo.Id
        };
        var ret = new Common.Action(actionInfo, 1 /* Read */, 4 /* concurrent */);
        ret.referencedObjectPath = obj._objectPath;
        context._pendingRequest.addAction(ret);
        context._pendingRequest.addReferencedObjectPath(obj._objectPath);
        context._pendingRequest.addActionResultHandler(ret, new InstantiateActionResultHandler(obj));
        return ret;
    };
    ActionFactory.createTraceAction = function (context, message, addTraceMessage) {
        var actionInfo = {
            Id: context._nextId(),
            ActionType: 5 /* Trace */,
            Name: 'Trace',
            ObjectPathId: 0
        };
        var ret = new Common.Action(actionInfo, 1 /* Read */, 4 /* concurrent */);
        context._pendingRequest.addAction(ret);
        if (addTraceMessage) {
            context._pendingRequest.addTrace(actionInfo.Id, message);
        }
        return ret;
    };
    ActionFactory.createTraceMarkerForCallback = function (context, callback) {
        var action = ActionFactory.createTraceAction(context, null, false);
        context._pendingRequest.addActionResultHandler(action, new TraceMarkerActionResultHandler(callback));
    };
    return ActionFactory;
}());
exports.ActionFactory = ActionFactory;
var ClientObject = /** @class */ (function (_super) {
    __extends(ClientObject, _super);
    function ClientObject(context, objectPath) {
        var _this = _super.call(this, context, objectPath) || this;
        Utility.checkArgumentNull(context, 'context');
        _this.m_context = context;
        if (_this._objectPath) {
            // If object is being created during a normal API flow (and NOT as part of processing load results),
            // create an instantiation action and call keepReference, if applicable
            if (!context._processingResult && context._pendingRequest) {
                ActionFactory.createInstantiateAction(context, _this);
                if (context._autoCleanup && _this._KeepReference) {
                    context.trackedObjects._autoAdd(_this);
                }
            }
        }
        return _this;
    }
    Object.defineProperty(ClientObject.prototype, "context", {
        get: function () {
            return this.m_context;
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(ClientObject.prototype, "isNull", {
        /**
         * Returns a boolean value for whether the corresponding object is a null object. You must call "context.sync()" before reading the isNull property.
         */
        get: function () {
            Utility.throwIfNotLoaded('isNull', this._isNull, null /*entityName*/, this._isNull);
            return this._isNull;
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(ClientObject.prototype, "isNullObject", {
        /**
         * Returns a boolean value for whether the corresponding object is a null object. You must call "context.sync()" before reading the isNullObject property.
         */
        get: function () {
            Utility.throwIfNotLoaded('isNullObject', this._isNull, null /*entityName*/, this._isNull);
            return this._isNull;
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(ClientObject.prototype, "_isNull", {
        get: function () {
            return this.m_isNull;
        },
        set: function (value) {
            this.m_isNull = value;
            if (value && this._objectPath) {
                this._objectPath._updateAsNullObject();
            }
        },
        enumerable: true,
        configurable: true
    });
    /** Handle the result returned by ctx.load()
     * The Utility.fixObjectPathIfNecessary() is not called here as all of the
     * derived class's _handleResult() called the Utility.fixObjectPathIfNecessary.
     */
    ClientObject.prototype._handleResult = function (value) {
        this._isNull = Utility.isNullOrUndefined(value);
        this.context.trackedObjects._autoTrackIfNecessaryWhenHandleObjectResultValue(this, value);
    };
    /** Used by InstantiateActionResultHandler to set the object's Id, _Id or _ReferenceId
     */
    ClientObject.prototype._handleIdResult = function (value) {
        this._isNull = Utility.isNullOrUndefined(value);
        Utility.fixObjectPathIfNecessary(this, value);
        this.context.trackedObjects._autoTrackIfNecessaryWhenHandleObjectResultValue(this, value);
    };
    /** Used by RetrieveResult to set the object's essential state, such as object path, id or _ReferenceId.
     */
    ClientObject.prototype._handleRetrieveResult = function (value, result) {
        this._handleIdResult(value);
    };
    /**
     * Sets properties on this object based on the passed-in source. Each type knows the writeable scalars and the object properties that are relevant for it.
     */
    ClientObject.prototype._recursivelySet = function (input, options, scalarWriteablePropertyNames /* property names for valid writeable scalar properties, in disp-id (application) order */, objectPropertyNames /* property names for valid navigation properties */, notAllowedToBeSetPropertyNames /* Properties that have a setter (if applicable), but are not allowed to be set */) {
        var isClientObject = input instanceof ClientObject;
        var originalInput = input;
        if (isClientObject) {
            // Extract the settable source data:
            if (Object.getPrototypeOf(this) === Object.getPrototypeOf(input)) {
                input = JSON.parse(JSON.stringify(input));
            }
            else {
                // If setting with a client object, can only set with an object of the same type
                throw Core._Internal.RuntimeError._createInvalidArgError({
                    argumentName: 'properties',
                    errorLocation: this._className + '.set'
                });
            }
        }
        try {
            var prop;
            // Now, using a clean JSON, start setting the object properties
            for (var i = 0; i < scalarWriteablePropertyNames.length; i++) {
                prop = scalarWriteablePropertyNames[i];
                if (input.hasOwnProperty(prop)) {
                    if (typeof input[prop] !== 'undefined') {
                        // For a scalar property, just set it. Setting to null should no-op,
                        // but let the property take care of that for itself.
                        // Note that don't need to worry about whether property exists, because it *has to*,
                        // by virtue of being part of the "scalarWriteablePropertyNames" array.
                        this[prop] = input[prop];
                    }
                }
            }
            for (var i = 0; i < objectPropertyNames.length; i++) {
                prop = objectPropertyNames[i];
                if (input.hasOwnProperty(prop)) {
                    if (typeof input[prop] !== 'undefined') {
                        // If original data was a client object, want to get its actual property rather than
                        // the stringified form.  That way, "set" will have leaway for read-only property
                        // on a client object, even though it wouldn't have on regular JSON.
                        var dataToPassToSet = isClientObject ? originalInput[prop] : input[prop];
                        this[prop].set(dataToPassToSet, options);
                    }
                }
            }
            // In general, properties not allowed to be set are indeed *not allowed to be set*.
            // However, if the passed-in object is a Client Object, which may have had some of these
            // properties loaded, we want to just skip these properties and *not* throw.
            // The "throwOnReadOnly" is thus applicable both to scalar read-only properties
            // and also to these not-allowed-to-be-set properties
            var throwOnReadOnly = !isClientObject;
            if (options && !Utility.isNullOrUndefined(throwOnReadOnly)) {
                throwOnReadOnly = options.throwOnReadOnly;
            }
            // BTW: The scenario where the developer may want to have options.throwOnReadOnly as "false"
            // is if he/she had previously serialized a client object (JSON.stringify(range)), stored it in a setting,
            // and now want to restore that state, skipping over the read-only properties (which would otherwise have thrown an error).
            // The check doesn't really make sense client objects, but keeping it in case the developer --
            // for whatever reason -- decided to set throwOnReadOnly to "true" on client objects (opposite of default behavior).
            for (var i = 0; i < notAllowedToBeSetPropertyNames.length; i++) {
                prop = notAllowedToBeSetPropertyNames[i];
                if (input.hasOwnProperty(prop)) {
                    if (typeof input[prop] !== 'undefined' && throwOnReadOnly) {
                        throw new Core._Internal.RuntimeError({
                            code: Core.CoreErrorCodes.invalidArgument,
                            message: Core.CoreUtility._getResourceString(ResourceStrings.cannotApplyPropertyThroughSetMethod, prop),
                            debugInfo: {
                                errorLocation: prop /* use property name as error location (will get wrapped into an outer Object.set error) */
                            }
                        });
                    }
                }
            }
            // Make sure that there aren't any "unused" properties on the source, which would indicate an error.
            // This check should throw by default for regular input, and NOT throw by default for client objects.
            for (prop in input) {
                if (scalarWriteablePropertyNames.indexOf(prop) < 0 && objectPropertyNames.indexOf(prop) < 0) {
                    // If the property is neither on the scalar-property nor object-property list, it either doesn't exist, or is read-only.
                    var propertyDescriptor = Object.getOwnPropertyDescriptor(Object.getPrototypeOf(this), prop);
                    if (!propertyDescriptor) {
                        throw new Core._Internal.RuntimeError({
                            code: Core.CoreErrorCodes.invalidArgument,
                            message: Core.CoreUtility._getResourceString(ResourceStrings.propertyDoesNotExist, prop),
                            debugInfo: {
                                errorLocation: prop /* use property name as error location (will get wrapped into an outer Object.set error) */
                            }
                        });
                    }
                    if (throwOnReadOnly && !propertyDescriptor.set) {
                        throw new Core._Internal.RuntimeError({
                            code: Core.CoreErrorCodes.invalidArgument,
                            message: Core.CoreUtility._getResourceString(ResourceStrings.attemptingToSetReadOnlyProperty, prop),
                            debugInfo: {
                                errorLocation: prop /* use property name as error location (will get wrapped into an outer Object.set error) */
                            }
                        });
                    }
                }
            }
        }
        catch (innerError) {
            throw new Core._Internal.RuntimeError({
                code: Core.CoreErrorCodes.invalidArgument,
                message: Core.CoreUtility._getResourceString(Core.CoreResourceStrings.invalidArgument, 'properties'),
                debugInfo: {
                    errorLocation: this._className + '.set'
                },
                innerError: innerError
            });
        }
    };
    /**
     * Update properties on this object based on the passed-in JSON. Each type knows the writeable scalars and the object properties that are relevant for it.
     */
    ClientObject.prototype._recursivelyUpdate = function (properties) {
        var shouldPolyfill = Common._internalConfig.alwaysPolyfillClientObjectUpdateMethod;
        if (!shouldPolyfill) {
            // check whether the host support RichApiRuntime 1.2
            shouldPolyfill = !Utility.isSetSupported('RichApiRuntime', '1.2');
        }
        try {
            // Scalar property names
            var scalarPropNames = this[Constants.scalarPropertyNames];
            if (!scalarPropNames) {
                scalarPropNames = [];
            }
            // Whether the scalar property is updateable
            var scalarPropUpdatable = this[Constants.scalarPropertyUpdateable];
            if (!scalarPropUpdatable) {
                scalarPropUpdatable = [];
                for (var i = 0; i < scalarPropNames.length; i++) {
                    scalarPropUpdatable.push(false);
                }
            }
            // Navigation property names
            var navigationPropNames = this[Constants.navigationPropertyNames];
            if (!navigationPropNames) {
                navigationPropNames = [];
            }
            var scalarProps = {};
            var navigationProps = {};
            var scalarPropCount = 0;
            // Collect the scalar properties into scalarProps and collect the
            // navigation properties into navigationProps.
            for (var propName in properties) {
                var index = scalarPropNames.indexOf(propName);
                if (index >= 0) {
                    if (!scalarPropUpdatable[index]) {
                        // The scalar property is not updateable.
                        throw new Core._Internal.RuntimeError({
                            code: Core.CoreErrorCodes.invalidArgument,
                            message: Core.CoreUtility._getResourceString(ResourceStrings.attemptingToSetReadOnlyProperty, propName),
                            debugInfo: {
                                errorLocation: propName /* use property name as error location (will get wrapped into an outer Object.update error) */
                            }
                        });
                    }
                    scalarProps[propName] = properties[propName];
                    ++scalarPropCount;
                }
                else if (navigationPropNames.indexOf(propName) >= 0) {
                    navigationProps[propName] = properties[propName];
                }
                else {
                    // It's unknown property.
                    throw new Core._Internal.RuntimeError({
                        code: Core.CoreErrorCodes.invalidArgument,
                        message: Core.CoreUtility._getResourceString(ResourceStrings.propertyDoesNotExist, propName),
                        debugInfo: {
                            errorLocation: propName /* use property name as error location (will get wrapped into an outer Object.update error) */
                        }
                    });
                }
            }
            if (scalarPropCount > 0) {
                if (shouldPolyfill) {
                    // The scalarPropNames is ordered by the DispId. We need to set the value in
                    // the order of DispId, which is the order of scalarPropNames.
                    for (var i = 0; i < scalarPropNames.length; i++) {
                        var propName = scalarPropNames[i];
                        var propValue = scalarProps[propName];
                        if (!Utility.isUndefined(propValue)) {
                            ActionFactory.createSetPropertyAction(this.context, this, propName, propValue);
                        }
                    }
                }
                else {
                    ActionFactory.createUpdateAction(this.context, this, scalarProps);
                }
            }
            for (var propName in navigationProps) {
                var navigationPropProxy = this[propName];
                var navigationPropValue = navigationProps[propName];
                navigationPropProxy._recursivelyUpdate(navigationPropValue);
            }
        }
        catch (innerError) {
            throw new Core._Internal.RuntimeError({
                code: Core.CoreErrorCodes.invalidArgument,
                message: Core.CoreUtility._getResourceString(Core.CoreResourceStrings.invalidArgument, 'properties'),
                debugInfo: {
                    errorLocation: this._className + '.update'
                },
                innerError: innerError
            });
        }
    };
    return ClientObject;
}(Common.ClientObjectBase));
exports.ClientObject = ClientObject;
var HostBridgeRequestExecutor = /** @class */ (function () {
    function HostBridgeRequestExecutor(session) {
        this.m_session = session;
    }
    HostBridgeRequestExecutor.prototype.executeAsync = function (customData, requestFlags, requestMessage) {
        var httpRequestInfo = {
            url: Core.CoreConstants.processQuery,
            method: 'POST',
            headers: requestMessage.Headers,
            body: requestMessage.Body
        };
        var message = {
            id: Core.HostBridge.nextId(),
            type: 1 /* request */,
            flags: requestFlags,
            message: httpRequestInfo
        };
        Core.CoreUtility.log(JSON.stringify(message));
        return this.m_session.sendMessageToHost(message).then(function (nativeBridgeResponse) {
            Core.CoreUtility.log('Received response: ' + JSON.stringify(nativeBridgeResponse));
            var responseInfo = nativeBridgeResponse.message;
            var response;
            if (responseInfo.statusCode === 200) {
                response = {
                    ErrorCode: null,
                    ErrorMessage: null,
                    Headers: responseInfo.headers,
                    Body: Core.CoreUtility._parseResponseBody(responseInfo)
                };
            }
            else {
                Core.CoreUtility.log('Error Response:' + responseInfo.body);
                var error = Core.CoreUtility._parseErrorResponse(responseInfo);
                response = {
                    ErrorCode: error.errorCode,
                    ErrorMessage: error.errorMessage,
                    Headers: responseInfo.headers,
                    Body: null
                };
            }
            return response;
        });
    };
    return HostBridgeRequestExecutor;
}());
var HostBridgeSession = /** @class */ (function (_super) {
    __extends(HostBridgeSession, _super);
    function HostBridgeSession(m_bridge) {
        var _this = _super.call(this) || this;
        _this.m_bridge = m_bridge;
        _this.m_bridge.addHostMessageHandler(function (message) {
            if (message.type === 3 /* genericMessage */) {
                GenericEventRegistration.getGenericEventRegistration()._handleRichApiMessage(message.message);
            }
        });
        return _this;
    }
    HostBridgeSession.getInstanceIfHostBridgeInited = function () {
        if (Core.HostBridge.instance) {
            if (Core.CoreUtility.isNullOrUndefined(HostBridgeSession.s_instance) ||
                HostBridgeSession.s_instance.m_bridge !== Core.HostBridge.instance) {
                HostBridgeSession.s_instance = new HostBridgeSession(Core.HostBridge.instance);
            }
            return HostBridgeSession.s_instance;
        }
        return null;
    };
    HostBridgeSession.prototype._resolveRequestUrlAndHeaderInfo = function () {
        return Core.CoreUtility._createPromiseFromResult(null);
    };
    HostBridgeSession.prototype._createRequestExecutorOrNull = function () {
        Core.CoreUtility.log('NativeBridgeSession::CreateRequestExecutor');
        return new HostBridgeRequestExecutor(this);
    };
    Object.defineProperty(HostBridgeSession.prototype, "eventRegistration", {
        get: function () {
            return GenericEventRegistration.getGenericEventRegistration();
        },
        enumerable: true,
        configurable: true
    });
    HostBridgeSession.prototype.sendMessageToHost = function (message) {
        return this.m_bridge.sendMessageToHostAndExpectResponse(message);
    };
    return HostBridgeSession;
}(Core.SessionBase));
var ClientRequestContext = /** @class */ (function (_super) {
    __extends(ClientRequestContext, _super);
    // It could also be of type IRequestUrlAndHeaderInfoResolver. But we do not
    // want to expose IRequestUrlAndHeaderInfoResolver publicly. So "any" is used
    function ClientRequestContext(url) {
        var _this = _super.call(this) || this;
        _this.m_customRequestHeaders = {};
        _this.m_batchMode = 0 /* implicit */;
        _this._onRunFinishedNotifiers = []; // Used in conjunction with _autoCleanup; see _runCommon below.
        if (Core.SessionBase._overrideSession) {
            _this.m_requestUrlAndHeaderInfoResolver = Core.SessionBase._overrideSession;
        }
        else {
            if (Utility.isNullOrUndefined(url) || (typeof url === 'string' && url.length === 0)) {
                url = ClientRequestContext.defaultRequestUrlAndHeaders;
                if (!url) {
                    url = { url: Core.CoreConstants.localDocument, headers: {} };
                }
            }
            if (typeof url === 'string') {
                _this.m_requestUrlAndHeaderInfo = { url: url, headers: {} };
            }
            else if (ClientRequestContext.isRequestUrlAndHeaderInfoResolver(url)) {
                _this.m_requestUrlAndHeaderInfoResolver = url;
            }
            else if (ClientRequestContext.isRequestUrlAndHeaderInfo(url)) {
                var requestInfo = url;
                _this.m_requestUrlAndHeaderInfo = { url: requestInfo.url, headers: {} };
                Core.CoreUtility._copyHeaders(requestInfo.headers, _this.m_requestUrlAndHeaderInfo.headers);
            }
            else {
                throw Core._Internal.RuntimeError._createInvalidArgError({ argumentName: 'url' });
            }
        }
        if (!_this.m_requestUrlAndHeaderInfoResolver &&
            _this.m_requestUrlAndHeaderInfo &&
            Core.CoreUtility._isLocalDocumentUrl(_this.m_requestUrlAndHeaderInfo.url) &&
            HostBridgeSession.getInstanceIfHostBridgeInited()) {
            _this.m_requestUrlAndHeaderInfo = null;
            _this.m_requestUrlAndHeaderInfoResolver = HostBridgeSession.getInstanceIfHostBridgeInited();
        }
        if (_this.m_requestUrlAndHeaderInfoResolver instanceof Core.SessionBase) {
            _this.m_session = _this.m_requestUrlAndHeaderInfoResolver;
        }
        _this._processingResult = false;
        _this._customData = Constants.iterativeExecutor;
        // Bind the sync function, to make it possible to call ".then(ctx.sync)"
        _this.sync = _this.sync.bind(_this);
        return _this;
    }
    Object.defineProperty(ClientRequestContext.prototype, "session", {
        get: function () {
            return this.m_session;
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(ClientRequestContext.prototype, "eventRegistration", {
        get: function () {
            if (this.m_session) {
                return this.m_session.eventRegistration;
            }
            return _Internal.officeJsEventRegistration;
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(ClientRequestContext.prototype, "_url", {
        get: function () {
            if (this.m_requestUrlAndHeaderInfo) {
                return this.m_requestUrlAndHeaderInfo.url;
            }
            return null;
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(ClientRequestContext.prototype, "_pendingRequest", {
        get: function () {
            if (this.m_pendingRequest == null) {
                this.m_pendingRequest = new ClientRequest(this);
            }
            return this.m_pendingRequest;
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(ClientRequestContext.prototype, "debugInfo", {
        get: function () {
            var prettyPrinter = new RequestPrettyPrinter(this._rootObjectPropertyName, this._pendingRequest._objectPaths, this._pendingRequest._actions, Common._internalConfig.showDisposeInfoInDebugInfo);
            var statements = prettyPrinter.process();
            return { pendingStatements: statements };
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(ClientRequestContext.prototype, "trackedObjects", {
        get: function () {
            if (!this.m_trackedObjects) {
                this.m_trackedObjects = new TrackedObjects(this);
            }
            return this.m_trackedObjects;
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(ClientRequestContext.prototype, "requestHeaders", {
        get: function () {
            return this.m_customRequestHeaders;
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(ClientRequestContext.prototype, "batchMode", {
        get: function () {
            return this.m_batchMode;
        },
        enumerable: true,
        configurable: true
    });
    ClientRequestContext.prototype.ensureInProgressBatchIfBatchMode = function () {
        if (this.m_batchMode === 1 /* explicit */ && !this.m_explicitBatchInProgress) {
            throw Utility.createRuntimeError(Core.CoreErrorCodes.generalException, Core.CoreUtility._getResourceString(ResourceStrings.notInsideBatch), null);
        }
    };
    ClientRequestContext.prototype.load = function (clientObj, option /* And could also be { [key: string]: any }, but we don't want this in the public signature */) {
        Utility.validateContext(this, clientObj);
        var queryOption = ClientRequestContext._parseQueryOption(option);
        var action = ActionFactory.createQueryAction(this, clientObj, queryOption);
        this._pendingRequest.addActionResultHandler(action, clientObj);
    };
    ClientRequestContext.isLoadOption = function (loadOption) {
        if (!Utility.isUndefined(loadOption.select) &&
            (typeof loadOption.select === 'string' || Array.isArray(loadOption.select)))
            return true;
        if (!Utility.isUndefined(loadOption.expand) &&
            (typeof loadOption.expand === 'string' || Array.isArray(loadOption.expand)))
            return true;
        if (!Utility.isUndefined(loadOption.top) && typeof loadOption.top === 'number')
            return true;
        if (!Utility.isUndefined(loadOption.skip) && typeof loadOption.skip === 'number')
            return true;
        for (var i in loadOption) {
            // it's not empty JSON
            return false;
        }
        return true;
    };
    ClientRequestContext.parseStrictLoadOption = function (option) {
        var ret = { Select: [] };
        ClientRequestContext.parseStrictLoadOptionHelper(ret, '', 'option', option);
        return ret;
    };
    ClientRequestContext.combineQueryPath = function (pathPrefix, key, separator) {
        if (pathPrefix.length === 0) {
            return key;
        }
        else {
            return pathPrefix + separator + key;
        }
    };
    ClientRequestContext.parseStrictLoadOptionHelper = function (queryInfo, pathPrefix, argPrefix, option) {
        for (var key in option) {
            var value = option[key];
            if (key === '$all') {
                if (typeof value !== 'boolean') {
                    throw Core._Internal.RuntimeError._createInvalidArgError({
                        argumentName: ClientRequestContext.combineQueryPath(argPrefix, key, '.')
                    });
                }
                if (value) {
                    queryInfo.Select.push(ClientRequestContext.combineQueryPath(pathPrefix, '*', '/'));
                }
            }
            else if (key === '$top') {
                if (typeof value !== 'number' || pathPrefix.length > 0) {
                    throw Core._Internal.RuntimeError._createInvalidArgError({
                        argumentName: ClientRequestContext.combineQueryPath(argPrefix, key, '.')
                    });
                }
                queryInfo.Top = value;
            }
            else if (key === '$skip') {
                if (typeof value !== 'number' || pathPrefix.length > 0) {
                    throw Core._Internal.RuntimeError._createInvalidArgError({
                        argumentName: ClientRequestContext.combineQueryPath(argPrefix, key, '.')
                    });
                }
                queryInfo.Skip = value;
            }
            else {
                if (typeof value === 'boolean') {
                    if (value) {
                        queryInfo.Select.push(ClientRequestContext.combineQueryPath(pathPrefix, key, '/'));
                    }
                }
                else if (typeof value === 'object') {
                    ClientRequestContext.parseStrictLoadOptionHelper(queryInfo, ClientRequestContext.combineQueryPath(pathPrefix, key, '/'), ClientRequestContext.combineQueryPath(argPrefix, key, '.'), value);
                }
                else {
                    throw Core._Internal.RuntimeError._createInvalidArgError({
                        argumentName: ClientRequestContext.combineQueryPath(argPrefix, key, '.')
                    });
                }
            }
        }
    };
    ClientRequestContext._parseQueryOption = function (option) {
        var queryOption = {};
        if (typeof option == 'string') {
            var select = option;
            queryOption.Select = Utility._parseSelectExpand(select);
        }
        else if (Array.isArray(option)) {
            queryOption.Select = option;
        }
        else if (typeof option === 'object') {
            // If it's an object, it's two possibilities.
            // Option A is that it's an older generic LoadOptions object.
            // Option B is that it's a "XYZLoadOptions" object that's specifically tailored for this object type. (PENDING)
            var loadOption = option;
            if (ClientRequestContext.isLoadOption(loadOption)) {
                if (typeof loadOption.select == 'string') {
                    queryOption.Select = Utility._parseSelectExpand(loadOption.select);
                }
                else if (Array.isArray(loadOption.select)) {
                    queryOption.Select = loadOption.select;
                }
                else if (!Utility.isNullOrUndefined(loadOption.select)) {
                    throw Core._Internal.RuntimeError._createInvalidArgError({ argumentName: 'option.select' });
                }
                if (typeof loadOption.expand == 'string') {
                    queryOption.Expand = Utility._parseSelectExpand(loadOption.expand);
                }
                else if (Array.isArray(loadOption.expand)) {
                    queryOption.Expand = loadOption.expand;
                }
                else if (!Utility.isNullOrUndefined(loadOption.expand)) {
                    throw Core._Internal.RuntimeError._createInvalidArgError({ argumentName: 'option.expand' });
                }
                if (typeof loadOption.top === 'number') {
                    queryOption.Top = loadOption.top;
                }
                else if (!Utility.isNullOrUndefined(loadOption.top)) {
                    throw Core._Internal.RuntimeError._createInvalidArgError({ argumentName: 'option.top' });
                }
                if (typeof loadOption.skip === 'number') {
                    queryOption.Skip = loadOption.skip;
                }
                else if (!Utility.isNullOrUndefined(loadOption.skip)) {
                    throw Core._Internal.RuntimeError._createInvalidArgError({ argumentName: 'option.skip' });
                }
            }
            else {
                queryOption = ClientRequestContext.parseStrictLoadOption(option);
            }
        }
        else if (!Utility.isNullOrUndefined(option)) {
            throw Core._Internal.RuntimeError._createInvalidArgError({ argumentName: 'option' });
        }
        return queryOption;
    };
    /**
     * Queues up a command to recursively load the specified properties of the object and its navigation properties.
     * You must call "context.sync()" before reading the properties.
     *
     * @param clientObj The object to be loaded
     * @param options The load options for the types
     * @param maxDepth The max recursive depth
     */
    ClientRequestContext.prototype.loadRecursive = function (clientObj, options, maxDepth) {
        if (!Utility.isPlainJsonObject(options)) {
            throw Core._Internal.RuntimeError._createInvalidArgError({ argumentName: 'options' });
        }
        var quries = {};
        for (var key in options) {
            quries[key] = ClientRequestContext._parseQueryOption(options[key]);
        }
        var action = ActionFactory.createRecursiveQueryAction(this, clientObj, { Queries: quries, MaxDepth: maxDepth });
        this._pendingRequest.addActionResultHandler(action, clientObj);
    };
    ClientRequestContext.prototype.trace = function (message) {
        ActionFactory.createTraceAction(this, message, true /*addTraceMessage*/);
    };
    // No extra processing by default. Subclasses can override if desired
    ClientRequestContext.prototype._processOfficeJsErrorResponse = function (officeJsErrorCode, response) { };
    ClientRequestContext.prototype.ensureRequestUrlAndHeaderInfo = function () {
        var _this = this;
        return Utility._createPromiseFromResult(null).then(function () {
            if (!_this.m_requestUrlAndHeaderInfo) {
                return _this.m_requestUrlAndHeaderInfoResolver._resolveRequestUrlAndHeaderInfo().then(function (value) {
                    _this.m_requestUrlAndHeaderInfo = value;
                    if (!_this.m_requestUrlAndHeaderInfo) {
                        _this.m_requestUrlAndHeaderInfo = { url: Core.CoreConstants.localDocument, headers: {} };
                    }
                    if (Utility.isNullOrEmptyString(_this.m_requestUrlAndHeaderInfo.url)) {
                        _this.m_requestUrlAndHeaderInfo.url = Core.CoreConstants.localDocument;
                    }
                    if (!_this.m_requestUrlAndHeaderInfo.headers) {
                        _this.m_requestUrlAndHeaderInfo.headers = {};
                    }
                    if (typeof _this.m_requestUrlAndHeaderInfoResolver._createRequestExecutorOrNull === 'function') {
                        var executor = _this.m_requestUrlAndHeaderInfoResolver._createRequestExecutorOrNull();
                        if (executor) {
                            _this._requestExecutor = executor;
                        }
                    }
                });
            }
        });
    };
    ClientRequestContext.prototype.syncPrivateMain = function () {
        var _this = this;
        return this.ensureRequestUrlAndHeaderInfo().then(function () {
            var req = _this._pendingRequest;
            _this.m_pendingRequest = null;
            return _this.processPreSyncPromises(req).then(function () { return _this.syncPrivate(req); });
        });
    };
    ClientRequestContext.prototype.syncPrivate = function (req) {
        var _this = this;
        // If there are no actions to dispatch, short-circuit without sending an empty request to the server
        if (!req.hasActions) {
            return this.processPendingEventHandlers(req);
        }
        var _a = req.buildRequestMessageBodyAndRequestFlags(), msgBody = _a.body, requestFlags = _a.flags;
        if (this._requestFlagModifier) {
            requestFlags |= this._requestFlagModifier;
        }
        if (!this._requestExecutor) {
            if (Core.CoreUtility._isLocalDocumentUrl(this.m_requestUrlAndHeaderInfo.url)) {
                this._requestExecutor = new OfficeJsRequestExecutor(this);
            }
            else {
                this._requestExecutor = new Common.HttpRequestExecutor();
            }
        }
        var requestExecutor = this._requestExecutor;
        var headers = {};
        Core.CoreUtility._copyHeaders(this.m_requestUrlAndHeaderInfo.headers, headers);
        Core.CoreUtility._copyHeaders(this.m_customRequestHeaders, headers);
        var requestExecutorRequestMessage = {
            Url: this.m_requestUrlAndHeaderInfo.url,
            Headers: headers,
            Body: msgBody
        };
        req.invalidatePendingInvalidObjectPaths();
        var errorFromResponse = null;
        var errorFromProcessEventHandlers = null;
        this._lastSyncStart = typeof performance === 'undefined' ? 0 : performance.now();
        this._lastRequestFlags = requestFlags;
        return requestExecutor
            .executeAsync(this._customData, requestFlags, requestExecutorRequestMessage)
            .then(function (response) {
            _this._lastSyncEnd = typeof performance === 'undefined' ? 0 : performance.now();
            errorFromResponse = _this.processRequestExecutorResponseMessage(req, response);
            return _this.processPendingEventHandlers(req).catch(function (ex) {
                Core.CoreUtility.log('Error in processPendingEventHandlers');
                Core.CoreUtility.log(JSON.stringify(ex));
                errorFromProcessEventHandlers = ex;
            });
        })
            .then(function () {
            if (errorFromResponse) {
                Core.CoreUtility.log('Throw error from response: ' + JSON.stringify(errorFromResponse));
                throw errorFromResponse;
            }
            if (errorFromProcessEventHandlers) {
                Core.CoreUtility.log('Throw error from ProcessEventHandler: ' + JSON.stringify(errorFromProcessEventHandlers));
                var transformedError = null;
                if (errorFromProcessEventHandlers instanceof Core._Internal.RuntimeError) {
                    transformedError = errorFromProcessEventHandlers;
                    transformedError.traceMessages = req._responseTraceMessages;
                }
                else {
                    var message = null;
                    if (typeof errorFromProcessEventHandlers === 'string') {
                        message = errorFromProcessEventHandlers;
                    }
                    else {
                        message = errorFromProcessEventHandlers.message;
                    }
                    if (Utility.isNullOrEmptyString(message)) {
                        message = Core.CoreUtility._getResourceString(ResourceStrings.cannotRegisterEvent);
                    }
                    transformedError = new Core._Internal.RuntimeError({
                        code: ErrorCodes.cannotRegisterEvent,
                        message: message,
                        traceMessages: req._responseTraceMessages
                    });
                }
                throw transformedError;
            }
        });
    };
    ClientRequestContext.prototype.processRequestExecutorResponseMessage = function (req, response) {
        if (response.Body && response.Body.TraceIds) {
            req._setResponseTraceIds(response.Body.TraceIds);
        }
        var traceMessages = req._responseTraceMessages;
        var errorStatementInfo = null;
        if (response.Body) {
            if (response.Body.Error && response.Body.Error.ActionIndex >= 0) {
                var prettyPrinter = new RequestPrettyPrinter(this._rootObjectPropertyName, req._objectPaths, req._actions, 
                /*showDispose*/ false, 
                /*removePII*/ true);
                var debugInfoStatementInfo = prettyPrinter.processForDebugStatementInfo(response.Body.Error.ActionIndex);
                errorStatementInfo = {
                    statement: debugInfoStatementInfo.statement,
                    surroundingStatements: debugInfoStatementInfo.surroundingStatements,
                    fullStatements: ['Please enable config.extendedErrorLogging to see full statements.']
                };
                if (Common.config.extendedErrorLogging) {
                    prettyPrinter = new RequestPrettyPrinter(this._rootObjectPropertyName, req._objectPaths, req._actions, 
                    /*showDispose*/ false, 
                    /*removePII*/ false);
                    errorStatementInfo.fullStatements = prettyPrinter.process();
                }
            }
            var actionResults = null;
            if (response.Body.Results) {
                actionResults = response.Body.Results;
            }
            else if (response.Body.ProcessedResults && response.Body.ProcessedResults.Results) {
                actionResults = response.Body.ProcessedResults.Results;
            }
            if (actionResults) {
                this._processingResult = true;
                try {
                    req.processResponse(actionResults);
                }
                finally {
                    this._processingResult = false;
                }
            }
        }
        if (!Utility.isNullOrEmptyString(response.ErrorCode)) {
            return new Core._Internal.RuntimeError({
                code: response.ErrorCode,
                message: response.ErrorMessage,
                traceMessages: traceMessages
            });
        }
        else if (response.Body && response.Body.Error) {
            var debugInfo = {
                errorLocation: response.Body.Error.Location
            };
            if (errorStatementInfo) {
                debugInfo.statement = errorStatementInfo.statement;
                debugInfo.surroundingStatements = errorStatementInfo.surroundingStatements;
                debugInfo.fullStatements = errorStatementInfo.fullStatements;
            }
            return new Core._Internal.RuntimeError({
                code: response.Body.Error.Code,
                message: response.Body.Error.Message,
                traceMessages: traceMessages,
                debugInfo: debugInfo
            });
        }
        return null;
    };
    ClientRequestContext.prototype.processPendingEventHandlers = function (req) {
        var ret = Utility._createPromiseFromResult(null);
        for (var i = 0; i < req._pendingProcessEventHandlers.length; i++) {
            var eventHandlers = req._pendingProcessEventHandlers[i];
            ret = ret.then(this.createProcessOneEventHandlersFunc(eventHandlers, req));
        }
        return ret;
    };
    ClientRequestContext.prototype.createProcessOneEventHandlersFunc = function (eventHandlers, req) {
        return function () { return eventHandlers._processRegistration(req); };
    };
    ClientRequestContext.prototype.processPreSyncPromises = function (req) {
        var ret = Utility._createPromiseFromResult(null);
        for (var i = 0; i < req._preSyncPromises.length; i++) {
            var p = req._preSyncPromises[i];
            ret = ret.then(this.createProcessOneProSyncFunc(p));
        }
        return ret;
    };
    ClientRequestContext.prototype.createProcessOneProSyncFunc = function (p) {
        return function () { return p; };
    };
    ClientRequestContext.prototype.sync = function (passThroughValue) {
        return this.syncPrivateMain().then(function () { return passThroughValue; });
    };
    ClientRequestContext.prototype.batch = function (batchBody) {
        var _this = this;
        if (this.m_batchMode !== 1 /* explicit */) {
            return Core.CoreUtility._createPromiseFromException(Utility.createRuntimeError(Core.CoreErrorCodes.generalException, null, null));
        }
        if (this.m_explicitBatchInProgress) {
            return Core.CoreUtility._createPromiseFromException(Utility.createRuntimeError(Core.CoreErrorCodes.generalException, Core.CoreUtility._getResourceString(ResourceStrings.pendingBatchInProgress), null));
        }
        if (Utility.isNullOrUndefined(batchBody)) {
            return Utility._createPromiseFromResult(null);
        }
        this.m_explicitBatchInProgress = true;
        var previousRequest = this.m_pendingRequest;
        this.m_pendingRequest = new ClientRequest(this);
        var batchBodyResult;
        try {
            batchBodyResult = batchBody(this._rootObject, this);
        }
        catch (ex) {
            this.m_explicitBatchInProgress = false;
            this.m_pendingRequest = previousRequest;
            return Core.CoreUtility._createPromiseFromException(ex);
        }
        var request;
        var batchBodyResultPromise;
        if (typeof batchBodyResult === 'object' && batchBodyResult && typeof batchBodyResult.then === 'function') {
            batchBodyResultPromise = Utility._createPromiseFromResult(null)
                .then(function () {
                return batchBodyResult;
            })
                .then(function (result) {
                _this.m_explicitBatchInProgress = false;
                request = _this.m_pendingRequest;
                _this.m_pendingRequest = previousRequest;
                return result;
            })
                .catch(function (ex) {
                _this.m_explicitBatchInProgress = false;
                request = _this.m_pendingRequest;
                _this.m_pendingRequest = previousRequest;
                return Core.CoreUtility._createPromiseFromException(ex);
            });
        }
        else {
            this.m_explicitBatchInProgress = false;
            request = this.m_pendingRequest;
            this.m_pendingRequest = previousRequest;
            batchBodyResultPromise = Utility._createPromiseFromResult(batchBodyResult);
        }
        return batchBodyResultPromise.then(function (result) {
            return _this.ensureRequestUrlAndHeaderInfo()
                .then(function () {
                return _this.syncPrivate(request);
            })
                .then(function () {
                return result;
            });
        });
    };
    /** Runs a (possibly multi-async-call) batch task, cleaning up tracked objects at the end
     *  This version of the processor, unlike _runBatch, will always create a new request context.
     *  It is therefor advised that all clients switch to the more powerful _runBatch, but
     *  leaving this one in for now due to branching RIs/FIs and timing delays.
     */
    ClientRequestContext._run = function (ctxInitializer, runBody, numCleanupAttempts, retryDelay, onCleanupSuccess, onCleanupFailure) {
        if (numCleanupAttempts === void 0) { numCleanupAttempts = 3; }
        if (retryDelay === void 0) { retryDelay = 5000; }
        return ClientRequestContext._runCommon('run', null, ctxInitializer, 0 /* implicit */, runBody, numCleanupAttempts, retryDelay, null, onCleanupSuccess, onCleanupFailure);
    };
    ClientRequestContext.isValidRequestInfo = function (value) {
        return (typeof value === 'string' ||
            ClientRequestContext.isRequestUrlAndHeaderInfo(value) ||
            ClientRequestContext.isRequestUrlAndHeaderInfoResolver(value));
    };
    ClientRequestContext.isRequestUrlAndHeaderInfo = function (value) {
        return (typeof value === 'object' &&
            value !== null &&
            Object.getPrototypeOf(value) === Object.getPrototypeOf({}) &&
            !Utility.isNullOrUndefined(value.url));
    };
    ClientRequestContext.isRequestUrlAndHeaderInfoResolver = function (value) {
        return typeof value === 'object' && value !== null && typeof value._resolveRequestUrlAndHeaderInfo === 'function';
    };
    /**  Runs a (possibly multi-async-call) batch task, cleaning up tracked objects at the end
     *   Note that this "_runBatch" method is more powerful than "_run", in that it allows users to pass in any of the following:
     *      .run(batch)
     *      .run(requestUrlAndHeaders, batch)
     *      .run(clientObj | contextObj, batch)
     *      .run(requestUrlAndHeaders, clientObj, batch)
     *      .run(clientObj[], batch)
     *      .run(requestUrlAndHeaders, clientObj[], batch)
     *      .run(options, batch)
     *  To avoid each client having to copy-paste very similar code, this "_runBatch" version takes the received "arguments"
     *  object from the caller and does the appropriate validation / extraction, before calling in to the ._runCommon method.
     */
    ClientRequestContext._runBatch = function (functionName, receivedRunArgs, ctxInitializer, onBeforeRun, numCleanupAttempts, retryDelay, onCleanupSuccess, onCleanupFailure) {
        if (numCleanupAttempts === void 0) { numCleanupAttempts = 3; }
        if (retryDelay === void 0) { retryDelay = 5000; }
        return ClientRequestContext._runBatchCommon(0 /* implicit */, functionName, receivedRunArgs, ctxInitializer, numCleanupAttempts, retryDelay, onBeforeRun, onCleanupSuccess, onCleanupFailure);
    };
    ClientRequestContext._runExplicitBatch = function (functionName, receivedRunArgs, ctxInitializer, onBeforeRun, numCleanupAttempts, retryDelay, onCleanupSuccess, onCleanupFailure) {
        if (numCleanupAttempts === void 0) { numCleanupAttempts = 3; }
        if (retryDelay === void 0) { retryDelay = 5000; }
        return ClientRequestContext._runBatchCommon(1 /* explicit */, functionName, receivedRunArgs, ctxInitializer, numCleanupAttempts, retryDelay, onBeforeRun, onCleanupSuccess, onCleanupFailure);
    };
    ClientRequestContext._runBatchCommon = function (batchMode, functionName, receivedRunArgs, ctxInitializer, numCleanupAttempts, retryDelay, onBeforeRun, onCleanupSuccess, onCleanupFailure) {
        if (numCleanupAttempts === void 0) { numCleanupAttempts = 3; }
        if (retryDelay === void 0) { retryDelay = 5000; }
        var ctxRetriever;
        var batch;
        var requestInfo = null;
        var previousObjects = null;
        var argOffset = 0;
        var options = null;
        // If options is specified as the first parameter, all arguments should be properties of RunOptions object
        // If options is not specified, the first parameter can be session, and there are also overloads for previousObjects and batch
        if (receivedRunArgs.length > 0) {
            if (ClientRequestContext.isValidRequestInfo(receivedRunArgs[0])) {
                requestInfo = receivedRunArgs[0];
                argOffset = 1;
            }
            else if (Utility.isPlainJsonObject(receivedRunArgs[0])) {
                options = receivedRunArgs[0];
                requestInfo = options.session;
                if (requestInfo != null && !ClientRequestContext.isValidRequestInfo(requestInfo)) {
                    return ClientRequestContext.createErrorPromise(functionName);
                }
                previousObjects = options.previousObjects;
                argOffset = 1;
            }
        }
        if (receivedRunArgs.length == argOffset + 1) {
            batch = receivedRunArgs[argOffset + 0];
        }
        else if (options == null && receivedRunArgs.length == argOffset + 2) {
            previousObjects = receivedRunArgs[argOffset + 0];
            batch = receivedRunArgs[argOffset + 1];
        }
        else {
            /* More argument than needed, or not enough */
            return ClientRequestContext.createErrorPromise(functionName);
        }
        if (previousObjects != null) {
            // Try to extract the context object out of previousObjects passed in or fail
            if (previousObjects instanceof ClientObject) {
                ctxRetriever = function () { return previousObjects.context; };
            }
            else if (previousObjects instanceof ClientRequestContext) {
                ctxRetriever = function () { return previousObjects; };
            }
            else if (Array.isArray(previousObjects)) {
                var array = previousObjects;
                if (array.length == 0) {
                    return ClientRequestContext.createErrorPromise(functionName);
                }
                for (var i = 0; i < array.length; i++) {
                    if (!(array[i] instanceof ClientObject)) {
                        return ClientRequestContext.createErrorPromise(functionName);
                    }
                    if (array[i].context != array[0].context) {
                        return ClientRequestContext.createErrorPromise(functionName, ResourceStrings.invalidRequestContext);
                    }
                }
                ctxRetriever = function () { return array[0].context; };
            }
            else {
                return ClientRequestContext.createErrorPromise(functionName);
            }
        }
        else {
            ctxRetriever = ctxInitializer;
        }
        var onBeforeRunWithOptions = null;
        if (onBeforeRun) {
            onBeforeRunWithOptions = function (context) { return onBeforeRun(options || {}, context); };
        }
        return ClientRequestContext._runCommon(functionName, requestInfo, ctxRetriever, batchMode, batch, numCleanupAttempts, retryDelay, onBeforeRunWithOptions, onCleanupSuccess, onCleanupFailure);
    };
    ClientRequestContext.createErrorPromise = function (functionName, code) {
        if (code === void 0) { code = Core.CoreResourceStrings.invalidArgument; }
        return Core.CoreUtility._createPromiseFromException(Utility.createRuntimeError(code, Core.CoreUtility._getResourceString(code), functionName /*errorLocation*/));
    };
    ClientRequestContext._runCommon = function (functionName, requestInfo, ctxRetriever, batchMode, runBody, numCleanupAttempts, retryDelay, onBeforeRun, onCleanupSuccess, onCleanupFailure) {
        if (Core.SessionBase._overrideSession) {
            requestInfo = Core.SessionBase._overrideSession;
        }
        // Create an empty promise that starts the whole chain (making sure that if anything
        // throws within the batch, even before the promise, it will still be caught):
        var starterPromise = Core.CoreUtility.createPromise(function (resolve, reject) {
            resolve();
        });
        var ctx;
        var succeeded = false;
        var resultOrError;
        var previousBatchMode;
        return starterPromise
            .then(function () {
            ctx = ctxRetriever(requestInfo); // Which might either return a brand-new context or an existing one.
            // If already in _autoCleanup mode, then another .run is already running. If so, begin waiting
            // on the other promise, which will in turn notify this promise once it's at the front of the queue.
            // Note that need to do this via a call-back-like notification mechanism rather than chaining
            // the promises, because want to have each of runs to run completely independenly of each other
            // (don't want to have them pipe into a single .then or .catch, intentionally breaking them apart.
            if (ctx._autoCleanup) {
                return new Promise(function (resolve, reject) {
                    ctx._onRunFinishedNotifiers.push(function () {
                        // Context is ready for action.  Set its _autoCleanup to true, so that other pending
                        // .run's don't steal it, and then resolve the promise to resume the paused execution.
                        ctx._autoCleanup = true;
                        resolve();
                    });
                });
            }
            else {
                ctx._autoCleanup = true;
            }
        })
            .then(function () {
            if (typeof runBody !== 'function') {
                return ClientRequestContext.createErrorPromise(functionName);
            }
            previousBatchMode = ctx.m_batchMode;
            ctx.m_batchMode = batchMode;
            // Give the host specific implementation a chance to adjust ClientRequestContext's state
            if (onBeforeRun) {
                onBeforeRun(ctx);
            }
            var runBodyResult;
            if (batchMode == 1 /* explicit */) {
                runBodyResult = runBody(ctx.batch.bind(ctx));
            }
            else {
                runBodyResult = runBody(ctx);
            }
            // Ensure that the developer didn't forget the "return" statement within their code:
            if (Utility.isNullOrUndefined(runBodyResult) || typeof runBodyResult.then !== 'function') {
                Utility.throwError(ResourceStrings.runMustReturnPromise);
            }
            return runBodyResult;
        })
            .then(function (runBodyResult) {
            if (batchMode === 1 /* explicit */) {
                return runBodyResult;
            }
            else {
                // Do a final flush of the command queue.
                // This both helps a developer who forgot to flush, (ensuring that any failed operations
                // get bubbled up to any following ".catch"), and also makes sure that the
                // tracked-object cleanup further down starts with a clean request queue.
                return ctx.sync(runBodyResult);
            }
        })
            .then(function (result) {
            succeeded = true;
            resultOrError = result;
        })
            .catch(function (error) {
            resultOrError = error;
        })
            .then(function () {
            var itemsToRemove = ctx.trackedObjects._retrieveAndClearAutoCleanupList();
            ctx._autoCleanup = false;
            ctx.m_batchMode = previousBatchMode;
            for (var key in itemsToRemove) {
                itemsToRemove[key]._objectPath.isValid = false;
            }
            // Note: explicitly not waiting on result of the cleanup, because there's nothing that the
            // developer can do about a failure. But do try to clean up, and re-try a few times on failure:
            var cleanupCounter = 0;
            if (Utility._synchronousCleanup || ClientRequestContext.isRequestUrlAndHeaderInfoResolver(requestInfo)) {
                return attemptCleanup();
            }
            else {
                attemptCleanup();
            }
            function attemptCleanup() {
                cleanupCounter++;
                var savedPendingRequest = ctx.m_pendingRequest;
                var savedBatchMode = ctx.m_batchMode;
                // always create new request to avoid messing up the existing pending request
                // as the attemptCleanup will be called from timeout.
                var request = new ClientRequest(ctx);
                ctx.m_pendingRequest = request;
                ctx.m_batchMode = 0 /* implicit */;
                try {
                    for (var key in itemsToRemove) {
                        ctx.trackedObjects.remove(itemsToRemove[key]);
                    }
                }
                finally {
                    ctx.m_batchMode = savedBatchMode;
                    ctx.m_pendingRequest = savedPendingRequest;
                }
                return ctx
                    .syncPrivate(request)
                    .then(function () {
                    if (onCleanupSuccess) {
                        onCleanupSuccess(cleanupCounter);
                    }
                })
                    .catch(function () {
                    if (onCleanupFailure) {
                        onCleanupFailure(cleanupCounter);
                    }
                    if (cleanupCounter < numCleanupAttempts) {
                        setTimeout(function () {
                            attemptCleanup();
                        }, retryDelay);
                    }
                });
            }
        })
            .then(function () {
            if (ctx._onRunFinishedNotifiers && ctx._onRunFinishedNotifiers.length > 0) {
                var func = ctx._onRunFinishedNotifiers.shift();
                func(); // Note, definitely not awaiting it, merely notifying and moving on to return result.
            }
            if (succeeded) {
                return resultOrError;
            }
            else {
                throw resultOrError;
            }
        });
    };
    return ClientRequestContext;
}(Common.ClientRequestContextBase));
exports.ClientRequestContext = ClientRequestContext;
var RetrieveResultImpl = /** @class */ (function () {
    function RetrieveResultImpl(m_proxy, m_shouldPolyfill) {
        this.m_proxy = m_proxy;
        this.m_shouldPolyfill = m_shouldPolyfill;
        var scalarPropertyNames = m_proxy[Constants.scalarPropertyNames];
        var navigationPropertyNames = m_proxy[Constants.navigationPropertyNames];
        var typeName = m_proxy[Constants.className];
        var isCollection = m_proxy[Constants.isCollection];
        if (scalarPropertyNames) {
            // mark all of the scalar properties to throw exception
            for (var i = 0; i < scalarPropertyNames.length; i++) {
                Utility.definePropertyThrowUnloadedException(this, typeName, scalarPropertyNames[i]);
            }
        }
        if (navigationPropertyNames) {
            // mark all of the navigation propertis to throw exception
            for (var i = 0; i < navigationPropertyNames.length; i++) {
                Utility.definePropertyThrowUnloadedException(this, typeName, navigationPropertyNames[i]);
            }
        }
        // mark the "items" property to throw exception
        if (isCollection) {
            Utility.definePropertyThrowUnloadedException(this, typeName, Constants.itemsLowerCase);
        }
    }
    Object.defineProperty(RetrieveResultImpl.prototype, "$proxy", {
        get: function () {
            return this.m_proxy;
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(RetrieveResultImpl.prototype, "$isNullObject", {
        get: function () {
            if (!this.m_isLoaded) {
                throw new Core._Internal.RuntimeError({
                    code: ErrorCodes.valueNotLoaded,
                    message: Core.CoreUtility._getResourceString(ResourceStrings.valueNotLoaded),
                    debugInfo: {
                        errorLocation: 'retrieveResult.$isNullObject'
                    }
                });
            }
            return this.m_isNullObject;
        },
        enumerable: true,
        configurable: true
    });
    RetrieveResultImpl.prototype.toJSON = function () {
        if (!this.m_isLoaded) {
            return undefined;
        }
        if (this.m_isNullObject) {
            return null;
        }
        if (Utility.isUndefined(this.m_json)) {
            this.m_json = this.purifyJson(this.m_value);
        }
        return this.m_json;
    };
    RetrieveResultImpl.prototype.toString = function () {
        return JSON.stringify(this.toJSON());
    };
    RetrieveResultImpl.prototype._handleResult = function (value) {
        this.m_isLoaded = true;
        if (value === null || (typeof value === 'object' && value && value._IsNull)) {
            this.m_isNullObject = true;
            value = null;
        }
        else {
            this.m_isNullObject = false;
        }
        if (this.m_shouldPolyfill) {
            value = this.changePropertyNameToCamelLowerCase(value);
        }
        this.m_value = value;
        this.m_proxy._handleRetrieveResult(value, this);
    };
    RetrieveResultImpl.prototype.changePropertyNameToCamelLowerCase = function (value) {
        var charCodeUnderscore = 95; // the charCode of '_'
        if (Array.isArray(value)) {
            var ret = [];
            for (var i = 0; i < value.length; i++) {
                ret.push(this.changePropertyNameToCamelLowerCase(value[i]));
            }
            return ret;
        }
        else if (typeof value === 'object' && value !== null) {
            var ret = {};
            for (var key in value) {
                var propValue = value[key];
                if (key === Constants.items) {
                    // it's a collection and we only need to have "items" property.
                    ret = {};
                    ret[Constants.itemsLowerCase] = this.changePropertyNameToCamelLowerCase(propValue);
                    break;
                }
                else {
                    var propName = Utility._toCamelLowerCase(key);
                    ret[propName] = this.changePropertyNameToCamelLowerCase(propValue);
                }
            }
            return ret;
        }
        else {
            return value;
        }
    };
    RetrieveResultImpl.prototype.purifyJson = function (value) {
        var charCodeUnderscore = 95; // the charCode of '_'
        if (Array.isArray(value)) {
            var ret = [];
            for (var i = 0; i < value.length; i++) {
                ret.push(this.purifyJson(value[i]));
            }
            return ret;
        }
        else if (typeof value === 'object' && value !== null) {
            var ret = {};
            for (var key in value) {
                if (key.charCodeAt(0) !== charCodeUnderscore) {
                    var propValue = value[key];
                    if (typeof propValue === 'object' && propValue !== null && Array.isArray(propValue['items'])) {
                        propValue = propValue['items'];
                    }
                    ret[key] = this.purifyJson(propValue);
                }
            }
            return ret;
        }
        else {
            return value;
        }
    };
    return RetrieveResultImpl;
}());
var Constants = /** @class */ (function (_super) {
    __extends(Constants, _super);
    function Constants() {
        return _super !== null && _super.apply(this, arguments) || this;
    }
    Constants.getItemAt = 'GetItemAt';
    Constants.index = '_Index';
    Constants.items = '_Items';
    Constants.iterativeExecutor = 'IterativeExecutor';
    Constants.isTracked = '_IsTracked';
    // The message category for event. Keep it in sync with Constants.EventMessageCategory in
    // %SRCROOT%\richapi\codegen\attributes\attributes.cs
    Constants.eventMessageCategory = 65536;
    Constants.eventWorkbookId = 'Workbook';
    // The remote value of event source enum. Keep it in sync with EventSource.Remote in
    // %SRCROOT%\xlshared\src\Api\Metadata\Current\ExcelApi.cs
    Constants.eventSourceRemote = 'Remote';
    Constants.itemsLowerCase = 'items';
    Constants.proxy = '$proxy';
    Constants.scalarPropertyNames = '_scalarPropertyNames';
    Constants.navigationPropertyNames = '_navigationPropertyNames';
    Constants.className = '_className';
    Constants.isCollection = '_isCollection';
    Constants.scalarPropertyUpdateable = '_scalarPropertyUpdateable';
    Constants.collectionPropertyPath = '_collectionPropertyPath';
    Constants.objectPathInfoDoNotKeepReferenceFieldName = 'D';
    return Constants;
}(Common.CommonConstants));
exports.Constants = Constants;
var ClientRequest = /** @class */ (function (_super) {
    __extends(ClientRequest, _super);
    function ClientRequest(context) {
        var _this = _super.call(this, context) || this;
        _this.m_context = context;
        _this.m_pendingProcessEventHandlers = [];
        _this.m_pendingEventHandlerActions = {};
        _this.m_traceInfos = {};
        _this.m_responseTraceIds = {};
        _this.m_responseTraceMessages = [];
        return _this;
    }
    Object.defineProperty(ClientRequest.prototype, "traceInfos", {
        get: function () {
            return this.m_traceInfos;
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(ClientRequest.prototype, "_responseTraceMessages", {
        get: function () {
            return this.m_responseTraceMessages;
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(ClientRequest.prototype, "_responseTraceIds", {
        get: function () {
            return this.m_responseTraceIds;
        },
        enumerable: true,
        configurable: true
    });
    ClientRequest.prototype._setResponseTraceIds = function (value) {
        if (value) {
            for (var i = 0; i < value.length; i++) {
                var traceId = value[i];
                this.m_responseTraceIds[traceId] = traceId;
                var message = this.m_traceInfos[traceId];
                if (!Core.CoreUtility.isNullOrUndefined(message)) {
                    this.m_responseTraceMessages.push(message);
                }
            }
        }
    };
    ClientRequest.prototype.addTrace = function (actionId, message) {
        this.m_traceInfos[actionId] = message;
    };
    // add eventHandler action.
    ClientRequest.prototype._addPendingEventHandlerAction = function (eventHandlers, action) {
        if (!this.m_pendingEventHandlerActions[eventHandlers._id]) {
            this.m_pendingEventHandlerActions[eventHandlers._id] = [];
            this.m_pendingProcessEventHandlers.push(eventHandlers);
        }
        this.m_pendingEventHandlerActions[eventHandlers._id].push(action);
    };
    Object.defineProperty(ClientRequest.prototype, "_pendingProcessEventHandlers", {
        get: function () {
            return this.m_pendingProcessEventHandlers;
        },
        enumerable: true,
        configurable: true
    });
    ClientRequest.prototype._getPendingEventHandlerActions = function (eventHandlers) {
        return this.m_pendingEventHandlerActions[eventHandlers._id];
    };
    return ClientRequest;
}(Common.ClientRequestBase));
exports.ClientRequest = ClientRequest;
var EventHandlers = /** @class */ (function () {
    function EventHandlers(context, parentObject, name, eventInfo) {
        var _this = this;
        this.m_id = context._nextId();
        this.m_context = context;
        this.m_name = name;
        this.m_handlers = [];
        this.m_registered = false;
        this.m_eventInfo = eventInfo;
        this.m_callback = function (args) {
            _this.m_eventInfo.eventArgsTransformFunc(args).then(function (newArgs) { return _this.fireEvent(newArgs); });
        };
    }
    Object.defineProperty(EventHandlers.prototype, "_registered", {
        get: function () {
            return this.m_registered;
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(EventHandlers.prototype, "_id", {
        get: function () {
            return this.m_id;
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(EventHandlers.prototype, "_handlers", {
        get: function () {
            return this.m_handlers;
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(EventHandlers.prototype, "_context", {
        get: function () {
            return this.m_context;
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(EventHandlers.prototype, "_callback", {
        get: function () {
            return this.m_callback;
        },
        enumerable: true,
        configurable: true
    });
    EventHandlers.prototype.add = function (handler) {
        var action = ActionFactory.createTraceAction(this.m_context, null /*message*/, false /*addTraceMessage*/);
        this.m_context._pendingRequest._addPendingEventHandlerAction(this, {
            id: action.actionInfo.Id,
            handler: handler,
            operation: 0 /* add */
        });
        return new EventHandlerResult(this.m_context, this, handler);
    };
    EventHandlers.prototype.remove = function (handler) {
        var action = ActionFactory.createTraceAction(this.m_context, null /*message*/, false /*addTraceMessage*/);
        this.m_context._pendingRequest._addPendingEventHandlerAction(this, {
            id: action.actionInfo.Id,
            handler: handler,
            operation: 1 /* remove */
        });
    };
    EventHandlers.prototype.removeAll = function () {
        var action = ActionFactory.createTraceAction(this.m_context, null /*message*/, false /*addTraceMessage*/);
        this.m_context._pendingRequest._addPendingEventHandlerAction(this, {
            id: action.actionInfo.Id,
            handler: null,
            operation: 2 /* removeAll */
        });
    };
    EventHandlers.prototype._processRegistration = function (req) {
        var _this = this;
        var ret = Core.CoreUtility._createPromiseFromResult(null);
        var actions = req._getPendingEventHandlerActions(this);
        if (!actions) {
            return ret;
        }
        var handlersResult = [];
        for (var i = 0; i < this.m_handlers.length; i++) {
            handlersResult.push(this.m_handlers[i]);
        }
        var hasChange = false;
        for (var i = 0; i < actions.length; i++) {
            // Check whether the action id is in the response's trace id.
            // If it is in the response strace id, it means that the corresponding
            // action is invoked
            if (req._responseTraceIds[actions[i].id]) {
                hasChange = true;
                switch (actions[i].operation) {
                    case 0 /* add */:
                        handlersResult.push(actions[i].handler);
                        break;
                    case 1 /* remove */:
                        for (var index = handlersResult.length - 1; index >= 0; index--) {
                            if (handlersResult[index] === actions[i].handler) {
                                handlersResult.splice(index, 1);
                                break;
                            }
                        }
                        break;
                    case 2 /* removeAll */:
                        handlersResult = [];
                        break;
                }
            }
        }
        if (hasChange) {
            if (!this.m_registered && handlersResult.length > 0) {
                ret = ret.then(function () { return _this.m_eventInfo.registerFunc(_this.m_callback); }).then(function () { return (_this.m_registered = true); });
            }
            else if (this.m_registered && handlersResult.length == 0) {
                ret = ret
                    .then(function () { return _this.m_eventInfo.unregisterFunc(_this.m_callback); })
                    .catch(function (ex) {
                    // Swallow error when unregister event.
                    Core.CoreUtility.log('Error when unregister event: ' + JSON.stringify(ex));
                })
                    .then(function () { return (_this.m_registered = false); });
            }
            ret = ret.then(function () { return (_this.m_handlers = handlersResult); });
        }
        return ret;
    };
    EventHandlers.prototype.fireEvent = function (args) {
        var promises = [];
        for (var i = 0; i < this.m_handlers.length; i++) {
            var handler = this.m_handlers[i];
            // Because event handler could return null promise, or even throw exception before
            // return promise, we will start with an empty promise.
            var p = Core.CoreUtility._createPromiseFromResult(null)
                .then(this.createFireOneEventHandlerFunc(handler, args))
                .catch(function (ex) {
                // Swallow error when invoke handler.
                Core.CoreUtility.log('Error when invoke handler: ' + JSON.stringify(ex));
            });
            promises.push(p);
        }
        Core.CoreUtility.Promise.all(promises);
    };
    EventHandlers.prototype.createFireOneEventHandlerFunc = function (handler, args) {
        return function () { return handler(args); };
    };
    return EventHandlers;
}());
exports.EventHandlers = EventHandlers;
var EventHandlerResult = /** @class */ (function () {
    function EventHandlerResult(context, handlers, handler) {
        this.m_context = context;
        this.m_allHandlers = handlers;
        this.m_handler = handler;
    }
    Object.defineProperty(EventHandlerResult.prototype, "context", {
        get: function () {
            return this.m_context;
        },
        enumerable: true,
        configurable: true
    });
    EventHandlerResult.prototype.remove = function () {
        if (this.m_allHandlers && this.m_handler) {
            this.m_allHandlers.remove(this.m_handler);
            this.m_allHandlers = null;
            this.m_handler = null;
        }
    };
    return EventHandlerResult;
}());
exports.EventHandlerResult = EventHandlerResult;
var _Internal;
(function (_Internal) {
    var OfficeJsEventRegistration = /** @class */ (function () {
        function OfficeJsEventRegistration() {
        }
        OfficeJsEventRegistration.prototype.register = function (eventId, targetId, handler) {
            switch (eventId) {
                case 4 /* BindingDataChangedEvent */:
                    return Utility.promisify(function (callback) { return Office.context.document.bindings.getByIdAsync(targetId, callback); }).then(function (officeBinding) {
                        return Utility.promisify(function (callback) {
                            return officeBinding.addHandlerAsync(Office.EventType.BindingDataChanged, handler, callback);
                        });
                    });
                case 3 /* BindingSelectionChangedEvent */:
                    return Utility.promisify(function (callback) { return Office.context.document.bindings.getByIdAsync(targetId, callback); }).then(function (officeBinding) {
                        return Utility.promisify(function (callback) {
                            return officeBinding.addHandlerAsync(Office.EventType.BindingSelectionChanged, handler, callback);
                        });
                    });
                case 2 /* DocumentSelectionChangedEvent */:
                    return Utility.promisify(function (callback) {
                        return Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, handler, callback);
                    });
                case 1 /* SettingsChangedEvent */:
                    return Utility.promisify(function (callback) {
                        return Office.context.document.settings.addHandlerAsync(Office.EventType.SettingsChanged, handler, callback);
                    });
                case 5 /* RichApiMessageEvent */:
                    return Utility.promisify(function (callback) {
                        return OSF.DDA.RichApi.richApiMessageManager.addHandlerAsync('richApiMessage', handler, callback);
                    });
                case 13 /* ObjectDeletedEvent */:
                    return Utility.promisify(function (callback) {
                        return Office.context.document.addHandlerAsync(Office.EventType.ObjectDeleted, handler, { id: targetId }, callback);
                    });
                case 14 /* ObjectSelectionChangedEvent */:
                    return Utility.promisify(function (callback) {
                        return Office.context.document.addHandlerAsync(Office.EventType.ObjectSelectionChanged, handler, { id: targetId }, callback);
                    });
                case 15 /* ObjectDataChangedEvent */:
                    return Utility.promisify(function (callback) {
                        return Office.context.document.addHandlerAsync(Office.EventType.ObjectDataChanged, handler, { id: targetId }, callback);
                    });
                case 16 /* ContentControlAddedEvent */:
                    return Utility.promisify(function (callback) {
                        return Office.context.document.addHandlerAsync(Office.EventType.ContentControlAdded, handler, { id: targetId }, callback);
                    });
                default:
                    throw Core._Internal.RuntimeError._createInvalidArgError({ argumentName: 'eventId' });
            }
        };
        OfficeJsEventRegistration.prototype.unregister = function (eventId, targetId, handler) {
            switch (eventId) {
                case 4 /* BindingDataChangedEvent */:
                    return Utility.promisify(function (callback) { return Office.context.document.bindings.getByIdAsync(targetId, callback); }).then(function (officeBinding) {
                        return Utility.promisify(function (callback) {
                            return officeBinding.removeHandlerAsync(Office.EventType.BindingDataChanged, { handler: handler }, callback);
                        });
                    });
                case 3 /* BindingSelectionChangedEvent */:
                    return Utility.promisify(function (callback) { return Office.context.document.bindings.getByIdAsync(targetId, callback); }).then(function (officeBinding) {
                        return Utility.promisify(function (callback) {
                            return officeBinding.removeHandlerAsync(Office.EventType.BindingSelectionChanged, { handler: handler }, callback);
                        });
                    });
                case 2 /* DocumentSelectionChangedEvent */:
                    return Utility.promisify(function (callback) {
                        return Office.context.document.removeHandlerAsync(Office.EventType.DocumentSelectionChanged, { handler: handler }, callback);
                    });
                case 1 /* SettingsChangedEvent */:
                    return Utility.promisify(function (callback) {
                        return Office.context.document.settings.removeHandlerAsync(Office.EventType.SettingsChanged, { handler: handler }, callback);
                    });
                case 5 /* RichApiMessageEvent */:
                    return Utility.promisify(function (callback) {
                        return OSF.DDA.RichApi.richApiMessageManager.removeHandlerAsync('richApiMessage', { handler: handler }, callback);
                    });
                case 13 /* ObjectDeletedEvent */:
                    return Utility.promisify(function (callback) {
                        return Office.context.document.removeHandlerAsync(Office.EventType.ObjectDeleted, { id: targetId, handler: handler }, callback);
                    });
                case 14 /* ObjectSelectionChangedEvent */:
                    return Utility.promisify(function (callback) {
                        return Office.context.document.removeHandlerAsync(Office.EventType.ObjectSelectionChanged, { id: targetId, handler: handler }, callback);
                    });
                case 15 /* ObjectDataChangedEvent */:
                    return Utility.promisify(function (callback) {
                        return Office.context.document.removeHandlerAsync(Office.EventType.ObjectDataChanged, { id: targetId, handler: handler }, callback);
                    });
                case 16 /* ContentControlAddedEvent */:
                    return Utility.promisify(function (callback) {
                        return Office.context.document.removeHandlerAsync(Office.EventType.ContentControlAdded, { id: targetId, handler: handler }, callback);
                    });
                default:
                    throw Core._Internal.RuntimeError._createInvalidArgError({ argumentName: 'eventId' });
            }
        };
        return OfficeJsEventRegistration;
    }());
    /**
     * Event registration using Office.js
     */
    _Internal.officeJsEventRegistration = new OfficeJsEventRegistration();
})(_Internal = exports._Internal || (exports._Internal = {}));
var EventRegistration = /** @class */ (function () {
    function EventRegistration(registerEventImpl, unregisterEventImpl) {
        this.m_handlersByEventByTarget = {};
        this.m_registerEventImpl = registerEventImpl;
        this.m_unregisterEventImpl = unregisterEventImpl;
    }
    EventRegistration.prototype.getHandlers = function (eventId, targetId) {
        if (Utility.isNullOrUndefined(targetId)) {
            targetId = '';
        }
        var handlersById = this.m_handlersByEventByTarget[eventId];
        if (!handlersById) {
            handlersById = {};
            this.m_handlersByEventByTarget[eventId] = handlersById;
        }
        var handlers = handlersById[targetId];
        if (!handlers) {
            handlers = [];
            handlersById[targetId] = handlers;
        }
        return handlers;
    };
    EventRegistration.prototype.register = function (eventId, targetId, handler) {
        if (!handler) {
            throw Core._Internal.RuntimeError._createInvalidArgError({ argumentName: 'handler' });
        }
        var handlers = this.getHandlers(eventId, targetId);
        handlers.push(handler);
        if (handlers.length === 1) {
            return this.m_registerEventImpl(eventId, targetId);
        }
        return Utility._createPromiseFromResult(null);
    };
    EventRegistration.prototype.unregister = function (eventId, targetId, handler) {
        if (!handler) {
            throw Core._Internal.RuntimeError._createInvalidArgError({ argumentName: 'handler' });
        }
        var handlers = this.getHandlers(eventId, targetId);
        for (var index = handlers.length - 1; index >= 0; index--) {
            if (handlers[index] === handler) {
                handlers.splice(index, 1);
                break;
            }
        }
        if (handlers.length === 0) {
            return this.m_unregisterEventImpl(eventId, targetId);
        }
        return Utility._createPromiseFromResult(null);
    };
    return EventRegistration;
}());
exports.EventRegistration = EventRegistration;
var GenericEventRegistration = /** @class */ (function () {
    function GenericEventRegistration() {
        this.m_eventRegistration = new EventRegistration(this._registerEventImpl.bind(this), this._unregisterEventImpl.bind(this));
        this.m_richApiMessageHandler = this._handleRichApiMessage.bind(this);
    }
    GenericEventRegistration.prototype.ready = function () {
        var _this = this;
        if (!this.m_ready) {
            if (GenericEventRegistration._testReadyImpl) {
                this.m_ready = GenericEventRegistration._testReadyImpl().then(function () {
                    _this.m_isReady = true;
                });
            }
            else if (Core.HostBridge.instance) {
                this.m_ready = Utility._createPromiseFromResult(null).then(function () {
                    _this.m_isReady = true;
                });
            }
            else {
                this.m_ready = _Internal.officeJsEventRegistration
                    .register(5 /* RichApiMessageEvent */, '', this.m_richApiMessageHandler)
                    .then(function () {
                    _this.m_isReady = true;
                });
            }
        }
        return this.m_ready;
    };
    Object.defineProperty(GenericEventRegistration.prototype, "isReady", {
        get: function () {
            return this.m_isReady;
        },
        enumerable: true,
        configurable: true
    });
    GenericEventRegistration.prototype.register = function (eventId, targetId, handler) {
        var _this = this;
        return this.ready().then(function () { return _this.m_eventRegistration.register(eventId, targetId, handler); });
    };
    GenericEventRegistration.prototype.unregister = function (eventId, targetId, handler) {
        var _this = this;
        return this.ready().then(function () { return _this.m_eventRegistration.unregister(eventId, targetId, handler); });
    };
    GenericEventRegistration.prototype._registerEventImpl = function (eventId, targetId) {
        return Utility._createPromiseFromResult(null);
    };
    GenericEventRegistration.prototype._unregisterEventImpl = function (eventId, targetId) {
        return Utility._createPromiseFromResult(null);
    };
    GenericEventRegistration.prototype._handleRichApiMessage = function (msg) {
        if (msg && msg.entries) {
            for (var entryIndex = 0; entryIndex < msg.entries.length; entryIndex++) {
                var entry = msg.entries[entryIndex];
                if (entry.messageCategory == Constants.eventMessageCategory) {
                    if (Core.CoreUtility._logEnabled) {
                        Core.CoreUtility.log(JSON.stringify(entry));
                    }
                    var funcs = this.m_eventRegistration.getHandlers(entry.messageType, entry.targetId);
                    if (funcs.length > 0) {
                        var arg = JSON.parse(entry.message);
                        // Now this is a Rich API event, so let's see whether we need to override the source field
                        if (entry.isRemoteOverride) {
                            arg.source = Constants.eventSourceRemote;
                        }
                        for (var i = 0; i < funcs.length; i++) {
                            funcs[i](arg);
                        }
                    }
                }
            }
        }
    };
    GenericEventRegistration.getGenericEventRegistration = function () {
        if (!GenericEventRegistration.s_genericEventRegistration) {
            GenericEventRegistration.s_genericEventRegistration = new GenericEventRegistration();
        }
        return GenericEventRegistration.s_genericEventRegistration;
    };
    GenericEventRegistration.richApiMessageEventCategory = 65536;
    return GenericEventRegistration;
}());
function _testSetRichApiMessageReadyImpl(impl) {
    GenericEventRegistration._testReadyImpl = impl;
}
exports._testSetRichApiMessageReadyImpl = _testSetRichApiMessageReadyImpl;
function _testTriggerRichApiMessageEvent(msg) {
    GenericEventRegistration.getGenericEventRegistration()._handleRichApiMessage(msg);
}
exports._testTriggerRichApiMessageEvent = _testTriggerRichApiMessageEvent;
var GenericEventHandlers = /** @class */ (function (_super) {
    __extends(GenericEventHandlers, _super);
    function GenericEventHandlers(context, parentObject, name, eventInfo) {
        var _this = _super.call(this, context, parentObject, name, eventInfo) || this;
        _this.m_genericEventInfo = eventInfo;
        return _this;
    }
    GenericEventHandlers.prototype.add = function (handler) {
        var _this = this;
        if (this._handlers.length == 0 && this.m_genericEventInfo.registerFunc) {
            this.m_genericEventInfo.registerFunc();
        }
        if (!GenericEventRegistration.getGenericEventRegistration().isReady) {
            this._context._pendingRequest._addPreSyncPromise(GenericEventRegistration.getGenericEventRegistration().ready());
        }
        ActionFactory.createTraceMarkerForCallback(this._context, function () {
            _this._handlers.push(handler);
            if (_this._handlers.length == 1) {
                GenericEventRegistration.getGenericEventRegistration().register(_this.m_genericEventInfo.eventType, _this.m_genericEventInfo.getTargetIdFunc(), _this._callback);
            }
        });
        return new EventHandlerResult(this._context, this, handler);
    };
    GenericEventHandlers.prototype.remove = function (handler) {
        var _this = this;
        if (this._handlers.length == 1 && this.m_genericEventInfo.unregisterFunc) {
            this.m_genericEventInfo.unregisterFunc();
        }
        ActionFactory.createTraceMarkerForCallback(this._context, function () {
            var handlers = _this._handlers;
            for (var index = handlers.length - 1; index >= 0; index--) {
                if (handlers[index] === handler) {
                    handlers.splice(index, 1);
                    break;
                }
            }
            if (handlers.length == 0) {
                GenericEventRegistration.getGenericEventRegistration().unregister(_this.m_genericEventInfo.eventType, _this.m_genericEventInfo.getTargetIdFunc(), _this._callback);
            }
        });
    };
    GenericEventHandlers.prototype.removeAll = function () { };
    return GenericEventHandlers;
}(EventHandlers));
exports.GenericEventHandlers = GenericEventHandlers;
var InstantiateActionResultHandler = /** @class */ (function () {
    function InstantiateActionResultHandler(clientObject) {
        this.m_clientObject = clientObject;
    }
    InstantiateActionResultHandler.prototype._handleResult = function (value) {
        this.m_clientObject._handleIdResult(value);
    };
    return InstantiateActionResultHandler;
}());
var ObjectPathFactory = /** @class */ (function () {
    function ObjectPathFactory() {
    }
    ObjectPathFactory.createGlobalObjectObjectPath = function (context) {
        var objectPathInfo = {
            Id: context._nextId(),
            ObjectPathType: 1 /* GlobalObject */,
            Name: ''
        };
        return new Common.ObjectPath(objectPathInfo, null, false /*isCollection*/, false /*isInvalidAfterRequest*/, 1 /* Read */, 4 /* concurrent */);
    };
    ObjectPathFactory.createNewObjectObjectPath = function (context, typeName, isCollection, flags) {
        var objectPathInfo = {
            Id: context._nextId(),
            ObjectPathType: 2 /* NewObject */,
            Name: typeName
        };
        var ret = new Common.ObjectPath(objectPathInfo, null, isCollection, false /*isInvalidAfterRequest*/, 1 /* Read */, Utility._fixupApiFlags(flags));
        return ret;
    };
    ObjectPathFactory.createPropertyObjectPath = function (context, parent, propertyName, isCollection, isInvalidAfterRequest, flags) {
        var objectPathInfo = {
            Id: context._nextId(),
            ObjectPathType: 4 /* Property */,
            Name: propertyName,
            ParentObjectPathId: parent._objectPath.objectPathInfo.Id
        };
        var ret = new Common.ObjectPath(objectPathInfo, parent._objectPath, isCollection, isInvalidAfterRequest, 1 /* Read */, Utility._fixupApiFlags(flags));
        return ret;
    };
    ObjectPathFactory.createIndexerObjectPath = function (context, parent, args) {
        var objectPathInfo = {
            Id: context._nextId(),
            ObjectPathType: 5 /* Indexer */,
            Name: '',
            ParentObjectPathId: parent._objectPath.objectPathInfo.Id,
            ArgumentInfo: {}
        };
        objectPathInfo.ArgumentInfo.Arguments = args;
        return new Common.ObjectPath(objectPathInfo, parent._objectPath, false /*isCollection*/, false /*isInvalidAfterRequest*/, 1 /* Read */, 4 /* concurrent */);
    };
    ObjectPathFactory.createIndexerObjectPathUsingParentPath = function (context, parentObjectPath, args) {
        var objectPathInfo = {
            Id: context._nextId(),
            ObjectPathType: 5 /* Indexer */,
            Name: '',
            ParentObjectPathId: parentObjectPath.objectPathInfo.Id,
            ArgumentInfo: {}
        };
        objectPathInfo.ArgumentInfo.Arguments = args;
        return new Common.ObjectPath(objectPathInfo, parentObjectPath, false /*isCollection*/, false /*isInvalidAfterRequest*/, 1 /* Read */, 4 /* concurrent */);
    };
    ObjectPathFactory.createMethodObjectPath = function (context, parent, methodName, operationType, args, isCollection, isInvalidAfterRequest, getByIdMethodName, flags) {
        var objectPathInfo = {
            Id: context._nextId(),
            ObjectPathType: 3 /* Method */,
            Name: methodName,
            ParentObjectPathId: parent._objectPath.objectPathInfo.Id,
            ArgumentInfo: {}
        };
        var argumentObjectPaths = Utility.setMethodArguments(context, objectPathInfo.ArgumentInfo, args);
        var ret = new Common.ObjectPath(objectPathInfo, parent._objectPath, isCollection, isInvalidAfterRequest, operationType, Utility._fixupApiFlags(flags));
        ret.argumentObjectPaths = argumentObjectPaths;
        ret.getByIdMethodName = getByIdMethodName;
        return ret;
    };
    ObjectPathFactory.createReferenceIdObjectPath = function (context, referenceId) {
        var objectPathInfo = {
            Id: context._nextId(),
            ObjectPathType: 6 /* ReferenceId */,
            Name: referenceId,
            ArgumentInfo: {}
        };
        var ret = new Common.ObjectPath(objectPathInfo, null, false, false, 1 /* Read */, 4 /* concurrent */);
        return ret;
    };
    ObjectPathFactory.createChildItemObjectPathUsingIndexerOrGetItemAt = function (hasIndexerMethod, context, parent, childItem, index) {
        var id = Utility.tryGetObjectIdFromLoadOrRetrieveResult(childItem);
        if (hasIndexerMethod && !Utility.isNullOrUndefined(id)) {
            return ObjectPathFactory.createChildItemObjectPathUsingIndexer(context, parent, childItem);
        }
        else {
            return ObjectPathFactory.createChildItemObjectPathUsingGetItemAt(context, parent, childItem, index);
        }
    };
    ObjectPathFactory.createChildItemObjectPathUsingIndexer = function (context, parent, childItem) {
        var id = Utility.tryGetObjectIdFromLoadOrRetrieveResult(childItem);
        var objectPathInfo = (objectPathInfo = {
            Id: context._nextId(),
            ObjectPathType: 5 /* Indexer */,
            Name: '',
            ParentObjectPathId: parent._objectPath.objectPathInfo.Id,
            ArgumentInfo: {}
        });
        objectPathInfo.ArgumentInfo.Arguments = [id];
        return new Common.ObjectPath(objectPathInfo, parent._objectPath, false /*isCollection*/, false /*isInvalidAfterRequest*/, 1 /* Read */, 4 /* concurrent */);
    };
    ObjectPathFactory.createChildItemObjectPathUsingGetItemAt = function (context, parent, childItem, index) {
        var indexFromServer = childItem[Constants.index];
        if (indexFromServer) {
            index = indexFromServer;
        }
        var objectPathInfo = {
            Id: context._nextId(),
            ObjectPathType: 3 /* Method */,
            Name: Constants.getItemAt,
            ParentObjectPathId: parent._objectPath.objectPathInfo.Id,
            ArgumentInfo: {}
        };
        objectPathInfo.ArgumentInfo.Arguments = [index];
        return new Common.ObjectPath(objectPathInfo, parent._objectPath, false /*isCollection*/, false /*isInvalidAfterRequest*/, 1 /* Read */, 4 /* concurrent */);
    };
    return ObjectPathFactory;
}());
exports.ObjectPathFactory = ObjectPathFactory;
var OfficeJsRequestExecutor = /** @class */ (function () {
    function OfficeJsRequestExecutor(context) {
        this.m_context = context;
    }
    OfficeJsRequestExecutor.prototype.executeAsync = function (customData, requestFlags, requestMessage) {
        var _this = this;
        var messageSafearray = Core.RichApiMessageUtility.buildMessageArrayForIRequestExecutor(customData, requestFlags, requestMessage, OfficeJsRequestExecutor.SourceLibHeaderValue);
        return new Promise(function (resolve, reject) {
            OSF.DDA.RichApi.executeRichApiRequestAsync(messageSafearray, function (result) {
                Core.CoreUtility.log('Response:');
                Core.CoreUtility.log(JSON.stringify(result));
                var response;
                if (result.status == 'succeeded') {
                    response = Core.RichApiMessageUtility.buildResponseOnSuccess(Core.RichApiMessageUtility.getResponseBody(result), Core.RichApiMessageUtility.getResponseHeaders(result));
                }
                else {
                    response = Core.RichApiMessageUtility.buildResponseOnError(result.error.code, result.error.message);
                    _this.m_context._processOfficeJsErrorResponse(result.error.code, response);
                }
                resolve(response);
            });
        });
    };
    OfficeJsRequestExecutor.SourceLibHeaderValue = 'officejs';
    return OfficeJsRequestExecutor;
}());
var TrackedObjects = /** @class */ (function () {
    function TrackedObjects(context) {
        // Objects that need to clean up after, if in auto-cleanup mode
        this._autoCleanupList = {};
        this.m_context = context;
    }
    TrackedObjects.prototype.add = function (param) {
        var _this = this;
        if (Array.isArray(param)) {
            param.forEach(function (item) { return _this._addCommon(item, true); });
        }
        else {
            this._addCommon(param, true);
        }
    };
    TrackedObjects.prototype._autoAdd = function (object) {
        this._addCommon(object, false);
        this._autoCleanupList[object._objectPath.objectPathInfo.Id] = object;
    };
    TrackedObjects.prototype._autoTrackIfNecessaryWhenHandleObjectResultValue = function (object, resultValue) {
        // Word's root object Document has _KeepReference() method and _ReferenceId property
        // However, when we need to autoCleanup, the code will call
        //		ctx._rootObject._RemoveReference(referenceId)
        // If we add the Word's root object Document to the autoCleanup list, the object path of the
        // root object Document will be marked as invalid and all of other calls
        //		ctx._rootObject._RemoveReference(referenceId)
        // will fail.
        // That's why we should check
        //		(object !== this.m_context._rootObject)
        // to not auto track the root object.
        var shouldAutoTrack = this.m_context._autoCleanup &&
            !object[Constants.isTracked] &&
            object !== this.m_context._rootObject &&
            resultValue &&
            !Utility.isNullOrEmptyString(resultValue[Constants.referenceId]);
        if (shouldAutoTrack) {
            this._autoCleanupList[object._objectPath.objectPathInfo.Id] = object;
            object[Constants.isTracked] = true;
        }
    };
    TrackedObjects.prototype._addCommon = function (object, isExplicitlyAdded) {
        if (object[Constants.isTracked]) {
            if (isExplicitlyAdded && this.m_context._autoCleanup) {
                // By explicitly calling ctx.references.add when in auto-cleanup mode, developer is indicating that he/she wants
                //  to take the object reference into their own hands -- so we should remove it from our auto-cleanup list.
                delete this._autoCleanupList[object._objectPath.objectPathInfo.Id];
            }
            // Beyond that, nothing else to do (and don't want to re-instantiate, etc). So exit.
            return;
        }
        var referenceId = object[Constants.referenceId];
        var donotKeepReference = object._objectPath.objectPathInfo[Constants.objectPathInfoDoNotKeepReferenceFieldName];
        if (donotKeepReference) {
            throw Utility.createRuntimeError(Core.CoreErrorCodes.generalException, Core.CoreUtility._getResourceString(ResourceStrings.objectIsUntracked), null /*location*/);
        }
        if (Utility.isNullOrEmptyString(referenceId) && object._KeepReference) {
            object._KeepReference();
            ActionFactory.createInstantiateAction(this.m_context, object);
            if (isExplicitlyAdded && this.m_context._autoCleanup) {
                // By explicitly calling ctx.references.add when in auto-cleanup mode, developer is indicating that he/she wants
                //  to take the object reference into their own hands -- so we should remove it from our auto-cleanup list.
                delete this._autoCleanupList[object._objectPath.objectPathInfo.Id];
            }
            object[Constants.isTracked] = true;
        }
    };
    TrackedObjects.prototype.remove = function (param) {
        var _this = this;
        if (Array.isArray(param)) {
            param.forEach(function (item) { return _this._removeCommon(item); });
        }
        else {
            this._removeCommon(param);
        }
    };
    TrackedObjects.prototype._removeCommon = function (object) {
        // mark the object path with DonotKeepReference = true
        object._objectPath.objectPathInfo[Constants.objectPathInfoDoNotKeepReferenceFieldName] = true;
        object.context._pendingRequest._removeKeepReferenceAction(object._objectPath.objectPathInfo.Id);
        var referenceId = object[Constants.referenceId];
        if (!Utility.isNullOrEmptyString(referenceId)) {
            var rootObject = this.m_context._rootObject;
            if (rootObject._RemoveReference) {
                rootObject._RemoveReference(referenceId);
            }
        }
        delete object[Constants.isTracked];
    };
    TrackedObjects.prototype._retrieveAndClearAutoCleanupList = function () {
        var list = this._autoCleanupList;
        this._autoCleanupList = {};
        return list;
    };
    return TrackedObjects;
}());
exports.TrackedObjects = TrackedObjects;
var RequestPrettyPrinter = /** @class */ (function () {
    function RequestPrettyPrinter(globalObjName, referencedObjectPaths, actions, showDispose, removePII) {
        if (!globalObjName) {
            globalObjName = 'root';
        }
        this.m_globalObjName = globalObjName;
        this.m_referencedObjectPaths = referencedObjectPaths;
        this.m_actions = actions;
        this.m_statements = [];
        this.m_variableNameForObjectPathMap = {};
        this.m_variableNameToObjectPathMap = {};
        this.m_declaredObjectPathMap = {};
        this.m_showDispose = showDispose;
        this.m_removePII = removePII;
    }
    RequestPrettyPrinter.prototype.process = function () {
        if (this.m_showDispose) {
            ClientRequest._calculateLastUsedObjectPathIds(this.m_actions);
        }
        for (var i = 0; i < this.m_actions.length; i++) {
            this.processOneAction(this.m_actions[i]);
        }
        return this.m_statements;
    };
    RequestPrettyPrinter.prototype.processForDebugStatementInfo = function (actionIndex) {
        if (this.m_showDispose) {
            ClientRequest._calculateLastUsedObjectPathIds(this.m_actions);
        }
        // The number of statements before the action and the after the action that we want to return.
        var surroundingCount = 5;
        this.m_statements = [];
        var oneStatement = '';
        var statementIndex = -1;
        for (var i = 0; i < this.m_actions.length; i++) {
            this.processOneAction(this.m_actions[i]);
            if (actionIndex == i) {
                statementIndex = this.m_statements.length - 1;
            }
            if (statementIndex >= 0 && this.m_statements.length > statementIndex + surroundingCount + 1) {
                // we got enough statements and we do not need to proceed further.
                break;
            }
        }
        if (statementIndex < 0) {
            return null;
        }
        var startIndex = statementIndex - surroundingCount;
        if (startIndex < 0) {
            startIndex = 0;
        }
        var endIndex = statementIndex + 1 + surroundingCount;
        if (endIndex > this.m_statements.length) {
            endIndex = this.m_statements.length;
        }
        var surroundingStatements = [];
        if (startIndex != 0) {
            surroundingStatements.push('...');
        }
        for (var i_1 = startIndex; i_1 < statementIndex; i_1++) {
            surroundingStatements.push(this.m_statements[i_1]);
        }
        surroundingStatements.push('// >>>>>');
        surroundingStatements.push(this.m_statements[statementIndex]);
        surroundingStatements.push('// <<<<<');
        for (var i_2 = statementIndex + 1; i_2 < endIndex; i_2++) {
            surroundingStatements.push(this.m_statements[i_2]);
        }
        if (endIndex < this.m_statements.length) {
            surroundingStatements.push('...');
        }
        return {
            statement: this.m_statements[statementIndex],
            surroundingStatements: surroundingStatements
        };
    };
    RequestPrettyPrinter.prototype.processOneAction = function (action) {
        var actionInfo = action.actionInfo;
        switch (actionInfo.ActionType) {
            case 1 /* Instantiate */:
                this.processInstantiateAction(action);
                break;
            case 3 /* Method */:
                this.processMethodAction(action);
                break;
            case 2 /* Query */:
                this.processQueryAction(action);
                break;
            case 7 /* QueryAsJson */:
                this.processQueryAsJsonAction(action);
                break;
            case 6 /* RecursiveQuery */:
                this.processRecursiveQueryAction(action);
                break;
            case 4 /* SetProperty */:
                this.processSetPropertyAction(action);
                break;
            case 5 /* Trace */:
                this.processTraceAction(action);
                break;
            case 8 /* EnsureUnchanged */:
                this.processEnsureUnchangedAction(action);
                break;
            case 9 /* Update */:
                this.processUpdateAction(action);
                break;
        }
    };
    RequestPrettyPrinter.prototype.processInstantiateAction = function (action) {
        var objId = action.actionInfo.ObjectPathId;
        var objPath = this.m_referencedObjectPaths[objId];
        var varName = this.getObjVarName(objId);
        // The Instantiate action for one object may showed multiple times within the action list
        // We should only declare one variable for that object.
        if (!this.m_declaredObjectPathMap[objId]) {
            var statement = 'var ' + varName + ' = ' + this.buildObjectPathExpressionWithParent(objPath) + ';';
            statement = this.appendDisposeCommentIfRelevant(statement, action);
            this.m_statements.push(statement);
            this.m_declaredObjectPathMap[objId] = varName;
        }
        else {
            var statement = '// Instantiate {' + varName + '}';
            statement = this.appendDisposeCommentIfRelevant(statement, action);
            this.m_statements.push(statement);
        }
    };
    RequestPrettyPrinter.prototype.processMethodAction = function (action) {
        var methodName = action.actionInfo.Name;
        if (methodName === '_KeepReference') {
            if (!Common._internalConfig.showInternalApiInDebugInfo) {
                return;
            }
            methodName = 'track';
        }
        var statement = this.getObjVarName(action.actionInfo.ObjectPathId) +
            '.' +
            Utility._toCamelLowerCase(methodName) +
            '(' +
            this.buildArgumentsExpression(action.actionInfo.ArgumentInfo) +
            ');';
        statement = this.appendDisposeCommentIfRelevant(statement, action);
        this.m_statements.push(statement);
    };
    RequestPrettyPrinter.prototype.processQueryAction = function (action) {
        var queryExp = this.buildQueryExpression(action);
        var statement = this.getObjVarName(action.actionInfo.ObjectPathId) + '.load(' + queryExp + ');';
        statement = this.appendDisposeCommentIfRelevant(statement, action);
        this.m_statements.push(statement);
    };
    RequestPrettyPrinter.prototype.processQueryAsJsonAction = function (action) {
        var queryExp = this.buildQueryExpression(action);
        var statement = this.getObjVarName(action.actionInfo.ObjectPathId) + '.retrieve(' + queryExp + ');';
        statement = this.appendDisposeCommentIfRelevant(statement, action);
        this.m_statements.push(statement);
    };
    RequestPrettyPrinter.prototype.processRecursiveQueryAction = function (action) {
        var queryExp = '';
        if (action.actionInfo.RecursiveQueryInfo) {
            queryExp = JSON.stringify(action.actionInfo.RecursiveQueryInfo);
        }
        var statement = this.getObjVarName(action.actionInfo.ObjectPathId) + '.loadRecursive(' + queryExp + ');';
        statement = this.appendDisposeCommentIfRelevant(statement, action);
        this.m_statements.push(statement);
    };
    RequestPrettyPrinter.prototype.processSetPropertyAction = function (action) {
        var statement = this.getObjVarName(action.actionInfo.ObjectPathId) +
            '.' +
            Utility._toCamelLowerCase(action.actionInfo.Name) +
            ' = ' +
            this.buildArgumentsExpression(action.actionInfo.ArgumentInfo) +
            ';';
        statement = this.appendDisposeCommentIfRelevant(statement, action);
        this.m_statements.push(statement);
    };
    RequestPrettyPrinter.prototype.processTraceAction = function (action) {
        var statement = 'context.trace();';
        statement = this.appendDisposeCommentIfRelevant(statement, action);
        this.m_statements.push(statement);
    };
    RequestPrettyPrinter.prototype.processEnsureUnchangedAction = function (action) {
        var statement = this.getObjVarName(action.actionInfo.ObjectPathId) +
            '.ensureUnchanged(' +
            JSON.stringify(action.actionInfo.ObjectState) +
            ');';
        statement = this.appendDisposeCommentIfRelevant(statement, action);
        this.m_statements.push(statement);
    };
    RequestPrettyPrinter.prototype.processUpdateAction = function (action) {
        var statement = this.getObjVarName(action.actionInfo.ObjectPathId) +
            '.update(' +
            JSON.stringify(action.actionInfo.ObjectState) +
            ');';
        statement = this.appendDisposeCommentIfRelevant(statement, action);
        this.m_statements.push(statement);
    };
    RequestPrettyPrinter.prototype.appendDisposeCommentIfRelevant = function (statement, action) {
        var _this = this;
        if (this.m_showDispose) {
            // action.actionInfo.LastUsedObjectPathIds
            var lastUsedObjectPathIds = action.actionInfo.L;
            if (lastUsedObjectPathIds && lastUsedObjectPathIds.length > 0) {
                var objectNamesToDispose = lastUsedObjectPathIds.map(function (item) { return _this.getObjVarName(item); }).join(', ');
                return statement + ' // And then dispose {' + objectNamesToDispose + '}';
            }
        }
        return statement;
    };
    RequestPrettyPrinter.prototype.buildQueryExpression = function (action) {
        if (action.actionInfo.QueryInfo) {
            var option = {};
            option.select = action.actionInfo.QueryInfo.Select;
            option.expand = action.actionInfo.QueryInfo.Expand;
            option.skip = action.actionInfo.QueryInfo.Skip;
            option.top = action.actionInfo.QueryInfo.Top;
            if (typeof option.top === 'undefined' &&
                typeof option.skip === 'undefined' &&
                typeof option.expand === 'undefined') {
                if (typeof option.select === 'undefined') {
                    return '';
                }
                else {
                    return JSON.stringify(option.select);
                }
            }
            else {
                return JSON.stringify(option);
            }
        }
        return '';
    };
    RequestPrettyPrinter.prototype.buildObjectPathExpressionWithParent = function (objPath) {
        var hasParent = objPath.objectPathInfo.ObjectPathType == 5 /* Indexer */ ||
            objPath.objectPathInfo.ObjectPathType == 3 /* Method */ ||
            objPath.objectPathInfo.ObjectPathType == 4 /* Property */;
        if (hasParent && objPath.objectPathInfo.ParentObjectPathId) {
            return (this.getObjVarName(objPath.objectPathInfo.ParentObjectPathId) + '.' + this.buildObjectPathExpression(objPath));
        }
        return this.buildObjectPathExpression(objPath);
    };
    RequestPrettyPrinter.prototype.buildObjectPathExpression = function (objPath) {
        var expr = this.buildObjectPathInfoExpression(objPath.objectPathInfo);
        var originalObjectPathInfo = objPath.originalObjectPathInfo;
        if (originalObjectPathInfo) {
            expr = expr + ' /* originally ' + this.buildObjectPathInfoExpression(originalObjectPathInfo) + ' */';
        }
        return expr;
    };
    RequestPrettyPrinter.prototype.buildObjectPathInfoExpression = function (objectPathInfo) {
        switch (objectPathInfo.ObjectPathType) {
            case 1 /* GlobalObject */:
                return 'context.' + this.m_globalObjName;
            case 5 /* Indexer */:
                return 'getItem(' + this.buildArgumentsExpression(objectPathInfo.ArgumentInfo) + ')';
            case 3 /* Method */:
                return (Utility._toCamelLowerCase(objectPathInfo.Name) +
                    '(' +
                    this.buildArgumentsExpression(objectPathInfo.ArgumentInfo) +
                    ')');
            case 2 /* NewObject */:
                return objectPathInfo.Name + '.newObject()';
            case 7 /* NullObject */:
                return 'null';
            case 4 /* Property */:
                return Utility._toCamelLowerCase(objectPathInfo.Name);
            case 6 /* ReferenceId */:
                return ('context.' + this.m_globalObjName + '._getObjectByReferenceId(' + JSON.stringify(objectPathInfo.Name) + ')');
        }
    };
    RequestPrettyPrinter.prototype.buildArgumentsExpression = function (args) {
        var ret = '';
        if (!args.Arguments || args.Arguments.length === 0) {
            return ret;
        }
        if (this.m_removePII) {
            if (typeof args.Arguments[0] === 'undefined') {
                return ret;
            }
            return '...';
        }
        for (var i = 0; i < args.Arguments.length; i++) {
            if (i > 0) {
                ret = ret + ', ';
            }
            ret =
                ret +
                    this.buildArgumentLiteral(args.Arguments[i], args.ReferencedObjectPathIds ? args.ReferencedObjectPathIds[i] : null);
        }
        if (ret === 'undefined') {
            // The argument is optional
            ret = '';
        }
        return ret;
    };
    RequestPrettyPrinter.prototype.buildArgumentLiteral = function (value, objectPathId) {
        if (typeof value == 'number' && value === objectPathId) {
            return this.getObjVarName(objectPathId);
        }
        else {
            return JSON.stringify(value);
        }
    };
    RequestPrettyPrinter.prototype.getObjVarNameBase = function (objectPathId) {
        var ret = 'v';
        var objPath = this.m_referencedObjectPaths[objectPathId];
        if (objPath) {
            switch (objPath.objectPathInfo.ObjectPathType) {
                case 1 /* GlobalObject */:
                    ret = this.m_globalObjName;
                    break;
                case 4 /* Property */:
                    ret = Utility._toCamelLowerCase(objPath.objectPathInfo.Name);
                    break;
                case 3 /* Method */:
                    var methodName = objPath.objectPathInfo.Name;
                    // For a variable corresponding to "getXyz", we would like to name it as "xyz".
                    if (methodName.length > 3 && methodName.substr(0, 3) === 'Get') {
                        methodName = methodName.substr(3);
                    }
                    ret = Utility._toCamelLowerCase(methodName);
                    break;
                case 5 /* Indexer */:
                    var parentName = this.getObjVarNameBase(objPath.objectPathInfo.ParentObjectPathId);
                    if (parentName.charAt(parentName.length - 1) === 's') {
                        ret = parentName.substr(0, parentName.length - 1);
                    }
                    else {
                        ret = parentName + 'Item';
                    }
                    break;
            }
        }
        return ret;
    };
    RequestPrettyPrinter.prototype.getObjVarName = function (objectPathId) {
        if (this.m_variableNameForObjectPathMap[objectPathId]) {
            return this.m_variableNameForObjectPathMap[objectPathId];
        }
        var ret = this.getObjVarNameBase(objectPathId);
        if (!this.m_variableNameToObjectPathMap[ret]) {
            this.m_variableNameForObjectPathMap[objectPathId] = ret;
            this.m_variableNameToObjectPathMap[ret] = objectPathId;
            return ret;
        }
        // the name is already been used. We need to append some integer to
        // get new name.
        var i = 1;
        while (this.m_variableNameToObjectPathMap[ret + i.toString()]) {
            i++;
        }
        ret = ret + i.toString();
        this.m_variableNameForObjectPathMap[objectPathId] = ret;
        this.m_variableNameToObjectPathMap[ret] = objectPathId;
        return ret;
    };
    return RequestPrettyPrinter;
}());
var ResourceStrings = /** @class */ (function (_super) {
    __extends(ResourceStrings, _super);
    function ResourceStrings() {
        return _super !== null && _super.apply(this, arguments) || this;
    }
    // IMPORTANT! Please add the default english resource string value to
    // both ResourceStringValues.ts and
    // %SRCROOT%\osfweb\jscript\office_strings.js.resx
    // Note that in the office_strings.js.resx file, each string will be
    // prefixed with "L_" (e.g., "L_PropertyDoesNotExist")
    ResourceStrings.cannotRegisterEvent = 'CannotRegisterEvent';
    ResourceStrings.connectionFailureWithStatus = 'ConnectionFailureWithStatus';
    ResourceStrings.connectionFailureWithDetails = 'ConnectionFailureWithDetails';
    ResourceStrings.propertyNotLoaded = 'PropertyNotLoaded';
    ResourceStrings.runMustReturnPromise = 'RunMustReturnPromise';
    ResourceStrings.propertyDoesNotExist = 'PropertyDoesNotExist';
    ResourceStrings.attemptingToSetReadOnlyProperty = 'AttemptingToSetReadOnlyProperty';
    ResourceStrings.moreInfoInnerError = 'MoreInfoInnerError';
    ResourceStrings.cannotApplyPropertyThroughSetMethod = 'CannotApplyPropertyThroughSetMethod';
    ResourceStrings.invalidOperationInCellEditMode = 'InvalidOperationInCellEditMode';
    ResourceStrings.objectIsUntracked = 'ObjectIsUntracked';
    ResourceStrings.customFunctionDefintionMissing = 'CustomFunctionDefintionMissing';
    ResourceStrings.customFunctionImplementationMissing = 'CustomFunctionImplementationMissing';
    ResourceStrings.customFunctionNameContainsBadChars = 'CustomFunctionNameContainsBadChars';
    ResourceStrings.customFunctionNameCannotSplit = 'CustomFunctionNameCannotSplit';
    ResourceStrings.customFunctionUnexpectedNumberOfEntriesInResultBatch = 'CustomFunctionUnexpectedNumberOfEntriesInResultBatch';
    ResourceStrings.customFunctionCancellationHandlerMissing = 'CustomFunctionCancellationHandlerMissing';
    ResourceStrings.customFunctionInvalidFunction = 'CustomFunctionInvalidFunction';
    ResourceStrings.customFunctionInvalidFunctionMapping = 'CustomFunctionInvalidFunctionMapping';
    ResourceStrings.customFunctionWindowMissing = 'CustomFunctionWindowMissing';
    ResourceStrings.customFunctionDefintionMissingOnWindow = 'CustomFunctionDefintionMissingOnWindow';
    ResourceStrings.pendingBatchInProgress = 'PendingBatchInProgress';
    ResourceStrings.notInsideBatch = 'NotInsideBatch';
    ResourceStrings.cannotUpdateReadOnlyProperty = 'CannotUpdateReadOnlyProperty';
    return ResourceStrings;
}(Core.CoreResourceStrings));
exports.ResourceStrings = ResourceStrings;
Core.CoreUtility.addResourceStringValues({
    CannotRegisterEvent: 'The event handler cannot be registered.',
    PropertyNotLoaded: "The property '{0}' is not available. Before reading the property's value, call the load method on the containing object and call \"context.sync()\" on the associated request context.",
    RunMustReturnPromise: 'The batch function passed to the ".run" method didn\'t return a promise. The function must return a promise, so that any automatically-tracked objects can be released at the completion of the batch operation. Typically, you return a promise by returning the response from "context.sync()".',
    InvalidOrTimedOutSessionMessage: 'Your Office Online session has expired or is invalid. To continue, refresh the page.',
    InvalidOperationInCellEditMode: 'Excel is in cell-editing mode. Please exit the edit mode by pressing ENTER or TAB or selecting another cell, and then try again.',
    CustomFunctionDefintionMissing: "A property with the name '{0}' that represents the function's definition must exist on Excel.Script.CustomFunctions.",
    CustomFunctionDefintionMissingOnWindow: "A property with the name '{0}' that represents the function's definition must exist on the window object.",
    CustomFunctionImplementationMissing: "The property with the name '{0}' on Excel.Script.CustomFunctions that represents the function's definition must contain a 'call' property that implements the function.",
    CustomFunctionNameContainsBadChars: 'The function name may only contain letters, digits, underscores, and periods.',
    CustomFunctionNameCannotSplit: 'The function name must contain a non-empty namespace and a non-empty short name.',
    CustomFunctionUnexpectedNumberOfEntriesInResultBatch: "The batching function returned a number of results that doesn't match the number of parameter value sets that were passed into it.",
    CustomFunctionCancellationHandlerMissing: 'The cancellation handler onCanceled is missing in the function. The handler must be present as the function is defined as cancelable.',
    CustomFunctionInvalidFunction: "The property with the name '{0}' that represents the function's definition is not a valid function.",
    CustomFunctionInvalidFunctionMapping: "The property with the name '{0}' on CustomFunctionMappings that represents the function's definition is not a valid function.",
    CustomFunctionWindowMissing: 'The window object was not found.',
    PendingBatchInProgress: 'There is a pending batch in progress. The batch method may not be called inside another batch, or simultaneously with another batch.',
    NotInsideBatch: 'Operations may not be invoked outside of a batch method.',
    CannotUpdateReadOnlyProperty: "The property '{0}' is read-only and it cannot be updated.",
    ObjectIsUntracked: 'The object is untracked.'
});
var Utility = /** @class */ (function (_super) {
    __extends(Utility, _super);
    function Utility() {
        return _super !== null && _super.apply(this, arguments) || this;
    }
    Utility.fixObjectPathIfNecessary = function (clientObject, value) {
        if (clientObject && clientObject._objectPath && value) {
            clientObject._objectPath.updateUsingObjectData(value, clientObject);
        }
    };
    Utility.validateObjectPath = function (clientObject) {
        var objectPath = clientObject._objectPath;
        while (objectPath) {
            if (!objectPath.isValid) {
                throw new Core._Internal.RuntimeError({
                    code: ErrorCodes.invalidObjectPath,
                    message: Core.CoreUtility._getResourceString(ResourceStrings.invalidObjectPath, Utility.getObjectPathExpression(objectPath)),
                    debugInfo: {
                        errorLocation: Utility.getObjectPathExpression(objectPath)
                    }
                });
            }
            objectPath = objectPath.parentObjectPath;
        }
    };
    Utility.validateReferencedObjectPaths = function (objectPaths) {
        if (objectPaths) {
            for (var i = 0; i < objectPaths.length; i++) {
                var objectPath = objectPaths[i];
                while (objectPath) {
                    if (!objectPath.isValid) {
                        throw new Core._Internal.RuntimeError({
                            code: ErrorCodes.invalidObjectPath,
                            message: Core.CoreUtility._getResourceString(ResourceStrings.invalidObjectPath, Utility.getObjectPathExpression(objectPath))
                        });
                    }
                    objectPath = objectPath.parentObjectPath;
                }
            }
        }
    };
    Utility.load = function (clientObj, option) {
        clientObj.context.load(clientObj, option);
        return clientObj;
    };
    Utility.loadAndSync = function (clientObj, option) {
        clientObj.context.load(clientObj, option);
        return clientObj.context.sync().then(function () { return clientObj; });
    };
    Utility.retrieve = function (clientObj, option) {
        var shouldPolyfill = Common._internalConfig.alwaysPolyfillClientObjectRetrieveMethod;
        if (!shouldPolyfill) {
            // check whether the host support RichApiRuntime 1.1
            shouldPolyfill = !Utility.isSetSupported('RichApiRuntime', '1.1');
        }
        var result = new RetrieveResultImpl(clientObj, shouldPolyfill);
        var queryOption = ClientRequestContext._parseQueryOption(option);
        var action;
        if (shouldPolyfill) {
            action = ActionFactory.createQueryAction(clientObj.context, clientObj, queryOption);
        }
        else {
            action = ActionFactory.createQueryAsJsonAction(clientObj.context, clientObj, queryOption);
        }
        clientObj.context._pendingRequest.addActionResultHandler(action, result);
        return result;
    };
    Utility.retrieveAndSync = function (clientObj, option) {
        var result = Utility.retrieve(clientObj, option);
        return clientObj.context.sync().then(function () { return result; });
    };
    Utility._parseSelectExpand = function (select) {
        var args = [];
        if (!Core.CoreUtility.isNullOrEmptyString(select)) {
            var propertyNames = select.split(',');
            for (var i = 0; i < propertyNames.length; i++) {
                var propertyName = propertyNames[i];
                propertyName = sanitizeForAnyItemsSlash(propertyName.trim());
                if (propertyName.length > 0) {
                    args.push(propertyName);
                }
            }
        }
        return args;
        /**
         * Because a lot of developers, when loading a collection, have a tendency to call "load" with
         * "items", or "items/name", or "worksheets/items/tables/items/name", sanitize the string "items",
         * or anything with an "items/" prefix or middle "/items/"(s) in input.
         *
         * @param propertyName: a propertyName string that is assumed to have already been trimmed
         */
        function sanitizeForAnyItemsSlash(propertyName) {
            var propertyNameLower = propertyName.toLowerCase();
            if (propertyNameLower === 'items' || propertyNameLower === 'items/') {
                return '*';
            }
            var itemsSlashLength = 6; // length of "items/"
            var isItemsSlashOrItemsDot = propertyNameLower.substr(0, itemsSlashLength) === 'items/' ||
                propertyNameLower.substr(0, itemsSlashLength) === 'items.';
            if (isItemsSlashOrItemsDot) {
                propertyName = propertyName.substr(itemsSlashLength);
            }
            return propertyName.replace(new RegExp('[/.]items[/.]', 'gi'), '/');
            // gi = global & case insensitive
        }
    };
    Utility.toJson = function (clientObj, scalarProperties, navigationProperties, collectionItemsIfAny) {
        var result = {};
        for (var prop in scalarProperties) {
            var value = scalarProperties[prop];
            if (typeof value !== 'undefined') {
                result[prop] = value;
            }
        }
        for (var prop in navigationProperties) {
            var value = navigationProperties[prop];
            if (typeof value !== 'undefined') {
                // For collections, skip a level, and only show the underlying array of items (i.e., toJSON-ing the collection)
                if (value[Utility.fieldName_isCollection] && typeof value[Utility.fieldName_m__items] !== 'undefined') {
                    result[prop] = value.toJSON()['items'];
                }
                else {
                    result[prop] = value.toJSON();
                }
            }
        }
        if (collectionItemsIfAny) {
            result['items'] = collectionItemsIfAny.map(function (item) { return item.toJSON(); });
        }
        return result;
    };
    // TODO OfficeMain #1121836: Remove these methods, throwError & createRuntimeError,
    // as soon as the other clients (Excel and Word) stop using it.
    Utility.throwError = function (resourceId, arg, errorLocation) {
        throw new Core._Internal.RuntimeError({
            code: resourceId,
            message: Core.CoreUtility._getResourceString(resourceId, arg),
            debugInfo: errorLocation ? { errorLocation: errorLocation } : undefined
        });
    };
    Utility.createRuntimeError = function (code, message, location) {
        return new Core._Internal.RuntimeError({
            code: code,
            message: message,
            debugInfo: { errorLocation: location }
        });
    };
    Utility.throwIfNotLoaded = function (propertyName, fieldValue, entityName, isNull) {
        if (!isNull &&
            Core.CoreUtility.isUndefined(fieldValue) &&
            propertyName.charCodeAt(0) != Utility.s_underscoreCharCode) {
            throw Utility.createPropertyNotLoadedException(entityName, propertyName);
        }
    };
    Utility.createPropertyNotLoadedException = function (entityName, propertyName) {
        return new Core._Internal.RuntimeError({
            code: ErrorCodes.propertyNotLoaded,
            message: Core.CoreUtility._getResourceString(ResourceStrings.propertyNotLoaded, propertyName),
            debugInfo: entityName ? { errorLocation: entityName + '.' + propertyName } : undefined
        });
    };
    Utility.createCannotUpdateReadOnlyPropertyException = function (entityName, propertyName) {
        return new Core._Internal.RuntimeError({
            code: ErrorCodes.cannotUpdateReadOnlyProperty,
            message: Core.CoreUtility._getResourceString(ResourceStrings.cannotUpdateReadOnlyProperty, propertyName),
            debugInfo: entityName ? { errorLocation: entityName + '.' + propertyName } : undefined
        });
    };
    Utility.promisify = function (action) {
        return new Promise(function (resolve, reject) {
            var callback = function (result) {
                if (result.status == 'failed') {
                    reject(result.error);
                }
                else {
                    resolve(result.value);
                }
            };
            action(callback);
        });
    };
    Utility._addActionResultHandler = function (clientObj, action, resultHandler) {
        clientObj.context._pendingRequest.addActionResultHandler(action, resultHandler);
    };
    Utility._handleNavigationPropertyResults = function (clientObj, objectValue, propertyNames) {
        for (var i = 0; i < propertyNames.length - 1; i += 2) {
            if (!Core.CoreUtility.isUndefined(objectValue[propertyNames[i + 1]])) {
                clientObj[propertyNames[i]]._handleResult(objectValue[propertyNames[i + 1]]);
            }
        }
    };
    // Please keep the logic same as ToCamelLowerCase() in %SRCROOT%\richapi\codegen\core\codegen.cs
    Utility._toCamelLowerCase = function (name) {
        if (Core.CoreUtility.isNullOrEmptyString(name)) {
            return name;
        }
        var index = 0;
        // 65 is the value for 'A' and 90 is the value for 'Z'
        while (index < name.length && name.charCodeAt(index) >= 65 && name.charCodeAt(index) <= 90) {
            index++;
        }
        if (index < name.length) {
            return name.substr(0, index).toLowerCase() + name.substr(index);
        }
        else {
            return name.toLowerCase();
        }
    };
    Utility._fixupApiFlags = function (flags) {
        // For older generated code, the flags is a boolean, which means isRestrictedResourceAccess.
        // Once everyone use new CodeGen, we will remove this hack.
        if (typeof flags === 'boolean') {
            if (flags) {
                flags = 1 /* restrictedResourceAccess */;
            }
            else {
                flags = 0 /* none */;
            }
        }
        return flags;
    };
    Utility.definePropertyThrowUnloadedException = function (obj, typeName, propertyName) {
        Object.defineProperty(obj, propertyName, {
            configurable: true,
            enumerable: true,
            get: function () {
                throw Utility.createPropertyNotLoadedException(typeName, propertyName);
            },
            set: function () {
                throw Utility.createCannotUpdateReadOnlyPropertyException(typeName, propertyName);
            }
        });
    };
    Utility.defineReadOnlyPropertyWithValue = function (obj, propertyName, value) {
        Object.defineProperty(obj, propertyName, {
            configurable: true,
            enumerable: true,
            get: function () {
                return value;
            },
            set: function () {
                throw Utility.createCannotUpdateReadOnlyPropertyException(null, propertyName);
            }
        });
    };
    Utility.processRetrieveResult = function (proxy, value, result, childItemCreateFunc) {
        if (Core.CoreUtility.isNullOrUndefined(value)) {
            return;
        }
        if (childItemCreateFunc) {
            var data = value[Constants.itemsLowerCase];
            if (Array.isArray(data)) {
                var itemsResult = [];
                for (var i = 0; i < data.length; i++) {
                    var itemProxy = childItemCreateFunc(data[i], i);
                    var itemResult = {};
                    itemResult[Constants.proxy] = itemProxy;
                    itemProxy._handleRetrieveResult(data[i], itemResult);
                    itemsResult.push(itemResult);
                }
                Utility.defineReadOnlyPropertyWithValue(result, Constants.itemsLowerCase, itemsResult);
            }
        }
        else {
            var scalarPropertyNames = proxy[Constants.scalarPropertyNames];
            var navigationPropertyNames = proxy[Constants.navigationPropertyNames];
            var typeName = proxy[Constants.className];
            if (scalarPropertyNames) {
                for (var i = 0; i < scalarPropertyNames.length; i++) {
                    var propName = scalarPropertyNames[i];
                    var propValue = value[propName];
                    if (Core.CoreUtility.isUndefined(propValue)) {
                        Utility.definePropertyThrowUnloadedException(result, typeName, propName);
                    }
                    else {
                        Utility.defineReadOnlyPropertyWithValue(result, propName, propValue);
                    }
                }
            }
            if (navigationPropertyNames) {
                for (var i = 0; i < navigationPropertyNames.length; i++) {
                    var propName = navigationPropertyNames[i];
                    var propValue = value[propName];
                    if (Core.CoreUtility.isUndefined(propValue)) {
                        Utility.definePropertyThrowUnloadedException(result, typeName, propName);
                    }
                    else {
                        var propProxy = proxy[propName];
                        var propResult = {};
                        propProxy._handleRetrieveResult(propValue, propResult);
                        propResult[Constants.proxy] = propProxy;
                        // change
                        // { items: [a, b, c] }
                        // to
                        // [a, b, c]
                        if (Array.isArray(propResult[Constants.itemsLowerCase])) {
                            propResult = propResult[Constants.itemsLowerCase];
                        }
                        Utility.defineReadOnlyPropertyWithValue(result, propName, propResult);
                    }
                }
            }
        }
    };
    Utility.fieldName_m__items = 'm__items';
    Utility.fieldName_isCollection = '_isCollection';
    Utility._synchronousCleanup = false;
    Utility.s_underscoreCharCode = '_'.charCodeAt(0);
    return Utility;
}(Common.CommonUtility));
exports.Utility = Utility;


/***/ }),
/* 6 */
/***/ (function(module, exports) {

(function(self) {
  'use strict';

  if (self.fetch) {
    return
  }

  var support = {
    searchParams: 'URLSearchParams' in self,
    iterable: 'Symbol' in self && 'iterator' in Symbol,
    blob: 'FileReader' in self && 'Blob' in self && (function() {
      try {
        new Blob()
        return true
      } catch(e) {
        return false
      }
    })(),
    formData: 'FormData' in self,
    arrayBuffer: 'ArrayBuffer' in self
  }

  if (support.arrayBuffer) {
    var viewClasses = [
      '[object Int8Array]',
      '[object Uint8Array]',
      '[object Uint8ClampedArray]',
      '[object Int16Array]',
      '[object Uint16Array]',
      '[object Int32Array]',
      '[object Uint32Array]',
      '[object Float32Array]',
      '[object Float64Array]'
    ]

    var isDataView = function(obj) {
      return obj && DataView.prototype.isPrototypeOf(obj)
    }

    var isArrayBufferView = ArrayBuffer.isView || function(obj) {
      return obj && viewClasses.indexOf(Object.prototype.toString.call(obj)) > -1
    }
  }

  function normalizeName(name) {
    if (typeof name !== 'string') {
      name = String(name)
    }
    if (/[^a-z0-9\-#$%&'*+.\^_`|~]/i.test(name)) {
      throw new TypeError('Invalid character in header field name')
    }
    return name.toLowerCase()
  }

  function normalizeValue(value) {
    if (typeof value !== 'string') {
      value = String(value)
    }
    return value
  }

  // Build a destructive iterator for the value list
  function iteratorFor(items) {
    var iterator = {
      next: function() {
        var value = items.shift()
        return {done: value === undefined, value: value}
      }
    }

    if (support.iterable) {
      iterator[Symbol.iterator] = function() {
        return iterator
      }
    }

    return iterator
  }

  function Headers(headers) {
    this.map = {}

    if (headers instanceof Headers) {
      headers.forEach(function(value, name) {
        this.append(name, value)
      }, this)
    } else if (Array.isArray(headers)) {
      headers.forEach(function(header) {
        this.append(header[0], header[1])
      }, this)
    } else if (headers) {
      Object.getOwnPropertyNames(headers).forEach(function(name) {
        this.append(name, headers[name])
      }, this)
    }
  }

  Headers.prototype.append = function(name, value) {
    name = normalizeName(name)
    value = normalizeValue(value)
    var oldValue = this.map[name]
    this.map[name] = oldValue ? oldValue+','+value : value
  }

  Headers.prototype['delete'] = function(name) {
    delete this.map[normalizeName(name)]
  }

  Headers.prototype.get = function(name) {
    name = normalizeName(name)
    return this.has(name) ? this.map[name] : null
  }

  Headers.prototype.has = function(name) {
    return this.map.hasOwnProperty(normalizeName(name))
  }

  Headers.prototype.set = function(name, value) {
    this.map[normalizeName(name)] = normalizeValue(value)
  }

  Headers.prototype.forEach = function(callback, thisArg) {
    for (var name in this.map) {
      if (this.map.hasOwnProperty(name)) {
        callback.call(thisArg, this.map[name], name, this)
      }
    }
  }

  Headers.prototype.keys = function() {
    var items = []
    this.forEach(function(value, name) { items.push(name) })
    return iteratorFor(items)
  }

  Headers.prototype.values = function() {
    var items = []
    this.forEach(function(value) { items.push(value) })
    return iteratorFor(items)
  }

  Headers.prototype.entries = function() {
    var items = []
    this.forEach(function(value, name) { items.push([name, value]) })
    return iteratorFor(items)
  }

  if (support.iterable) {
    Headers.prototype[Symbol.iterator] = Headers.prototype.entries
  }

  function consumed(body) {
    if (body.bodyUsed) {
      return Promise.reject(new TypeError('Already read'))
    }
    body.bodyUsed = true
  }

  function fileReaderReady(reader) {
    return new Promise(function(resolve, reject) {
      reader.onload = function() {
        resolve(reader.result)
      }
      reader.onerror = function() {
        reject(reader.error)
      }
    })
  }

  function readBlobAsArrayBuffer(blob) {
    var reader = new FileReader()
    var promise = fileReaderReady(reader)
    reader.readAsArrayBuffer(blob)
    return promise
  }

  function readBlobAsText(blob) {
    var reader = new FileReader()
    var promise = fileReaderReady(reader)
    reader.readAsText(blob)
    return promise
  }

  function readArrayBufferAsText(buf) {
    var view = new Uint8Array(buf)
    var chars = new Array(view.length)

    for (var i = 0; i < view.length; i++) {
      chars[i] = String.fromCharCode(view[i])
    }
    return chars.join('')
  }

  function bufferClone(buf) {
    if (buf.slice) {
      return buf.slice(0)
    } else {
      var view = new Uint8Array(buf.byteLength)
      view.set(new Uint8Array(buf))
      return view.buffer
    }
  }

  function Body() {
    this.bodyUsed = false

    this._initBody = function(body) {
      this._bodyInit = body
      if (!body) {
        this._bodyText = ''
      } else if (typeof body === 'string') {
        this._bodyText = body
      } else if (support.blob && Blob.prototype.isPrototypeOf(body)) {
        this._bodyBlob = body
      } else if (support.formData && FormData.prototype.isPrototypeOf(body)) {
        this._bodyFormData = body
      } else if (support.searchParams && URLSearchParams.prototype.isPrototypeOf(body)) {
        this._bodyText = body.toString()
      } else if (support.arrayBuffer && support.blob && isDataView(body)) {
        this._bodyArrayBuffer = bufferClone(body.buffer)
        // IE 10-11 can't handle a DataView body.
        this._bodyInit = new Blob([this._bodyArrayBuffer])
      } else if (support.arrayBuffer && (ArrayBuffer.prototype.isPrototypeOf(body) || isArrayBufferView(body))) {
        this._bodyArrayBuffer = bufferClone(body)
      } else {
        throw new Error('unsupported BodyInit type')
      }

      if (!this.headers.get('content-type')) {
        if (typeof body === 'string') {
          this.headers.set('content-type', 'text/plain;charset=UTF-8')
        } else if (this._bodyBlob && this._bodyBlob.type) {
          this.headers.set('content-type', this._bodyBlob.type)
        } else if (support.searchParams && URLSearchParams.prototype.isPrototypeOf(body)) {
          this.headers.set('content-type', 'application/x-www-form-urlencoded;charset=UTF-8')
        }
      }
    }

    if (support.blob) {
      this.blob = function() {
        var rejected = consumed(this)
        if (rejected) {
          return rejected
        }

        if (this._bodyBlob) {
          return Promise.resolve(this._bodyBlob)
        } else if (this._bodyArrayBuffer) {
          return Promise.resolve(new Blob([this._bodyArrayBuffer]))
        } else if (this._bodyFormData) {
          throw new Error('could not read FormData body as blob')
        } else {
          return Promise.resolve(new Blob([this._bodyText]))
        }
      }

      this.arrayBuffer = function() {
        if (this._bodyArrayBuffer) {
          return consumed(this) || Promise.resolve(this._bodyArrayBuffer)
        } else {
          return this.blob().then(readBlobAsArrayBuffer)
        }
      }
    }

    this.text = function() {
      var rejected = consumed(this)
      if (rejected) {
        return rejected
      }

      if (this._bodyBlob) {
        return readBlobAsText(this._bodyBlob)
      } else if (this._bodyArrayBuffer) {
        return Promise.resolve(readArrayBufferAsText(this._bodyArrayBuffer))
      } else if (this._bodyFormData) {
        throw new Error('could not read FormData body as text')
      } else {
        return Promise.resolve(this._bodyText)
      }
    }

    if (support.formData) {
      this.formData = function() {
        return this.text().then(decode)
      }
    }

    this.json = function() {
      return this.text().then(JSON.parse)
    }

    return this
  }

  // HTTP methods whose capitalization should be normalized
  var methods = ['DELETE', 'GET', 'HEAD', 'OPTIONS', 'POST', 'PUT']

  function normalizeMethod(method) {
    var upcased = method.toUpperCase()
    return (methods.indexOf(upcased) > -1) ? upcased : method
  }

  function Request(input, options) {
    options = options || {}
    var body = options.body

    if (input instanceof Request) {
      if (input.bodyUsed) {
        throw new TypeError('Already read')
      }
      this.url = input.url
      this.credentials = input.credentials
      if (!options.headers) {
        this.headers = new Headers(input.headers)
      }
      this.method = input.method
      this.mode = input.mode
      if (!body && input._bodyInit != null) {
        body = input._bodyInit
        input.bodyUsed = true
      }
    } else {
      this.url = String(input)
    }

    this.credentials = options.credentials || this.credentials || 'omit'
    if (options.headers || !this.headers) {
      this.headers = new Headers(options.headers)
    }
    this.method = normalizeMethod(options.method || this.method || 'GET')
    this.mode = options.mode || this.mode || null
    this.referrer = null

    if ((this.method === 'GET' || this.method === 'HEAD') && body) {
      throw new TypeError('Body not allowed for GET or HEAD requests')
    }
    this._initBody(body)
  }

  Request.prototype.clone = function() {
    return new Request(this, { body: this._bodyInit })
  }

  function decode(body) {
    var form = new FormData()
    body.trim().split('&').forEach(function(bytes) {
      if (bytes) {
        var split = bytes.split('=')
        var name = split.shift().replace(/\+/g, ' ')
        var value = split.join('=').replace(/\+/g, ' ')
        form.append(decodeURIComponent(name), decodeURIComponent(value))
      }
    })
    return form
  }

  function parseHeaders(rawHeaders) {
    var headers = new Headers()
    rawHeaders.split(/\r?\n/).forEach(function(line) {
      var parts = line.split(':')
      var key = parts.shift().trim()
      if (key) {
        var value = parts.join(':').trim()
        headers.append(key, value)
      }
    })
    return headers
  }

  Body.call(Request.prototype)

  function Response(bodyInit, options) {
    if (!options) {
      options = {}
    }

    this.type = 'default'
    this.status = 'status' in options ? options.status : 200
    this.ok = this.status >= 200 && this.status < 300
    this.statusText = 'statusText' in options ? options.statusText : 'OK'
    this.headers = new Headers(options.headers)
    this.url = options.url || ''
    this._initBody(bodyInit)
  }

  Body.call(Response.prototype)

  Response.prototype.clone = function() {
    return new Response(this._bodyInit, {
      status: this.status,
      statusText: this.statusText,
      headers: new Headers(this.headers),
      url: this.url
    })
  }

  Response.error = function() {
    var response = new Response(null, {status: 0, statusText: ''})
    response.type = 'error'
    return response
  }

  var redirectStatuses = [301, 302, 303, 307, 308]

  Response.redirect = function(url, status) {
    if (redirectStatuses.indexOf(status) === -1) {
      throw new RangeError('Invalid status code')
    }

    return new Response(null, {status: status, headers: {location: url}})
  }

  self.Headers = Headers
  self.Request = Request
  self.Response = Response

  self.fetch = function(input, init) {
    return new Promise(function(resolve, reject) {
      var request = new Request(input, init)
      var xhr = new XMLHttpRequest()

      xhr.onload = function() {
        var options = {
          status: xhr.status,
          statusText: xhr.statusText,
          headers: parseHeaders(xhr.getAllResponseHeaders() || '')
        }
        options.url = 'responseURL' in xhr ? xhr.responseURL : options.headers.get('X-Request-URL')
        var body = 'response' in xhr ? xhr.response : xhr.responseText
        resolve(new Response(body, options))
      }

      xhr.onerror = function() {
        reject(new TypeError('Network request failed'))
      }

      xhr.ontimeout = function() {
        reject(new TypeError('Network request failed'))
      }

      xhr.open(request.method, request.url, true)

      if (request.credentials === 'include') {
        xhr.withCredentials = true
      }

      if ('responseType' in xhr && support.blob) {
        xhr.responseType = 'blob'
      }

      request.headers.forEach(function(value, name) {
        xhr.setRequestHeader(name, value)
      })

      xhr.send(typeof request._bodyInit === 'undefined' ? null : request._bodyInit)
    })
  }
  self.fetch.polyfill = true
})(typeof self !== 'undefined' ? self : this);


/***/ })
/******/ ]);
//# sourceMappingURL=customfunctions.g.js.map