var OSFPerformance;
(function (OSFPerformance) {
    OSFPerformance.officeExecuteStart = 0;
    OSFPerformance.officeExecuteEnd = 0;
    OSFPerformance.hostInitializationStart = 0;
    OSFPerformance.hostInitializationEnd = 0;
    OSFPerformance.createOMEnd = 0;
    OSFPerformance.hostSpecificFileName = "";
    function now() {
        if (performance && performance.now) {
            return performance.now();
        }
        else {
            return 0;
        }
    }
    OSFPerformance.now = now;
})(OSFPerformance || (OSFPerformance = {}));
;
OSFPerformance.officeExecuteStart = OSFPerformance.now();



/* Office JavaScript API library */

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
        "android": "16.00",
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
    },
    "visio": {
        "web": "16.00",
        "win32": "16.00"
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
var OTel;
(function (OTel) {
    var CDN_PATH_OTELJS_AGAVE = 'telemetry/oteljs_agave.js';
    var OTelLogger = (function () {
        function OTelLogger() {
        }
        OTelLogger.loaded = function () {
            return !(OTelLogger.logger === undefined);
        };
        OTelLogger.getOtelSinkCDNLocation = function () {
            return (OSF._OfficeAppFactory.getLoadScriptHelper().getOfficeJsBasePath() + CDN_PATH_OTELJS_AGAVE);
        };
        OTelLogger.getMapName = function (map, name) {
            if (name !== undefined && map.hasOwnProperty(name)) {
                return map[name];
            }
            return name;
        };
        OTelLogger.getHost = function () {
            var host = OSF._OfficeAppFactory.getHostInfo()["hostType"];
            var map = {
                "excel": "Excel",
                "onenote": "OneNote",
                "outlook": "Outlook",
                "powerpoint": "PowerPoint",
                "project": "Project",
                "visio": "Visio",
                "word": "Word"
            };
            var mappedName = OTelLogger.getMapName(map, host);
            return mappedName;
        };
        OTelLogger.getFlavor = function () {
            var flavor = OSF._OfficeAppFactory.getHostInfo()["hostPlatform"];
            var map = {
                "android": "Android",
                "ios": "iOS",
                "mac": "Mac",
                "universal": "Universal",
                "web": "Web",
                "win32": "Win32"
            };
            var mappedName = OTelLogger.getMapName(map, flavor);
            return mappedName;
        };
        OTelLogger.ensureValue = function (value, alternative) {
            if (!value) {
                return alternative;
            }
            return value;
        };
        OTelLogger.create = function (info) {
            var contract = {
                id: info.appId,
                assetId: info.assetId,
                officeJsVersion: info.officeJSVersion,
                hostJsVersion: info.hostJSVersion,
                browserToken: info.clientId,
                instanceId: info.appInstanceId,
                name: info.name,
                sessionId: info.sessionId
            };
            var fields = oteljs.Contracts.Office.System.SDX.getFields("SDX", contract);
            var host = OTelLogger.getHost();
            var flavor = OTelLogger.getFlavor();
            var version = (flavor === "Web" && info.hostVersion.slice(0, 2) === "0.") ? "16.0.0.0" : info.hostVersion;
            var context = {
                'App.Name': host,
                'App.Platform': flavor,
                'App.Version': version,
                'Session.Id': OTelLogger.ensureValue(info.correlationId, "00000000-0000-0000-0000-000000000000")
            };
            var sink = oteljs_agave.AgaveSink.createInstance(context);
            var namespace = "Office.Extensibility.OfficeJs";
            var ariaTenantToken = 'db334b301e7b474db5e0f02f07c51a47-a1b5bc36-1bbe-482f-a64a-c2d9cb606706-7439';
            var nexusTenantToken = 1755;
            var logger = new oteljs.TelemetryLogger(undefined, fields);
            logger.addSink(sink);
            logger.setTenantToken(namespace, ariaTenantToken, nexusTenantToken);
            return logger;
        };
        OTelLogger.initialize = function (info) {
            if (!OTelLogger.Enabled) {
                OTelLogger.promises = [];
                return;
            }
            var timeoutAfterOneSecond = 1000;
            var afterOnReady = function () {
                if ((typeof oteljs === "undefined") || (typeof oteljs_agave === "undefined")) {
                    return;
                }
                if (!OTelLogger.loaded()) {
                    OTelLogger.logger = OTelLogger.create(info);
                }
                if (OTelLogger.loaded()) {
                    OTelLogger.promises.forEach(function (resolve) {
                        resolve();
                    });
                }
            };
            var afterLoadOtelSink = function () {
                Microsoft.Office.WebExtension.onReadyInternal().then(function () { return afterOnReady(); });
            };
            OSF.OUtil.loadScript(OTelLogger.getOtelSinkCDNLocation(), afterLoadOtelSink, timeoutAfterOneSecond);
        };
        OTelLogger.sendTelemetryEvent = function (telemetryEvent) {
            OTelLogger.onTelemetryLoaded(function () {
                try {
                    OTelLogger.logger.sendTelemetryEvent(telemetryEvent);
                }
                catch (e) {
                }
            });
        };
        OTelLogger.onTelemetryLoaded = function (resolve) {
            if (!OTelLogger.Enabled) {
                return;
            }
            if (OTelLogger.loaded()) {
                resolve();
            }
            else {
                OTelLogger.promises.push(resolve);
            }
        };
        OTelLogger.promises = [];
        OTelLogger.Enabled = true;
        return OTelLogger;
    })();
    OTel.OTelLogger = OTelLogger;
})(OTel || (OTel = {}));
var OfficeExt;
(function (OfficeExt) {
    var Association = (function () {
        function Association() {
            this.m_mappings = {};
            this.m_onchangeHandlers = [];
        }
        Association.prototype.associate = function (arg1, arg2) {
            function consoleWarn(message) {
                if (typeof console !== 'undefined' && console.warn) {
                    console.warn(message);
                }
            }
            if (arguments.length == 1 && typeof arguments[0] === 'object' && arguments[0]) {
                var mappings = arguments[0];
                for (var key in mappings) {
                    this.associate(key, mappings[key]);
                }
            }
            else if (arguments.length == 2) {
                var name_1 = arguments[0];
                var func = arguments[1];
                if (typeof name_1 !== 'string') {
                    consoleWarn('[InvalidArg] Function=associate');
                    return;
                }
                if (typeof func !== 'function') {
                    consoleWarn('[InvalidArg] Function=associate');
                    return;
                }
                var nameUpperCase = name_1.toUpperCase();
                if (this.m_mappings[nameUpperCase]) {
                    consoleWarn('[DuplicatedName] Function=' + name_1);
                }
                this.m_mappings[nameUpperCase] = func;
                for (var i = 0; i < this.m_onchangeHandlers.length; i++) {
                    this.m_onchangeHandlers[i]();
                }
            }
            else {
                consoleWarn('[InvalidArg] Function=associate');
            }
        };
        Association.prototype.onchange = function (handler) {
            if (handler) {
                this.m_onchangeHandlers.push(handler);
            }
        };
        Object.defineProperty(Association.prototype, "mappings", {
            get: function () {
                return this.m_mappings;
            },
            enumerable: true,
            configurable: true
        });
        return Association;
    })();
    OfficeExt.Association = Association;
})(OfficeExt || (OfficeExt = {}));
var CustomFunctionMappings = window.CustomFunctionMappings || {};
var CustomFunctions;
(function (CustomFunctions) {
    function delayInitialization() {
        CustomFunctionMappings['__delay__'] = true;
    }
    CustomFunctions.delayInitialization = delayInitialization;
    ;
    CustomFunctions._association = new OfficeExt.Association();
    function associate() {
        CustomFunctions._association.associate.apply(CustomFunctions._association, arguments);
        delete CustomFunctionMappings['__delay__'];
    }
    CustomFunctions.associate = associate;
    ;
})(CustomFunctions || (CustomFunctions = {}));
var Office;
(function (Office) {
    var actions;
    (function (actions) {
        actions._association = new OfficeExt.Association();
        function associate() {
            actions._association.associate.apply(actions._association, arguments);
        }
        actions.associate = associate;
        ;
    })(actions = Office.actions || (Office.actions = {}));
})(Office || (Office = {}));
(function () {
    var previousConstantNames = OSF.ConstantNames || {};
    OSF.ConstantNames = {
        FileVersion: "0.0.0.0",
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
OSF.HostInfoFlags = {
    SharedApp: 1,
    CustomFunction: 2
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
    if (typeof (window.Office) === 'object') {
        for (var p in window.Office) {
            if (window.Office.hasOwnProperty(p)) {
                Microsoft.Office.WebExtension[p] = window.Office[p];
            }
        }
    }
    window.Office = Microsoft.Office.WebExtension;
    var initialDisplayModeMappings = {
        0: "Unknown",
        1: "Hidden",
        2: "Taskpane",
        3: "Dialog"
    };
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
        Access: "Access",
        Visio: "Visio"
    };
    var _context = {};
    var _settings = {};
    var _hostFacade = {};
    var _WebAppState = { id: null, webAppUrl: null, conversationID: null, clientEndPoint: null, wnd: window.parent, focused: false };
    var _hostInfo = { isO15: true, isRichClient: true, hostType: "", hostPlatform: "", hostSpecificFileVersion: "", hostLocale: "", osfControlAppCorrelationId: "", isDialog: false, disableLogging: false, flags: 0 };
    var _isLoggingAllowed = true;
    var _initializationHelper = {};
    var _appInstanceId = null;
    var _isOfficeJsLoaded = false;
    var _officeOnReadyPendingResolves = [];
    var _isOfficeOnReadyCalled = false;
    var _officeOnReadyHostAndPlatformInfo = { host: null, platform: null, addin: null };
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
    var setOfficeJsAsLoadedAndDispatchPendingOnReadyCallbacks = function (_a) {
        var host = _a.host, platform = _a.platform, addin = _a.addin;
        _isOfficeJsLoaded = true;
        if (typeof OSFPerformance !== "undefined") {
            OSFPerformance.officeOnReady = OSFPerformance.now();
        }
        _officeOnReadyHostAndPlatformInfo = { host: host, platform: platform, addin: addin };
        while (_officeOnReadyPendingResolves.length > 0) {
            _officeOnReadyPendingResolves.shift()(_officeOnReadyHostAndPlatformInfo);
        }
    };
    Microsoft.Office.WebExtension.FeatureGates = {};
    Microsoft.Office.WebExtension.sendTelemetryEvent = function Microsoft_Office_WebExtension_sendTelemetryEvent(telemetryEvent) {
        OTel.OTelLogger.sendTelemetryEvent(telemetryEvent);
    };
    Microsoft.Office.WebExtension.onReadyInternal = function Microsoft_Office_WebExtension_onReadyInternal(callback) {
        if (_isOfficeJsLoaded) {
            var host = _officeOnReadyHostAndPlatformInfo.host, platform = _officeOnReadyHostAndPlatformInfo.platform, addin = _officeOnReadyHostAndPlatformInfo.addin;
            if (callback) {
                var result = callback({ host: host, platform: platform, addin: addin });
                if (result && result.then && typeof result.then === "function") {
                    return result.then(function () { return Office.Promise.resolve({ host: host, platform: platform, addin: addin }); });
                }
            }
            return Office.Promise.resolve({ host: host, platform: platform, addin: addin });
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
    Microsoft.Office.WebExtension.onReady = function Microsoft_Office_WebExtension_onReady(callback) {
        _isOfficeOnReadyCalled = true;
        return Microsoft.Office.WebExtension.onReadyInternal(callback);
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
            _hostInfo.flags = (((typeof items[7]) === "string") && items[7].length > 0) ? parseInt(items[7]) : 0;
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
        if ((_hostInfo.flags & (OSF.HostInfoFlags.SharedApp | OSF.HostInfoFlags.CustomFunction)) !== 0) {
            if (typeof (window.Promise) === 'undefined') {
                window.Promise = window.Office.Promise;
            }
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
                if (typeof OSFPerformance !== "undefined") {
                    OSFPerformance.getAppContextStart = OSFPerformance.now();
                }
                getAppContextAsync(_WebAppState.wnd, function (appContext) {
                    if (typeof OSFPerformance !== "undefined") {
                        OSFPerformance.getAppContextEnd = OSFPerformance.now();
                    }
                    if (OSF.AppTelemetry && OSF.AppTelemetry.logAppCommonMessage) {
                        OSF.AppTelemetry.logAppCommonMessage("getAppContextAsync callback start");
                    }
                    _appInstanceId = appContext._appInstanceId;
                    if (appContext.get_featureGates) {
                        var featureGates = appContext.get_featureGates();
                        if (featureGates) {
                            Microsoft.Office.WebExtension.FeatureGates = featureGates;
                        }
                    }
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
                                    if (!Microsoft.Office.WebExtension.initialize) {
                                        Microsoft.Office.WebExtension.initialize = function () { };
                                    }
                                    _initializationHelper.prepareRightBeforeWebExtensionInitialize(appContext);
                                }
                                _initializationHelper.prepareRightAfterWebExtensionInitialize && _initializationHelper.prepareRightAfterWebExtensionInitialize();
                                var appNumber = appContext.get_appName();
                                var addinInfo = null;
                                if ((_hostInfo.flags & OSF.HostInfoFlags.SharedApp) !== 0) {
                                    addinInfo = {
                                        visibilityMode: initialDisplayModeMappings[(appContext.get_initialDisplayMode && typeof appContext.get_initialDisplayMode === 'function') ? appContext.get_initialDisplayMode() : 0]
                                    };
                                }
                                setOfficeJsAsLoadedAndDispatchPendingOnReadyCallbacks({
                                    host: OfficeExt.HostName.Host.getInstance().getHost(appNumber),
                                    platform: OfficeExt.HostName.Host.getInstance().getPlatform(appNumber),
                                    addin: addinInfo
                                });
                            }
                            else {
                                throw new Error("Office.js has not fully loaded. Your app must call \"Office.onReady()\" as part of it's loading sequence (or set the \"Office.initialize\" function). If your app has this functionality, try reloading this page.");
                            }
                        }, 400, 50);
                    };
                    if (!_loadScriptHelper.isScriptLoading(OSF.ConstantNames.OfficeStringsId)) {
                        loadLocaleStrings(appContext.get_appUILocale());
                    }
                    _loadScriptHelper.waitForScripts([OSF.ConstantNames.OfficeStringsId], function () {
                        if (officeStrings && !Strings.OfficeOM) {
                            Strings.OfficeOM = officeStrings;
                        }
                        _initializationHelper.loadAppSpecificScriptAndCreateOM(appContext, appReady, basePath);
                        if (typeof OSFPerformance !== "undefined") {
                            OSFPerformance.createOMEnd = OSFPerformance.now();
                        }
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
                            setOfficeJsAsLoadedAndDispatchPendingOnReadyCallbacks({
                                host: null,
                                platform: null,
                                addin: null
                            });
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
            if (typeof OSFPerformance !== "undefined") {
                OSFPerformance.hostSpecificFileName = hostSpecificFileName;
            }
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
    if (window.addEventListener) {
        window.addEventListener('DOMContentLoaded', function (event) {
            Microsoft.Office.WebExtension.onReadyInternal(function () {
                if (typeof OSFPerfUtil !== 'undefined') {
                    OSFPerfUtil.sendPerformanceTelemetry();
                }
            });
        });
    }
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



var oteljs=function(e){var t={};function n(i){if(t[i])return t[i].exports;var r=t[i]={i:i,l:!1,exports:{}};return e[i].call(r.exports,r,r.exports,n),r.l=!0,r.exports}return n.m=e,n.c=t,n.d=function(e,t,i){n.o(e,t)||Object.defineProperty(e,t,{enumerable:!0,get:i})},n.r=function(e){"undefined"!=typeof Symbol&&Symbol.toStringTag&&Object.defineProperty(e,Symbol.toStringTag,{value:"Module"}),Object.defineProperty(e,"__esModule",{value:!0})},n.t=function(e,t){if(1&t&&(e=n(e)),8&t)return e;if(4&t&&"object"==typeof e&&e&&e.__esModule)return e;var i=Object.create(null);if(n.r(i),Object.defineProperty(i,"default",{enumerable:!0,value:e}),2&t&&"string"!=typeof e)for(var r in e)n.d(i,r,function(t){return e[t]}.bind(null,r));return i},n.n=function(e){var t=e&&e.__esModule?function(){return e.default}:function(){return e};return n.d(t,"a",t),t},n.o=function(e,t){return Object.prototype.hasOwnProperty.call(e,t)},n.p="",n(n.s=19)}([function(e,t,n){"use strict";n.d(t,"a",(function(){return o})),n.d(t,"d",(function(){return a})),n.d(t,"b",(function(){return c})),n.d(t,"e",(function(){return s})),n.d(t,"c",(function(){return u}));var i=n(3),r=n(4);function o(e,t){return{name:e,dataType:i.a.Boolean,value:t,classification:r.a.SystemMetadata}}function a(e,t){return{name:e,dataType:i.a.Int64,value:t,classification:r.a.SystemMetadata}}function c(e,t){return{name:e,dataType:i.a.Double,value:t,classification:r.a.SystemMetadata}}function s(e,t){return{name:e,dataType:i.a.String,value:t,classification:r.a.SystemMetadata}}function u(e,t){return{name:e,dataType:i.a.Guid,value:t,classification:r.a.SystemMetadata}}},function(e,t,n){"use strict";n.d(t,"b",(function(){return i})),n.d(t,"a",(function(){return r})),n.d(t,"e",(function(){return a})),n.d(t,"d",(function(){return c})),n.d(t,"c",(function(){return s}));var i,r,o=new(n(9).a);function a(){return o}function c(e,t,n){o.fireEvent({level:e,category:t,message:n})}function s(e,t,n){c(i.Error,e,(function(){var e=n instanceof Error?n.message:"";return t+": "+e}))}!function(e){e[e.Error=0]="Error",e[e.Warning=1]="Warning",e[e.Info=2]="Info",e[e.Verbose=3]="Verbose"}(i||(i={})),function(e){e[e.Core=0]="Core",e[e.Sink=1]="Sink",e[e.Transport=2]="Transport"}(r||(r={}))},function(e,t,n){"use strict";n.d(t,"a",(function(){return r}));var i=n(0);function r(e,t,n){e.push(Object(i.e)("zC."+t,n))}},function(e,t,n){"use strict";var i;n.d(t,"a",(function(){return i})),function(e){e[e.String=0]="String",e[e.Boolean=1]="Boolean",e[e.Int64=2]="Int64",e[e.Double=3]="Double",e[e.Guid=4]="Guid"}(i||(i={}))},function(e,t,n){"use strict";var i;n.d(t,"a",(function(){return i})),function(e){e[e.EssentialServiceMetadata=1]="EssentialServiceMetadata",e[e.AccountData=2]="AccountData",e[e.SystemMetadata=4]="SystemMetadata",e[e.OrganizationIdentifiableInformation=8]="OrganizationIdentifiableInformation",e[e.EndUserIdentifiableInformation=16]="EndUserIdentifiableInformation",e[e.CustomerContent=32]="CustomerContent",e[e.AccessControl=64]="AccessControl"}(i||(i={}))},function(e,t,n){"use strict";var i,r,o,a,c;n.d(t,"e",(function(){return i})),n.d(t,"d",(function(){return r})),n.d(t,"a",(function(){return o})),n.d(t,"b",(function(){return a})),n.d(t,"c",(function(){return c})),function(e){e[e.NotSet=0]="NotSet",e[e.Measure=1]="Measure",e[e.Diagnostics=2]="Diagnostics",e[e.CriticalBusinessImpact=191]="CriticalBusinessImpact",e[e.CriticalCensus=192]="CriticalCensus",e[e.CriticalExperimentation=193]="CriticalExperimentation",e[e.CriticalUsage=194]="CriticalUsage"}(i||(i={})),function(e){e[e.NotSet=0]="NotSet",e[e.Normal=1]="Normal",e[e.High=2]="High"}(r||(r={})),function(e){e[e.NotSet=0]="NotSet",e[e.Normal=1]="Normal",e[e.High=2]="High"}(o||(o={})),function(e){e[e.NotSet=0]="NotSet",e[e.SoftwareSetup=1]="SoftwareSetup",e[e.ProductServiceUsage=2]="ProductServiceUsage",e[e.ProductServicePerformance=4]="ProductServicePerformance",e[e.DeviceConfiguration=8]="DeviceConfiguration",e[e.InkingTypingSpeech=16]="InkingTypingSpeech"}(a||(a={})),function(e){e[e.ReservedDoNotUse=0]="ReservedDoNotUse",e[e.BasicEvent=10]="BasicEvent",e[e.FullEvent=100]="FullEvent",e[e.NecessaryServiceDataEvent=110]="NecessaryServiceDataEvent",e[e.AlwaysOnNecessaryServiceDataEvent=120]="AlwaysOnNecessaryServiceDataEvent"}(c||(c={}))},function(e,t,n){"use strict";n.d(t,"a",(function(){return p}));var i,r,o,a,c,s,u,d,f,l=n(0),v=n(2);(i||(i={})).getFields=function(e,t){var n=[];return n.push(Object(l.d)(e+".Code",t.code)),void 0!==t.type&&n.push(Object(l.e)(e+".Type",t.type)),void 0!==t.tag&&n.push(Object(l.d)(e+".Tag",t.tag)),void 0!==t.isExpected&&n.push(Object(l.a)(e+".IsExpected",t.isExpected)),Object(v.a)(n,e,"Office.System.Result"),n},(o=r||(r={})).contractName="Office.System.Activity",o.getFields=function(e){var t=[];return void 0!==e.cV&&t.push(Object(l.e)("Activity.CV",e.cV)),t.push(Object(l.d)("Activity.Duration",e.duration)),t.push(Object(l.d)("Activity.Count",e.count)),t.push(Object(l.d)("Activity.AggMode",e.aggMode)),void 0!==e.success&&t.push(Object(l.a)("Activity.Success",e.success)),void 0!==e.result&&t.push.apply(t,i.getFields("Activity.Result",e.result)),Object(v.a)(t,"Activity",o.contractName),t},(a||(a={})).getFields=function(e,t){var n=[];return void 0!==t.id&&n.push(Object(l.e)(e+".Id",t.id)),void 0!==t.version&&n.push(Object(l.e)(e+".Version",t.version)),void 0!==t.sessionId&&n.push(Object(l.e)(e+".SessionId",t.sessionId)),Object(v.a)(n,e,"Office.System.Host"),n},(c||(c={})).getFields=function(e,t){var n=[];return void 0!==t.alias&&n.push(Object(l.e)(e+".Alias",t.alias)),void 0!==t.primaryIdentityHash&&n.push(Object(l.e)(e+".PrimaryIdentityHash",t.primaryIdentityHash)),void 0!==t.primaryIdentitySpace&&n.push(Object(l.e)(e+".PrimaryIdentitySpace",t.primaryIdentitySpace)),void 0!==t.tenantId&&n.push(Object(l.e)(e+".TenantId",t.tenantId)),void 0!==t.tenantGroup&&n.push(Object(l.e)(e+".TenantGroup",t.tenantGroup)),void 0!==t.isAnonymous&&n.push(Object(l.a)(e+".IsAnonymous",t.isAnonymous)),Object(v.a)(n,e,"Office.System.User"),n},(s||(s={})).getFields=function(e,t){var n=[];return void 0!==t.id&&n.push(Object(l.e)(e+".Id",t.id)),void 0!==t.version&&n.push(Object(l.e)(e+".Version",t.version)),void 0!==t.instanceId&&n.push(Object(l.e)(e+".InstanceId",t.instanceId)),void 0!==t.name&&n.push(Object(l.e)(e+".Name",t.name)),void 0!==t.marketplaceType&&n.push(Object(l.e)(e+".MarketplaceType",t.marketplaceType)),void 0!==t.sessionId&&n.push(Object(l.e)(e+".SessionId",t.sessionId)),void 0!==t.browserToken&&n.push(Object(l.e)(e+".BrowserToken",t.browserToken)),void 0!==t.osfRuntimeVersion&&n.push(Object(l.e)(e+".OsfRuntimeVersion",t.osfRuntimeVersion)),void 0!==t.officeJsVersion&&n.push(Object(l.e)(e+".OfficeJsVersion",t.officeJsVersion)),void 0!==t.hostJsVersion&&n.push(Object(l.e)(e+".HostJsVersion",t.hostJsVersion)),void 0!==t.assetId&&n.push(Object(l.e)(e+".AssetId",t.assetId)),void 0!==t.providerName&&n.push(Object(l.e)(e+".ProviderName",t.providerName)),void 0!==t.type&&n.push(Object(l.e)(e+".Type",t.type)),Object(v.a)(n,e,"Office.System.SDX"),n},(u||(u={})).getFields=function(e,t){var n=[];return void 0!==t.name&&n.push(Object(l.e)(e+".Name",t.name)),void 0!==t.state&&n.push(Object(l.e)(e+".State",t.state)),Object(v.a)(n,e,"Office.System.Funnel"),n},(d||(d={})).getFields=function(e,t){var n=[];return void 0!==t.id&&n.push(Object(l.d)(e+".Id",t.id)),void 0!==t.name&&n.push(Object(l.e)(e+".Name",t.name)),void 0!==t.commandSurface&&n.push(Object(l.e)(e+".CommandSurface",t.commandSurface)),void 0!==t.parentName&&n.push(Object(l.e)(e+".ParentName",t.parentName)),void 0!==t.triggerMethod&&n.push(Object(l.e)(e+".TriggerMethod",t.triggerMethod)),void 0!==t.timeOffsetMs&&n.push(Object(l.d)(e+".TimeOffsetMs",t.timeOffsetMs)),Object(v.a)(n,e,"Office.System.UserAction"),n},function(e){e.getFields=function(e,t){var n=[];return n.push(Object(l.e)(e+".ErrorGroup",t.errorGroup)),n.push(Object(l.d)(e+".Tag",t.tag)),void 0!==t.code&&n.push(Object(l.d)(e+".Code",t.code)),void 0!==t.id&&n.push(Object(l.d)(e+".Id",t.id)),void 0!==t.count&&n.push(Object(l.d)(e+".Count",t.count)),Object(v.a)(n,e,"Office.System.Error"),n}}(f||(f={}));var p,y=r,g=i,h=f,m=u,b=a,F=s,S=d,O=c;!function(e){!function(e){!function(e){e.Activity=y,e.Result=g,e.Error=h,e.Funnel=m,e.Host=b,e.SDX=F,e.User=O,e.UserAction=S}(e.System||(e.System={}))}(e.Office||(e.Office={}))}(p||(p={}))},function(e,t,n){"use strict";n.d(t,"b",(function(){return f})),n.d(t,"a",(function(){return l}));var i,r,o=n(1);!function(e){e[e.Aria=0]="Aria",e[e.Nexus=1]="Nexus"}(i||(i={})),function(e){var t={},n={},r={};function a(e){if("object"!=typeof e)throw new Error("tokenTree must be an object");r=function e(t,n){if("object"!=typeof n)return n;for(var i=0,r=Object.keys(n);i<r.length;i++){var o=r[i];o in t&&(t[o],1)?t[o]=e(t[o],n[o]):t[o]=n[o]}return t}(r,e)}function c(e){if(t[e])return t[e];var n=u(e,i.Aria);return"string"==typeof n?(t[e]=n,n):void 0}function s(e){if(n[e])return n[e];var t=u(e,i.Nexus);return"number"==typeof t?(n[e]=t,t):void 0}function u(e,t){var n=e.split("."),o=r,a=void 0;if(o){for(var c=0;c<n.length-1;c++)o[n[c]]&&(o=o[n[c]],t===i.Aria&&"string"==typeof o.ariaTenantToken?a=o.ariaTenantToken:t===i.Nexus&&"number"==typeof o.nexusTenantToken&&(a=o.nexusTenantToken));return a}}e.setTenantToken=function(e,t,n){var i=e.split(".");if(i.length<2||"Office"!==i[0])Object(o.d)(o.b.Error,o.a.Core,(function(){return"Invalid namespace: "+e}));else{var r=Object.create(Object.prototype);t&&(r.ariaTenantToken=t),n&&(r.nexusTenantToken=n);var c,s=r;for(c=i.length-1;c>=0;--c){var u=Object.create(Object.prototype);u[i[c]]=s,s=u}a(s)}},e.setTenantTokens=a,e.getTenantTokens=function(e){var t=c(e),n=s(e);if(!n||!t)throw new Error("Could not find tenant token for "+e);return{ariaTenantToken:t,nexusTenantToken:n}},e.getAriaTenantToken=c,e.getNexusTenantToken=s,e.clear=function(){t={},n={},r={}}}(r||(r={}));var a,c=n(3);!function(e){var t=/^[A-Z][a-zA-Z0-9]*$/,n=/^[a-zA-Z0-9_\.]*$/;function i(e){return void 0!==e&&n.test(e)}function r(e){if(!((t=e.name)&&i(t)&&t.length+5<100))throw new Error("Invalid dataField name");var t;e.dataType===c.a.Int64&&o(e.value)}function o(e){if("number"!=typeof e||!isFinite(e)||Math.floor(e)!==e||e<-9007199254740991||e>9007199254740991)throw new Error("Invalid integer "+JSON.stringify(e))}e.validateTelemetryEvent=function(e){if(!function(e){if(!e||e.length>98)return!1;var n=e.split("."),i=n[n.length-1];return function(e){return!!e&&e.length>=3&&"Office"===e[0]}(n)&&(r=i,void 0!==r&&t.test(r));var r}(e.eventName))throw new Error("Invalid eventName");if(e.eventContract&&!i(e.eventContract.name))throw new Error("Invalid eventContract");if(null!=e.dataFields)for(var n=0;n<e.dataFields.length;n++)r(e.dataFields[n])},e.validateInt=o}(a||(a={}));var s=n(9),u=n(0),d=function(){return(d=Object.assign||function(e){for(var t,n=1,i=arguments.length;n<i;n++)for(var r in t=arguments[n])Object.prototype.hasOwnProperty.call(t,r)&&(e[r]=t[r]);return e}).apply(this,arguments)},f=-1,l=function(){function e(e,t,n){var i,r;this.onSendEvent=new s.a,this.persistentDataFields=[],this.config=n||{},e?(this.onSendEvent=e.onSendEvent,(i=this.persistentDataFields).push.apply(i,e.persistentDataFields),this.config=d(d({},e.getConfig()),this.config)):this.persistentDataFields.push(Object(u.e)("OTelJS.Version","3.1.48")),t&&(r=this.persistentDataFields).push.apply(r,t)}return e.prototype.sendTelemetryEvent=function(e){var t;try{if(0===this.onSendEvent.getListenerCount())return void Object(o.d)(o.b.Warning,o.a.Core,(function(){return"No telemetry sinks are attached."}));t=this.cloneEvent(e),this.processTelemetryEvent(t)}catch(e){return void Object(o.c)(o.a.Core,"SendTelemetryEvent",e)}try{this.onSendEvent.fireEvent(t)}catch(e){}},e.prototype.processTelemetryEvent=function(e){var t;e.telemetryProperties||(e.telemetryProperties=r.getTenantTokens(e.eventName)),e.dataFields&&this.persistentDataFields&&(t=e.dataFields).unshift.apply(t,this.persistentDataFields),this.config.disableValidation||a.validateTelemetryEvent(e)},e.prototype.addSink=function(e){this.onSendEvent.addListener((function(t){return e.sendTelemetryEvent(t)}))},e.prototype.setTenantToken=function(e,t,n){r.setTenantToken(e,t,n)},e.prototype.setTenantTokens=function(e){r.setTenantTokens(e)},e.prototype.cloneEvent=function(e){var t={eventName:e.eventName,eventFlags:e.eventFlags};return e.telemetryProperties&&(t.telemetryProperties={ariaTenantToken:e.telemetryProperties.ariaTenantToken,nexusTenantToken:e.telemetryProperties.nexusTenantToken}),e.eventContract&&(t.eventContract={name:e.eventContract.name,dataFields:e.eventContract.dataFields.slice()}),t.dataFields=e.dataFields?e.dataFields.slice():[],t},e.prototype.getConfig=function(){return this.config},e}()},function(e,t,n){"use strict";var i;n.d(t,"a",(function(){return s})),function(e){var t,n=0;e.getNext=function(){return void 0===t&&(t=function(){for(var e="ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/",t=[],n=0;n<22;n++)t.push(e.charAt(Math.floor(Math.random()*e.length)));return t.join("")}()),new i(t,++n)},e.getNextChild=function(e){return new i(e.getString(),++e.nextChild)};var i=function(){function e(e,t){this.base=e,this.id=t,this.nextChild=0}return e.prototype.getString=function(){return this.base+"."+this.id},e}();e.CV=i}(i||(i={}));var r=n(1),o=function(e,t,n,i){return new(n||(n=Promise))((function(r,o){function a(e){try{s(i.next(e))}catch(e){o(e)}}function c(e){try{s(i.throw(e))}catch(e){o(e)}}function s(e){var t;e.done?r(e.value):(t=e.value,t instanceof n?t:new n((function(e){e(t)}))).then(a,c)}s((i=i.apply(e,t||[])).next())}))},a=function(e,t){var n,i,r,o,a={label:0,sent:function(){if(1&r[0])throw r[1];return r[1]},trys:[],ops:[]};return o={next:c(0),throw:c(1),return:c(2)},"function"==typeof Symbol&&(o[Symbol.iterator]=function(){return this}),o;function c(o){return function(c){return function(o){if(n)throw new TypeError("Generator is already executing.");for(;a;)try{if(n=1,i&&(r=2&o[0]?i.return:o[0]?i.throw||((r=i.return)&&r.call(i),0):i.next)&&!(r=r.call(i,o[1])).done)return r;switch(i=0,r&&(o=[2&o[0],r.value]),o[0]){case 0:case 1:r=o;break;case 4:return a.label++,{value:o[1],done:!1};case 5:a.label++,i=o[1],o=[0];continue;case 7:o=a.ops.pop(),a.trys.pop();continue;default:if(!(r=a.trys,(r=r.length>0&&r[r.length-1])||6!==o[0]&&2!==o[0])){a=0;continue}if(3===o[0]&&(!r||o[1]>r[0]&&o[1]<r[3])){a.label=o[1];break}if(6===o[0]&&a.label<r[1]){a.label=r[1],r=o;break}if(r&&a.label<r[2]){a.label=r[2],a.ops.push(o);break}r[2]&&a.ops.pop(),a.trys.pop();continue}o=t.call(e,a)}catch(e){o=[6,e],i=0}finally{n=r=0}if(5&o[0])throw o[1];return{value:o[0]?o[1]:void 0,done:!0}}([o,c])}}},c=function(){return 1e3*Date.now()};"object"==typeof window.performance&&"now"in window.performance&&(c=function(){return 1e3*Math.floor(window.performance.now())});var s=function(){function e(e,t,n){this._optionalEventFlags={},this._ended=!1,this._telemetryLogger=e,this._activityName=t,this._cv=n?i.getNextChild(n._cv):i.getNext(),this._dataFields=[],this._success=void 0,this._startTime=c()}return e.createNew=function(t,n){return new e(t,n)},e.prototype.createChildActivity=function(t){return new e(this._telemetryLogger,t,this)},e.prototype.setEventFlags=function(e){this._optionalEventFlags=e},e.prototype.addDataField=function(e){this._dataFields.push(e)},e.prototype.addDataFields=function(e){var t;(t=this._dataFields).push.apply(t,e)},e.prototype.setSuccess=function(e){this._success=e},e.prototype.setResult=function(e,t,n){this._result={code:e,type:t,tag:n}},e.prototype.endNow=function(){if(!this._ended){void 0===this._success&&void 0===this._result&&Object(r.d)(r.b.Warning,r.a.Core,(function(){return"Activity does not have success or result set"}));var e=c()-this._startTime;this._ended=!0;var t={duration:e,count:1,aggMode:0,cV:this._cv.getString(),success:this._success,result:this._result};return this._telemetryLogger.sendActivity(this._activityName,t,this._dataFields,this._optionalEventFlags)}Object(r.d)(r.b.Error,r.a.Core,(function(){return"Activity has already ended"}))},e.prototype.executeAsync=function(e){return o(this,void 0,void 0,(function(){var t=this;return a(this,(function(n){return[2,e(this).then((function(e){return t.endNow(),e})).catch((function(e){throw t.endNow(),e}))]}))}))},e.prototype.executeSync=function(e){try{var t=e(this);return this.endNow(),t}catch(e){throw this.endNow(),e}},e.prototype.executeChildActivityAsync=function(e,t){return o(this,void 0,void 0,(function(){return a(this,(function(n){return[2,this.createChildActivity(e).executeAsync(t)]}))}))},e.prototype.executeChildActivitySync=function(e,t){return this.createChildActivity(e).executeSync(t)},e}()},function(e,t,n){"use strict";n.d(t,"a",(function(){return i}));var i=function(){function e(){this._listeners=[]}return e.prototype.fireEvent=function(e){this._listeners.forEach((function(t){return t(e)}))},e.prototype.addListener=function(e){e&&this._listeners.push(e)},e.prototype.removeListener=function(e){this._listeners=this._listeners.filter((function(t){return t!==e}))},e.prototype.getListenerCount=function(){return this._listeners.length},e}()},function(e,t,n){"use strict";n.r(t);var i=n(6);n.d(t,"Contracts",(function(){return i.a}));var r=n(8);n.d(t,"ActivityScope",(function(){return r.a}));var o=n(2);n.d(t,"addContractField",(function(){return o.a}));var a=n(11);n.d(t,"getFieldsForContract",(function(){return a.a}));var c=n(4);n.d(t,"DataClassification",(function(){return c.a}));var s=n(12);for(var u in s)["Contracts","ActivityScope","addContractField","getFieldsForContract","DataClassification","default"].indexOf(u)<0&&function(e){n.d(t,e,(function(){return s[e]}))}(u);var d=n(0);n.d(t,"makeBooleanDataField",(function(){return d.a})),n.d(t,"makeInt64DataField",(function(){return d.d})),n.d(t,"makeDoubleDataField",(function(){return d.b})),n.d(t,"makeStringDataField",(function(){return d.e})),n.d(t,"makeGuidDataField",(function(){return d.c}));var f=n(3);n.d(t,"DataFieldType",(function(){return f.a}));var l=n(13);n.d(t,"getEffectiveEventFlags",(function(){return l.a}));var v=n(5);n.d(t,"SamplingPolicy",(function(){return v.e})),n.d(t,"PersistencePriority",(function(){return v.d})),n.d(t,"CostPriority",(function(){return v.a})),n.d(t,"DataCategories",(function(){return v.b})),n.d(t,"DiagnosticLevel",(function(){return v.c}));var p=n(14);for(var u in p)["Contracts","ActivityScope","addContractField","getFieldsForContract","DataClassification","makeBooleanDataField","makeInt64DataField","makeDoubleDataField","makeStringDataField","makeGuidDataField","DataFieldType","getEffectiveEventFlags","SamplingPolicy","PersistencePriority","CostPriority","DataCategories","DiagnosticLevel","default"].indexOf(u)<0&&function(e){n.d(t,e,(function(){return p[e]}))}(u);var y=n(1);n.d(t,"LogLevel",(function(){return y.b})),n.d(t,"Category",(function(){return y.a})),n.d(t,"onNotification",(function(){return y.e})),n.d(t,"logNotification",(function(){return y.d})),n.d(t,"logError",(function(){return y.c}));var g=n(7);n.d(t,"SuppressNexus",(function(){return g.b})),n.d(t,"SimpleTelemetryLogger",(function(){return g.a}));var h=n(15);n.d(t,"TelemetryLogger",(function(){return h.a}));var m=n(16);for(var u in m)["Contracts","ActivityScope","addContractField","getFieldsForContract","DataClassification","makeBooleanDataField","makeInt64DataField","makeDoubleDataField","makeStringDataField","makeGuidDataField","DataFieldType","getEffectiveEventFlags","SamplingPolicy","PersistencePriority","CostPriority","DataCategories","DiagnosticLevel","LogLevel","Category","onNotification","logNotification","logError","SuppressNexus","SimpleTelemetryLogger","TelemetryLogger","default"].indexOf(u)<0&&function(e){n.d(t,e,(function(){return m[e]}))}(u);var b=n(17);for(var u in b)["Contracts","ActivityScope","addContractField","getFieldsForContract","DataClassification","makeBooleanDataField","makeInt64DataField","makeDoubleDataField","makeStringDataField","makeGuidDataField","DataFieldType","getEffectiveEventFlags","SamplingPolicy","PersistencePriority","CostPriority","DataCategories","DiagnosticLevel","LogLevel","Category","onNotification","logNotification","logError","SuppressNexus","SimpleTelemetryLogger","TelemetryLogger","default"].indexOf(u)<0&&function(e){n.d(t,e,(function(){return b[e]}))}(u);var F=n(18);for(var u in F)["Contracts","ActivityScope","addContractField","getFieldsForContract","DataClassification","makeBooleanDataField","makeInt64DataField","makeDoubleDataField","makeStringDataField","makeGuidDataField","DataFieldType","getEffectiveEventFlags","SamplingPolicy","PersistencePriority","CostPriority","DataCategories","DiagnosticLevel","LogLevel","Category","onNotification","logNotification","logError","SuppressNexus","SimpleTelemetryLogger","TelemetryLogger","default"].indexOf(u)<0&&function(e){n.d(t,e,(function(){return F[e]}))}(u)},function(e,t,n){"use strict";n.d(t,"a",(function(){return r}));var i=n(2);function r(e,t,n){var r=n.map((function(t){return{name:e+"."+t.name,value:t.value,dataType:t.dataType}}));return Object(i.a)(r,e,t),r}},function(e,t){},function(e,t,n){"use strict";n.d(t,"a",(function(){return o}));var i=n(5),r=n(1);function o(e){var t={costPriority:i.a.Normal,samplingPolicy:i.e.Measure,persistencePriority:i.d.Normal,dataCategories:i.b.NotSet,diagnosticLevel:i.c.FullEvent};return e.eventFlags&&e.eventFlags.dataCategories||Object(r.d)(r.b.Error,r.a.Core,(function(){return"Event is missing DataCategories event flag"})),e.eventFlags?(e.eventFlags.costPriority&&(t.costPriority=e.eventFlags.costPriority),e.eventFlags.samplingPolicy&&(t.samplingPolicy=e.eventFlags.samplingPolicy),e.eventFlags.persistencePriority&&(t.persistencePriority=e.eventFlags.persistencePriority),e.eventFlags.dataCategories&&(t.dataCategories=e.eventFlags.dataCategories),e.eventFlags.diagnosticLevel&&(t.diagnosticLevel=e.eventFlags.diagnosticLevel),t):t}},function(e,t){},function(e,t,n){"use strict";n.d(t,"a",(function(){return d}));var i,r=n(7),o=n(8),a=n(6),c=(i=function(e,t){return(i=Object.setPrototypeOf||{__proto__:[]}instanceof Array&&function(e,t){e.__proto__=t}||function(e,t){for(var n in t)t.hasOwnProperty(n)&&(e[n]=t[n])})(e,t)},function(e,t){function n(){this.constructor=e}i(e,t),e.prototype=null===t?Object.create(t):(n.prototype=t.prototype,new n)}),s=function(e,t,n,i){return new(n||(n=Promise))((function(r,o){function a(e){try{s(i.next(e))}catch(e){o(e)}}function c(e){try{s(i.throw(e))}catch(e){o(e)}}function s(e){var t;e.done?r(e.value):(t=e.value,t instanceof n?t:new n((function(e){e(t)}))).then(a,c)}s((i=i.apply(e,t||[])).next())}))},u=function(e,t){var n,i,r,o,a={label:0,sent:function(){if(1&r[0])throw r[1];return r[1]},trys:[],ops:[]};return o={next:c(0),throw:c(1),return:c(2)},"function"==typeof Symbol&&(o[Symbol.iterator]=function(){return this}),o;function c(o){return function(c){return function(o){if(n)throw new TypeError("Generator is already executing.");for(;a;)try{if(n=1,i&&(r=2&o[0]?i.return:o[0]?i.throw||((r=i.return)&&r.call(i),0):i.next)&&!(r=r.call(i,o[1])).done)return r;switch(i=0,r&&(o=[2&o[0],r.value]),o[0]){case 0:case 1:r=o;break;case 4:return a.label++,{value:o[1],done:!1};case 5:a.label++,i=o[1],o=[0];continue;case 7:o=a.ops.pop(),a.trys.pop();continue;default:if(!(r=a.trys,(r=r.length>0&&r[r.length-1])||6!==o[0]&&2!==o[0])){a=0;continue}if(3===o[0]&&(!r||o[1]>r[0]&&o[1]<r[3])){a.label=o[1];break}if(6===o[0]&&a.label<r[1]){a.label=r[1],r=o;break}if(r&&a.label<r[2]){a.label=r[2],a.ops.push(o);break}r[2]&&a.ops.pop(),a.trys.pop();continue}o=t.call(e,a)}catch(e){o=[6,e],i=0}finally{n=r=0}if(5&o[0])throw o[1];return{value:o[0]?o[1]:void 0,done:!0}}([o,c])}}},d=function(e){function t(){return null!==e&&e.apply(this,arguments)||this}return c(t,e),t.prototype.executeActivityAsync=function(e,t){return s(this,void 0,void 0,(function(){return u(this,(function(n){return[2,this.createNewActivity(e).executeAsync(t)]}))}))},t.prototype.executeActivitySync=function(e,t){return this.createNewActivity(e).executeSync(t)},t.prototype.createNewActivity=function(e){return o.a.createNew(this,e)},t.prototype.sendActivity=function(e,t,n,i){return this.sendTelemetryEvent({eventName:e,eventContract:{name:a.a.Office.System.Activity.contractName,dataFields:a.a.Office.System.Activity.getFields(t)},dataFields:n,eventFlags:i})},t.prototype.sendError=function(e){var t=a.a.Office.System.Error.getFields("Error",e.error);return null!=e.dataFields&&t.push.apply(t,e.dataFields),this.sendTelemetryEvent({eventName:e.eventName,dataFields:t,eventFlags:e.eventFlags})},t}(r.a)},function(e,t){},function(e,t){},function(e,t){},function(e,t,n){e.exports=n(10)}]);



OSFPerformance.officeExecuteEnd = OSFPerformance.now();