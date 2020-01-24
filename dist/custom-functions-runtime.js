/* Office JavaScript API library - Custom Functions */

/*
	Copyright (c) Microsoft Corporation.  All rights reserved.
*/

/*
	This file incorporates the "whatwg-fetch" implementation, version 2.0.3, licensed under MIT with the following licensing notice:
	(See github.com/github/fetch/blob/master/LICENSE)

		Copyright (c) 2014-2016 GitHub, Inc.

		Permission is hereby granted, free of charge, to any person obtaining
		a copy of this software and associated documentation files (the
		"Software"), to deal in the Software without restriction, including
		without limitation the rights to use, copy, modify, merge, publish,
		distribute, sublicense, and/or sell copies of the Software, and to
		permit persons to whom the Software is furnished to do so, subject to
		the following conditions:

		The above copyright notice and this permission notice shall be
		included in all copies or substantial portions of the Software.

		THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
		EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
		MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
		NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE
		LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION
		OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION
		WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

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
    FileVersion: "0.0.0.0",
    OfficeJS: "custom-functions-runtime.js",
    OfficeDebugJS: "custom-functions-runtime.debug.js",
    HostFileScriptSuffix: "core",
    IsCustomFunctionsRuntime: true
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
                getAppContextAsync(_WebAppState.wnd, function (appContext) {
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
                                        visibilityMode: initialDisplayModeMappings[appContext.get_initialDisplayMode()]
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



!function(e){var t={};function n(r){if(t[r])return t[r].exports;var o=t[r]={i:r,l:!1,exports:{}};return e[r].call(o.exports,o,o.exports,n),o.l=!0,o.exports}n.m=e,n.c=t,n.d=function(e,t,r){n.o(e,t)||Object.defineProperty(e,t,{enumerable:!0,get:r})},n.r=function(e){"undefined"!=typeof Symbol&&Symbol.toStringTag&&Object.defineProperty(e,Symbol.toStringTag,{value:"Module"}),Object.defineProperty(e,"__esModule",{value:!0})},n.t=function(e,t){if(1&t&&(e=n(e)),8&t)return e;if(4&t&&"object"==typeof e&&e&&e.__esModule)return e;var r=Object.create(null);if(n.r(r),Object.defineProperty(r,"default",{enumerable:!0,value:e}),2&t&&"string"!=typeof e)for(var o in e)n.d(r,o,function(t){return e[t]}.bind(null,o));return r},n.n=function(e){var t=e&&e.__esModule?function(){return e.default}:function(){return e};return n.d(t,"a",t),t},n.o=function(e,t){return Object.prototype.hasOwnProperty.call(e,t)},n.p="",n(n.s=3)}([function(e,t,n){"use strict";var r,o=this&&this.__extends||(r=function(e,t){return(r=Object.setPrototypeOf||{__proto__:[]}instanceof Array&&function(e,t){e.__proto__=t}||function(e,t){for(var n in t)t.hasOwnProperty(n)&&(e[n]=t[n])})(e,t)},function(e,t){function n(){this.constructor=e}r(e,t),e.prototype=null===t?Object.create(t):(n.prototype=t.prototype,new n)});Object.defineProperty(t,"__esModule",{value:!0});var i=function(){function e(){}return e.prototype._resolveRequestUrlAndHeaderInfo=function(){return h._createPromiseFromResult(null)},e.prototype._createRequestExecutorOrNull=function(){return null},Object.defineProperty(e.prototype,"eventRegistration",{get:function(){return null},enumerable:!0,configurable:!0}),e}();t.SessionBase=i;var s=function(){function e(){}return e.setCustomSendRequestFunc=function(t){e.s_customSendRequestFunc=t},e.xhrSendRequestFunc=function(e){return h.createPromise((function(t,n){var r=new XMLHttpRequest;if(r.open(e.method,e.url),r.onload=function(){var e={statusCode:r.status,headers:h._parseHttpResponseHeaders(r.getAllResponseHeaders()),body:r.responseText};t(e)},r.onerror=function(){n(new a.RuntimeError({code:u.connectionFailure,message:h._getResourceString(l.connectionFailureWithStatus,r.statusText)}))},e.headers)for(var o in e.headers)r.setRequestHeader(o,e.headers[o]);r.send(h._getRequestBodyText(e))}))},e.sendRequest=function(t){e.validateAndNormalizeRequest(t);var n=e.s_customSendRequestFunc;return n||(n=e.xhrSendRequestFunc),n(t)},e.setCustomSendLocalDocumentRequestFunc=function(t){e.s_customSendLocalDocumentRequestFunc=t},e.sendLocalDocumentRequest=function(t){return e.validateAndNormalizeRequest(t),(e.s_customSendLocalDocumentRequestFunc||e.officeJsSendLocalDocumentRequestFunc)(t)},e.officeJsSendLocalDocumentRequestFunc=function(e){e=h._validateLocalDocumentRequest(e);var t=h._buildRequestMessageSafeArray(e);return h.createPromise((function(e,n){OSF.DDA.RichApi.executeRichApiRequestAsync(t,(function(t){var n;n="succeeded"==t.status?{statusCode:f.getResponseStatusCode(t),headers:f.getResponseHeaders(t),body:f.getResponseBody(t)}:f.buildHttpResponseFromOfficeJsError(t.error.code,t.error.message),h.log("Response:"),h.log(JSON.stringify(n)),e(n)}))}))},e.validateAndNormalizeRequest=function(e){if(h.isNullOrUndefined(e))throw a.RuntimeError._createInvalidArgError({argumentName:"request"});h.isNullOrEmptyString(e.method)&&(e.method="GET"),e.method=e.method.toUpperCase()},e.logRequest=function(t){if(h._logEnabled){if(h.log("---HTTP Request---"),h.log(t.method+" "+t.url),t.headers)for(var n in t.headers)h.log(n+": "+t.headers[n]);e._logBodyEnabled&&h.log(h._getRequestBodyText(t))}},e.logResponse=function(t){if(h._logEnabled){if(h.log("---HTTP Response---"),h.log(""+t.statusCode),t.headers)for(var n in t.headers)h.log(n+": "+t.headers[n]);e._logBodyEnabled&&h.log(t.body)}},e._logBodyEnabled=!1,e}();t.HttpUtility=s;var a,c=function(){function e(e){var t=this;this.m_bridge=e,this.m_promiseResolver={},this.m_handlers=[],this.m_bridge.onMessageFromHost=function(e){var n=JSON.parse(e);if(3==n.type){var r=n.message;if(r&&r.entries)for(var o=0;o<r.entries.length;o++){var i=r.entries[o];if(Array.isArray(i)){var s={messageCategory:i[0],messageType:i[1],targetId:i[2],message:i[3],id:i[4]};r.entries[o]=s}}}t.dispatchMessage(n)}}return e.init=function(t){if("object"==typeof t&&t){var n=new e(t);e.s_instance=n,s.setCustomSendLocalDocumentRequestFunc((function(t){t=h._validateLocalDocumentRequest(t);var r=0;h.isReadonlyRestRequest(t.method)||(r=1);var o=t.url.indexOf("?");if(o>=0){var i=t.url.substr(o+1),s=h._parseRequestFlagsAndCustomDataFromQueryStringIfAny(i);s.flags>=0&&(r=s.flags)}var a={id:e.nextId(),type:1,flags:r,message:t};return n.sendMessageToHostAndExpectResponse(a).then((function(e){return e.message}))}));for(var r=0;r<e.s_onInitedHandlers.length;r++)e.s_onInitedHandlers[r](n)}},Object.defineProperty(e,"instance",{get:function(){return e.s_instance},enumerable:!0,configurable:!0}),e.prototype.sendMessageToHost=function(e){this.m_bridge.sendMessageToHost(JSON.stringify(e))},e.prototype.sendMessageToHostAndExpectResponse=function(e){var t=this,n=h.createPromise((function(n,r){t.m_promiseResolver[e.id]=n}));return this.m_bridge.sendMessageToHost(JSON.stringify(e)),n},e.prototype.addHostMessageHandler=function(e){this.m_handlers.push(e)},e.prototype.removeHostMessageHandler=function(e){var t=this.m_handlers.indexOf(e);t>=0&&this.m_handlers.splice(t,1)},e.onInited=function(t){e.s_onInitedHandlers.push(t),e.s_instance&&t(e.s_instance)},e.prototype.dispatchMessage=function(e){if("number"==typeof e.id){var t=this.m_promiseResolver[e.id];if(t)return t(e),void delete this.m_promiseResolver[e.id]}for(var n=0;n<this.m_handlers.length;n++)this.m_handlers[n](e)},e.nextId=function(){return e.s_nextId++},e.s_onInitedHandlers=[],e.s_nextId=1,e}();t.HostBridge=c,"object"==typeof _richApiNativeBridge&&_richApiNativeBridge&&c.init(_richApiNativeBridge),function(e){var t=function(t){function n(e){var r=t.call(this,"string"==typeof e?e:e.message)||this;return Object.setPrototypeOf(r,n.prototype),r.name="RichApi.Error","string"==typeof e?r.message=e:(r.code=e.code,r.message=e.message,r.traceMessages=e.traceMessages||[],r.innerError=e.innerError||null,r.debugInfo=r._createDebugInfo(e.debugInfo||{})),r}return o(n,t),n.prototype.toString=function(){return this.code+": "+this.message},n.prototype._createDebugInfo=function(t){var n={code:this.code,message:this.message,toString:function(){return JSON.stringify(this)}};for(var r in t)n[r]=t[r];return this.innerError&&(this.innerError instanceof e.RuntimeError?n.innerError=this.innerError.debugInfo:n.innerError=this.innerError),n},n._createInvalidArgError=function(t){return new e.RuntimeError({code:u.invalidArgument,message:h.isNullOrEmptyString(t.argumentName)?h._getResourceString(l.invalidArgumentGeneric):h._getResourceString(l.invalidArgument,t.argumentName),debugInfo:t.errorLocation?{errorLocation:t.errorLocation}:{},innerError:t.innerError})},n}(Error);e.RuntimeError=t}(a=t._Internal||(t._Internal={})),t.Error=a.RuntimeError;var u=function(){function e(){}return e.apiNotFound="ApiNotFound",e.accessDenied="AccessDenied",e.generalException="GeneralException",e.activityLimitReached="ActivityLimitReached",e.invalidArgument="InvalidArgument",e.connectionFailure="ConnectionFailure",e.timeout="Timeout",e.invalidOrTimedOutSession="InvalidOrTimedOutSession",e.invalidObjectPath="InvalidObjectPath",e.invalidRequestContext="InvalidRequestContext",e.valueNotLoaded="ValueNotLoaded",e}();t.CoreErrorCodes=u;var l=function(){function e(){}return e.apiNotFoundDetails="ApiNotFoundDetails",e.connectionFailureWithStatus="ConnectionFailureWithStatus",e.connectionFailureWithDetails="ConnectionFailureWithDetails",e.invalidArgument="InvalidArgument",e.invalidArgumentGeneric="InvalidArgumentGeneric",e.timeout="Timeout",e.invalidOrTimedOutSessionMessage="InvalidOrTimedOutSessionMessage",e.invalidObjectPath="InvalidObjectPath",e.invalidRequestContext="InvalidRequestContext",e.valueNotLoaded="ValueNotLoaded",e}();t.CoreResourceStrings=l;var d=function(){function e(){}return e.flags="flags",e.sourceLibHeader="SdkVersion",e.processQuery="ProcessQuery",e.localDocument="http://document.localhost/",e.localDocumentApiPrefix="http://document.localhost/_api/",e.customData="customdata",e}();t.CoreConstants=d;var f=function(){function e(){}return e.buildMessageArrayForIRequestExecutor=function(t,n,r,o){var i=JSON.stringify(r.Body);h.log("Request:"),h.log(i);var s={};return s[d.sourceLibHeader]=o,e.buildRequestMessageSafeArray(t,n,"POST",d.processQuery,s,i)},e.buildResponseOnSuccess=function(e,t){var n={ErrorCode:"",ErrorMessage:"",Headers:null,Body:null};return n.Body=JSON.parse(e),n.Headers=t,n},e.buildResponseOnError=function(t,n){var r={ErrorCode:"",ErrorMessage:"",Headers:null,Body:null};return r.ErrorCode=u.generalException,r.ErrorMessage=n,t==e.OfficeJsErrorCode_ooeNoCapability?r.ErrorCode=u.accessDenied:t==e.OfficeJsErrorCode_ooeActivityLimitReached?r.ErrorCode=u.activityLimitReached:t==e.OfficeJsErrorCode_ooeInvalidOrTimedOutSession&&(r.ErrorCode=u.invalidOrTimedOutSession,r.ErrorMessage=h._getResourceString(l.invalidOrTimedOutSessionMessage)),r},e.buildHttpResponseFromOfficeJsError=function(t,n){var r=500,o={error:{}};return o.error.code=u.generalException,o.error.message=n,t===e.OfficeJsErrorCode_ooeNoCapability?(r=403,o.error.code=u.accessDenied):t===e.OfficeJsErrorCode_ooeActivityLimitReached&&(r=429,o.error.code=u.activityLimitReached),{statusCode:r,headers:{},body:JSON.stringify(o)}},e.buildRequestMessageSafeArray=function(e,t,n,r,o,i){var s=[];if(o)for(var a in o)s.push(a),s.push(o[a]);return[e,n,r,s,i,0,t,"","",""]},e.getResponseBody=function(t){return e.getResponseBodyFromSafeArray(t.value.data)},e.getResponseHeaders=function(t){return e.getResponseHeadersFromSafeArray(t.value.data)},e.getResponseBodyFromSafeArray=function(e){var t=e[2];return"string"==typeof t?t:t.join("")},e.getResponseHeadersFromSafeArray=function(e){var t=e[1];if(!t)return null;for(var n={},r=0;r<t.length-1;r+=2)n[t[r]]=t[r+1];return n},e.getResponseStatusCode=function(t){return e.getResponseStatusCodeFromSafeArray(t.value.data)},e.getResponseStatusCodeFromSafeArray=function(e){return e[0]},e.OfficeJsErrorCode_ooeInvalidOrTimedOutSession=5012,e.OfficeJsErrorCode_ooeActivityLimitReached=5102,e.OfficeJsErrorCode_ooeNoCapability=7e3,e}();t.RichApiMessageUtility=f,function(e){e.getPromiseType=function(){if("undefined"!=typeof Promise)return Promise;if("undefined"!=typeof Office&&Office.Promise)return Office.Promise;if("undefined"!=typeof OfficeExtension&&OfficeExtension.Promise)return OfficeExtension.Promise;throw new e.Error("No Promise implementation found")}}(a=t._Internal||(t._Internal={}));var h=function(){function e(){}return e.log=function(t){e._logEnabled&&"undefined"!=typeof console&&console.log&&console.log(t)},e.checkArgumentNull=function(t,n){if(e.isNullOrUndefined(t))throw a.RuntimeError._createInvalidArgError({argumentName:n})},e.isNullOrUndefined=function(e){return null==e},e.isUndefined=function(e){return void 0===e},e.isNullOrEmptyString=function(e){return null==e||0==e.length},e.isPlainJsonObject=function(t){if(e.isNullOrUndefined(t))return!1;if("object"!=typeof t)return!1;if("[object Object]"!==Object.prototype.toString.apply(t))return!1;if(t.constructor&&!Object.prototype.hasOwnProperty.call(t,"constructor")&&!Object.prototype.hasOwnProperty.call(t.constructor.prototype,"hasOwnProperty"))return!1;for(var n in t)if(!Object.prototype.hasOwnProperty.call(t,n))return!1;return!0},e.trim=function(e){return e.replace(new RegExp("^\\s+|\\s+$","g"),"")},e.caseInsensitiveCompareString=function(t,n){return e.isNullOrUndefined(t)?e.isNullOrUndefined(n):!e.isNullOrUndefined(n)&&t.toUpperCase()==n.toUpperCase()},e.isReadonlyRestRequest=function(t){return e.caseInsensitiveCompareString(t,"GET")},e._getResourceString=function(t,n){var r;if("undefined"!=typeof window&&window.Strings&&window.Strings.OfficeOM){var o="L_"+t,i=window.Strings.OfficeOM[o];i&&(r=i)}if(r||(r=e.s_resourceStringValues[t]),r||(r=t),!e.isNullOrUndefined(n))if(Array.isArray(n)){var s=n;r=e._formatString(r,s)}else r=r.replace("{0}",n);return r},e._formatString=function(e,t){return e.replace(/\{\d\}/g,(function(e){var n=parseInt(e.substr(1,e.length-2));if(n<t.length)return t[n];throw a.RuntimeError._createInvalidArgError({argumentName:"format"})}))},Object.defineProperty(e,"Promise",{get:function(){return a.getPromiseType()},enumerable:!0,configurable:!0}),e.createPromise=function(t){return new e.Promise(t)},e._createPromiseFromResult=function(t){return e.createPromise((function(e,n){e(t)}))},e._createPromiseFromException=function(t){return e.createPromise((function(e,n){n(t)}))},e._createTimeoutPromise=function(t){return e.createPromise((function(e,n){setTimeout((function(){e(null)}),t)}))},e._createInvalidArgError=function(e){return a.RuntimeError._createInvalidArgError(e)},e._isLocalDocumentUrl=function(t){return e._getLocalDocumentUrlPrefixLength(t)>0},e._getLocalDocumentUrlPrefixLength=function(e){for(var t=["http://document.localhost","https://document.localhost","//document.localhost"],n=e.toLowerCase().trim(),r=0;r<t.length;r++){if(n===t[r])return t[r].length;if(n.substr(0,t[r].length+1)===t[r]+"/")return t[r].length+1}return 0},e._validateLocalDocumentRequest=function(t){var n=e._getLocalDocumentUrlPrefixLength(t.url);if(n<=0)throw a.RuntimeError._createInvalidArgError({argumentName:"request"});var r=t.url.substr(n),o=r.toLowerCase();return"_api"===o?r="":"_api/"===o.substr(0,"_api/".length)&&(r=r.substr("_api/".length)),{method:t.method,url:r,headers:t.headers,body:t.body}},e._parseRequestFlagsAndCustomDataFromQueryStringIfAny=function(e){for(var t={flags:-1,customData:""},n=e.split("&"),r=0;r<n.length;r++){var o=n[r].split("=");if(o[0].toLowerCase()===d.flags){var i=parseInt(o[1]);i&=4095,t.flags=i}else o[0].toLowerCase()===d.customData&&(t.customData=decodeURIComponent(o[1]))}return t},e._getRequestBodyText=function(e){var t="";return"string"==typeof e.body?t=e.body:e.body&&"object"==typeof e.body&&(t=JSON.stringify(e.body)),t},e._parseResponseBody=function(t){if("string"==typeof t.body){var n=e.trim(t.body);return JSON.parse(n)}return t.body},e._buildRequestMessageSafeArray=function(t){var n=0;e.isReadonlyRestRequest(t.method)||(n=1);var r="";if(t.url.substr(0,d.processQuery.length).toLowerCase()===d.processQuery.toLowerCase()){var o=t.url.indexOf("?");if(o>0){var i=t.url.substr(o+1),s=e._parseRequestFlagsAndCustomDataFromQueryStringIfAny(i);s.flags>=0&&(n=s.flags),r=s.customData}}return f.buildRequestMessageSafeArray(r,n,t.method,t.url,t.headers,e._getRequestBodyText(t))},e._parseHttpResponseHeaders=function(t){var n={};if(!e.isNullOrEmptyString(t))for(var r=new RegExp("\r?\n"),o=t.split(r),i=0;i<o.length;i++){var s=o[i];if(null!=s){var a=s.indexOf(":");if(a>0){var c=s.substr(0,a),u=s.substr(a+1);c=e.trim(c),u=e.trim(u),n[c.toUpperCase()]=u}}}return n},e._parseErrorResponse=function(t){var n,r,o=null;if(e.isPlainJsonObject(t.body))o=t.body;else if(!e.isNullOrEmptyString(t.body)){var i=e.trim(t.body);try{o=JSON.parse(i)}catch(t){e.log("Error when parse "+i)}}return!e.isNullOrUndefined(o)&&"object"==typeof o&&o.error?(r=o.error.code,n=e._getResourceString(l.connectionFailureWithDetails,[t.statusCode.toString(),o.error.code,o.error.message])):n=e._getResourceString(l.connectionFailureWithStatus,t.statusCode.toString()),e.isNullOrEmptyString(r)&&(r=u.connectionFailure),{errorCode:r,errorMessage:n}},e._copyHeaders=function(e,t){if(e&&t)for(var n in e)t[n]=e[n]},e.addResourceStringValues=function(t){for(var n in t)e.s_resourceStringValues[n]=t[n]},e._logEnabled=!1,e.s_resourceStringValues={ApiNotFoundDetails:"The method or property {0} is part of the {1} requirement set, which is not available in your version of {2}.",ConnectionFailureWithStatus:"The request failed with status code of {0}.",ConnectionFailureWithDetails:"The request failed with status code of {0}, error code {1} and the following error message: {2}",InvalidArgument:"The argument '{0}' doesn't work for this situation, is missing, or isn't in the right format.",InvalidObjectPath:'The object path \'{0}\' isn\'t working for what you\'re trying to do. If you\'re using the object across multiple "context.sync" calls and outside the sequential execution of a ".run" batch, please use the "context.trackedObjects.add()" and "context.trackedObjects.remove()" methods to manage the object\'s lifetime.',InvalidRequestContext:"Cannot use the object across different request contexts.",Timeout:"The operation has timed out.",ValueNotLoaded:'The value of the result object has not been loaded yet. Before reading the value property, call "context.sync()" on the associated request context.'},e}();t.CoreUtility=h;var p=function(){function e(){}return e.setMock=function(t){e.s_isMock=t},e.isMock=function(){return e.s_isMock},e}();t.TestUtility=p},function(e,t,n){"use strict";var r,o=this&&this.__extends||(r=function(e,t){return(r=Object.setPrototypeOf||{__proto__:[]}instanceof Array&&function(e,t){e.__proto__=t}||function(e,t){for(var n in t)t.hasOwnProperty(n)&&(e[n]=t[n])})(e,t)},function(e,t){function n(){this.constructor=e}r(e,t),e.prototype=null===t?Object.create(t):(n.prototype=t.prototype,new n)});Object.defineProperty(t,"__esModule",{value:!0});var i=n(0);!function(e){for(var n in e)t.hasOwnProperty(n)||(t[n]=e[n])}(n(0)),t._internalConfig={showDisposeInfoInDebugInfo:!1,showInternalApiInDebugInfo:!1,enableEarlyDispose:!0,alwaysPolyfillClientObjectUpdateMethod:!1,alwaysPolyfillClientObjectRetrieveMethod:!1,enableConcurrentFlag:!0,enableUndoableFlag:!0},t.config={extendedErrorLogging:!1};var s=function(){function e(){}return e.createSetPropertyAction=function(e,t,n,r,o){y.validateObjectPath(t);var i={Id:e._nextId(),ActionType:4,Name:n,ObjectPathId:t._objectPath.objectPathInfo.Id,ArgumentInfo:{}},s=[r],a=y.setMethodArguments(e,i.ArgumentInfo,s);y.validateReferencedObjectPaths(a);var u=new c(i,0,o);return u.referencedObjectPath=t._objectPath,u.referencedArgumentObjectPaths=a,t._addAction(u)},e.createQueryAction=function(e,t,n,r){y.validateObjectPath(t);var o={Id:e._nextId(),ActionType:2,Name:"",ObjectPathId:t._objectPath.objectPathInfo.Id,QueryInfo:n},i=new c(o,1,4);return i.referencedObjectPath=t._objectPath,t._addAction(i,r)},e.createQueryAsJsonAction=function(e,t,n,r){y.validateObjectPath(t);var o={Id:e._nextId(),ActionType:7,Name:"",ObjectPathId:t._objectPath.objectPathInfo.Id,QueryInfo:n},i=new c(o,1,4);return i.referencedObjectPath=t._objectPath,t._addAction(i,r)},e.createUpdateAction=function(e,t,n){y.validateObjectPath(t);var r={Id:e._nextId(),ActionType:9,Name:"",ObjectPathId:t._objectPath.objectPathInfo.Id,ObjectState:n},o=new c(r,0,0);return o.referencedObjectPath=t._objectPath,t._addAction(o)},e}();t.CommonActionFactory=s;var a=function(){function e(e,t){this.m_contextBase=e,this.m_objectPath=t}return Object.defineProperty(e.prototype,"_objectPath",{get:function(){return this.m_objectPath},enumerable:!0,configurable:!0}),Object.defineProperty(e.prototype,"_context",{get:function(){return this.m_contextBase},enumerable:!0,configurable:!0}),e.prototype._addAction=function(e,t){var n=this;return void 0===t&&(t=null),i.CoreUtility.createPromise((function(r,o){n._context._addServiceApiAction(e,t,r,o)}))},e.prototype._retrieve=function(e,n){var r=t._internalConfig.alwaysPolyfillClientObjectRetrieveMethod;r||(r=!y.isSetSupported("RichApiRuntime","1.1"));var o=l._parseQueryOption(e);return r?s.createQueryAction(this._context,this,o,n):s.createQueryAsJsonAction(this._context,this,o,n)},e.prototype._recursivelyUpdate=function(e){var n=t._internalConfig.alwaysPolyfillClientObjectUpdateMethod;n||(n=!y.isSetSupported("RichApiRuntime","1.2"));try{var r=this[g.scalarPropertyNames];r||(r=[]);var o=this[g.scalarPropertyUpdateable];if(!o){o=[];for(var a=0;a<r.length;a++)o.push(!1)}var c=this[g.navigationPropertyNames];c||(c=[]);var u={},l={},d=0;for(var f in e){var h=r.indexOf(f);if(h>=0){if(!o[h])throw new i._Internal.RuntimeError({code:i.CoreErrorCodes.invalidArgument,message:i.CoreUtility._getResourceString(_.attemptingToSetReadOnlyProperty,f),debugInfo:{errorLocation:f}});u[f]=e[f],++d}else{if(!(c.indexOf(f)>=0))throw new i._Internal.RuntimeError({code:i.CoreErrorCodes.invalidArgument,message:i.CoreUtility._getResourceString(_.propertyDoesNotExist,f),debugInfo:{errorLocation:f}});l[f]=e[f]}}if(d>0)if(n)for(a=0;a<r.length;a++){var p=u[f=r[a]];y.isUndefined(p)||s.createSetPropertyAction(this._context,this,f,p)}else s.createUpdateAction(this._context,this,u);for(var f in l){var m=this[f],b=l[f];m._recursivelyUpdate(b)}}catch(e){throw new i._Internal.RuntimeError({code:i.CoreErrorCodes.invalidArgument,message:i.CoreUtility._getResourceString(i.CoreResourceStrings.invalidArgument,"properties"),debugInfo:{errorLocation:this._className+".update"},innerError:e})}},e}();t.ClientObjectBase=a;var c=function(){function e(e,t,n){this.m_actionInfo=e,this.m_operationType=t,this.m_flags=n}return Object.defineProperty(e.prototype,"actionInfo",{get:function(){return this.m_actionInfo},enumerable:!0,configurable:!0}),Object.defineProperty(e.prototype,"operationType",{get:function(){return this.m_operationType},enumerable:!0,configurable:!0}),Object.defineProperty(e.prototype,"flags",{get:function(){return this.m_flags},enumerable:!0,configurable:!0}),e}();t.Action=c;var u=function(){function e(e,t,n,r,o,i){this.m_objectPathInfo=e,this.m_parentObjectPath=t,this.m_isCollection=n,this.m_isInvalidAfterRequest=r,this.m_isValid=!0,this.m_operationType=o,this.m_flags=i}return Object.defineProperty(e.prototype,"id",{get:function(){var e=this.m_objectPathInfo.ArgumentInfo;if(e){var t=e.Arguments;if(t)return t[0]}},enumerable:!0,configurable:!0}),Object.defineProperty(e.prototype,"parent",{get:function(){var e=this.m_parentObjectPath;if(e)return e},enumerable:!0,configurable:!0}),Object.defineProperty(e.prototype,"parentId",{get:function(){return this.parent?this.parent.id:void 0},enumerable:!0,configurable:!0}),Object.defineProperty(e.prototype,"objectPathInfo",{get:function(){return this.m_objectPathInfo},enumerable:!0,configurable:!0}),Object.defineProperty(e.prototype,"operationType",{get:function(){return this.m_operationType},enumerable:!0,configurable:!0}),Object.defineProperty(e.prototype,"flags",{get:function(){return this.m_flags},enumerable:!0,configurable:!0}),Object.defineProperty(e.prototype,"isCollection",{get:function(){return this.m_isCollection},enumerable:!0,configurable:!0}),Object.defineProperty(e.prototype,"isInvalidAfterRequest",{get:function(){return this.m_isInvalidAfterRequest},enumerable:!0,configurable:!0}),Object.defineProperty(e.prototype,"parentObjectPath",{get:function(){return this.m_parentObjectPath},enumerable:!0,configurable:!0}),Object.defineProperty(e.prototype,"argumentObjectPaths",{get:function(){return this.m_argumentObjectPaths},set:function(e){this.m_argumentObjectPaths=e},enumerable:!0,configurable:!0}),Object.defineProperty(e.prototype,"isValid",{get:function(){return this.m_isValid},set:function(t){this.m_isValid=t,!t&&6===this.m_objectPathInfo.ObjectPathType&&this.m_savedObjectPathInfo&&(e.copyObjectPathInfo(this.m_savedObjectPathInfo.pathInfo,this.m_objectPathInfo),this.m_parentObjectPath=this.m_savedObjectPathInfo.parent,this.m_isValid=!0,this.m_savedObjectPathInfo=null)},enumerable:!0,configurable:!0}),Object.defineProperty(e.prototype,"originalObjectPathInfo",{get:function(){return this.m_originalObjectPathInfo},enumerable:!0,configurable:!0}),Object.defineProperty(e.prototype,"getByIdMethodName",{get:function(){return this.m_getByIdMethodName},set:function(e){this.m_getByIdMethodName=e},enumerable:!0,configurable:!0}),e.prototype._updateAsNullObject=function(){this.resetForUpdateUsingObjectData(),this.m_objectPathInfo.ObjectPathType=7,this.m_objectPathInfo.Name="",this.m_parentObjectPath=null},e.prototype.saveOriginalObjectPathInfo=function(){t.config.extendedErrorLogging&&!this.m_originalObjectPathInfo&&(this.m_originalObjectPathInfo={},e.copyObjectPathInfo(this.m_objectPathInfo,this.m_originalObjectPathInfo))},e.prototype.updateUsingObjectData=function(t,n){var r=t[g.referenceId];if(!i.CoreUtility.isNullOrEmptyString(r)){if(!this.m_savedObjectPathInfo&&!this.isInvalidAfterRequest&&e.isRestorableObjectPath(this.m_objectPathInfo.ObjectPathType)){var o={};e.copyObjectPathInfo(this.m_objectPathInfo,o),this.m_savedObjectPathInfo={pathInfo:o,parent:this.m_parentObjectPath}}return this.saveOriginalObjectPathInfo(),this.resetForUpdateUsingObjectData(),this.m_objectPathInfo.ObjectPathType=6,this.m_objectPathInfo.Name=r,delete this.m_objectPathInfo.ParentObjectPathId,void(this.m_parentObjectPath=null)}if(n){var s=n[g.collectionPropertyPath];if(!i.CoreUtility.isNullOrEmptyString(s)&&n.context){var a=y.tryGetObjectIdFromLoadOrRetrieveResult(t);if(!i.CoreUtility.isNullOrUndefined(a)){for(var c=s.split("."),u=n.context[c[0]],l=1;l<c.length;l++)u=u[c[l]];return this.saveOriginalObjectPathInfo(),this.resetForUpdateUsingObjectData(),this.m_parentObjectPath=u._objectPath,this.m_objectPathInfo.ParentObjectPathId=this.m_parentObjectPath.objectPathInfo.Id,this.m_objectPathInfo.ObjectPathType=5,this.m_objectPathInfo.Name="",void(this.m_objectPathInfo.ArgumentInfo.Arguments=[a])}}}var d=this.parentObjectPath&&this.parentObjectPath.isCollection,f=this.getByIdMethodName;if((d||!i.CoreUtility.isNullOrEmptyString(f))&&(a=y.tryGetObjectIdFromLoadOrRetrieveResult(t),!i.CoreUtility.isNullOrUndefined(a)))return this.saveOriginalObjectPathInfo(),this.resetForUpdateUsingObjectData(),i.CoreUtility.isNullOrEmptyString(f)?(this.m_objectPathInfo.ObjectPathType=5,this.m_objectPathInfo.Name=""):(this.m_objectPathInfo.ObjectPathType=3,this.m_objectPathInfo.Name=f),void(this.m_objectPathInfo.ArgumentInfo.Arguments=[a])},e.prototype.resetForUpdateUsingObjectData=function(){this.m_isInvalidAfterRequest=!1,this.m_isValid=!0,this.m_operationType=1,this.m_flags=4,this.m_objectPathInfo.ArgumentInfo={},this.m_argumentObjectPaths=null,this.m_getByIdMethodName=null},e.isRestorableObjectPath=function(e){return 1===e||5===e||3===e||4===e},e.copyObjectPathInfo=function(e,t){t.Id=e.Id,t.ArgumentInfo=e.ArgumentInfo,t.Name=e.Name,t.ObjectPathType=e.ObjectPathType,t.ParentObjectPathId=e.ParentObjectPathId},e}();t.ObjectPath=u;var l=function(){function e(){this.m_nextId=0}return e.prototype._nextId=function(){return++this.m_nextId},e.prototype._addServiceApiAction=function(e,t,n,r){this.m_serviceApiQueue||(this.m_serviceApiQueue=new p(this)),this.m_serviceApiQueue.add(e,t,n,r)},e._parseQueryOption=function(t){var n={};if("string"==typeof t){var r=t;n.Select=y._parseSelectExpand(r)}else if(Array.isArray(t))n.Select=t;else if("object"==typeof t){var o=t;if(e.isLoadOption(o)){if("string"==typeof o.select)n.Select=y._parseSelectExpand(o.select);else if(Array.isArray(o.select))n.Select=o.select;else if(!y.isNullOrUndefined(o.select))throw i._Internal.RuntimeError._createInvalidArgError({argumentName:"option.select"});if("string"==typeof o.expand)n.Expand=y._parseSelectExpand(o.expand);else if(Array.isArray(o.expand))n.Expand=o.expand;else if(!y.isNullOrUndefined(o.expand))throw i._Internal.RuntimeError._createInvalidArgError({argumentName:"option.expand"});if("number"==typeof o.top)n.Top=o.top;else if(!y.isNullOrUndefined(o.top))throw i._Internal.RuntimeError._createInvalidArgError({argumentName:"option.top"});if("number"==typeof o.skip)n.Skip=o.skip;else if(!y.isNullOrUndefined(o.skip))throw i._Internal.RuntimeError._createInvalidArgError({argumentName:"option.skip"})}else n=e.parseStrictLoadOption(t)}else if(!y.isNullOrUndefined(t))throw i._Internal.RuntimeError._createInvalidArgError({argumentName:"option"});return n},e.isLoadOption=function(e){if(!y.isUndefined(e.select)&&("string"==typeof e.select||Array.isArray(e.select)))return!0;if(!y.isUndefined(e.expand)&&("string"==typeof e.expand||Array.isArray(e.expand)))return!0;if(!y.isUndefined(e.top)&&"number"==typeof e.top)return!0;if(!y.isUndefined(e.skip)&&"number"==typeof e.skip)return!0;for(var t in e)return!1;return!0},e.parseStrictLoadOption=function(t){var n={Select:[]};return e.parseStrictLoadOptionHelper(n,"","option",t),n},e.combineQueryPath=function(e,t,n){return 0===e.length?t:e+n+t},e.parseStrictLoadOptionHelper=function(t,n,r,o){for(var s in o){var a=o[s];if("$all"===s){if("boolean"!=typeof a)throw i._Internal.RuntimeError._createInvalidArgError({argumentName:e.combineQueryPath(r,s,".")});a&&t.Select.push(e.combineQueryPath(n,"*","/"))}else if("$top"===s){if("number"!=typeof a||n.length>0)throw i._Internal.RuntimeError._createInvalidArgError({argumentName:e.combineQueryPath(r,s,".")});t.Top=a}else if("$skip"===s){if("number"!=typeof a||n.length>0)throw i._Internal.RuntimeError._createInvalidArgError({argumentName:e.combineQueryPath(r,s,".")});t.Skip=a}else if("boolean"==typeof a)a&&t.Select.push(e.combineQueryPath(n,s,"/"));else{if("object"!=typeof a)throw i._Internal.RuntimeError._createInvalidArgError({argumentName:e.combineQueryPath(r,s,".")});e.parseStrictLoadOptionHelper(t,e.combineQueryPath(n,s,"/"),e.combineQueryPath(r,s,"."),a)}}},e}();t.ClientRequestContextBase=l;var d=function(){function e(e){this.m_objectPath=e}return e.prototype._handleResult=function(e){i.CoreUtility.isNullOrUndefined(e)?this.m_objectPath._updateAsNullObject():this.m_objectPath.updateUsingObjectData(e,null)},e}(),f=function(){function e(e){this.m_contextBase=e,this.m_actions=[],this.m_actionResultHandler={},this.m_referencedObjectPaths={},this.m_instantiatedObjectPaths={},this.m_preSyncPromises=[]}return e.prototype.addAction=function(e){this.m_actions.push(e),1==e.actionInfo.ActionType&&(this.m_instantiatedObjectPaths[e.actionInfo.ObjectPathId]=e)},Object.defineProperty(e.prototype,"hasActions",{get:function(){return this.m_actions.length>0},enumerable:!0,configurable:!0}),e.prototype._getLastAction=function(){return this.m_actions[this.m_actions.length-1]},e.prototype.ensureInstantiateObjectPath=function(e){if(e){if(this.m_instantiatedObjectPaths[e.objectPathInfo.Id])return;if(this.ensureInstantiateObjectPath(e.parentObjectPath),this.ensureInstantiateObjectPaths(e.argumentObjectPaths),!this.m_instantiatedObjectPaths[e.objectPathInfo.Id]){var t={Id:this.m_contextBase._nextId(),ActionType:1,Name:"",ObjectPathId:e.objectPathInfo.Id},n=new c(t,1,4);n.referencedObjectPath=e,this.addReferencedObjectPath(e),this.addAction(n);var r=new d(e);this.addActionResultHandler(n,r)}}},e.prototype.ensureInstantiateObjectPaths=function(e){if(e)for(var t=0;t<e.length;t++)this.ensureInstantiateObjectPath(e[t])},e.prototype.addReferencedObjectPath=function(e){if(e&&!this.m_referencedObjectPaths[e.objectPathInfo.Id]){if(!e.isValid)throw new i._Internal.RuntimeError({code:i.CoreErrorCodes.invalidObjectPath,message:i.CoreUtility._getResourceString(i.CoreResourceStrings.invalidObjectPath,y.getObjectPathExpression(e)),debugInfo:{errorLocation:y.getObjectPathExpression(e)}});for(;e;)this.m_referencedObjectPaths[e.objectPathInfo.Id]=e,3==e.objectPathInfo.ObjectPathType&&this.addReferencedObjectPaths(e.argumentObjectPaths),e=e.parentObjectPath}},e.prototype.addReferencedObjectPaths=function(e){if(e)for(var t=0;t<e.length;t++)this.addReferencedObjectPath(e[t])},e.prototype.addActionResultHandler=function(e,t){this.m_actionResultHandler[e.actionInfo.Id]=t},e.prototype.aggregrateRequestFlags=function(e,t,n){return 0===t&&(e|=1,0==(2&n)&&(e&=-17),0==(8&n)&&(e&=-257),e&=-5),1&n&&(e|=2),0==(4&n)&&(e&=-5),e},e.prototype.finallyNormalizeFlags=function(e){return 0==(1&e)&&(e&=-17,e&=-257),t._internalConfig.enableConcurrentFlag||(e&=-5),t._internalConfig.enableUndoableFlag||(e&=-17),y.isSetSupported("RichApiRuntimeFlag","1.1")||(e&=-5,e&=-17),y.isSetSupported("RichApiRuntimeFlag","1.2")||(e&=-257),"number"==typeof this.m_flagsForTesting&&(e=this.m_flagsForTesting),e},e.prototype.buildRequestMessageBodyAndRequestFlags=function(){t._internalConfig.enableEarlyDispose&&e._calculateLastUsedObjectPathIds(this.m_actions);var n=276,r={};for(var o in this.m_referencedObjectPaths)n=this.aggregrateRequestFlags(n,this.m_referencedObjectPaths[o].operationType,this.m_referencedObjectPaths[o].flags),r[o]=this.m_referencedObjectPaths[o].objectPathInfo;for(var i=[],s=!1,a=0;a<this.m_actions.length;a++){var c=this.m_actions[a];3===c.actionInfo.ActionType&&c.actionInfo.Name===g.keepReference&&(s=!0),n=this.aggregrateRequestFlags(n,c.operationType,c.flags),i.push(c.actionInfo)}return n=this.finallyNormalizeFlags(n),{body:{AutoKeepReference:this.m_contextBase._autoCleanup&&s,Actions:i,ObjectPaths:r},flags:n}},e.prototype.processResponse=function(e){if(e)for(var t=0;t<e.length;t++){var n=e[t],r=this.m_actionResultHandler[n.ActionId];r&&r._handleResult(n.Value)}},e.prototype.invalidatePendingInvalidObjectPaths=function(){for(var e in this.m_referencedObjectPaths)this.m_referencedObjectPaths[e].isInvalidAfterRequest&&(this.m_referencedObjectPaths[e].isValid=!1)},e.prototype._addPreSyncPromise=function(e){this.m_preSyncPromises.push(e)},Object.defineProperty(e.prototype,"_preSyncPromises",{get:function(){return this.m_preSyncPromises},enumerable:!0,configurable:!0}),Object.defineProperty(e.prototype,"_actions",{get:function(){return this.m_actions},enumerable:!0,configurable:!0}),Object.defineProperty(e.prototype,"_objectPaths",{get:function(){return this.m_referencedObjectPaths},enumerable:!0,configurable:!0}),e.prototype._removeKeepReferenceAction=function(e){for(var t=this.m_actions.length-1;t>=0;t--){var n=this.m_actions[t].actionInfo;if(n.ObjectPathId===e&&3===n.ActionType&&n.Name===g.keepReference){this.m_actions.splice(t,1);break}}},e._updateLastUsedActionIdOfObjectPathId=function(t,n,r){for(;n;){if(t[n.objectPathInfo.Id])return;t[n.objectPathInfo.Id]=r;var o=n.argumentObjectPaths;if(o)for(var i=o.length,s=0;s<i;s++)e._updateLastUsedActionIdOfObjectPathId(t,o[s],r);n=n.parentObjectPath}},e._calculateLastUsedObjectPathIds=function(t){for(var n={},r=t.length,o=r-1;o>=0;--o){var i=(f=t[o]).actionInfo.Id;f.referencedObjectPath&&e._updateLastUsedActionIdOfObjectPathId(n,f.referencedObjectPath,i);var s=f.referencedArgumentObjectPaths;if(s)for(var a=s.length,c=0;c<a;c++)e._updateLastUsedActionIdOfObjectPathId(n,s[c],i)}var u={};for(var l in n){var d=u[i=n[l]];d||(d=[],u[i]=d),d.push(parseInt(l))}for(o=0;o<r;o++){var f,h=u[(f=t[o]).actionInfo.Id];h&&h.length>0?f.actionInfo.L=h:f.actionInfo.L&&delete f.actionInfo.L}},e}();t.ClientRequestBase=f;var h=function(){function e(e){this.m_type=e}return Object.defineProperty(e.prototype,"value",{get:function(){if(!this.m_isLoaded)throw new i._Internal.RuntimeError({code:i.CoreErrorCodes.valueNotLoaded,message:i.CoreUtility._getResourceString(i.CoreResourceStrings.valueNotLoaded),debugInfo:{errorLocation:"clientResult.value"}});return this.m_value},enumerable:!0,configurable:!0}),e.prototype._handleResult=function(e){this.m_isLoaded=!0,"object"==typeof e&&e&&e._IsNull||(1===this.m_type?this.m_value=y.adjustToDateTime(e):this.m_value=e)},e}();t.ClientResult=h;var p=function(){function e(e){this.m_context=e,this.m_actions=[]}return e.prototype.add=function(e,t,n,r){var o=this;this.m_actions.push({action:e,resultHandler:t,resolve:n,reject:r}),1===this.m_actions.length&&setTimeout((function(){return o.processActions()}),0)},e.prototype.processActions=function(){var e=this;if(0!==this.m_actions.length){var t=this.m_actions;this.m_actions=[];for(var n=new f(this.m_context),r=0;r<t.length;r++){var o=t[r];n.ensureInstantiateObjectPath(o.action.referencedObjectPath),n.ensureInstantiateObjectPaths(o.action.referencedArgumentObjectPaths),n.addAction(o.action),n.addReferencedObjectPath(o.action.referencedObjectPath),n.addReferencedObjectPaths(o.action.referencedArgumentObjectPaths)}var s=n.buildRequestMessageBodyAndRequestFlags(),a=s.body,c=s.flags,u={Url:i.CoreConstants.localDocumentApiPrefix,Headers:null,Body:a};i.CoreUtility.log("Request:"),i.CoreUtility.log(JSON.stringify(a)),(new m).executeAsync(this.m_context._customData,c,u).then((function(r){e.processResponse(n,t,r)})).catch((function(e){for(var n=0;n<t.length;n++)t[n].reject(e)}))}},e.prototype.processResponse=function(e,t,n){var r=this.getErrorFromResponse(n),o=null;n.Body.Results?o=n.Body.Results:n.Body.ProcessedResults&&n.Body.ProcessedResults.Results&&(o=n.Body.ProcessedResults.Results),o||(o=[]),this.processActionResults(e,t,o,r)},e.prototype.getErrorFromResponse=function(e){return i.CoreUtility.isNullOrEmptyString(e.ErrorCode)?e.Body&&e.Body.Error?new i._Internal.RuntimeError({code:e.Body.Error.Code,message:e.Body.Error.Message}):null:new i._Internal.RuntimeError({code:e.ErrorCode,message:e.ErrorMessage})},e.prototype.processActionResults=function(e,t,n,r){e.processResponse(n);for(var o=0;o<t.length;o++){for(var i=t[o],s=i.action.actionInfo.Id,a=!1,c=0;c<n.length;c++)if(s==n[c].ActionId){var u=n[c].Value;i.resultHandler&&(i.resultHandler._handleResult(u),u=i.resultHandler.value),i.resolve&&i.resolve(u),a=!0;break}!a&&i.reject&&(r?i.reject(r):i.reject("No response for the action."))}},e}(),m=function(){function e(){}return e.prototype.getRequestUrl=function(e,t){return"/"!=e.charAt(e.length-1)&&(e+="/"),(e+=i.CoreConstants.processQuery)+"?"+i.CoreConstants.flags+"="+t.toString()},e.prototype.executeAsync=function(t,n,r){var o={method:"POST",url:this.getRequestUrl(r.Url,n),headers:{},body:r.Body};if(o.headers[i.CoreConstants.sourceLibHeader]=e.SourceLibHeaderValue,o.headers["CONTENT-TYPE"]="application/json",r.Headers)for(var s in r.Headers)o.headers[s]=r.Headers[s];return(i.CoreUtility._isLocalDocumentUrl(o.url)?i.HttpUtility.sendLocalDocumentRequest:i.HttpUtility.sendRequest)(o).then((function(e){var t;if(200===e.statusCode)t={ErrorCode:null,ErrorMessage:null,Headers:e.headers,Body:i.CoreUtility._parseResponseBody(e)};else{i.CoreUtility.log("Error Response:"+e.body);var n=i.CoreUtility._parseErrorResponse(e);t={ErrorCode:n.errorCode,ErrorMessage:n.errorMessage,Headers:e.headers,Body:null}}return t}))},e.SourceLibHeaderValue="officejs-rest",e}();t.HttpRequestExecutor=m;var g=function(e){function t(){return null!==e&&e.apply(this,arguments)||this}return o(t,e),t.collectionPropertyPath="_collectionPropertyPath",t.id="Id",t.idLowerCase="id",t.idPrivate="_Id",t.keepReference="_KeepReference",t.objectPathIdPrivate="_ObjectPathId",t.referenceId="_ReferenceId",t.items="_Items",t.itemsLowerCase="items",t.scalarPropertyNames="_scalarPropertyNames",t.scalarPropertyOriginalNames="_scalarPropertyOriginalNames",t.navigationPropertyNames="_navigationPropertyNames",t.scalarPropertyUpdateable="_scalarPropertyUpdateable",t}(i.CoreConstants);t.CommonConstants=g;var y=function(e){function t(){return null!==e&&e.apply(this,arguments)||this}return o(t,e),t.validateObjectPath=function(e){for(var n=e._objectPath;n;){if(!n.isValid)throw new i._Internal.RuntimeError({code:i.CoreErrorCodes.invalidObjectPath,message:i.CoreUtility._getResourceString(i.CoreResourceStrings.invalidObjectPath,t.getObjectPathExpression(n)),debugInfo:{errorLocation:t.getObjectPathExpression(n)}});n=n.parentObjectPath}},t.validateReferencedObjectPaths=function(e){if(e)for(var n=0;n<e.length;n++)for(var r=e[n];r;){if(!r.isValid)throw new i._Internal.RuntimeError({code:i.CoreErrorCodes.invalidObjectPath,message:i.CoreUtility._getResourceString(i.CoreResourceStrings.invalidObjectPath,t.getObjectPathExpression(r))});r=r.parentObjectPath}},t._toCamelLowerCase=function(e){if(i.CoreUtility.isNullOrEmptyString(e))return e;for(var t=0;t<e.length&&e.charCodeAt(t)>=65&&e.charCodeAt(t)<=90;)t++;return t<e.length?e.substr(0,t).toLowerCase()+e.substr(t):e.toLowerCase()},t.adjustToDateTime=function(e){if(i.CoreUtility.isNullOrUndefined(e))return null;if("string"==typeof e)return new Date(e);if(Array.isArray(e)){for(var n=e,r=0;r<n.length;r++)n[r]=t.adjustToDateTime(n[r]);return n}throw i.CoreUtility._createInvalidArgError({argumentName:"date"})},t.tryGetObjectIdFromLoadOrRetrieveResult=function(e){var t=e[g.id];return i.CoreUtility.isNullOrUndefined(t)&&(t=e[g.idLowerCase]),i.CoreUtility.isNullOrUndefined(t)&&(t=e[g.idPrivate]),t},t.getObjectPathExpression=function(e){for(var n="";e;){switch(e.objectPathInfo.ObjectPathType){case 1:n=n;break;case 2:n="new()"+(n.length>0?".":"")+n;break;case 3:n=t.normalizeName(e.objectPathInfo.Name)+"()"+(n.length>0?".":"")+n;break;case 4:n=t.normalizeName(e.objectPathInfo.Name)+(n.length>0?".":"")+n;break;case 5:n="getItem()"+(n.length>0?".":"")+n;break;case 6:n="_reference()"+(n.length>0?".":"")+n}e=e.parentObjectPath}return n},t.setMethodArguments=function(e,n,r){if(i.CoreUtility.isNullOrUndefined(r))return null;var o=new Array,s=new Array,a=t.collectObjectPathInfos(e,r,o,s);return n.Arguments=r,a&&(n.ReferencedObjectPathIds=s),o},t.validateContext=function(e,t){if(e&&t&&t._context!==e)throw new i._Internal.RuntimeError({code:i.CoreErrorCodes.invalidRequestContext,message:i.CoreUtility._getResourceString(i.CoreResourceStrings.invalidRequestContext)})},t.isSetSupported=function(e,t){return!("undefined"!=typeof window&&window.Office&&window.Office.context&&window.Office.context.requirements)||window.Office.context.requirements.isSetSupported(e,t)},t.throwIfApiNotSupported=function(e,n,r,o){if(t._doApiNotSupportedCheck&&!t.isSetSupported(n,r)){var s=i.CoreUtility._getResourceString(i.CoreResourceStrings.apiNotFoundDetails,[e,n+" "+r,o]);throw new i._Internal.RuntimeError({code:i.CoreErrorCodes.apiNotFound,message:s,debugInfo:{errorLocation:e}})}},t._parseSelectExpand=function(e){var t=[];if(!i.CoreUtility.isNullOrEmptyString(e))for(var n=e.split(","),r=0;r<n.length;r++){var o=n[r];(o=s(o.trim())).length>0&&t.push(o)}return t;function s(e){var t=e.toLowerCase();return"items"===t||"items/"===t?"*":(("items/"===t.substr(0,6)||"items."===t.substr(0,6))&&(e=e.substr(6)),e.replace(new RegExp("[/.]items[/.]","gi"),"/"))}},t.changePropertyNameToCamelLowerCase=function(e){if(Array.isArray(e)){for(var n=[],r=0;r<e.length;r++)n.push(this.changePropertyNameToCamelLowerCase(e[r]));return n}if("object"==typeof e&&null!==e){for(var o in n={},e){var i=e[o];if(o===g.items){(n={})[g.itemsLowerCase]=this.changePropertyNameToCamelLowerCase(i);break}n[t._toCamelLowerCase(o)]=this.changePropertyNameToCamelLowerCase(i)}return n}return e},t.purifyJson=function(e){if(Array.isArray(e)){for(var t=[],n=0;n<e.length;n++)t.push(this.purifyJson(e[n]));return t}if("object"==typeof e&&null!==e){for(var r in t={},e)if(95!==r.charCodeAt(0)){var o=e[r];"object"==typeof o&&null!==o&&Array.isArray(o.items)&&(o=o.items),t[r]=this.purifyJson(o)}return t}return e},t.collectObjectPathInfos=function(e,n,r,o){for(var s=!1,c=0;c<n.length;c++)if(n[c]instanceof a){var u=n[c];t.validateContext(e,u),n[c]=u._objectPath.objectPathInfo.Id,o.push(u._objectPath.objectPathInfo.Id),r.push(u._objectPath),s=!0}else if(Array.isArray(n[c])){var l=new Array;t.collectObjectPathInfos(e,n[c],r,l)?(o.push(l),s=!0):o.push(0)}else i.CoreUtility.isPlainJsonObject(n[c])?(o.push(0),t.replaceClientObjectPropertiesWithObjectPathIds(n[c],r)):o.push(0);return s},t.replaceClientObjectPropertiesWithObjectPathIds=function(e,n){var r,o;for(var s in e){var c=e[s];if(c instanceof a)n.push(c._objectPath),e[s]=((r={})[g.objectPathIdPrivate]=c._objectPath.objectPathInfo.Id,r);else if(Array.isArray(c))for(var u=0;u<c.length;u++)if(c[u]instanceof a){var l=c[u];n.push(l._objectPath),c[u]=((o={})[g.objectPathIdPrivate]=l._objectPath.objectPathInfo.Id,o)}else i.CoreUtility.isPlainJsonObject(c[u])&&t.replaceClientObjectPropertiesWithObjectPathIds(c[u],n);else i.CoreUtility.isPlainJsonObject(c)&&t.replaceClientObjectPropertiesWithObjectPathIds(c,n)}},t.normalizeName=function(e){return e.substr(0,1).toLowerCase()+e.substr(1)},t._doApiNotSupportedCheck=!1,t}(i.CoreUtility);t.CommonUtility=y;var _=function(e){function t(){return null!==e&&e.apply(this,arguments)||this}return o(t,e),t.propertyDoesNotExist="PropertyDoesNotExist",t.attemptingToSetReadOnlyProperty="AttemptingToSetReadOnlyProperty",t}(i.CoreResourceStrings);t.CommonResourceStrings=_},function(e,t,n){"use strict";Object.defineProperty(t,"__esModule",{value:!0});var r=n(0);t.CoreUtility=r.CoreUtility,t.Error=r.Error,t.HttpUtility=r.HttpUtility,t.SessionBase=r.SessionBase;var o=n(1);t.CommonUtility=o.CommonUtility,t.ClientResult=o.ClientResult;var i=n(4);t.ClientRequestContext=i.ClientRequestContext,t.ClientObject=i.ClientObject,t.config=i.config,t.Constants=i.Constants,t.ErrorCodes=i.ErrorCodes,t.EventHandlers=i.EventHandlers,t.GenericEventHandlers=i.GenericEventHandlers,t.ResourceStrings=i.ResourceStrings,t.Utility=i.Utility,t._internalConfig=i._internalConfig;var s=function(){function e(){}return e.invokeMethod=function(e,t,n,r,s,a){var c=i.ActionFactory.createMethodAction(e.context,e,t,n,r,s),u=new o.ClientResult(a);return i.Utility._addActionResultHandler(e,c,u),u},e.invokeEnsureUnchanged=function(e,t){i.ActionFactory.createEnsureUnchangedAction(e.context,e,t)},e.invokeSetProperty=function(e,t,n,r){i.ActionFactory.createSetPropertyAction(e.context,e,t,n,r)},e.createRootServiceObject=function(e,t){return new e(t,i.ObjectPathFactory.createGlobalObjectObjectPath(t))},e.createObjectFromReferenceId=function(e,t,n){return new e(t,i.ObjectPathFactory.createReferenceIdObjectPath(t,n))},e.createTopLevelServiceObject=function(e,t,n,r,o){return new e(t,i.ObjectPathFactory.createNewObjectObjectPath(t,n,r,o))},e.createPropertyObject=function(e,t,n,r,o){var s=i.ObjectPathFactory.createPropertyObjectPath(t.context,t,n,r,!1,o);return new e(t.context,s)},e.createIndexerObject=function(e,t,n){var r=i.ObjectPathFactory.createIndexerObjectPath(t.context,t,n);return new e(t.context,r)},e.createMethodObject=function(e,t,n,r,o,s,a,c,u){var l=i.ObjectPathFactory.createMethodObjectPath(t.context,t,n,r,o,s,a,c,u);return new e(t.context,l)},e.createChildItemObject=function(e,t,n,r,o){var s=i.ObjectPathFactory.createChildItemObjectPathUsingIndexerOrGetItemAt(t,n.context,n,r,o);return new e(n.context,s)},e}();t.BatchApiHelper=s},function(e,t,n){"use strict";Object.defineProperty(t,"__esModule",{value:!0});var r=n(2);n(5),n(6),window.OfficeExtensionBatch=r,"undefined"==typeof CustomFunctionMappings&&(window.CustomFunctionMappings={}),"undefined"==typeof Promise&&(window.Promise=Office.Promise),window.OfficeExtension={Promise:Promise,Error:r.Error,ErrorCodes:r.ErrorCodes},n(7).default(!0)},function(e,t,n){"use strict";var r,o=this&&this.__extends||(r=function(e,t){return(r=Object.setPrototypeOf||{__proto__:[]}instanceof Array&&function(e,t){e.__proto__=t}||function(e,t){for(var n in t)t.hasOwnProperty(n)&&(e[n]=t[n])})(e,t)},function(e,t){function n(){this.constructor=e}r(e,t),e.prototype=null===t?Object.create(t):(n.prototype=t.prototype,new n)});Object.defineProperty(t,"__esModule",{value:!0});var i=n(0),s=n(1);!function(e){for(var n in e)t.hasOwnProperty(n)||(t[n]=e[n])}(n(1));var a=function(e){function t(){return null!==e&&e.apply(this,arguments)||this}return o(t,e),t.propertyNotLoaded="PropertyNotLoaded",t.runMustReturnPromise="RunMustReturnPromise",t.cannotRegisterEvent="CannotRegisterEvent",t.invalidOrTimedOutSession="InvalidOrTimedOutSession",t.cannotUpdateReadOnlyProperty="CannotUpdateReadOnlyProperty",t}(i.CoreErrorCodes);t.ErrorCodes=a;var c=function(){function e(e){this.m_callback=e}return e.prototype._handleResult=function(e){this.m_callback&&this.m_callback()},e}(),u=function(e){function t(){return null!==e&&e.apply(this,arguments)||this}return o(t,e),t.createMethodAction=function(e,t,n,r,o,i){x.validateObjectPath(t);var a={Id:e._nextId(),ActionType:3,Name:n,ObjectPathId:t._objectPath.objectPathInfo.Id,ArgumentInfo:{}},c=x.setMethodArguments(e,a.ArgumentInfo,o);x.validateReferencedObjectPaths(c);var u=new s.Action(a,r,x._fixupApiFlags(i));return u.referencedObjectPath=t._objectPath,u.referencedArgumentObjectPaths=c,t._addAction(u),u},t.createRecursiveQueryAction=function(e,t,n){x.validateObjectPath(t);var r={Id:e._nextId(),ActionType:6,Name:"",ObjectPathId:t._objectPath.objectPathInfo.Id,RecursiveQueryInfo:n},o=new s.Action(r,1,4);return o.referencedObjectPath=t._objectPath,t._addAction(o),o},t.createEnsureUnchangedAction=function(e,t,n){x.validateObjectPath(t);var r={Id:e._nextId(),ActionType:8,Name:"",ObjectPathId:t._objectPath.objectPathInfo.Id,ObjectState:n},o=new s.Action(r,1,4);return o.referencedObjectPath=t._objectPath,t._addAction(o),o},t.createInstantiateAction=function(e,t){x.validateObjectPath(t),e._pendingRequest.ensureInstantiateObjectPath(t._objectPath.parentObjectPath),e._pendingRequest.ensureInstantiateObjectPaths(t._objectPath.argumentObjectPaths);var n={Id:e._nextId(),ActionType:1,Name:"",ObjectPathId:t._objectPath.objectPathInfo.Id},r=new s.Action(n,1,4);return r.referencedObjectPath=t._objectPath,t._addAction(r,new I(t),!0),r},t.createTraceAction=function(e,t,n){var r={Id:e._nextId(),ActionType:5,Name:"Trace",ObjectPathId:0},o=new s.Action(r,1,4);return e._pendingRequest.addAction(o),n&&e._pendingRequest.addTrace(r.Id,t),o},t.createTraceMarkerForCallback=function(e,n){var r=t.createTraceAction(e,null,!1);e._pendingRequest.addActionResultHandler(r,new c(n))},t}(s.CommonActionFactory);t.ActionFactory=u;var l=function(e){function t(t,n){var r=e.call(this,t,n)||this;return x.checkArgumentNull(t,"context"),r.m_context=t,r._objectPath&&!t._processingResult&&t._pendingRequest&&(u.createInstantiateAction(t,r),t._autoCleanup&&r._KeepReference&&t.trackedObjects._autoAdd(r)),r}return o(t,e),Object.defineProperty(t.prototype,"context",{get:function(){return this.m_context},enumerable:!0,configurable:!0}),Object.defineProperty(t.prototype,"isNull",{get:function(){return(void 0!==this.m_isNull||!i.TestUtility.isMock())&&(x.throwIfNotLoaded("isNull",this._isNull,null,this._isNull),this._isNull)},enumerable:!0,configurable:!0}),Object.defineProperty(t.prototype,"isNullObject",{get:function(){return(void 0!==this.m_isNull||!i.TestUtility.isMock())&&(x.throwIfNotLoaded("isNullObject",this._isNull,null,this._isNull),this._isNull)},enumerable:!0,configurable:!0}),Object.defineProperty(t.prototype,"_isNull",{get:function(){return this.m_isNull},set:function(e){this.m_isNull=e,e&&this._objectPath&&this._objectPath._updateAsNullObject()},enumerable:!0,configurable:!0}),t.prototype._addAction=function(e,t,n){return void 0===t&&(t=null),n||(this.context._pendingRequest.ensureInstantiateObjectPath(this._objectPath),this.context._pendingRequest.ensureInstantiateObjectPaths(e.referencedArgumentObjectPaths)),this.context._pendingRequest.addAction(e),this.context._pendingRequest.addReferencedObjectPath(this._objectPath),this.context._pendingRequest.addReferencedObjectPaths(e.referencedArgumentObjectPaths),this.context._pendingRequest.addActionResultHandler(e,t),i.CoreUtility._createPromiseFromResult(null)},t.prototype._handleResult=function(e){this._isNull=x.isNullOrUndefined(e),this.context.trackedObjects._autoTrackIfNecessaryWhenHandleObjectResultValue(this,e)},t.prototype._handleIdResult=function(e){this._isNull=x.isNullOrUndefined(e),x.fixObjectPathIfNecessary(this,e),this.context.trackedObjects._autoTrackIfNecessaryWhenHandleObjectResultValue(this,e)},t.prototype._handleRetrieveResult=function(e,t){this._handleIdResult(e)},t.prototype._recursivelySet=function(e,n,r,o,a){var c=e instanceof t,u=e;if(c){if(Object.getPrototypeOf(this)!==Object.getPrototypeOf(e))throw i._Internal.RuntimeError._createInvalidArgError({argumentName:"properties",errorLocation:this._className+".set"});e=JSON.parse(JSON.stringify(e))}try{for(var l,d=0;d<r.length;d++)l=r[d],e.hasOwnProperty(l)&&void 0!==e[l]&&(this[l]=e[l]);for(d=0;d<o.length;d++)if(l=o[d],e.hasOwnProperty(l)&&void 0!==e[l]){var f=c?u[l]:e[l];this[l].set(f,n)}var h=!c;for(n&&!x.isNullOrUndefined(h)&&(h=n.throwOnReadOnly),d=0;d<a.length;d++)if(l=a[d],e.hasOwnProperty(l)&&void 0!==e[l]&&h)throw new i._Internal.RuntimeError({code:i.CoreErrorCodes.invalidArgument,message:i.CoreUtility._getResourceString(A.cannotApplyPropertyThroughSetMethod,l),debugInfo:{errorLocation:l}});for(l in e)if(r.indexOf(l)<0&&o.indexOf(l)<0){var p=Object.getOwnPropertyDescriptor(Object.getPrototypeOf(this),l);if(!p)throw new i._Internal.RuntimeError({code:i.CoreErrorCodes.invalidArgument,message:i.CoreUtility._getResourceString(s.CommonResourceStrings.propertyDoesNotExist,l),debugInfo:{errorLocation:l}});if(h&&!p.set)throw new i._Internal.RuntimeError({code:i.CoreErrorCodes.invalidArgument,message:i.CoreUtility._getResourceString(s.CommonResourceStrings.attemptingToSetReadOnlyProperty,l),debugInfo:{errorLocation:l}})}}catch(e){throw new i._Internal.RuntimeError({code:i.CoreErrorCodes.invalidArgument,message:i.CoreUtility._getResourceString(i.CoreResourceStrings.invalidArgument,"properties"),debugInfo:{errorLocation:this._className+".set"},innerError:e})}},t}(s.ClientObjectBase);t.ClientObject=l;var d=function(){function e(e){this.m_session=e}return e.prototype.executeAsync=function(e,t,n){var r={url:i.CoreConstants.processQuery,method:"POST",headers:n.Headers,body:n.Body},o={id:i.HostBridge.nextId(),type:1,flags:t,message:r};return i.CoreUtility.log(JSON.stringify(o)),this.m_session.sendMessageToHost(o).then((function(e){i.CoreUtility.log("Received response: "+JSON.stringify(e));var t,n=e.message;if(200===n.statusCode)t={ErrorCode:null,ErrorMessage:null,Headers:n.headers,Body:i.CoreUtility._parseResponseBody(n)};else{i.CoreUtility.log("Error Response:"+n.body);var r=i.CoreUtility._parseErrorResponse(n);t={ErrorCode:r.errorCode,ErrorMessage:r.errorMessage,Headers:n.headers,Body:null}}return t}))},e}(),f=function(e){function t(t){var n=e.call(this)||this;return n.m_bridge=t,n.m_bridge.addHostMessageHandler((function(e){3===e.type&&P.getGenericEventRegistration()._handleRichApiMessage(e.message)})),n}return o(t,e),t.getInstanceIfHostBridgeInited=function(){return i.HostBridge.instance?((i.CoreUtility.isNullOrUndefined(t.s_instance)||t.s_instance.m_bridge!==i.HostBridge.instance)&&(t.s_instance=new t(i.HostBridge.instance)),t.s_instance):null},t.prototype._resolveRequestUrlAndHeaderInfo=function(){return i.CoreUtility._createPromiseFromResult(null)},t.prototype._createRequestExecutorOrNull=function(){return i.CoreUtility.log("NativeBridgeSession::CreateRequestExecutor"),new d(this)},Object.defineProperty(t.prototype,"eventRegistration",{get:function(){return P.getGenericEventRegistration()},enumerable:!0,configurable:!0}),t.prototype.sendMessageToHost=function(e){return this.m_bridge.sendMessageToHostAndExpectResponse(e)},t}(i.SessionBase),h=function(e){function t(n){var r=e.call(this)||this;if(r.m_customRequestHeaders={},r.m_batchMode=0,r._onRunFinishedNotifiers=[],i.SessionBase._overrideSession)r.m_requestUrlAndHeaderInfoResolver=i.SessionBase._overrideSession;else if((x.isNullOrUndefined(n)||"string"==typeof n&&0===n.length)&&((n=t.defaultRequestUrlAndHeaders)||(n={url:i.CoreConstants.localDocument,headers:{}})),"string"==typeof n)r.m_requestUrlAndHeaderInfo={url:n,headers:{}};else if(t.isRequestUrlAndHeaderInfoResolver(n))r.m_requestUrlAndHeaderInfoResolver=n;else{if(!t.isRequestUrlAndHeaderInfo(n))throw i._Internal.RuntimeError._createInvalidArgError({argumentName:"url"});var o=n;r.m_requestUrlAndHeaderInfo={url:o.url,headers:{}},i.CoreUtility._copyHeaders(o.headers,r.m_requestUrlAndHeaderInfo.headers)}return!r.m_requestUrlAndHeaderInfoResolver&&r.m_requestUrlAndHeaderInfo&&i.CoreUtility._isLocalDocumentUrl(r.m_requestUrlAndHeaderInfo.url)&&f.getInstanceIfHostBridgeInited()&&(r.m_requestUrlAndHeaderInfo=null,r.m_requestUrlAndHeaderInfoResolver=f.getInstanceIfHostBridgeInited()),r.m_requestUrlAndHeaderInfoResolver instanceof i.SessionBase&&(r.m_session=r.m_requestUrlAndHeaderInfoResolver),r._processingResult=!1,r._customData=m.iterativeExecutor,r.sync=r.sync.bind(r),r}return o(t,e),Object.defineProperty(t.prototype,"session",{get:function(){return this.m_session},enumerable:!0,configurable:!0}),Object.defineProperty(t.prototype,"eventRegistration",{get:function(){return this.m_session?this.m_session.eventRegistration:_.officeJsEventRegistration},enumerable:!0,configurable:!0}),Object.defineProperty(t.prototype,"_url",{get:function(){return this.m_requestUrlAndHeaderInfo?this.m_requestUrlAndHeaderInfo.url:null},enumerable:!0,configurable:!0}),Object.defineProperty(t.prototype,"_pendingRequest",{get:function(){return null==this.m_pendingRequest&&(this.m_pendingRequest=new g(this)),this.m_pendingRequest},enumerable:!0,configurable:!0}),Object.defineProperty(t.prototype,"debugInfo",{get:function(){return{pendingStatements:new E(this._rootObjectPropertyName,this._pendingRequest._objectPaths,this._pendingRequest._actions,s._internalConfig.showDisposeInfoInDebugInfo).process()}},enumerable:!0,configurable:!0}),Object.defineProperty(t.prototype,"trackedObjects",{get:function(){return this.m_trackedObjects||(this.m_trackedObjects=new j(this)),this.m_trackedObjects},enumerable:!0,configurable:!0}),Object.defineProperty(t.prototype,"requestHeaders",{get:function(){return this.m_customRequestHeaders},enumerable:!0,configurable:!0}),Object.defineProperty(t.prototype,"batchMode",{get:function(){return this.m_batchMode},enumerable:!0,configurable:!0}),t.prototype.ensureInProgressBatchIfBatchMode=function(){if(1===this.m_batchMode&&!this.m_explicitBatchInProgress)throw x.createRuntimeError(i.CoreErrorCodes.generalException,i.CoreUtility._getResourceString(A.notInsideBatch),null)},t.prototype.load=function(e,n){x.validateContext(this,e);var r=t._parseQueryOption(n);s.CommonActionFactory.createQueryAction(this,e,r,e)},t.prototype.loadRecursive=function(e,n,r){if(!x.isPlainJsonObject(n))throw i._Internal.RuntimeError._createInvalidArgError({argumentName:"options"});var o={};for(var s in n)o[s]=t._parseQueryOption(n[s]);var a=u.createRecursiveQueryAction(this,e,{Queries:o,MaxDepth:r});this._pendingRequest.addActionResultHandler(a,e)},t.prototype.trace=function(e){u.createTraceAction(this,e,!0)},t.prototype._processOfficeJsErrorResponse=function(e,t){},t.prototype.ensureRequestUrlAndHeaderInfo=function(){var e=this;return x._createPromiseFromResult(null).then((function(){if(!e.m_requestUrlAndHeaderInfo)return e.m_requestUrlAndHeaderInfoResolver._resolveRequestUrlAndHeaderInfo().then((function(t){if(e.m_requestUrlAndHeaderInfo=t,e.m_requestUrlAndHeaderInfo||(e.m_requestUrlAndHeaderInfo={url:i.CoreConstants.localDocument,headers:{}}),x.isNullOrEmptyString(e.m_requestUrlAndHeaderInfo.url)&&(e.m_requestUrlAndHeaderInfo.url=i.CoreConstants.localDocument),e.m_requestUrlAndHeaderInfo.headers||(e.m_requestUrlAndHeaderInfo.headers={}),"function"==typeof e.m_requestUrlAndHeaderInfoResolver._createRequestExecutorOrNull){var n=e.m_requestUrlAndHeaderInfoResolver._createRequestExecutorOrNull();n&&(e._requestExecutor=n)}}))}))},t.prototype.syncPrivateMain=function(){var e=this;return this.ensureRequestUrlAndHeaderInfo().then((function(){var t=e._pendingRequest;return e.m_pendingRequest=null,e.processPreSyncPromises(t).then((function(){return e.syncPrivate(t)}))}))},t.prototype.syncPrivate=function(e){var t=this;if(i.TestUtility.isMock())return i.CoreUtility._createPromiseFromResult(null);if(!e.hasActions)return this.processPendingEventHandlers(e);var n=e.buildRequestMessageBodyAndRequestFlags(),r=n.body,o=n.flags;this._requestFlagModifier&&(o|=this._requestFlagModifier),this._requestExecutor||(i.CoreUtility._isLocalDocumentUrl(this.m_requestUrlAndHeaderInfo.url)?this._requestExecutor=new C(this):this._requestExecutor=new s.HttpRequestExecutor);var c=this._requestExecutor,u={};i.CoreUtility._copyHeaders(this.m_requestUrlAndHeaderInfo.headers,u),i.CoreUtility._copyHeaders(this.m_customRequestHeaders,u);var l={Url:this.m_requestUrlAndHeaderInfo.url,Headers:u,Body:r};e.invalidatePendingInvalidObjectPaths();var d=null,f=null;return this._lastSyncStart="undefined"==typeof performance?0:performance.now(),this._lastRequestFlags=o,c.executeAsync(this._customData,o,l).then((function(n){return t._lastSyncEnd="undefined"==typeof performance?0:performance.now(),d=t.processRequestExecutorResponseMessage(e,n),t.processPendingEventHandlers(e).catch((function(e){i.CoreUtility.log("Error in processPendingEventHandlers"),i.CoreUtility.log(JSON.stringify(e)),f=e}))})).then((function(){if(d)throw i.CoreUtility.log("Throw error from response: "+JSON.stringify(d)),d;if(f){i.CoreUtility.log("Throw error from ProcessEventHandler: "+JSON.stringify(f));var t=null;if(f instanceof i._Internal.RuntimeError)(t=f).traceMessages=e._responseTraceMessages;else{var n=null;n="string"==typeof f?f:f.message,x.isNullOrEmptyString(n)&&(n=i.CoreUtility._getResourceString(A.cannotRegisterEvent)),t=new i._Internal.RuntimeError({code:a.cannotRegisterEvent,message:n,traceMessages:e._responseTraceMessages})}throw t}}))},t.prototype.processRequestExecutorResponseMessage=function(e,t){t.Body&&t.Body.TraceIds&&e._setResponseTraceIds(t.Body.TraceIds);var n=e._responseTraceMessages,r=null;if(t.Body){if(t.Body.Error&&t.Body.Error.ActionIndex>=0){var o=new E(this._rootObjectPropertyName,e._objectPaths,e._actions,!1,!0),a=o.processForDebugStatementInfo(t.Body.Error.ActionIndex);r={statement:a.statement,surroundingStatements:a.surroundingStatements,fullStatements:["Please enable config.extendedErrorLogging to see full statements."]},s.config.extendedErrorLogging&&(o=new E(this._rootObjectPropertyName,e._objectPaths,e._actions,!1,!1),r.fullStatements=o.process())}var c=null;if(t.Body.Results?c=t.Body.Results:t.Body.ProcessedResults&&t.Body.ProcessedResults.Results&&(c=t.Body.ProcessedResults.Results),c){this._processingResult=!0;try{e.processResponse(c)}finally{this._processingResult=!1}}}if(!x.isNullOrEmptyString(t.ErrorCode))return new i._Internal.RuntimeError({code:t.ErrorCode,message:t.ErrorMessage,traceMessages:n});if(t.Body&&t.Body.Error){var u={errorLocation:t.Body.Error.Location};return r&&(u.statement=r.statement,u.surroundingStatements=r.surroundingStatements,u.fullStatements=r.fullStatements),new i._Internal.RuntimeError({code:t.Body.Error.Code,message:t.Body.Error.Message,traceMessages:n,debugInfo:u})}return null},t.prototype.processPendingEventHandlers=function(e){for(var t=x._createPromiseFromResult(null),n=0;n<e._pendingProcessEventHandlers.length;n++){var r=e._pendingProcessEventHandlers[n];t=t.then(this.createProcessOneEventHandlersFunc(r,e))}return t},t.prototype.createProcessOneEventHandlersFunc=function(e,t){return function(){return e._processRegistration(t)}},t.prototype.processPreSyncPromises=function(e){for(var t=x._createPromiseFromResult(null),n=0;n<e._preSyncPromises.length;n++){var r=e._preSyncPromises[n];t=t.then(this.createProcessOneProSyncFunc(r))}return t},t.prototype.createProcessOneProSyncFunc=function(e){return function(){return e}},t.prototype.sync=function(e){return i.TestUtility.isMock()?i.CoreUtility._createPromiseFromResult(e):this.syncPrivateMain().then((function(){return e}))},t.prototype.batch=function(e){var t=this;if(1!==this.m_batchMode)return i.CoreUtility._createPromiseFromException(x.createRuntimeError(i.CoreErrorCodes.generalException,null,null));if(this.m_explicitBatchInProgress)return i.CoreUtility._createPromiseFromException(x.createRuntimeError(i.CoreErrorCodes.generalException,i.CoreUtility._getResourceString(A.pendingBatchInProgress),null));if(x.isNullOrUndefined(e))return x._createPromiseFromResult(null);this.m_explicitBatchInProgress=!0;var n,r,o,s=this.m_pendingRequest;this.m_pendingRequest=new g(this);try{n=e(this._rootObject,this)}catch(e){return this.m_explicitBatchInProgress=!1,this.m_pendingRequest=s,i.CoreUtility._createPromiseFromException(e)}return"object"==typeof n&&n&&"function"==typeof n.then?o=x._createPromiseFromResult(null).then((function(){return n})).then((function(e){return t.m_explicitBatchInProgress=!1,r=t.m_pendingRequest,t.m_pendingRequest=s,e})).catch((function(e){return t.m_explicitBatchInProgress=!1,r=t.m_pendingRequest,t.m_pendingRequest=s,i.CoreUtility._createPromiseFromException(e)})):(this.m_explicitBatchInProgress=!1,r=this.m_pendingRequest,this.m_pendingRequest=s,o=x._createPromiseFromResult(n)),o.then((function(e){return t.ensureRequestUrlAndHeaderInfo().then((function(){return t.syncPrivate(r)})).then((function(){return e}))}))},t._run=function(e,n,r,o,i,s){return void 0===r&&(r=3),void 0===o&&(o=5e3),t._runCommon("run",null,e,0,n,r,o,null,i,s)},t.isValidRequestInfo=function(e){return"string"==typeof e||t.isRequestUrlAndHeaderInfo(e)||t.isRequestUrlAndHeaderInfoResolver(e)},t.isRequestUrlAndHeaderInfo=function(e){return"object"==typeof e&&null!==e&&Object.getPrototypeOf(e)===Object.getPrototypeOf({})&&!x.isNullOrUndefined(e.url)},t.isRequestUrlAndHeaderInfoResolver=function(e){return"object"==typeof e&&null!==e&&"function"==typeof e._resolveRequestUrlAndHeaderInfo},t._runBatch=function(e,n,r,o,i,s,a,c){return void 0===i&&(i=3),void 0===s&&(s=5e3),t._runBatchCommon(0,e,n,r,i,s,o,a,c)},t._runExplicitBatch=function(e,n,r,o,i,s,a,c){return void 0===i&&(i=3),void 0===s&&(s=5e3),t._runBatchCommon(1,e,n,r,i,s,o,a,c)},t._runBatchCommon=function(e,n,r,o,i,s,a,c,u){var d,f;void 0===i&&(i=3),void 0===s&&(s=5e3);var h=null,p=null,m=0,g=null;if(r.length>0)if(t.isValidRequestInfo(r[0]))h=r[0],m=1;else if(x.isPlainJsonObject(r[0])){if(null!=(h=(g=r[0]).session)&&!t.isValidRequestInfo(h))return t.createErrorPromise(n);p=g.previousObjects,m=1}if(r.length==m+1)f=r[m+0];else{if(null!=g||r.length!=m+2)return t.createErrorPromise(n);p=r[m+0],f=r[m+1]}if(null!=p)if(p instanceof l)d=function(){return p.context};else if(p instanceof t)d=function(){return p};else{if(!Array.isArray(p))return t.createErrorPromise(n);var y=p;if(0==y.length)return t.createErrorPromise(n);for(var _=0;_<y.length;_++){if(!(y[_]instanceof l))return t.createErrorPromise(n);if(y[_].context!=y[0].context)return t.createErrorPromise(n,A.invalidRequestContext)}d=function(){return y[0].context}}else d=o;var b=null;return a&&(b=function(e){return a(g||{},e)}),t._runCommon(n,h,d,e,f,i,s,b,c,u)},t.createErrorPromise=function(e,t){return void 0===t&&(t=i.CoreResourceStrings.invalidArgument),i.CoreUtility._createPromiseFromException(x.createRuntimeError(t,i.CoreUtility._getResourceString(t),e))},t._runCommon=function(e,n,r,o,s,a,c,u,l,d){i.SessionBase._overrideSession&&(n=i.SessionBase._overrideSession);var f,h,p,m=!1;return i.CoreUtility.createPromise((function(e,t){e()})).then((function(){if((f=r(n))._autoCleanup)return new Promise((function(e,t){f._onRunFinishedNotifiers.push((function(){f._autoCleanup=!0,e()}))}));f._autoCleanup=!0})).then((function(){return"function"!=typeof s?t.createErrorPromise(e):(p=f.m_batchMode,f.m_batchMode=o,u&&u(f),n=s(1==o?f.batch.bind(f):f),(x.isNullOrUndefined(n)||"function"!=typeof n.then)&&x.throwError(A.runMustReturnPromise),n);var n})).then((function(e){return 1===o?e:f.sync(e)})).then((function(e){m=!0,h=e})).catch((function(e){h=e})).then((function(){var e=f.trackedObjects._retrieveAndClearAutoCleanupList();for(var r in f._autoCleanup=!1,f.m_batchMode=p,e)e[r]._objectPath.isValid=!1;var o=0;if(x._synchronousCleanup||t.isRequestUrlAndHeaderInfoResolver(n))return i();function i(){o++;var t=f.m_pendingRequest,n=f.m_batchMode,r=new g(f);f.m_pendingRequest=r,f.m_batchMode=0;try{for(var s in e)f.trackedObjects.remove(e[s])}finally{f.m_batchMode=n,f.m_pendingRequest=t}return f.syncPrivate(r).then((function(){l&&l(o)})).catch((function(){d&&d(o),o<a&&setTimeout((function(){i()}),c)}))}i()})).then((function(){if(f._onRunFinishedNotifiers&&f._onRunFinishedNotifiers.length>0&&f._onRunFinishedNotifiers.shift()(),m)return h;throw h}))},t}(s.ClientRequestContextBase);t.ClientRequestContext=h;var p=function(){function e(e,t){this.m_proxy=e,this.m_shouldPolyfill=t;var n=e[m.scalarPropertyNames],r=e[m.navigationPropertyNames],o=e[m.className],i=e[m.isCollection];if(n)for(var s=0;s<n.length;s++)x.definePropertyThrowUnloadedException(this,o,n[s]);if(r)for(s=0;s<r.length;s++)x.definePropertyThrowUnloadedException(this,o,r[s]);i&&x.definePropertyThrowUnloadedException(this,o,m.itemsLowerCase)}return Object.defineProperty(e.prototype,"$proxy",{get:function(){return this.m_proxy},enumerable:!0,configurable:!0}),Object.defineProperty(e.prototype,"$isNullObject",{get:function(){if(!this.m_isLoaded)throw new i._Internal.RuntimeError({code:a.valueNotLoaded,message:i.CoreUtility._getResourceString(A.valueNotLoaded),debugInfo:{errorLocation:"retrieveResult.$isNullObject"}});return this.m_isNullObject},enumerable:!0,configurable:!0}),e.prototype.toJSON=function(){if(this.m_isLoaded)return this.m_isNullObject?null:(x.isUndefined(this.m_json)&&(this.m_json=x.purifyJson(this.m_value)),this.m_json)},e.prototype.toString=function(){return JSON.stringify(this.toJSON())},e.prototype._handleResult=function(e){this.m_isLoaded=!0,null===e||"object"==typeof e&&e&&e._IsNull?(this.m_isNullObject=!0,e=null):this.m_isNullObject=!1,this.m_shouldPolyfill&&(e=x.changePropertyNameToCamelLowerCase(e)),this.m_value=e,this.m_proxy._handleRetrieveResult(e,this)},e}(),m=function(e){function t(){return null!==e&&e.apply(this,arguments)||this}return o(t,e),t.getItemAt="GetItemAt",t.index="_Index",t.iterativeExecutor="IterativeExecutor",t.isTracked="_IsTracked",t.eventMessageCategory=65536,t.eventWorkbookId="Workbook",t.eventSourceRemote="Remote",t.proxy="$proxy",t.className="_className",t.isCollection="_isCollection",t.collectionPropertyPath="_collectionPropertyPath",t.objectPathInfoDoNotKeepReferenceFieldName="D",t}(s.CommonConstants);t.Constants=m;var g=function(e){function t(t){var n=e.call(this,t)||this;return n.m_context=t,n.m_pendingProcessEventHandlers=[],n.m_pendingEventHandlerActions={},n.m_traceInfos={},n.m_responseTraceIds={},n.m_responseTraceMessages=[],n}return o(t,e),Object.defineProperty(t.prototype,"traceInfos",{get:function(){return this.m_traceInfos},enumerable:!0,configurable:!0}),Object.defineProperty(t.prototype,"_responseTraceMessages",{get:function(){return this.m_responseTraceMessages},enumerable:!0,configurable:!0}),Object.defineProperty(t.prototype,"_responseTraceIds",{get:function(){return this.m_responseTraceIds},enumerable:!0,configurable:!0}),t.prototype._setResponseTraceIds=function(e){if(e)for(var t=0;t<e.length;t++){var n=e[t];this.m_responseTraceIds[n]=n;var r=this.m_traceInfos[n];i.CoreUtility.isNullOrUndefined(r)||this.m_responseTraceMessages.push(r)}},t.prototype.addTrace=function(e,t){this.m_traceInfos[e]=t},t.prototype._addPendingEventHandlerAction=function(e,t){this.m_pendingEventHandlerActions[e._id]||(this.m_pendingEventHandlerActions[e._id]=[],this.m_pendingProcessEventHandlers.push(e)),this.m_pendingEventHandlerActions[e._id].push(t)},Object.defineProperty(t.prototype,"_pendingProcessEventHandlers",{get:function(){return this.m_pendingProcessEventHandlers},enumerable:!0,configurable:!0}),t.prototype._getPendingEventHandlerActions=function(e){return this.m_pendingEventHandlerActions[e._id]},t}(s.ClientRequestBase);t.ClientRequest=g;var y=function(){function e(e,t,n,r){var o=this;this.m_id=e._nextId(),this.m_context=e,this.m_name=n,this.m_handlers=[],this.m_registered=!1,this.m_eventInfo=r,this.m_callback=function(e){o.m_eventInfo.eventArgsTransformFunc(e).then((function(e){return o.fireEvent(e)}))}}return Object.defineProperty(e.prototype,"_registered",{get:function(){return this.m_registered},enumerable:!0,configurable:!0}),Object.defineProperty(e.prototype,"_id",{get:function(){return this.m_id},enumerable:!0,configurable:!0}),Object.defineProperty(e.prototype,"_handlers",{get:function(){return this.m_handlers},enumerable:!0,configurable:!0}),Object.defineProperty(e.prototype,"_context",{get:function(){return this.m_context},enumerable:!0,configurable:!0}),Object.defineProperty(e.prototype,"_callback",{get:function(){return this.m_callback},enumerable:!0,configurable:!0}),e.prototype.add=function(e){var t=u.createTraceAction(this.m_context,null,!1);return this.m_context._pendingRequest._addPendingEventHandlerAction(this,{id:t.actionInfo.Id,handler:e,operation:0}),new b(this.m_context,this,e)},e.prototype.remove=function(e){var t=u.createTraceAction(this.m_context,null,!1);this.m_context._pendingRequest._addPendingEventHandlerAction(this,{id:t.actionInfo.Id,handler:e,operation:1})},e.prototype.removeAll=function(){var e=u.createTraceAction(this.m_context,null,!1);this.m_context._pendingRequest._addPendingEventHandlerAction(this,{id:e.actionInfo.Id,handler:null,operation:2})},e.prototype._processRegistration=function(e){var t=this,n=i.CoreUtility._createPromiseFromResult(null),r=e._getPendingEventHandlerActions(this);if(!r)return n;for(var o=[],s=0;s<this.m_handlers.length;s++)o.push(this.m_handlers[s]);var a=!1;for(s=0;s<r.length;s++)if(e._responseTraceIds[r[s].id])switch(a=!0,r[s].operation){case 0:o.push(r[s].handler);break;case 1:for(var c=o.length-1;c>=0;c--)if(o[c]===r[s].handler){o.splice(c,1);break}break;case 2:o=[]}return a&&(!this.m_registered&&o.length>0?n=n.then((function(){return t.m_eventInfo.registerFunc(t.m_callback)})).then((function(){return t.m_registered=!0})):this.m_registered&&0==o.length&&(n=n.then((function(){return t.m_eventInfo.unregisterFunc(t.m_callback)})).catch((function(e){i.CoreUtility.log("Error when unregister event: "+JSON.stringify(e))})).then((function(){return t.m_registered=!1}))),n=n.then((function(){return t.m_handlers=o}))),n},e.prototype.fireEvent=function(e){for(var t=[],n=0;n<this.m_handlers.length;n++){var r=this.m_handlers[n],o=i.CoreUtility._createPromiseFromResult(null).then(this.createFireOneEventHandlerFunc(r,e)).catch((function(e){i.CoreUtility.log("Error when invoke handler: "+JSON.stringify(e))}));t.push(o)}i.CoreUtility.Promise.all(t)},e.prototype.createFireOneEventHandlerFunc=function(e,t){return function(){return e(t)}},e}();t.EventHandlers=y;var _,b=function(){function e(e,t,n){this.m_context=e,this.m_allHandlers=t,this.m_handler=n}return Object.defineProperty(e.prototype,"context",{get:function(){return this.m_context},enumerable:!0,configurable:!0}),e.prototype.remove=function(){this.m_allHandlers&&this.m_handler&&(this.m_allHandlers.remove(this.m_handler),this.m_allHandlers=null,this.m_handler=null)},e}();t.EventHandlerResult=b,function(e){var t=function(){function e(){}return e.prototype.register=function(e,t,n){switch(e){case 4:return x.promisify((function(e){return Office.context.document.bindings.getByIdAsync(t,e)})).then((function(e){return x.promisify((function(t){return e.addHandlerAsync(Office.EventType.BindingDataChanged,n,t)}))}));case 3:return x.promisify((function(e){return Office.context.document.bindings.getByIdAsync(t,e)})).then((function(e){return x.promisify((function(t){return e.addHandlerAsync(Office.EventType.BindingSelectionChanged,n,t)}))}));case 2:return x.promisify((function(e){return Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged,n,e)}));case 1:return x.promisify((function(e){return Office.context.document.settings.addHandlerAsync(Office.EventType.SettingsChanged,n,e)}));case 5:return x.promisify((function(e){return OSF.DDA.RichApi.richApiMessageManager.addHandlerAsync("richApiMessage",n,e)}));case 13:return x.promisify((function(e){return Office.context.document.addHandlerAsync(Office.EventType.ObjectDeleted,n,{id:t},e)}));case 14:return x.promisify((function(e){return Office.context.document.addHandlerAsync(Office.EventType.ObjectSelectionChanged,n,{id:t},e)}));case 15:return x.promisify((function(e){return Office.context.document.addHandlerAsync(Office.EventType.ObjectDataChanged,n,{id:t},e)}));case 16:return x.promisify((function(e){return Office.context.document.addHandlerAsync(Office.EventType.ContentControlAdded,n,{id:t},e)}));default:throw i._Internal.RuntimeError._createInvalidArgError({argumentName:"eventId"})}},e.prototype.unregister=function(e,t,n){switch(e){case 4:return x.promisify((function(e){return Office.context.document.bindings.getByIdAsync(t,e)})).then((function(e){return x.promisify((function(t){return e.removeHandlerAsync(Office.EventType.BindingDataChanged,{handler:n},t)}))}));case 3:return x.promisify((function(e){return Office.context.document.bindings.getByIdAsync(t,e)})).then((function(e){return x.promisify((function(t){return e.removeHandlerAsync(Office.EventType.BindingSelectionChanged,{handler:n},t)}))}));case 2:return x.promisify((function(e){return Office.context.document.removeHandlerAsync(Office.EventType.DocumentSelectionChanged,{handler:n},e)}));case 1:return x.promisify((function(e){return Office.context.document.settings.removeHandlerAsync(Office.EventType.SettingsChanged,{handler:n},e)}));case 5:return x.promisify((function(e){return OSF.DDA.RichApi.richApiMessageManager.removeHandlerAsync("richApiMessage",{handler:n},e)}));case 13:return x.promisify((function(e){return Office.context.document.removeHandlerAsync(Office.EventType.ObjectDeleted,{id:t,handler:n},e)}));case 14:return x.promisify((function(e){return Office.context.document.removeHandlerAsync(Office.EventType.ObjectSelectionChanged,{id:t,handler:n},e)}));case 15:return x.promisify((function(e){return Office.context.document.removeHandlerAsync(Office.EventType.ObjectDataChanged,{id:t,handler:n},e)}));case 16:return x.promisify((function(e){return Office.context.document.removeHandlerAsync(Office.EventType.ContentControlAdded,{id:t,handler:n},e)}));default:throw i._Internal.RuntimeError._createInvalidArgError({argumentName:"eventId"})}},e}();e.officeJsEventRegistration=new t}(_=t._Internal||(t._Internal={}));var v=function(){function e(e,t){this.m_handlersByEventByTarget={},this.m_registerEventImpl=e,this.m_unregisterEventImpl=t}return e.getTargetIdOrDefault=function(e){return x.isNullOrUndefined(e)?"":e},e.prototype.getHandlers=function(t,n){n=e.getTargetIdOrDefault(n);var r=this.m_handlersByEventByTarget[t];r||(r={},this.m_handlersByEventByTarget[t]=r);var o=r[n];return o||(o=[],r[n]=o),o},e.prototype.callHandlers=function(e,t,n){for(var r=this.getHandlers(e,t),o=0;o<r.length;o++)r[o](n)},e.prototype.hasHandlers=function(e,t){return this.getHandlers(e,t).length>0},e.prototype.register=function(e,t,n){if(!n)throw i._Internal.RuntimeError._createInvalidArgError({argumentName:"handler"});var r=this.getHandlers(e,t);return r.push(n),1===r.length?this.m_registerEventImpl(e,t):x._createPromiseFromResult(null)},e.prototype.unregister=function(e,t,n){if(!n)throw i._Internal.RuntimeError._createInvalidArgError({argumentName:"handler"});for(var r=this.getHandlers(e,t),o=r.length-1;o>=0;o--)if(r[o]===n){r.splice(o,1);break}return 0===r.length?this.m_unregisterEventImpl(e,t):x._createPromiseFromResult(null)},e}();t.EventRegistration=v;var P=function(){function e(){this.m_eventRegistration=new v(this._registerEventImpl.bind(this),this._unregisterEventImpl.bind(this)),this.m_richApiMessageHandler=this._handleRichApiMessage.bind(this)}return e.prototype.ready=function(){var t=this;return this.m_ready||(e._testReadyImpl?this.m_ready=e._testReadyImpl().then((function(){t.m_isReady=!0})):i.HostBridge.instance?this.m_ready=x._createPromiseFromResult(null).then((function(){t.m_isReady=!0})):this.m_ready=_.officeJsEventRegistration.register(5,"",this.m_richApiMessageHandler).then((function(){t.m_isReady=!0}))),this.m_ready},Object.defineProperty(e.prototype,"isReady",{get:function(){return this.m_isReady},enumerable:!0,configurable:!0}),e.prototype.register=function(e,t,n){var r=this;return this.ready().then((function(){return r.m_eventRegistration.register(e,t,n)}))},e.prototype.unregister=function(e,t,n){var r=this;return this.ready().then((function(){return r.m_eventRegistration.unregister(e,t,n)}))},e.prototype._registerEventImpl=function(e,t){return x._createPromiseFromResult(null)},e.prototype._unregisterEventImpl=function(e,t){return x._createPromiseFromResult(null)},e.prototype._handleRichApiMessage=function(e){if(e&&e.entries)for(var t=0;t<e.entries.length;t++){var n=e.entries[t];if(n.messageCategory==m.eventMessageCategory){i.CoreUtility._logEnabled&&i.CoreUtility.log(JSON.stringify(n));var r=n.messageType,o=n.targetId;if(this.m_eventRegistration.hasHandlers(r,o)){var s=JSON.parse(n.message);n.isRemoteOverride&&(s.source=m.eventSourceRemote),this.m_eventRegistration.callHandlers(r,o,s)}}}},e.getGenericEventRegistration=function(){return e.s_genericEventRegistration||(e.s_genericEventRegistration=new e),e.s_genericEventRegistration},e.richApiMessageEventCategory=65536,e}();t.GenericEventRegistration=P,t._testSetRichApiMessageReadyImpl=function(e){P._testReadyImpl=e},t._testTriggerRichApiMessageEvent=function(e){P.getGenericEventRegistration()._handleRichApiMessage(e)};var O=function(e){function t(t,n,r,o){var i=e.call(this,t,n,r,o)||this;return i.m_genericEventInfo=o,i}return o(t,e),t.prototype.add=function(e){var t=this;return 0==this._handlers.length&&this.m_genericEventInfo.registerFunc&&this.m_genericEventInfo.registerFunc(),P.getGenericEventRegistration().isReady||this._context._pendingRequest._addPreSyncPromise(P.getGenericEventRegistration().ready()),u.createTraceMarkerForCallback(this._context,(function(){t._handlers.push(e),1==t._handlers.length&&P.getGenericEventRegistration().register(t.m_genericEventInfo.eventType,t.m_genericEventInfo.getTargetIdFunc(),t._callback)})),new b(this._context,this,e)},t.prototype.remove=function(e){var t=this;1==this._handlers.length&&this.m_genericEventInfo.unregisterFunc&&this.m_genericEventInfo.unregisterFunc(),u.createTraceMarkerForCallback(this._context,(function(){for(var n=t._handlers,r=n.length-1;r>=0;r--)if(n[r]===e){n.splice(r,1);break}0==n.length&&P.getGenericEventRegistration().unregister(t.m_genericEventInfo.eventType,t.m_genericEventInfo.getTargetIdFunc(),t._callback)}))},t.prototype.removeAll=function(){},t}(y);t.GenericEventHandlers=O;var I=function(){function e(e){this.m_clientObject=e}return e.prototype._handleResult=function(e){this.m_clientObject._handleIdResult(e)},e}(),R=function(){function e(){}return e.createGlobalObjectObjectPath=function(e){var t={Id:e._nextId(),ObjectPathType:1,Name:""};return new s.ObjectPath(t,null,!1,!1,1,4)},e.createNewObjectObjectPath=function(e,t,n,r){var o={Id:e._nextId(),ObjectPathType:2,Name:t};return new s.ObjectPath(o,null,n,!1,1,x._fixupApiFlags(r))},e.createPropertyObjectPath=function(e,t,n,r,o,i){var a={Id:e._nextId(),ObjectPathType:4,Name:n,ParentObjectPathId:t._objectPath.objectPathInfo.Id};return new s.ObjectPath(a,t._objectPath,r,o,1,x._fixupApiFlags(i))},e.createIndexerObjectPath=function(e,t,n){var r={Id:e._nextId(),ObjectPathType:5,Name:"",ParentObjectPathId:t._objectPath.objectPathInfo.Id,ArgumentInfo:{}};return r.ArgumentInfo.Arguments=n,new s.ObjectPath(r,t._objectPath,!1,!1,1,4)},e.createIndexerObjectPathUsingParentPath=function(e,t,n){var r={Id:e._nextId(),ObjectPathType:5,Name:"",ParentObjectPathId:t.objectPathInfo.Id,ArgumentInfo:{}};return r.ArgumentInfo.Arguments=n,new s.ObjectPath(r,t,!1,!1,1,4)},e.createMethodObjectPath=function(e,t,n,r,o,i,a,c,u){var l={Id:e._nextId(),ObjectPathType:3,Name:n,ParentObjectPathId:t._objectPath.objectPathInfo.Id,ArgumentInfo:{}},d=x.setMethodArguments(e,l.ArgumentInfo,o),f=new s.ObjectPath(l,t._objectPath,i,a,r,x._fixupApiFlags(u));return f.argumentObjectPaths=d,f.getByIdMethodName=c,f},e.createReferenceIdObjectPath=function(e,t){var n={Id:e._nextId(),ObjectPathType:6,Name:t,ArgumentInfo:{}};return new s.ObjectPath(n,null,!1,!1,1,4)},e.createChildItemObjectPathUsingIndexerOrGetItemAt=function(t,n,r,o,i){var s=x.tryGetObjectIdFromLoadOrRetrieveResult(o);return t&&!x.isNullOrUndefined(s)?e.createChildItemObjectPathUsingIndexer(n,r,o):e.createChildItemObjectPathUsingGetItemAt(n,r,o,i)},e.createChildItemObjectPathUsingIndexer=function(e,t,n){var r=x.tryGetObjectIdFromLoadOrRetrieveResult(n),o=o={Id:e._nextId(),ObjectPathType:5,Name:"",ParentObjectPathId:t._objectPath.objectPathInfo.Id,ArgumentInfo:{}};return o.ArgumentInfo.Arguments=[r],new s.ObjectPath(o,t._objectPath,!1,!1,1,4)},e.createChildItemObjectPathUsingGetItemAt=function(e,t,n,r){var o=n[m.index];o&&(r=o);var i={Id:e._nextId(),ObjectPathType:3,Name:m.getItemAt,ParentObjectPathId:t._objectPath.objectPathInfo.Id,ArgumentInfo:{}};return i.ArgumentInfo.Arguments=[r],new s.ObjectPath(i,t._objectPath,!1,!1,1,4)},e}();t.ObjectPathFactory=R;var C=function(){function e(e){this.m_context=e}return e.prototype.executeAsync=function(t,n,r){var o=this,s=i.RichApiMessageUtility.buildMessageArrayForIRequestExecutor(t,n,r,e.SourceLibHeaderValue);return new Promise((function(e,t){OSF.DDA.RichApi.executeRichApiRequestAsync(s,(function(t){var n;i.CoreUtility.log("Response:"),i.CoreUtility.log(JSON.stringify(t)),"succeeded"==t.status?n=i.RichApiMessageUtility.buildResponseOnSuccess(i.RichApiMessageUtility.getResponseBody(t),i.RichApiMessageUtility.getResponseHeaders(t)):(n=i.RichApiMessageUtility.buildResponseOnError(t.error.code,t.error.message),o.m_context._processOfficeJsErrorResponse(t.error.code,n)),e(n)}))}))},e.SourceLibHeaderValue="officejs",e}(),j=function(){function e(e){this._autoCleanupList={},this.m_context=e}return e.prototype.add=function(e){var t=this;Array.isArray(e)?e.forEach((function(e){return t._addCommon(e,!0)})):this._addCommon(e,!0)},e.prototype._autoAdd=function(e){this._addCommon(e,!1),this._autoCleanupList[e._objectPath.objectPathInfo.Id]=e},e.prototype._autoTrackIfNecessaryWhenHandleObjectResultValue=function(e,t){this.m_context._autoCleanup&&!e[m.isTracked]&&e!==this.m_context._rootObject&&t&&!x.isNullOrEmptyString(t[m.referenceId])&&(this._autoCleanupList[e._objectPath.objectPathInfo.Id]=e,e[m.isTracked]=!0)},e.prototype._addCommon=function(e,t){if(e[m.isTracked])t&&this.m_context._autoCleanup&&delete this._autoCleanupList[e._objectPath.objectPathInfo.Id];else{var n=e[m.referenceId];if(e._objectPath.objectPathInfo[m.objectPathInfoDoNotKeepReferenceFieldName])throw x.createRuntimeError(i.CoreErrorCodes.generalException,i.CoreUtility._getResourceString(A.objectIsUntracked),null);x.isNullOrEmptyString(n)&&e._KeepReference&&(e._KeepReference(),u.createInstantiateAction(this.m_context,e),t&&this.m_context._autoCleanup&&delete this._autoCleanupList[e._objectPath.objectPathInfo.Id],e[m.isTracked]=!0)}},e.prototype.remove=function(e){var t=this;Array.isArray(e)?e.forEach((function(e){return t._removeCommon(e)})):this._removeCommon(e)},e.prototype._removeCommon=function(e){e._objectPath.objectPathInfo[m.objectPathInfoDoNotKeepReferenceFieldName]=!0,e.context._pendingRequest._removeKeepReferenceAction(e._objectPath.objectPathInfo.Id);var t=e[m.referenceId];if(!x.isNullOrEmptyString(t)){var n=this.m_context._rootObject;n._RemoveReference&&n._RemoveReference(t)}delete e[m.isTracked]},e.prototype._retrieveAndClearAutoCleanupList=function(){var e=this._autoCleanupList;return this._autoCleanupList={},e},e}();t.TrackedObjects=j;var E=function(){function e(e,t,n,r,o){e||(e="root"),this.m_globalObjName=e,this.m_referencedObjectPaths=t,this.m_actions=n,this.m_statements=[],this.m_variableNameForObjectPathMap={},this.m_variableNameToObjectPathMap={},this.m_declaredObjectPathMap={},this.m_showDispose=r,this.m_removePII=o}return e.prototype.process=function(){this.m_showDispose&&g._calculateLastUsedObjectPathIds(this.m_actions);for(var e=0;e<this.m_actions.length;e++)this.processOneAction(this.m_actions[e]);return this.m_statements},e.prototype.processForDebugStatementInfo=function(e){this.m_showDispose&&g._calculateLastUsedObjectPathIds(this.m_actions),this.m_statements=[];for(var t=-1,n=0;n<this.m_actions.length&&(this.processOneAction(this.m_actions[n]),e==n&&(t=this.m_statements.length-1),!(t>=0&&this.m_statements.length>t+5+1));n++);if(t<0)return null;var r=t-5;r<0&&(r=0);var o=t+1+5;o>this.m_statements.length&&(o=this.m_statements.length);var i=[];0!=r&&i.push("...");for(var s=r;s<t;s++)i.push(this.m_statements[s]);i.push("// >>>>>"),i.push(this.m_statements[t]),i.push("// <<<<<");for(var a=t+1;a<o;a++)i.push(this.m_statements[a]);return o<this.m_statements.length&&i.push("..."),{statement:this.m_statements[t],surroundingStatements:i}},e.prototype.processOneAction=function(e){switch(e.actionInfo.ActionType){case 1:this.processInstantiateAction(e);break;case 3:this.processMethodAction(e);break;case 2:this.processQueryAction(e);break;case 7:this.processQueryAsJsonAction(e);break;case 6:this.processRecursiveQueryAction(e);break;case 4:this.processSetPropertyAction(e);break;case 5:this.processTraceAction(e);break;case 8:this.processEnsureUnchangedAction(e);break;case 9:this.processUpdateAction(e)}},e.prototype.processInstantiateAction=function(e){var t=e.actionInfo.ObjectPathId,n=this.m_referencedObjectPaths[t],r=this.getObjVarName(t);if(this.m_declaredObjectPathMap[t])o="// Instantiate {"+r+"}",o=this.appendDisposeCommentIfRelevant(o,e),this.m_statements.push(o);else{var o="var "+r+" = "+this.buildObjectPathExpressionWithParent(n)+";";o=this.appendDisposeCommentIfRelevant(o,e),this.m_statements.push(o),this.m_declaredObjectPathMap[t]=r}},e.prototype.processMethodAction=function(e){var t=e.actionInfo.Name;if("_KeepReference"===t){if(!s._internalConfig.showInternalApiInDebugInfo)return;t="track"}var n=this.getObjVarName(e.actionInfo.ObjectPathId)+"."+x._toCamelLowerCase(t)+"("+this.buildArgumentsExpression(e.actionInfo.ArgumentInfo)+");";n=this.appendDisposeCommentIfRelevant(n,e),this.m_statements.push(n)},e.prototype.processQueryAction=function(e){var t=this.buildQueryExpression(e),n=this.getObjVarName(e.actionInfo.ObjectPathId)+".load("+t+");";n=this.appendDisposeCommentIfRelevant(n,e),this.m_statements.push(n)},e.prototype.processQueryAsJsonAction=function(e){var t=this.buildQueryExpression(e),n=this.getObjVarName(e.actionInfo.ObjectPathId)+".retrieve("+t+");";n=this.appendDisposeCommentIfRelevant(n,e),this.m_statements.push(n)},e.prototype.processRecursiveQueryAction=function(e){var t="";e.actionInfo.RecursiveQueryInfo&&(t=JSON.stringify(e.actionInfo.RecursiveQueryInfo));var n=this.getObjVarName(e.actionInfo.ObjectPathId)+".loadRecursive("+t+");";n=this.appendDisposeCommentIfRelevant(n,e),this.m_statements.push(n)},e.prototype.processSetPropertyAction=function(e){var t=this.getObjVarName(e.actionInfo.ObjectPathId)+"."+x._toCamelLowerCase(e.actionInfo.Name)+" = "+this.buildArgumentsExpression(e.actionInfo.ArgumentInfo)+";";t=this.appendDisposeCommentIfRelevant(t,e),this.m_statements.push(t)},e.prototype.processTraceAction=function(e){var t="context.trace();";t=this.appendDisposeCommentIfRelevant(t,e),this.m_statements.push(t)},e.prototype.processEnsureUnchangedAction=function(e){var t=this.getObjVarName(e.actionInfo.ObjectPathId)+".ensureUnchanged("+JSON.stringify(e.actionInfo.ObjectState)+");";t=this.appendDisposeCommentIfRelevant(t,e),this.m_statements.push(t)},e.prototype.processUpdateAction=function(e){var t=this.getObjVarName(e.actionInfo.ObjectPathId)+".update("+JSON.stringify(e.actionInfo.ObjectState)+");";t=this.appendDisposeCommentIfRelevant(t,e),this.m_statements.push(t)},e.prototype.appendDisposeCommentIfRelevant=function(e,t){var n=this;if(this.m_showDispose){var r=t.actionInfo.L;if(r&&r.length>0)return e+" // And then dispose {"+r.map((function(e){return n.getObjVarName(e)})).join(", ")+"}"}return e},e.prototype.buildQueryExpression=function(e){if(e.actionInfo.QueryInfo){var t={};return t.select=e.actionInfo.QueryInfo.Select,t.expand=e.actionInfo.QueryInfo.Expand,t.skip=e.actionInfo.QueryInfo.Skip,t.top=e.actionInfo.QueryInfo.Top,void 0===t.top&&void 0===t.skip&&void 0===t.expand?void 0===t.select?"":JSON.stringify(t.select):JSON.stringify(t)}return""},e.prototype.buildObjectPathExpressionWithParent=function(e){return 5!=e.objectPathInfo.ObjectPathType&&3!=e.objectPathInfo.ObjectPathType&&4!=e.objectPathInfo.ObjectPathType||!e.objectPathInfo.ParentObjectPathId?this.buildObjectPathExpression(e):this.getObjVarName(e.objectPathInfo.ParentObjectPathId)+"."+this.buildObjectPathExpression(e)},e.prototype.buildObjectPathExpression=function(e){var t=this.buildObjectPathInfoExpression(e.objectPathInfo),n=e.originalObjectPathInfo;return n&&(t=t+" /* originally "+this.buildObjectPathInfoExpression(n)+" */"),t},e.prototype.buildObjectPathInfoExpression=function(e){switch(e.ObjectPathType){case 1:return"context."+this.m_globalObjName;case 5:return"getItem("+this.buildArgumentsExpression(e.ArgumentInfo)+")";case 3:return x._toCamelLowerCase(e.Name)+"("+this.buildArgumentsExpression(e.ArgumentInfo)+")";case 2:return e.Name+".newObject()";case 7:return"null";case 4:return x._toCamelLowerCase(e.Name);case 6:return"context."+this.m_globalObjName+"._getObjectByReferenceId("+JSON.stringify(e.Name)+")"}},e.prototype.buildArgumentsExpression=function(e){var t="";if(!e.Arguments||0===e.Arguments.length)return t;if(this.m_removePII)return void 0===e.Arguments[0]?t:"...";for(var n=0;n<e.Arguments.length;n++)n>0&&(t+=", "),t+=this.buildArgumentLiteral(e.Arguments[n],e.ReferencedObjectPathIds?e.ReferencedObjectPathIds[n]:null);return"undefined"===t&&(t=""),t},e.prototype.buildArgumentLiteral=function(e,t){return"number"==typeof e&&e===t?this.getObjVarName(t):JSON.stringify(e)},e.prototype.getObjVarNameBase=function(e){var t="v",n=this.m_referencedObjectPaths[e];if(n)switch(n.objectPathInfo.ObjectPathType){case 1:t=this.m_globalObjName;break;case 4:t=x._toCamelLowerCase(n.objectPathInfo.Name);break;case 3:var r=n.objectPathInfo.Name;r.length>3&&"Get"===r.substr(0,3)&&(r=r.substr(3)),t=x._toCamelLowerCase(r);break;case 5:var o=this.getObjVarNameBase(n.objectPathInfo.ParentObjectPathId);t="s"===o.charAt(o.length-1)?o.substr(0,o.length-1):o+"Item"}return t},e.prototype.getObjVarName=function(e){if(this.m_variableNameForObjectPathMap[e])return this.m_variableNameForObjectPathMap[e];var t=this.getObjVarNameBase(e);if(!this.m_variableNameToObjectPathMap[t])return this.m_variableNameForObjectPathMap[e]=t,this.m_variableNameToObjectPathMap[t]=e,t;for(var n=1;this.m_variableNameToObjectPathMap[t+n.toString()];)n++;return t+=n.toString(),this.m_variableNameForObjectPathMap[e]=t,this.m_variableNameToObjectPathMap[t]=e,t},e}(),A=function(e){function t(){return null!==e&&e.apply(this,arguments)||this}return o(t,e),t.cannotRegisterEvent="CannotRegisterEvent",t.connectionFailureWithStatus="ConnectionFailureWithStatus",t.connectionFailureWithDetails="ConnectionFailureWithDetails",t.propertyNotLoaded="PropertyNotLoaded",t.runMustReturnPromise="RunMustReturnPromise",t.moreInfoInnerError="MoreInfoInnerError",t.cannotApplyPropertyThroughSetMethod="CannotApplyPropertyThroughSetMethod",t.invalidOperationInCellEditMode="InvalidOperationInCellEditMode",t.objectIsUntracked="ObjectIsUntracked",t.customFunctionDefintionMissing="CustomFunctionDefintionMissing",t.customFunctionImplementationMissing="CustomFunctionImplementationMissing",t.customFunctionNameContainsBadChars="CustomFunctionNameContainsBadChars",t.customFunctionNameCannotSplit="CustomFunctionNameCannotSplit",t.customFunctionUnexpectedNumberOfEntriesInResultBatch="CustomFunctionUnexpectedNumberOfEntriesInResultBatch",t.customFunctionCancellationHandlerMissing="CustomFunctionCancellationHandlerMissing",t.customFunctionInvalidFunction="CustomFunctionInvalidFunction",t.customFunctionInvalidFunctionMapping="CustomFunctionInvalidFunctionMapping",t.customFunctionWindowMissing="CustomFunctionWindowMissing",t.customFunctionDefintionMissingOnWindow="CustomFunctionDefintionMissingOnWindow",t.pendingBatchInProgress="PendingBatchInProgress",t.notInsideBatch="NotInsideBatch",t.cannotUpdateReadOnlyProperty="CannotUpdateReadOnlyProperty",t}(s.CommonResourceStrings);t.ResourceStrings=A,i.CoreUtility.addResourceStringValues({CannotRegisterEvent:"The event handler cannot be registered.",PropertyNotLoaded:"The property '{0}' is not available. Before reading the property's value, call the load method on the containing object and call \"context.sync()\" on the associated request context.",RunMustReturnPromise:'The batch function passed to the ".run" method didn\'t return a promise. The function must return a promise, so that any automatically-tracked objects can be released at the completion of the batch operation. Typically, you return a promise by returning the response from "context.sync()".',InvalidOrTimedOutSessionMessage:"Your Office Online session has expired or is invalid. To continue, refresh the page.",InvalidOperationInCellEditMode:"Excel is in cell-editing mode. Please exit the edit mode by pressing ENTER or TAB or selecting another cell, and then try again.",CustomFunctionDefintionMissing:"A property with the name '{0}' that represents the function's definition must exist on Excel.Script.CustomFunctions.",CustomFunctionDefintionMissingOnWindow:"A property with the name '{0}' that represents the function's definition must exist on the window object.",CustomFunctionImplementationMissing:"The property with the name '{0}' on Excel.Script.CustomFunctions that represents the function's definition must contain a 'call' property that implements the function.",CustomFunctionNameContainsBadChars:"The function name may only contain letters, digits, underscores, and periods.",CustomFunctionNameCannotSplit:"The function name must contain a non-empty namespace and a non-empty short name.",CustomFunctionUnexpectedNumberOfEntriesInResultBatch:"The batching function returned a number of results that doesn't match the number of parameter value sets that were passed into it.",CustomFunctionCancellationHandlerMissing:"The cancellation handler onCanceled is missing in the function. The handler must be present as the function is defined as cancelable.",CustomFunctionInvalidFunction:"The property with the name '{0}' that represents the function's definition is not a valid function.",CustomFunctionInvalidFunctionMapping:"The property with the name '{0}' on CustomFunctionMappings that represents the function's definition is not a valid function.",CustomFunctionWindowMissing:"The window object was not found.",PendingBatchInProgress:"There is a pending batch in progress. The batch method may not be called inside another batch, or simultaneously with another batch.",NotInsideBatch:"Operations may not be invoked outside of a batch method.",CannotUpdateReadOnlyProperty:"The property '{0}' is read-only and it cannot be updated.",ObjectIsUntracked:"The object is untracked."});var x=function(e){function t(){return null!==e&&e.apply(this,arguments)||this}return o(t,e),t.fixObjectPathIfNecessary=function(e,t){e&&e._objectPath&&t&&e._objectPath.updateUsingObjectData(t,e)},t.load=function(e,t){return e.context.load(e,t),e},t.loadAndSync=function(e,t){return e.context.load(e,t),e.context.sync().then((function(){return e}))},t.retrieve=function(e,n){var r=s._internalConfig.alwaysPolyfillClientObjectRetrieveMethod;r||(r=!t.isSetSupported("RichApiRuntime","1.1"));var o=new p(e,r);return e._retrieve(n,o),o},t.retrieveAndSync=function(e,n){var r=t.retrieve(e,n);return e.context.sync().then((function(){return r}))},t.toJson=function(e,n,r,o){var i={};for(var s in n)void 0!==(a=n[s])&&(i[s]=a);for(var s in r){var a;void 0!==(a=r[s])&&(a[t.fieldName_isCollection]&&void 0!==a[t.fieldName_m__items]?i[s]=a.toJSON().items:i[s]=a.toJSON())}return o&&(i.items=o.map((function(e){return e.toJSON()}))),i},t.throwError=function(e,t,n){throw new i._Internal.RuntimeError({code:e,message:i.CoreUtility._getResourceString(e,t),debugInfo:n?{errorLocation:n}:void 0})},t.createRuntimeError=function(e,t,n){return new i._Internal.RuntimeError({code:e,message:t,debugInfo:{errorLocation:n}})},t.throwIfNotLoaded=function(e,n,r,o){if(!o&&i.CoreUtility.isUndefined(n)&&e.charCodeAt(0)!=t.s_underscoreCharCode)throw t.createPropertyNotLoadedException(r,e)},t.createPropertyNotLoadedException=function(e,t){return new i._Internal.RuntimeError({code:a.propertyNotLoaded,message:i.CoreUtility._getResourceString(A.propertyNotLoaded,t),debugInfo:e?{errorLocation:e+"."+t}:void 0})},t.createCannotUpdateReadOnlyPropertyException=function(e,t){return new i._Internal.RuntimeError({code:a.cannotUpdateReadOnlyProperty,message:i.CoreUtility._getResourceString(A.cannotUpdateReadOnlyProperty,t),debugInfo:e?{errorLocation:e+"."+t}:void 0})},t.promisify=function(e){return new Promise((function(t,n){e((function(e){"failed"==e.status?n(e.error):t(e.value)}))}))},t._addActionResultHandler=function(e,t,n){e.context._pendingRequest.addActionResultHandler(t,n)},t._handleNavigationPropertyResults=function(e,t,n){for(var r=0;r<n.length-1;r+=2)i.CoreUtility.isUndefined(t[n[r+1]])||e[n[r]]._handleResult(t[n[r+1]])},t._fixupApiFlags=function(e){return"boolean"==typeof e&&(e=e?1:0),e},t.definePropertyThrowUnloadedException=function(e,n,r){Object.defineProperty(e,r,{configurable:!0,enumerable:!0,get:function(){throw t.createPropertyNotLoadedException(n,r)},set:function(){throw t.createCannotUpdateReadOnlyPropertyException(n,r)}})},t.defineReadOnlyPropertyWithValue=function(e,n,r){Object.defineProperty(e,n,{configurable:!0,enumerable:!0,get:function(){return r},set:function(){throw t.createCannotUpdateReadOnlyPropertyException(null,n)}})},t.processRetrieveResult=function(e,n,r,o){if(!i.CoreUtility.isNullOrUndefined(n))if(o){var s=n[m.itemsLowerCase];if(Array.isArray(s)){for(var a=[],c=0;c<s.length;c++){var u=o(s[c],c),l={};l[m.proxy]=u,u._handleRetrieveResult(s[c],l),a.push(l)}t.defineReadOnlyPropertyWithValue(r,m.itemsLowerCase,a)}}else{var d=e[m.scalarPropertyNames],f=e[m.navigationPropertyNames],h=e[m.className];if(d)for(c=0;c<d.length;c++){var p=n[g=d[c]];i.CoreUtility.isUndefined(p)?t.definePropertyThrowUnloadedException(r,h,g):t.defineReadOnlyPropertyWithValue(r,g,p)}if(f)for(c=0;c<f.length;c++){var g;if(p=n[g=f[c]],i.CoreUtility.isUndefined(p))t.definePropertyThrowUnloadedException(r,h,g);else{var y=e[g],_={};y._handleRetrieveResult(p,_),_[m.proxy]=y,Array.isArray(_[m.itemsLowerCase])&&(_=_[m.itemsLowerCase]),t.defineReadOnlyPropertyWithValue(r,g,_)}}}},t.setMockData=function(e,n,r,o){if(i.CoreUtility.isNullOrUndefined(n))e._handleResult(n);else{if(e[m.scalarPropertyOriginalNames]){for(var s={},a=e[m.scalarPropertyOriginalNames],c=e[m.scalarPropertyNames],u=0;u<c.length;u++)void 0!==n[c[u]]&&(s[a[u]]=n[c[u]]);e._handleResult(s)}if(e[m.navigationPropertyNames]){var l=e[m.navigationPropertyNames];for(u=0;u<l.length;u++)if(void 0!==n[l[u]]){var d=e[l[u]];d.setMockData&&d.setMockData(n[l[u]])}}if(e[m.isCollection]&&r){var f=Array.isArray(n)?n:n[m.itemsLowerCase];if(Array.isArray(f)){var h=[];for(u=0;u<f.length;u++){var p=r(f,u);t.setMockData(p,f[u]),h.push(p)}o(h)}}}},t.fieldName_m__items="m__items",t.fieldName_isCollection="_isCollection",t._synchronousCleanup=!1,t.s_underscoreCharCode="_".charCodeAt(0),t}(s.CommonUtility);t.Utility=x},function(e,t){var n,r=this&&this.__assign||function(){return(r=Object.assign||function(e){for(var t,n=1,r=arguments.length;n<r;n++)for(var o in t=arguments[n])Object.prototype.hasOwnProperty.call(t,o)&&(e[o]=t[o]);return e}).apply(this,arguments)};function o(e){return window[e.platform===Office.PlatformType.OfficeOnline?"_OfficeRuntimeWeb":"_OfficeRuntimeNative"]}function i(e){return function(t){return function(e,t){return s((function(n){return n[e][t]}))}(e,t)}}function s(e){return function(){var t=this,n=arguments;return Office.onReady().then((function(r){return e(o(r)).apply(t,n)}))}}Office.onReady((function(e){window.OfficeRuntime=r({},window.OfficeRuntime,o(e))})),window.OfficeRuntime={AsyncStorage:function(){return{getItem:e("getItem"),setItem:e("setItem"),removeItem:e("removeItem"),getAllKeys:e("getAllKeys"),multiSet:e("multiSet"),multiRemove:e("multiRemove"),multiGet:e("multiGet")};function e(e){return s((function(t){return t.storage[e]}))}}(),displayWebDialog:s((function(e){return e.displayWebDialog})),storage:function(){return{getItem:e("getItem"),setItem:e("setItem"),removeItem:e("removeItem"),getKeys:e("getKeys"),setItems:e("setItems"),removeItems:e("removeItems"),getItems:e("getItems")};function e(e){return s((function(t){return t.storage[e]}))}}(),experimentation:function(){return{getBooleanFeatureGateAsync:e("getBooleanFeatureGateAsync"),getIntFeatureGateAsync:e("getIntFeatureGateAsync"),getStringFeatureGateAsync:e("getStringFeatureGateAsync")};function e(e){return s((function(t){return t.experimentation[e]}))}}(),apiInformation:{isSetSupported:function(e,t){return Office.context.requirements.isSetSupported(e,Number(t))}},message:(n=i("message"),{on:n("on"),off:n("off"),emit:n("emit")}),auth:{getAccessToken:i("auth")("getAccessToken")},ui:{getRibbon:i("ui")("getRibbon")}}},function(e,t){!function(e){"use strict";if(!e.fetch){var t={searchParams:"URLSearchParams"in e,iterable:"Symbol"in e&&"iterator"in Symbol,blob:"FileReader"in e&&"Blob"in e&&function(){try{return new Blob,!0}catch(e){return!1}}(),formData:"FormData"in e,arrayBuffer:"ArrayBuffer"in e};if(t.arrayBuffer)var n=["[object Int8Array]","[object Uint8Array]","[object Uint8ClampedArray]","[object Int16Array]","[object Uint16Array]","[object Int32Array]","[object Uint32Array]","[object Float32Array]","[object Float64Array]"],r=function(e){return e&&DataView.prototype.isPrototypeOf(e)},o=ArrayBuffer.isView||function(e){return e&&n.indexOf(Object.prototype.toString.call(e))>-1};l.prototype.append=function(e,t){e=a(e),t=c(t);var n=this.map[e];this.map[e]=n?n+","+t:t},l.prototype.delete=function(e){delete this.map[a(e)]},l.prototype.get=function(e){return e=a(e),this.has(e)?this.map[e]:null},l.prototype.has=function(e){return this.map.hasOwnProperty(a(e))},l.prototype.set=function(e,t){this.map[a(e)]=c(t)},l.prototype.forEach=function(e,t){for(var n in this.map)this.map.hasOwnProperty(n)&&e.call(t,this.map[n],n,this)},l.prototype.keys=function(){var e=[];return this.forEach((function(t,n){e.push(n)})),u(e)},l.prototype.values=function(){var e=[];return this.forEach((function(t){e.push(t)})),u(e)},l.prototype.entries=function(){var e=[];return this.forEach((function(t,n){e.push([n,t])})),u(e)},t.iterable&&(l.prototype[Symbol.iterator]=l.prototype.entries);var i=["DELETE","GET","HEAD","OPTIONS","POST","PUT"];g.prototype.clone=function(){return new g(this,{body:this._bodyInit})},m.call(g.prototype),m.call(_.prototype),_.prototype.clone=function(){return new _(this._bodyInit,{status:this.status,statusText:this.statusText,headers:new l(this.headers),url:this.url})},_.error=function(){var e=new _(null,{status:0,statusText:""});return e.type="error",e};var s=[301,302,303,307,308];_.redirect=function(e,t){if(-1===s.indexOf(t))throw new RangeError("Invalid status code");return new _(null,{status:t,headers:{location:e}})},e.Headers=l,e.Request=g,e.Response=_,e.fetch=function(e,n){return new Promise((function(r,o){var i=new g(e,n),s=new XMLHttpRequest;s.onload=function(){var e,t,n={status:s.status,statusText:s.statusText,headers:(e=s.getAllResponseHeaders()||"",t=new l,e.split(/\r?\n/).forEach((function(e){var n=e.split(":"),r=n.shift().trim();if(r){var o=n.join(":").trim();t.append(r,o)}})),t)};n.url="responseURL"in s?s.responseURL:n.headers.get("X-Request-URL");var o="response"in s?s.response:s.responseText;r(new _(o,n))},s.onerror=function(){o(new TypeError("Network request failed"))},s.ontimeout=function(){o(new TypeError("Network request failed"))},s.open(i.method,i.url,!0),"include"===i.credentials&&(s.withCredentials=!0),"responseType"in s&&t.blob&&(s.responseType="blob"),i.headers.forEach((function(e,t){s.setRequestHeader(t,e)})),s.send(void 0===i._bodyInit?null:i._bodyInit)}))},e.fetch.polyfill=!0}function a(e){if("string"!=typeof e&&(e=String(e)),/[^a-z0-9\-#$%&'*+.\^_`|~]/i.test(e))throw new TypeError("Invalid character in header field name");return e.toLowerCase()}function c(e){return"string"!=typeof e&&(e=String(e)),e}function u(e){var n={next:function(){var t=e.shift();return{done:void 0===t,value:t}}};return t.iterable&&(n[Symbol.iterator]=function(){return n}),n}function l(e){this.map={},e instanceof l?e.forEach((function(e,t){this.append(t,e)}),this):Array.isArray(e)?e.forEach((function(e){this.append(e[0],e[1])}),this):e&&Object.getOwnPropertyNames(e).forEach((function(t){this.append(t,e[t])}),this)}function d(e){if(e.bodyUsed)return Promise.reject(new TypeError("Already read"));e.bodyUsed=!0}function f(e){return new Promise((function(t,n){e.onload=function(){t(e.result)},e.onerror=function(){n(e.error)}}))}function h(e){var t=new FileReader,n=f(t);return t.readAsArrayBuffer(e),n}function p(e){if(e.slice)return e.slice(0);var t=new Uint8Array(e.byteLength);return t.set(new Uint8Array(e)),t.buffer}function m(){return this.bodyUsed=!1,this._initBody=function(e){if(this._bodyInit=e,e)if("string"==typeof e)this._bodyText=e;else if(t.blob&&Blob.prototype.isPrototypeOf(e))this._bodyBlob=e;else if(t.formData&&FormData.prototype.isPrototypeOf(e))this._bodyFormData=e;else if(t.searchParams&&URLSearchParams.prototype.isPrototypeOf(e))this._bodyText=e.toString();else if(t.arrayBuffer&&t.blob&&r(e))this._bodyArrayBuffer=p(e.buffer),this._bodyInit=new Blob([this._bodyArrayBuffer]);else{if(!t.arrayBuffer||!ArrayBuffer.prototype.isPrototypeOf(e)&&!o(e))throw new Error("unsupported BodyInit type");this._bodyArrayBuffer=p(e)}else this._bodyText="";this.headers.get("content-type")||("string"==typeof e?this.headers.set("content-type","text/plain;charset=UTF-8"):this._bodyBlob&&this._bodyBlob.type?this.headers.set("content-type",this._bodyBlob.type):t.searchParams&&URLSearchParams.prototype.isPrototypeOf(e)&&this.headers.set("content-type","application/x-www-form-urlencoded;charset=UTF-8"))},t.blob&&(this.blob=function(){var e=d(this);if(e)return e;if(this._bodyBlob)return Promise.resolve(this._bodyBlob);if(this._bodyArrayBuffer)return Promise.resolve(new Blob([this._bodyArrayBuffer]));if(this._bodyFormData)throw new Error("could not read FormData body as blob");return Promise.resolve(new Blob([this._bodyText]))},this.arrayBuffer=function(){return this._bodyArrayBuffer?d(this)||Promise.resolve(this._bodyArrayBuffer):this.blob().then(h)}),this.text=function(){var e,t,n,r=d(this);if(r)return r;if(this._bodyBlob)return e=this._bodyBlob,n=f(t=new FileReader),t.readAsText(e),n;if(this._bodyArrayBuffer)return Promise.resolve(function(e){for(var t=new Uint8Array(e),n=new Array(t.length),r=0;r<t.length;r++)n[r]=String.fromCharCode(t[r]);return n.join("")}(this._bodyArrayBuffer));if(this._bodyFormData)throw new Error("could not read FormData body as text");return Promise.resolve(this._bodyText)},t.formData&&(this.formData=function(){return this.text().then(y)}),this.json=function(){return this.text().then(JSON.parse)},this}function g(e,t){var n,r,o=(t=t||{}).body;if(e instanceof g){if(e.bodyUsed)throw new TypeError("Already read");this.url=e.url,this.credentials=e.credentials,t.headers||(this.headers=new l(e.headers)),this.method=e.method,this.mode=e.mode,o||null==e._bodyInit||(o=e._bodyInit,e.bodyUsed=!0)}else this.url=String(e);if(this.credentials=t.credentials||this.credentials||"omit",!t.headers&&this.headers||(this.headers=new l(t.headers)),this.method=(r=(n=t.method||this.method||"GET").toUpperCase(),i.indexOf(r)>-1?r:n),this.mode=t.mode||this.mode||null,this.referrer=null,("GET"===this.method||"HEAD"===this.method)&&o)throw new TypeError("Body not allowed for GET or HEAD requests");this._initBody(o)}function y(e){var t=new FormData;return e.trim().split("&").forEach((function(e){if(e){var n=e.split("="),r=n.shift().replace(/\+/g," "),o=n.join("=").replace(/\+/g," ");t.append(decodeURIComponent(r),decodeURIComponent(o))}})),t}function _(e,t){t||(t={}),this.type="default",this.status="status"in t?t.status:200,this.ok=this.status>=200&&this.status<300,this.statusText="statusText"in t?t.statusText:"OK",this.headers=new l(t.headers),this.url=t.url||"",this._initBody(e)}}("undefined"!=typeof self?self:this)},function(e,t,n){"use strict";Object.defineProperty(t,"__esModule",{value:!0});var r=n(8);t.default=function(e){function t(){Office.onReady((function(e){e.host===Office.HostType.Excel?function e(){CustomFunctionMappings&&CustomFunctionMappings.__delay__?setTimeout(e,50):r.CustomFunctions.initialize()}():console.warn("Warning: Expected to be loaded inside of an Excel add-in.")}))}window.CustomFunctions=window.CustomFunctions||{},window.CustomFunctions.setCustomFunctionInvoker=r.setCustomFunctionInvoker,window.CustomFunctions.Error=r.CustomFunctionError,window.CustomFunctions.ErrorCode=r.ErrorCode,r.setCustomFunctionAssociation(window.CustomFunctions._association),e&&("loading"===document.readyState?document.addEventListener("DOMContentLoaded",t):t())}},function(e,t,n){"use strict";var r,o=this&&this.__extends||(r=function(e,t){return(r=Object.setPrototypeOf||{__proto__:[]}instanceof Array&&function(e,t){e.__proto__=t}||function(e,t){for(var n in t)t.hasOwnProperty(n)&&(e[n]=t[n])})(e,t)},function(e,t){function n(){this.constructor=e}r(e,t),e.prototype=null===t?Object.create(t):(n.prototype=t.prototype,new n)});Object.defineProperty(t,"__esModule",{value:!0});var i=n(2),s=n(0),a=i.BatchApiHelper.createPropertyObject,c=(i.BatchApiHelper.createMethodObject,i.BatchApiHelper.createIndexerObject,i.BatchApiHelper.createRootServiceObject),u=i.BatchApiHelper.createTopLevelServiceObject,l=(i.BatchApiHelper.createChildItemObject,i.BatchApiHelper.invokeMethod),d=(i.BatchApiHelper.invokeEnsureUnchanged,i.BatchApiHelper.invokeSetProperty,i.Utility.isNullOrUndefined),f=(i.Utility.isUndefined,i.Utility.throwIfNotLoaded,i.Utility.throwIfApiNotSupported),h=i.Utility.load,p=(i.Utility.retrieve,i.Utility.toJson),m=i.Utility.fixObjectPathIfNecessary,g=i.Utility._handleNavigationPropertyResults,y=(i.Utility.adjustToDateTime,i.Utility.processRetrieveResult),_=(i.Utility.setMockData,function(e){function t(t){var n=e.call(this,t)||this;return n.m_customFunctions=j.newObject(n),n.m_container=c(A,n),n._rootObject=n.m_container,n._rootObjectPropertyName="customFunctionsContainer",n._requestFlagModifier=2176,n}return o(t,e),Object.defineProperty(t.prototype,"customFunctions",{get:function(){return this.m_customFunctions},enumerable:!0,configurable:!0}),Object.defineProperty(t.prototype,"customFunctionsContainer",{get:function(){return this.m_container},enumerable:!0,configurable:!0}),t.prototype._processOfficeJsErrorResponse=function(e,t){5004===e&&(t.ErrorCode=E.invalidOperationInCellEditMode,t.ErrorMessage=i.Utility._getResourceString(i.ResourceStrings.invalidOperationInCellEditMode))},t}(i.ClientRequestContext));t.Script={_CustomFunctionMetadata:{}};var b,v=function(){function e(e,t,n,r){this._functionName=e,d(t)||(this._address=t),this.setResult=n,this.setError=r}return Object.defineProperty(e.prototype,"onCanceled",{get:function(){if(!d(this._onCanceled)&&"function"==typeof this._onCanceled)return this._onCanceled},set:function(e){this._onCanceled=e},enumerable:!0,configurable:!0}),Object.defineProperty(e.prototype,"functionName",{get:function(){return this._functionName},enumerable:!0,configurable:!0}),Object.defineProperty(e.prototype,"address",{get:function(){return this._address},enumerable:!0,configurable:!0}),e}();t.InvocationContext=v,function(e){e.Info="Medium",e.Error="Unexpected",e.Verbose="Verbose"}(b||(b={}));var P,O=function(e,t){this.Severity=e,this.Message=t},I=function(){function e(){}return e.logEvent=function(t,n,r){if(e.s_shouldLog||i.CoreUtility._logEnabled){var o=t.Severity+" "+t.Message+n;if(r&&(o=o+" "+r),i.Utility.log(o),e.s_shouldLog)switch(t.Severity){case b.Verbose:null!==console.log&&console.log(o);break;case b.Info:null!==console.info&&console.info(o);break;case b.Error:null!==console.error&&console.error(o)}}},e.shouldLog=function(){try{return!d(console)&&!d(window)&&window.name&&"string"==typeof window.name&&JSON.parse(window.name)[e.CustomFunctionLoggingFlag]}catch(e){return i.Utility.log(JSON.stringify(e)),!1}},e.CustomFunctionLoggingFlag="CustomFunctionsRuntimeLogging",e.s_shouldLog=e.shouldLog(),e}();!function(e){e.invalidValue="#VALUE!",e.notAvailable="#N/A",e.divisionByZero="#DIV/0!",e.invalidNumber="#NUM!",e.nullReference="#NULL!"}(P=t.ErrorCode||(t.ErrorCode={}));var R=function(e){function t(n,r){var o=e.call(this,r)||this;return Object.setPrototypeOf(o,t.prototype),o._code=n||P.invalidValue,o._message=r||"",o}return o(t,e),Object.defineProperty(t.prototype,"code",{get:function(){return this._code},enumerable:!0,configurable:!0}),Object.defineProperty(t.prototype,"message",{get:function(){return this._message},enumerable:!0,configurable:!0}),t}(Error);t.CustomFunctionError=R;var C=function(){function e(){this._whenInit=void 0,this._isInit=!1,this._setResultsDelayMillis=50,this._setResultsOverdueDelayMillis=2e3,this._maxContextSyncExecutionDurationMills=15e3,this._minContextSyncIntervalMills=500,this._setResultsLifeMillis=6e4,this._ensureInitRetryDelayMillis=500,this._resultEntryBuffer={},this._isSetResultsTaskScheduled=!1,this._setResultsTaskOverdueTime=0,this._inProgressContextSyncExpectedFinishTime=0,this._batchQuotaMillis=1e3,this._invocationContextMap={}}return e.prototype._initSettings=function(){if("object"==typeof t.Script&&"object"==typeof t.Script._CustomFunctionSettings){if("number"==typeof t.Script._CustomFunctionSettings.setResultsDelayMillis){var e=t.Script._CustomFunctionSettings.setResultsDelayMillis;e=Math.max(0,e),e=Math.min(1e3,e),this._setResultsDelayMillis=e}if("number"==typeof t.Script._CustomFunctionSettings.ensureInitRetryDelayMillis){var n=t.Script._CustomFunctionSettings.ensureInitRetryDelayMillis;n=Math.max(0,n),n=Math.min(2e3,n),this._ensureInitRetryDelayMillis=n}if("number"==typeof t.Script._CustomFunctionSettings.setResultsLifeMillis){var r=t.Script._CustomFunctionSettings.setResultsLifeMillis;r=Math.max(0,r),r=Math.min(6e5,r),this._setResultsLifeMillis=r}if("number"==typeof t.Script._CustomFunctionSettings.batchQuotaMillis){var o=t.Script._CustomFunctionSettings.batchQuotaMillis;o=Math.max(0,o),o=Math.min(1e3,o),this._batchQuotaMillis=o}}},e.prototype.ensureInit=function(e){var t=this;return this._initSettings(),void 0===this._whenInit&&(this._whenInit=i.Utility._createPromiseFromResult(null).then((function(){if(!t._isInit)return e.eventRegistration.register(5,"",t._handleMessage.bind(t))})).then((function(){t._isInit=!0}))),this._isInit||e._pendingRequest._addPreSyncPromise(this._whenInit),this._whenInit},e.prototype.setCustomFunctionInvoker=function(e){"object"==typeof CustomFunctionMappings&&delete CustomFunctionMappings.__delay__,this._invoker=e},e.prototype.setCustomFunctionAssociation=function(e){var t=this;this._customFunctionMappingsUpperCase=void 0,this._association=e,this._association&&this._association.onchange((function(){t._customFunctionMappingsUpperCase=void 0}))},e.prototype._initFromHostBridge=function(e){var t=this;this._initSettings(),e.addHostMessageHandler((function(e){3===e.type?t._handleMessage(e.message):4===e.type&&t._handleSettings(e.message)})),this._isInit=!0,this._whenInit=i.CoreUtility.Promise.resolve()},e.prototype._handleSettings=function(e){i.Utility.log("CustomFunctionProxy._handleSettings:"+JSON.stringify(e)),e&&"object"==typeof e&&(I.s_shouldLog=e[I.CustomFunctionLoggingFlag])},e.prototype._handleMessage=function(t){try{i.Utility.log("CustomFunctionProxy._handleMessage"),i.Utility.checkArgumentNull(t,"args");for(var n=t.entries,r=[],o=[],s=[],a=0;a<n.length;a++)1===n[a].messageCategory&&("string"==typeof n[a].message&&(n[a].message=JSON.parse(n[a].message)),1e3===n[a].messageType?r.push(n[a]):1001===n[a].messageType?o.push(n[a]):1002===n[a].messageType?s.push(n[a]):i.Utility.log("CustomFunctionProxy._handleMessage unknown message type "+n[a].messageType));if(s.length>0&&this._handleMetadataEntries(s),r.length>0){var c=this._batchInvocationEntries(r);c.length>0&&this._invokeRemainingBatchEntries(c,0)}o.length>0&&this._handleCancellationEntries(o)}catch(t){throw e._tryLog(t),t}return i.Utility._createPromiseFromResult(null)},e.toLogMessage=function(e){var t="Unknown Error";if(e)try{e.toString&&(t=e.toString()),t=t+" "+JSON.stringify(e)}catch(e){t="Unexpected Error"}return t},e._tryLog=function(t){var n=e.toLogMessage(t);i.Utility.log(n)},e.prototype._handleMetadataEntries=function(e){for(var n=0;n<e.length;n++){var r=e[n].message;if(d(r))throw i.Utility.createRuntimeError(E.generalException,"message","CustomFunctionProxy._handleMetadataEntries");t.Script._CustomFunctionMetadata[r.functionName]={options:{stream:r.isStream,cancelable:r.isCancelable}}}},e.prototype._handleCancellationEntries=function(t){for(var n=0;n<t.length;n++){var r=t[n].message;if(d(r))throw i.Utility.createRuntimeError(E.generalException,"message","CustomFunctionProxy._handleCancellationEntries");var o=r.invocationId,s=this._invocationContextMap[o];d(s)||(delete this._invocationContextMap[o],I.logEvent(e.CustomFunctionCancellation,s.functionName),d(s.onCanceled)||s.onCanceled())}},e.prototype._batchInvocationEntries=function(n){for(var r=this,o=[],s=function(s){var c,u=n[s].message;if(c=Array.isArray(u)?{invocationId:u[0],functionName:u[1],parameterValues:u[2],address:u[3],flags:u[4]}:u,d(c))throw i.Utility.createRuntimeError(E.generalException,"message","CustomFunctionProxy._batchInvocationEntries");if(d(c.invocationId)||c.invocationId<0)throw i.Utility.createRuntimeError(E.generalException,"invocationId","CustomFunctionProxy._batchInvocationEntries");if(d(c.functionName))throw i.Utility.createRuntimeError(E.generalException,"functionName","CustomFunctionProxy._batchInvocationEntries");var l=null,f=!1,h=!1;if("number"==typeof c.flags)f=0!=(1&c.flags),h=0!=(2&c.flags);else{var p=t.Script._CustomFunctionMetadata[c.functionName];if(d(p))return I.logEvent(e.CustomFunctionExecutionNotFoundLog,c.functionName),a._setError(c.invocationId,"N/A",1),"continue";f=p.options.cancelable,h=p.options.stream}if(a._invoker&&!a._customFunctionMappingsContains(c.functionName))return a._invokeFunctionUsingInvoker(c),"continue";try{l=a._getFunction(c.functionName)}catch(t){return I.logEvent(e.CustomFunctionExecutionNotFoundLog,c.functionName),a._setError(c.invocationId,t,1),"continue"}var m=void 0;if(h||f){var g=void 0,y=void 0;h&&(g=function(t){r._invocationContextMap[c.invocationId]?r._setResult(c.invocationId,t):I.logEvent(e.CustomFunctionAlreadyCancelled,c.functionName)},y=function(t){r._invocationContextMap[c.invocationId]?r._setError(c.invocationId,t.message,r._getCustomFunctionResultErrorCodeFromErrorCode(t.code)):I.logEvent(e.CustomFunctionAlreadyCancelled,c.functionName)}),m=new v(c.functionName,c.address,g,y),a._invocationContextMap[c.invocationId]=m}else m=new v(c.functionName,c.address);c.parameterValues.push(m),o.push({call:l,isBatching:!1,isStreaming:h,invocationIds:[c.invocationId],parameterValueSets:[c.parameterValues],functionName:c.functionName})},a=this,c=0;c<n.length;c++)s(c);return o},e.prototype._invokeFunctionUsingInvoker=function(e){var t=this,n=0!=(1&e.flags),r=0!=(2&e.flags),o=e.invocationId,i=void 0,s=void 0;if(r)i=function(e){t._invocationContextMap[o]&&t._setResult(o,e)},s=function(e){t._invocationContextMap[o]&&t._setError(o,e.message,t._getCustomFunctionResultErrorCodeFromErrorCode(e.code))};else{var a=!1;i=function(e){a||t._setResult(o,e),a=!0},s=function(e){a||t._setError(o,e.message,t._getCustomFunctionResultErrorCodeFromErrorCode(e.code)),a=!0}}var c=new v(e.functionName,e.address,i,s);(r||n)&&(this._invocationContextMap[o]=c),this._invoker.invoke(e.functionName,e.parameterValues,c)},e.prototype._ensureCustomFunctionMappingsUpperCase=function(){if(d(this._customFunctionMappingsUpperCase)){if(this._customFunctionMappingsUpperCase={},"object"==typeof CustomFunctionMappings)for(var t in i.CoreUtility.log("CustomFunctionMappings.Keys="+JSON.stringify(Object.keys(CustomFunctionMappings))),CustomFunctionMappings)this._customFunctionMappingsUpperCase[t.toUpperCase()]&&I.logEvent(e.CustomFunctionDuplicatedName,t),this._customFunctionMappingsUpperCase[t.toUpperCase()]=CustomFunctionMappings[t];if(this._association)for(var t in i.CoreUtility.log("CustomFunctionAssociateMappings.Keys="+JSON.stringify(Object.keys(this._association.mappings))),this._association.mappings)this._customFunctionMappingsUpperCase[t.toUpperCase()]&&I.logEvent(e.CustomFunctionDuplicatedName,t),this._customFunctionMappingsUpperCase[t.toUpperCase()]=this._association.mappings[t]}},e.prototype._customFunctionMappingsContains=function(e){this._ensureCustomFunctionMappingsUpperCase();var t=e.toUpperCase();if(!d(this._customFunctionMappingsUpperCase[t]))return!0;if("undefined"!=typeof window){for(var n=window,r=e.split("."),o=0;o<r.length-1;o++)if(n=n[r[o]],d(n)||"object"!=typeof n)return!1;if("function"==typeof n[r[r.length-1]])return!0}return!1},e.prototype._getCustomFunctionMappings=function(e){this._ensureCustomFunctionMappingsUpperCase();var t=e.toUpperCase();if(!d(this._customFunctionMappingsUpperCase[t])){if("function"==typeof this._customFunctionMappingsUpperCase[t])return this._customFunctionMappingsUpperCase[t];throw i.Utility.createRuntimeError(E.invalidOperation,i.Utility._getResourceString(i.ResourceStrings.customFunctionInvalidFunctionMapping,e),"CustomFunctionProxy._getCustomFunctionMappings")}},e.prototype._getFunction=function(e){var t=this._getCustomFunctionMappings(e);if(!d(t))return t;if(d(window))throw i.Utility.createRuntimeError(E.invalidOperation,i.Utility._getResourceString(i.ResourceStrings.customFunctionWindowMissing),"CustomFunctionProxy._getFunction");for(var n=window,r=e.split("."),o=0;o<r.length-1;o++)if(n=n[r[o]],d(n)||"object"!=typeof n)throw i.Utility.createRuntimeError(E.invalidOperation,i.Utility._getResourceString(i.ResourceStrings.customFunctionDefintionMissingOnWindow,e),"CustomFunctionProxy._getFunction");if("function"!=typeof(t=n[r[r.length-1]]))throw i.Utility.createRuntimeError(E.invalidOperation,i.Utility._getResourceString(i.ResourceStrings.customFunctionInvalidFunction,e),"CustomFunctionProxy._getFunction");return t},e.prototype._invokeRemainingBatchEntries=function(e,t){i.Utility.log("CustomFunctionProxy._invokeRemainingBatchEntries");for(var n=Date.now(),r=t;r<e.length;r++){if(!(Date.now()-n<this._batchQuotaMillis)){i.Utility.log("setTimeout(CustomFunctionProxy._invokeRemainingBatchEntries)"),setTimeout(this._invokeRemainingBatchEntries.bind(this),0,e,r);break}this._invokeFunctionAndSetResult(e[r])}},e.prototype._invokeFunctionAndSetResult=function(t){var n,r=this;I.logEvent(e.CustomFunctionExecutionStartLog,t.functionName);try{n=t.isBatching?t.call.call(null,t.parameterValueSets):[t.call.apply(null,t.parameterValueSets[0])]}catch(n){for(var o=0;o<t.invocationIds.length;o++)n instanceof R?this._setError(t.invocationIds[o],n.message,this._getCustomFunctionResultErrorCodeFromErrorCode(n.code)):this._setError(t.invocationIds[o],n,2);return void I.logEvent(e.CustomFunctionExecutionExceptionThrownLog,t.functionName,e.toLogMessage(n))}if(t.isStreaming);else if(n.length===t.parameterValueSets.length){var s=function(o){d(n[o])||"object"!=typeof n[o]||"function"!=typeof n[o].then?(I.logEvent(e.CustomFunctionExecutionFinishLog,t.functionName),a._setResult(t.invocationIds[o],n[o])):n[o].then((function(n){I.logEvent(e.CustomFunctionExecutionFinishLog,t.functionName),r._setResult(t.invocationIds[o],n)}),(function(n){I.logEvent(e.CustomFunctionExecutionRejectedPromoseLog,t.functionName,e.toLogMessage(n)),r._setError(t.invocationIds[o],n,3)}))},a=this;for(o=0;o<n.length;o++)s(o)}else for(I.logEvent(e.CustomFunctionExecutionBatchMismatchLog,t.functionName),o=0;o<t.invocationIds.length;o++)this._setError(t.invocationIds[o],i.Utility._getResourceString(i.ResourceStrings.customFunctionUnexpectedNumberOfEntriesInResultBatch),4)},e.prototype._setResult=function(t,n){var r={id:t,value:n};"number"==typeof n?isNaN(n)?(r.failed=!0,r.value="NaN"):isFinite(n)||(r.failed=!0,r.value="Infinity",r.errorCode=6):n instanceof Error&&(r.failed=!0,r.value=e.toLogMessage(n),r.errorCode=0);var o=Date.now();this._resultEntryBuffer[t]={timeCreated:o,result:r},this._ensureSetResultsTaskIsScheduled(o)},e.prototype._setError=function(e,t,n){var r="";d(t)||(t instanceof R&&!d(t.message)?r=t.message:"string"==typeof t&&(r=t));var o={id:e,failed:!0,value:r,errorCode:n},i=Date.now();this._resultEntryBuffer[e]={timeCreated:i,result:o},this._ensureSetResultsTaskIsScheduled(i)},e.prototype._getCustomFunctionResultErrorCodeFromErrorCode=function(e){var t;switch(e){case P.notAvailable:t=1;break;case P.divisionByZero:t=5;break;case P.invalidValue:t=7;break;case P.invalidNumber:t=6;break;case P.nullReference:t=8;break;default:t=7}return t},e.prototype._ensureSetResultsTaskIsScheduled=function(e){if(this._setResultsTaskOverdueTime>0&&e>this._setResultsTaskOverdueTime)return i.Utility.log("SetResultsTask overdue"),void this._executeSetResultsTask();this._isSetResultsTaskScheduled||(i.Utility.log("setTimeout(CustomFunctionProxy._executeSetResultsTask)"),setTimeout(this._executeSetResultsTask.bind(this),this._setResultsDelayMillis),this._isSetResultsTaskScheduled=!0,this._setResultsTaskOverdueTime=e+this._setResultsDelayMillis+this._setResultsOverdueDelayMillis)},e.prototype._convertCustomFunctionInvocationResultToArray=function(e){var t=[];return t.push(e.id),t.push(!e.failed),i.CoreUtility.isUndefined(e.value)?t.push(null):t.push(e.value),e.failed&&(i.CoreUtility.isUndefined(e.errorCode)?t.push(0):t.push(e.errorCode)),t},e.prototype._executeSetResultsTask=function(){var e=this;i.Utility.log("CustomFunctionProxy._executeSetResultsTask");var t=Date.now();if(this._inProgressContextSyncExpectedFinishTime>0&&this._inProgressContextSyncExpectedFinishTime>t)return i.Utility.log("context.sync() is in progress. setTimeout(CustomFunctionProxy._executeSetResultsTask)"),setTimeout(this._executeSetResultsTask.bind(this),this._setResultsDelayMillis),void(this._setResultsTaskOverdueTime=t+this._setResultsDelayMillis+this._setResultsOverdueDelayMillis);this._isSetResultsTaskScheduled=!1,this._setResultsTaskOverdueTime=0;var n=this._resultEntryBuffer;this._resultEntryBuffer={};var r=i.Utility.isSetSupported("CustomFunctions","1.7"),o=[];for(var s in n)r?o.push(this._convertCustomFunctionInvocationResultToArray(n[s].result)):o.push(n[s].result);if(0!==o.length){var a=new _;r?a.customFunctions.setInvocationArrayResults(o):a.customFunctions.setInvocationResults(o);var c=Date.now();this._inProgressContextSyncExpectedFinishTime=c+this._maxContextSyncExecutionDurationMills,a.sync().then((function(t){e._clearInProgressContextSyncExpectedFinishTimeAfterMinInterval(Date.now()-c)}),(function(t){var r=Date.now();e._clearInProgressContextSyncExpectedFinishTimeAfterMinInterval(r-c),e._restoreResultEntries(r,n),e._ensureSetResultsTaskIsScheduled(r)}))}},e.prototype._restoreResultEntries=function(e,t){for(var n in t){var r=t[n];e-r.timeCreated<=this._setResultsLifeMillis&&(this._resultEntryBuffer[n]||(this._resultEntryBuffer[n]=r))}},e.prototype._clearInProgressContextSyncExpectedFinishTimeAfterMinInterval=function(e){var t=this,n=Math.max(this._minContextSyncIntervalMills,2*e);i.Utility.log("setTimeout(clearInProgressContestSyncExpectedFinishedTime,"+n+")"),setTimeout((function(){i.Utility.log("clearInProgressContestSyncExpectedFinishedTime"),t._inProgressContextSyncExpectedFinishTime=0}),n)},e.CustomFunctionExecutionStartLog=new O(b.Verbose,"CustomFunctions [Execution] [Begin] Function="),e.CustomFunctionExecutionFailureLog=new O(b.Error,"CustomFunctions [Execution] [End] [Failure] Function="),e.CustomFunctionExecutionRejectedPromoseLog=new O(b.Error,"CustomFunctions [Execution] [End] [Failure] [RejectedPromise] Function="),e.CustomFunctionExecutionExceptionThrownLog=new O(b.Error,"CustomFunctions [Execution] [End] [Failure] [ExceptionThrown] Function="),e.CustomFunctionExecutionBatchMismatchLog=new O(b.Error,"CustomFunctions [Execution] [End] [Failure] [BatchMismatch] Function="),e.CustomFunctionExecutionFinishLog=new O(b.Info,"CustomFunctions [Execution] [End] [Success] Function="),e.CustomFunctionExecutionNotFoundLog=new O(b.Error,"CustomFunctions [Execution] [NotFound] Function="),e.CustomFunctionCancellation=new O(b.Info,"CustomFunctions [Cancellation] Function="),e.CustomFunctionAlreadyCancelled=new O(b.Info,"CustomFunctions [AlreadyCancelled] Function="),e.CustomFunctionDuplicatedName=new O(b.Error,"CustomFunctions [DuplicatedName] Function="),e.CustomFunctionInvalidArg=new O(b.Error,"CustomFunctions [InvalidArg] Name="),e}();t.CustomFunctionProxy=C,t.customFunctionProxy=new C,t.setCustomFunctionAssociation=t.customFunctionProxy.setCustomFunctionAssociation.bind(t.customFunctionProxy),t.setCustomFunctionInvoker=t.customFunctionProxy.setCustomFunctionInvoker.bind(t.customFunctionProxy),s.HostBridge.onInited((function(e){t.customFunctionProxy._initFromHostBridge(e)}));var j=function(e){function n(){return null!==e&&e.apply(this,arguments)||this}return o(n,e),Object.defineProperty(n.prototype,"_className",{get:function(){return"CustomFunctions"},enumerable:!0,configurable:!0}),n.initialize=function(){var e=new _;return t.customFunctionProxy.ensureInit(e).then((function(){return e.customFunctions._SetOsfControlContainerReadyForCustomFunctions(),e._customData="SetOsfControlContainerReadyForCustomFunctions",e.sync().catch((function(e){return function(e,r){var o=e instanceof i.Error&&e.code===E.invalidOperationInCellEditMode;if(i.CoreUtility.log("Error on starting custom functions: "+e),o){i.CoreUtility.log("Was in cell-edit mode, will try again");var s=t.customFunctionProxy._ensureInitRetryDelayMillis;return new i.CoreUtility.Promise((function(e){return setTimeout(e,s)})).then((function(){return n.initialize()}))}throw e}(e)}))}))},n.prototype.setInvocationArrayResults=function(e){f("CustomFunctions.setInvocationArrayResults","CustomFunctions","1.4","Excel"),l(this,"SetInvocationArrayResults",0,[e],2,0)},n.prototype.setInvocationResults=function(e){l(this,"SetInvocationResults",0,[e],2,0)},n.prototype._SetInvocationError=function(e,t){l(this,"_SetInvocationError",0,[e,t],2,0)},n.prototype._SetInvocationResult=function(e,t){l(this,"_SetInvocationResult",0,[e,t],2,0)},n.prototype._SetOsfControlContainerReadyForCustomFunctions=function(){l(this,"_SetOsfControlContainerReadyForCustomFunctions",0,[],10,0)},n.prototype._handleResult=function(t){e.prototype._handleResult.call(this,t),d(t)||m(this,t)},n.prototype._handleRetrieveResult=function(t,n){e.prototype._handleRetrieveResult.call(this,t,n),y(this,t,n)},n.newObject=function(e){return u(n,e,"Microsoft.ExcelServices.CustomFunctions",!1,4)},n.prototype.toJSON=function(){return p(this,{},{})},n}(i.ClientObject);t.CustomFunctions=j;var E,A=function(e){function t(){return null!==e&&e.apply(this,arguments)||this}return o(t,e),Object.defineProperty(t.prototype,"_className",{get:function(){return"CustomFunctionsContainer"},enumerable:!0,configurable:!0}),Object.defineProperty(t.prototype,"_navigationPropertyNames",{get:function(){return["customFunctions"]},enumerable:!0,configurable:!0}),Object.defineProperty(t.prototype,"customFunctions",{get:function(){return f("CustomFunctionsContainer.customFunctions","CustomFunctions","1.2","Excel"),this._C||(this._C=a(j,this,"CustomFunctions",!1,4)),this._C},enumerable:!0,configurable:!0}),t.prototype._handleResult=function(t){if(e.prototype._handleResult.call(this,t),!d(t)){var n=t;m(this,n),g(this,n,["customFunctions","CustomFunctions"])}},t.prototype.load=function(e){return h(this,e)},t.prototype._handleRetrieveResult=function(t,n){e.prototype._handleRetrieveResult.call(this,t,n),y(this,t,n)},t.prototype.toJSON=function(){return p(this,{},{})},t}(i.ClientObject);t.CustomFunctionsContainer=A,function(e){e.generalException="GeneralException",e.invalidOperation="InvalidOperation",e.invalidOperationInCellEditMode="InvalidOperationInCellEditMode"}(E||(E={}))}]);



var oteljs=function(e){var t={};function n(r){if(t[r])return t[r].exports;var i=t[r]={i:r,l:!1,exports:{}};return e[r].call(i.exports,i,i.exports,n),i.l=!0,i.exports}return n.m=e,n.c=t,n.d=function(e,t,r){n.o(e,t)||Object.defineProperty(e,t,{enumerable:!0,get:r})},n.r=function(e){"undefined"!=typeof Symbol&&Symbol.toStringTag&&Object.defineProperty(e,Symbol.toStringTag,{value:"Module"}),Object.defineProperty(e,"__esModule",{value:!0})},n.t=function(e,t){if(1&t&&(e=n(e)),8&t)return e;if(4&t&&"object"==typeof e&&e&&e.__esModule)return e;var r=Object.create(null);if(n.r(r),Object.defineProperty(r,"default",{enumerable:!0,value:e}),2&t&&"string"!=typeof e)for(var i in e)n.d(r,i,function(t){return e[t]}.bind(null,i));return r},n.n=function(e){var t=e&&e.__esModule?function(){return e.default}:function(){return e};return n.d(t,"a",t),t},n.o=function(e,t){return Object.prototype.hasOwnProperty.call(e,t)},n.p="",n(n.s=0)}([function(e,t,n){e.exports=n(1)},function(e,t,n){"use strict";var r,i,o,a,s,u,c,f,l,d,p,v;function y(e,t){return{name:e,dataType:r.Boolean,value:t}}function h(e,t){return{name:e,dataType:r.Int64,value:t}}function g(e,t){return{name:e,dataType:r.Double,value:t}}function m(e,t){return{name:e,dataType:r.String,value:t}}function S(e,t,n){var r=n.map((function(t){return{name:e+"."+t.name,value:t.value,dataType:t.dataType}}));return T(r,e,t),r}function T(e,t,n){e.push(m("zC."+t,n))}n.r(t),function(e){e[e.String=0]="String",e[e.Boolean=1]="Boolean",e[e.Int64=2]="Int64",e[e.Double=3]="Double"}(r||(r={})),o=i||(i={}),a="Office.System.Result",o.getFields=function(e,t){var n=[];return n.push(h(e+".Code",t.code)),void 0!==t.type&&n.push(m(e+".Type",t.type)),void 0!==t.tag&&n.push(h(e+".Tag",t.tag)),void 0!==t.isExpected&&n.push(y(e+".IsExpected",t.isExpected)),T(n,e,a),n},(u=s||(s={})).contractName="Office.System.Activity",u.getFields=function(e){var t=[];return void 0!==e.cV&&t.push(m("Activity.CV",e.cV)),t.push(h("Activity.Duration",e.duration)),t.push(h("Activity.Count",e.count)),t.push(h("Activity.AggMode",e.aggMode)),void 0!==e.success&&t.push(y("Activity.Success",e.success)),void 0!==e.result&&t.push.apply(t,i.getFields("Activity.Result",e.result)),T(t,"Activity",u.contractName),t},function(e){var t="Office.System.Host";e.getFields=function(e,n){var r=[];return void 0!==n.id&&r.push(m(e+".Id",n.id)),void 0!==n.version&&r.push(m(e+".Version",n.version)),void 0!==n.sessionId&&r.push(m(e+".SessionId",n.sessionId)),T(r,e,t),r}}(c||(c={})),function(e){var t="Office.System.User";e.getFields=function(e,n){var r=[];return void 0!==n.alias&&r.push(m(e+".Alias",n.alias)),void 0!==n.primaryIdentityHash&&r.push(m(e+".PrimaryIdentityHash",n.primaryIdentityHash)),void 0!==n.primaryIdentitySpace&&r.push(m(e+".PrimaryIdentitySpace",n.primaryIdentitySpace)),void 0!==n.tenantId&&r.push(m(e+".TenantId",n.tenantId)),void 0!==n.tenantGroup&&r.push(m(e+".TenantGroup",n.tenantGroup)),void 0!==n.isAnonymous&&r.push(y(e+".IsAnonymous",n.isAnonymous)),T(r,e,t),r}}(f||(f={})),function(e){var t="Office.System.SDX";e.getFields=function(e,n){var r=[];return void 0!==n.id&&r.push(m(e+".Id",n.id)),void 0!==n.version&&r.push(m(e+".Version",n.version)),void 0!==n.instanceId&&r.push(m(e+".InstanceId",n.instanceId)),void 0!==n.name&&r.push(m(e+".Name",n.name)),void 0!==n.marketplaceType&&r.push(m(e+".MarketplaceType",n.marketplaceType)),void 0!==n.sessionId&&r.push(m(e+".SessionId",n.sessionId)),void 0!==n.browserToken&&r.push(m(e+".BrowserToken",n.browserToken)),void 0!==n.osfRuntimeVersion&&r.push(m(e+".OsfRuntimeVersion",n.osfRuntimeVersion)),void 0!==n.officeJsVersion&&r.push(m(e+".OfficeJsVersion",n.officeJsVersion)),void 0!==n.hostJsVersion&&r.push(m(e+".HostJsVersion",n.hostJsVersion)),void 0!==n.assetId&&r.push(m(e+".AssetId",n.assetId)),void 0!==n.providerName&&r.push(m(e+".ProviderName",n.providerName)),void 0!==n.type&&r.push(m(e+".Type",n.type)),T(r,e,t),r}}(l||(l={})),function(e){var t="Office.System.Funnel";e.getFields=function(e,n){var r=[];return void 0!==n.name&&r.push(m(e+".Name",n.name)),void 0!==n.state&&r.push(m(e+".State",n.state)),T(r,e,t),r}}(d||(d={})),function(e){var t="Office.System.UserAction";e.getFields=function(e,n){var r=[];return void 0!==n.id&&r.push(h(e+".Id",n.id)),void 0!==n.name&&r.push(m(e+".Name",n.name)),void 0!==n.commandSurface&&r.push(m(e+".CommandSurface",n.commandSurface)),void 0!==n.parentName&&r.push(m(e+".ParentName",n.parentName)),void 0!==n.triggerMethod&&r.push(m(e+".TriggerMethod",n.triggerMethod)),void 0!==n.timeOffsetMs&&r.push(h(e+".TimeOffsetMs",n.timeOffsetMs)),T(r,e,t),r}}(p||(p={})),function(e){var t="Office.System.Error";e.getFields=function(e,n){var r=[];return r.push(m(e+".ErrorGroup",n.errorGroup)),r.push(h(e+".Tag",n.tag)),void 0!==n.code&&r.push(h(e+".Code",n.code)),void 0!==n.id&&r.push(h(e+".Id",n.id)),void 0!==n.count&&r.push(h(e+".Count",n.count)),T(r,e,t),r}}(v||(v={}));var F,b,w=s,C=i,E=v,N=d,A=c,I=l,k=p,x=f;!function(e){!function(e){!function(e){e.Activity=w,e.Result=C,e.Error=E,e.Funnel=N,e.Host=A,e.SDX=I,e.User=x,e.UserAction=k}(e.System||(e.System={}))}(e.Office||(e.Office={}))}(F||(F={})),function(e){var t,n=0;e.getNext=function(){return void 0===t&&(t=function(){for(var e="ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/",t=[],n=0;n<22;n++)t.push(e.charAt(Math.floor(Math.random()*e.length)));return t.join("")}()),new r(t,++n)},e.getNextChild=function(e){return new r(e.getString(),++e.nextChild)};var r=function(){function e(e,t){this.base=e,this.id=t,this.nextChild=0}return e.prototype.getString=function(){return this.base+"."+this.id},e}();e.CV=r}(b||(b={}));var _,O,P=function(){function e(){this._listeners=[]}return e.prototype.fireEvent=function(e){this._listeners.forEach((function(t){return t(e)}))},e.prototype.addListener=function(e){e&&this._listeners.push(e)},e.prototype.removeListener=function(e){this._listeners=this._listeners.filter((function(t){return t!==e}))},e.prototype.getListenerCount=function(){return this._listeners.length},e}(),D=new P;function M(){return D}function j(e,t,n){D.fireEvent({level:e,category:t,message:n})}function V(e,t,n){j(_.Error,e,(function(){var e=n instanceof Error?n.message:"";return t+": "+e}))}!function(e){e[e.Error=0]="Error",e[e.Warning=1]="Warning",e[e.Info=2]="Info",e[e.Verbose=3]="Verbose"}(_||(_={})),function(e){e[e.Core=0]="Core",e[e.Sink=1]="Sink",e[e.Transport=2]="Transport"}(O||(O={}));var L=function(e,t,n,r){return new(n||(n=Promise))((function(i,o){function a(e){try{u(r.next(e))}catch(e){o(e)}}function s(e){try{u(r.throw(e))}catch(e){o(e)}}function u(e){var t;e.done?i(e.value):(t=e.value,t instanceof n?t:new n((function(e){e(t)}))).then(a,s)}u((r=r.apply(e,t||[])).next())}))},U=function(e,t){var n,r,i,o,a={label:0,sent:function(){if(1&i[0])throw i[1];return i[1]},trys:[],ops:[]};return o={next:s(0),throw:s(1),return:s(2)},"function"==typeof Symbol&&(o[Symbol.iterator]=function(){return this}),o;function s(o){return function(s){return function(o){if(n)throw new TypeError("Generator is already executing.");for(;a;)try{if(n=1,r&&(i=2&o[0]?r.return:o[0]?r.throw||((i=r.return)&&i.call(r),0):r.next)&&!(i=i.call(r,o[1])).done)return i;switch(r=0,i&&(o=[2&o[0],i.value]),o[0]){case 0:case 1:i=o;break;case 4:return a.label++,{value:o[1],done:!1};case 5:a.label++,r=o[1],o=[0];continue;case 7:o=a.ops.pop(),a.trys.pop();continue;default:if(!(i=(i=a.trys).length>0&&i[i.length-1])&&(6===o[0]||2===o[0])){a=0;continue}if(3===o[0]&&(!i||o[1]>i[0]&&o[1]<i[3])){a.label=o[1];break}if(6===o[0]&&a.label<i[1]){a.label=i[1],i=o;break}if(i&&a.label<i[2]){a.label=i[2],a.ops.push(o);break}i[2]&&a.ops.pop(),a.trys.pop();continue}o=t.call(e,a)}catch(e){o=[6,e],r=0}finally{n=i=0}if(5&o[0])throw o[1];return{value:o[0]?o[1]:void 0,done:!0}}([o,s])}}},H=function(){return 1e3*Date.now()};"object"==typeof window.performance&&"now"in window.performance&&(H=function(){return 1e3*Math.floor(window.performance.now())});var B,R,J,G,z,W,Z,X,$,q=function(){function e(e,t,n){this._optionalEventFlags={},this._ended=!1,this._telemetryLogger=e,this._activityName=t,this._cv=n?b.getNextChild(n._cv):b.getNext(),this._dataFields=[],this._success=void 0,this._startTime=H()}return e.createNew=function(t,n){return new e(t,n)},e.prototype.createChildActivity=function(t){return new e(this._telemetryLogger,t,this)},e.prototype.setEventFlags=function(e){this._optionalEventFlags=e},e.prototype.addDataField=function(e){this._dataFields.push(e)},e.prototype.addDataFields=function(e){var t;(t=this._dataFields).push.apply(t,e)},e.prototype.setSuccess=function(e){this._success=e},e.prototype.setResult=function(e,t,n){this._result={code:e,type:t,tag:n}},e.prototype.endNow=function(){if(!this._ended){void 0===this._success&&void 0===this._result&&j(_.Warning,O.Core,(function(){return"Activity does not have success or result set"}));var e=H()-this._startTime;this._ended=!0;var t={duration:e,count:1,aggMode:0,cV:this._cv.getString(),success:this._success,result:this._result};return this._telemetryLogger.sendActivity(this._activityName,t,this._dataFields,this._optionalEventFlags)}j(_.Error,O.Core,(function(){return"Activity has already ended"}))},e.prototype.executeAsync=function(e){return L(this,void 0,void 0,(function(){var t=this;return U(this,(function(n){return[2,e(this).then((function(e){return t.endNow(),e})).catch((function(e){throw t.endNow(),e}))]}))}))},e.prototype.executeSync=function(e){try{var t=e(this);return this.endNow(),t}catch(e){throw this.endNow(),e}},e.prototype.executeChildActivityAsync=function(e,t){return L(this,void 0,void 0,(function(){return U(this,(function(n){return[2,this.createChildActivity(e).executeAsync(t)]}))}))},e.prototype.executeChildActivitySync=function(e,t){return this.createChildActivity(e).executeSync(t)},e}();function K(e){var t={costPriority:G.Normal,samplingPolicy:R.Measure,persistencePriority:J.Normal,dataCategories:z.NotSet,diagnosticLevel:W.FullEvent};return e.eventFlags&&e.eventFlags.dataCategories||j(_.Error,O.Core,(function(){return"Event is missing DataCategories event flag"})),e.eventFlags?(e.eventFlags.costPriority&&(t.costPriority=e.eventFlags.costPriority),e.eventFlags.samplingPolicy&&(t.samplingPolicy=e.eventFlags.samplingPolicy),e.eventFlags.persistencePriority&&(t.persistencePriority=e.eventFlags.persistencePriority),e.eventFlags.dataCategories&&(t.dataCategories=e.eventFlags.dataCategories),e.eventFlags.diagnosticLevel&&(t.diagnosticLevel=e.eventFlags.diagnosticLevel),t):t}!function(e){e[e.EssentialServiceMetadata=1]="EssentialServiceMetadata",e[e.AccountData=2]="AccountData",e[e.SystemMetadata=4]="SystemMetadata",e[e.OrganizationIdentifiableInformation=8]="OrganizationIdentifiableInformation",e[e.EndUserIdentifiableInformation=16]="EndUserIdentifiableInformation",e[e.CustomerContent=32]="CustomerContent",e[e.AccessControl=64]="AccessControl"}(B||(B={})),function(e){e[e.NotSet=0]="NotSet",e[e.Measure=1]="Measure",e[e.Diagnostics=2]="Diagnostics",e[e.CriticalBusinessImpact=191]="CriticalBusinessImpact",e[e.CriticalCensus=192]="CriticalCensus",e[e.CriticalExperimentation=193]="CriticalExperimentation",e[e.CriticalUsage=194]="CriticalUsage"}(R||(R={})),function(e){e[e.NotSet=0]="NotSet",e[e.Normal=1]="Normal",e[e.High=2]="High"}(J||(J={})),function(e){e[e.NotSet=0]="NotSet",e[e.Normal=1]="Normal",e[e.High=2]="High"}(G||(G={})),function(e){e[e.NotSet=0]="NotSet",e[e.SoftwareSetup=1]="SoftwareSetup",e[e.ProductServiceUsage=2]="ProductServiceUsage",e[e.ProductServicePerformance=4]="ProductServicePerformance",e[e.DeviceConfiguration=8]="DeviceConfiguration",e[e.InkingTypingSpeech=16]="InkingTypingSpeech"}(z||(z={})),function(e){e[e.ReservedDoNotUse=0]="ReservedDoNotUse",e[e.BasicEvent=10]="BasicEvent",e[e.FullEvent=100]="FullEvent",e[e.NecessaryServiceDataEvent=110]="NecessaryServiceDataEvent",e[e.AlwaysOnNecessaryServiceDataEvent=120]="AlwaysOnNecessaryServiceDataEvent"}(W||(W={})),function(e){e[e.Aria=0]="Aria",e[e.Nexus=1]="Nexus"}(Z||(Z={})),function(e){var t={},n={},r={};function i(e){if("object"!=typeof e)throw new Error("tokenTree must be an object");r=function e(t,n){if("object"!=typeof n)return n;for(var r=0,i=Object.keys(n);r<i.length;r++){var o=i[r];o in t&&(t[o],1)?t[o]=e(t[o],n[o]):t[o]=n[o]}return t}(r,e)}function o(e){if(t[e])return t[e];var n=s(e,Z.Aria);return"string"==typeof n?(t[e]=n,n):void 0}function a(e){if(n[e])return n[e];var t=s(e,Z.Nexus);return"number"==typeof t?(n[e]=t,t):void 0}function s(e,t){var n=e.split("."),i=r,o=void 0;if(i){for(var a=0;a<n.length-1;a++)i[n[a]]&&(i=i[n[a]],t===Z.Aria&&"string"==typeof i.ariaTenantToken?o=i.ariaTenantToken:t===Z.Nexus&&"number"==typeof i.nexusTenantToken&&(o=i.nexusTenantToken));return o}}e.setTenantToken=function(e,t,n){var r=e.split(".");if(r.length<2||"Office"!==r[0])j(_.Error,O.Core,(function(){return"Invalid namespace: "+e}));else{var o=Object.create(Object.prototype);t&&(o.ariaTenantToken=t),n&&(o.nexusTenantToken=n);var a,s=o;for(a=r.length-1;a>=0;--a){var u=Object.create(Object.prototype);u[r[a]]=s,s=u}i(s)}},e.setTenantTokens=i,e.getTenantTokens=function(e){var t=o(e),n=a(e);if(!n||!t)throw new Error("Could not find tenant token for "+e);return{ariaTenantToken:t,nexusTenantToken:n}},e.getAriaTenantToken=o,e.getNexusTenantToken=a,e.clear=function(){t={},n={},r={}}}(X||(X={})),function(e){var t=-9007199254740991,n=9007199254740991,i=/^[A-Z][a-zA-Z0-9]*$/,o=/^[a-zA-Z0-9_\.]*$/;function a(e){return void 0!==e&&o.test(e)}function s(e){if(!((t=e.name)&&a(t)&&t.length+5<100))throw new Error("Invalid dataField name");var t;e.dataType===r.Int64&&u(e.value)}function u(e){if("number"!=typeof e||!isFinite(e)||Math.floor(e)!==e||e<t||e>n)throw new Error("Invalid integer "+JSON.stringify(e))}e.validateTelemetryEvent=function(e){if(!function(e){if(!e||e.length>98)return!1;var t=e.split("."),n=t[t.length-1];return function(e){return!!e&&e.length>=3&&"Office"===e[0]}(t)&&(r=n,void 0!==r&&i.test(r));var r}(e.eventName))throw new Error("Invalid eventName");if(e.eventContract&&!a(e.eventContract.name))throw new Error("Invalid eventContract");if(null!=e.dataFields)for(var t=0;t<e.dataFields.length;t++)s(e.dataFields[t])},e.validateInt=u}($||($={}));var Q,Y="3.1.30",ee=function(){return(ee=Object.assign||function(e){for(var t,n=1,r=arguments.length;n<r;n++)for(var i in t=arguments[n])Object.prototype.hasOwnProperty.call(t,i)&&(e[i]=t[i]);return e}).apply(this,arguments)},te=function(){function e(e,t,n){var r,i;this.onSendEvent=new P,this.persistentDataFields=[],this.config=n||{},e?(this.onSendEvent=e.onSendEvent,(r=this.persistentDataFields).push.apply(r,e.persistentDataFields),this.config=ee(ee({},e.getConfig()),this.config)):this.persistentDataFields.push(m("OTelJS.Version",Y)),t&&(i=this.persistentDataFields).push.apply(i,t)}return e.prototype.sendTelemetryEvent=function(e){var t;try{if(0===this.onSendEvent.getListenerCount())return void j(_.Warning,O.Core,(function(){return"No telemetry sinks are attached."}));t=this.cloneEvent(e),this.processTelemetryEvent(t)}catch(e){return void V(O.Core,"SendTelemetryEvent",e)}try{this.onSendEvent.fireEvent(t)}catch(e){}},e.prototype.processTelemetryEvent=function(e){var t;e.telemetryProperties||(e.telemetryProperties=X.getTenantTokens(e.eventName)),(t=e.dataFields).push.apply(t,this.persistentDataFields),this.config.disableValidation||$.validateTelemetryEvent(e)},e.prototype.addSink=function(e){this.onSendEvent.addListener((function(t){return e.sendTelemetryEvent(t)}))},e.prototype.setTenantToken=function(e,t,n){X.setTenantToken(e,t,n)},e.prototype.setTenantTokens=function(e){X.setTenantTokens(e)},e.prototype.cloneEvent=function(e){var t={eventName:e.eventName,eventFlags:e.eventFlags};return e.telemetryProperties&&(t.telemetryProperties={ariaTenantToken:e.telemetryProperties.ariaTenantToken,nexusTenantToken:e.telemetryProperties.nexusTenantToken}),e.eventContract&&(t.eventContract={name:e.eventContract.name,dataFields:e.eventContract.dataFields.slice()}),t.dataFields=e.dataFields?e.dataFields.slice():[],t},e.prototype.getConfig=function(){return this.config},e}(),ne=(Q=function(e,t){return(Q=Object.setPrototypeOf||{__proto__:[]}instanceof Array&&function(e,t){e.__proto__=t}||function(e,t){for(var n in t)t.hasOwnProperty(n)&&(e[n]=t[n])})(e,t)},function(e,t){function n(){this.constructor=e}Q(e,t),e.prototype=null===t?Object.create(t):(n.prototype=t.prototype,new n)}),re=function(e,t,n,r){return new(n||(n=Promise))((function(i,o){function a(e){try{u(r.next(e))}catch(e){o(e)}}function s(e){try{u(r.throw(e))}catch(e){o(e)}}function u(e){var t;e.done?i(e.value):(t=e.value,t instanceof n?t:new n((function(e){e(t)}))).then(a,s)}u((r=r.apply(e,t||[])).next())}))},ie=function(e,t){var n,r,i,o,a={label:0,sent:function(){if(1&i[0])throw i[1];return i[1]},trys:[],ops:[]};return o={next:s(0),throw:s(1),return:s(2)},"function"==typeof Symbol&&(o[Symbol.iterator]=function(){return this}),o;function s(o){return function(s){return function(o){if(n)throw new TypeError("Generator is already executing.");for(;a;)try{if(n=1,r&&(i=2&o[0]?r.return:o[0]?r.throw||((i=r.return)&&i.call(r),0):r.next)&&!(i=i.call(r,o[1])).done)return i;switch(r=0,i&&(o=[2&o[0],i.value]),o[0]){case 0:case 1:i=o;break;case 4:return a.label++,{value:o[1],done:!1};case 5:a.label++,r=o[1],o=[0];continue;case 7:o=a.ops.pop(),a.trys.pop();continue;default:if(!(i=(i=a.trys).length>0&&i[i.length-1])&&(6===o[0]||2===o[0])){a=0;continue}if(3===o[0]&&(!i||o[1]>i[0]&&o[1]<i[3])){a.label=o[1];break}if(6===o[0]&&a.label<i[1]){a.label=i[1],i=o;break}if(i&&a.label<i[2]){a.label=i[2],a.ops.push(o);break}i[2]&&a.ops.pop(),a.trys.pop();continue}o=t.call(e,a)}catch(e){o=[6,e],r=0}finally{n=i=0}if(5&o[0])throw o[1];return{value:o[0]?o[1]:void 0,done:!0}}([o,s])}}},oe=function(e){function t(){return null!==e&&e.apply(this,arguments)||this}return ne(t,e),t.prototype.executeActivityAsync=function(e,t){return re(this,void 0,void 0,(function(){return ie(this,(function(n){return[2,this.createNewActivity(e).executeAsync(t)]}))}))},t.prototype.executeActivitySync=function(e,t){return this.createNewActivity(e).executeSync(t)},t.prototype.createNewActivity=function(e){return q.createNew(this,e)},t.prototype.sendActivity=function(e,t,n,r){return this.sendTelemetryEvent({eventName:e,eventContract:{name:F.Office.System.Activity.contractName,dataFields:F.Office.System.Activity.getFields(t)},dataFields:n,eventFlags:r})},t.prototype.sendError=function(e){var t=v.getFields("Error",e.error);return null!=e.dataFields&&t.push.apply(t,e.dataFields),this.sendTelemetryEvent({eventName:e.eventName,dataFields:t,eventFlags:e.eventFlags})},t}(te);n.d(t,"Contracts",(function(){return F})),n.d(t,"ActivityScope",(function(){return q})),n.d(t,"getFieldsForContract",(function(){return S})),n.d(t,"addContractField",(function(){return T})),n.d(t,"DataClassification",(function(){return B})),n.d(t,"makeBooleanDataField",(function(){return y})),n.d(t,"makeInt64DataField",(function(){return h})),n.d(t,"makeDoubleDataField",(function(){return g})),n.d(t,"makeStringDataField",(function(){return m})),n.d(t,"DataFieldType",(function(){return r})),n.d(t,"getEffectiveEventFlags",(function(){return K})),n.d(t,"SamplingPolicy",(function(){return R})),n.d(t,"PersistencePriority",(function(){return J})),n.d(t,"CostPriority",(function(){return G})),n.d(t,"DataCategories",(function(){return z})),n.d(t,"DiagnosticLevel",(function(){return W})),n.d(t,"LogLevel",(function(){return _})),n.d(t,"Category",(function(){return O})),n.d(t,"onNotification",(function(){return M})),n.d(t,"logNotification",(function(){return j})),n.d(t,"logError",(function(){return V})),n.d(t,"SuppressNexus",(function(){return-1})),n.d(t,"SimpleTelemetryLogger",(function(){return te})),n.d(t,"TelemetryLogger",(function(){return oe}))}]);