/* Excel specific API library (Core APIs only) */
/* Version: 16.0.10726.30000 */
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
var __extends = (this && this.__extends) || function (d, b) {
    for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p];
    function __() { this.constructor = d; }
    d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
};
var OfficeExt;
(function (OfficeExt) {
    var MicrosoftAjaxFactory = (function () {
        function MicrosoftAjaxFactory() {
        }
        MicrosoftAjaxFactory.prototype.isMsAjaxLoaded = function () {
            if (typeof (Sys) !== 'undefined' && typeof (Type) !== 'undefined' &&
                Sys.StringBuilder && typeof (Sys.StringBuilder) === "function" &&
                Type.registerNamespace && typeof (Type.registerNamespace) === "function" &&
                Type.registerClass && typeof (Type.registerClass) === "function" &&
                typeof (Function._validateParams) === "function" &&
                Sys.Serialization && Sys.Serialization.JavaScriptSerializer && typeof (Sys.Serialization.JavaScriptSerializer.serialize) === "function") {
                return true;
            }
            else {
                return false;
            }
        };
        MicrosoftAjaxFactory.prototype.loadMsAjaxFull = function (callback) {
            var msAjaxCDNPath = (window.location.protocol.toLowerCase() === 'https:' ? 'https:' : 'http:') + '//ajax.aspnetcdn.com/ajax/3.5/MicrosoftAjax.js';
            OSF.OUtil.loadScript(msAjaxCDNPath, callback);
        };
        Object.defineProperty(MicrosoftAjaxFactory.prototype, "msAjaxError", {
            get: function () {
                if (this._msAjaxError == null && this.isMsAjaxLoaded()) {
                    this._msAjaxError = Error;
                }
                return this._msAjaxError;
            },
            set: function (errorClass) {
                this._msAjaxError = errorClass;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(MicrosoftAjaxFactory.prototype, "msAjaxString", {
            get: function () {
                if (this._msAjaxString == null && this.isMsAjaxLoaded()) {
                    this._msAjaxString = String;
                }
                return this._msAjaxString;
            },
            set: function (stringClass) {
                this._msAjaxString = stringClass;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(MicrosoftAjaxFactory.prototype, "msAjaxDebug", {
            get: function () {
                if (this._msAjaxDebug == null && this.isMsAjaxLoaded()) {
                    this._msAjaxDebug = Sys.Debug;
                }
                return this._msAjaxDebug;
            },
            set: function (debugClass) {
                this._msAjaxDebug = debugClass;
            },
            enumerable: true,
            configurable: true
        });
        return MicrosoftAjaxFactory;
    })();
    OfficeExt.MicrosoftAjaxFactory = MicrosoftAjaxFactory;
})(OfficeExt || (OfficeExt = {}));
var OsfMsAjaxFactory = new OfficeExt.MicrosoftAjaxFactory();
var OSF = OSF || {};
var OfficeExt;
(function (OfficeExt) {
    var SafeStorage = (function () {
        function SafeStorage(_internalStorage) {
            this._internalStorage = _internalStorage;
        }
        SafeStorage.prototype.getItem = function (key) {
            try {
                return this._internalStorage && this._internalStorage.getItem(key);
            }
            catch (e) {
                return null;
            }
        };
        SafeStorage.prototype.setItem = function (key, data) {
            try {
                this._internalStorage && this._internalStorage.setItem(key, data);
            }
            catch (e) {
            }
        };
        SafeStorage.prototype.clear = function () {
            try {
                this._internalStorage && this._internalStorage.clear();
            }
            catch (e) {
            }
        };
        SafeStorage.prototype.removeItem = function (key) {
            try {
                this._internalStorage && this._internalStorage.removeItem(key);
            }
            catch (e) {
            }
        };
        SafeStorage.prototype.getKeysWithPrefix = function (keyPrefix) {
            var keyList = [];
            try {
                var len = this._internalStorage && this._internalStorage.length || 0;
                for (var i = 0; i < len; i++) {
                    var key = this._internalStorage.key(i);
                    if (key.indexOf(keyPrefix) === 0) {
                        keyList.push(key);
                    }
                }
            }
            catch (e) {
            }
            return keyList;
        };
        return SafeStorage;
    })();
    OfficeExt.SafeStorage = SafeStorage;
})(OfficeExt || (OfficeExt = {}));
OSF.XdmFieldName = {
    ConversationUrl: "ConversationUrl",
    AppId: "AppId"
};
OSF.WindowNameItemKeys = {
    BaseFrameName: "baseFrameName",
    HostInfo: "hostInfo",
    XdmInfo: "xdmInfo",
    SerializerVersion: "serializerVersion",
    AppContext: "appContext"
};
OSF.OUtil = (function () {
    var _uniqueId = -1;
    var _xdmInfoKey = '&_xdm_Info=';
    var _serializerVersionKey = '&_serializer_version=';
    var _xdmSessionKeyPrefix = '_xdm_';
    var _serializerVersionKeyPrefix = '_serializer_version=';
    var _fragmentSeparator = '#';
    var _fragmentInfoDelimiter = '&';
    var _classN = "class";
    var _loadedScripts = {};
    var _defaultScriptLoadingTimeout = 30000;
    var _safeSessionStorage = null;
    var _safeLocalStorage = null;
    var _rndentropy = new Date().getTime();
    function _random() {
        var nextrand = 0x7fffffff * (Math.random());
        nextrand ^= _rndentropy ^ ((new Date().getMilliseconds()) << Math.floor(Math.random() * (31 - 10)));
        return nextrand.toString(16);
    }
    ;
    function _getSessionStorage() {
        if (!_safeSessionStorage) {
            try {
                var sessionStorage = window.sessionStorage;
            }
            catch (ex) {
                sessionStorage = null;
            }
            _safeSessionStorage = new OfficeExt.SafeStorage(sessionStorage);
        }
        return _safeSessionStorage;
    }
    ;
    function _reOrderTabbableElements(elements) {
        var bucket0 = [];
        var bucketPositive = [];
        var i;
        var len = elements.length;
        var ele;
        for (i = 0; i < len; i++) {
            ele = elements[i];
            if (ele.tabIndex) {
                if (ele.tabIndex > 0) {
                    bucketPositive.push(ele);
                }
                else if (ele.tabIndex === 0) {
                    bucket0.push(ele);
                }
            }
            else {
                bucket0.push(ele);
            }
        }
        bucketPositive = bucketPositive.sort(function (left, right) {
            var diff = left.tabIndex - right.tabIndex;
            if (diff === 0) {
                diff = bucketPositive.indexOf(left) - bucketPositive.indexOf(right);
            }
            return diff;
        });
        return [].concat(bucketPositive, bucket0);
    }
    ;
    return {
        set_entropy: function OSF_OUtil$set_entropy(entropy) {
            if (typeof entropy == "string") {
                for (var i = 0; i < entropy.length; i += 4) {
                    var temp = 0;
                    for (var j = 0; j < 4 && i + j < entropy.length; j++) {
                        temp = (temp << 8) + entropy.charCodeAt(i + j);
                    }
                    _rndentropy ^= temp;
                }
            }
            else if (typeof entropy == "number") {
                _rndentropy ^= entropy;
            }
            else {
                _rndentropy ^= 0x7fffffff * Math.random();
            }
            _rndentropy &= 0x7fffffff;
        },
        extend: function OSF_OUtil$extend(child, parent) {
            var F = function () { };
            F.prototype = parent.prototype;
            child.prototype = new F();
            child.prototype.constructor = child;
            child.uber = parent.prototype;
            if (parent.prototype.constructor === Object.prototype.constructor) {
                parent.prototype.constructor = parent;
            }
        },
        setNamespace: function OSF_OUtil$setNamespace(name, parent) {
            if (parent && name && !parent[name]) {
                parent[name] = {};
            }
        },
        unsetNamespace: function OSF_OUtil$unsetNamespace(name, parent) {
            if (parent && name && parent[name]) {
                delete parent[name];
            }
        },
        serializeSettings: function OSF_OUtil$serializeSettings(settingsCollection) {
            var ret = {};
            for (var key in settingsCollection) {
                var value = settingsCollection[key];
                try {
                    if (JSON) {
                        value = JSON.stringify(value, function dateReplacer(k, v) {
                            return OSF.OUtil.isDate(this[k]) ? OSF.DDA.SettingsManager.DateJSONPrefix + this[k].getTime() + OSF.DDA.SettingsManager.DataJSONSuffix : v;
                        });
                    }
                    else {
                        value = Sys.Serialization.JavaScriptSerializer.serialize(value);
                    }
                    ret[key] = value;
                }
                catch (ex) {
                }
            }
            return ret;
        },
        deserializeSettings: function OSF_OUtil$deserializeSettings(serializedSettings) {
            var ret = {};
            serializedSettings = serializedSettings || {};
            for (var key in serializedSettings) {
                var value = serializedSettings[key];
                try {
                    if (JSON) {
                        value = JSON.parse(value, function dateReviver(k, v) {
                            var d;
                            if (typeof v === 'string' && v && v.length > 6 && v.slice(0, 5) === OSF.DDA.SettingsManager.DateJSONPrefix && v.slice(-1) === OSF.DDA.SettingsManager.DataJSONSuffix) {
                                d = new Date(parseInt(v.slice(5, -1)));
                                if (d) {
                                    return d;
                                }
                            }
                            return v;
                        });
                    }
                    else {
                        value = Sys.Serialization.JavaScriptSerializer.deserialize(value, true);
                    }
                    ret[key] = value;
                }
                catch (ex) {
                }
            }
            return ret;
        },
        loadScript: function OSF_OUtil$loadScript(url, callback, timeoutInMs) {
            if (url && callback) {
                var doc = window.document;
                var _loadedScriptEntry = _loadedScripts[url];
                if (!_loadedScriptEntry) {
                    var script = doc.createElement("script");
                    script.type = "text/javascript";
                    _loadedScriptEntry = { loaded: false, pendingCallbacks: [callback], timer: null };
                    _loadedScripts[url] = _loadedScriptEntry;
                    var onLoadCallback = function OSF_OUtil_loadScript$onLoadCallback() {
                        if (_loadedScriptEntry.timer != null) {
                            clearTimeout(_loadedScriptEntry.timer);
                            delete _loadedScriptEntry.timer;
                        }
                        _loadedScriptEntry.loaded = true;
                        var pendingCallbackCount = _loadedScriptEntry.pendingCallbacks.length;
                        for (var i = 0; i < pendingCallbackCount; i++) {
                            var currentCallback = _loadedScriptEntry.pendingCallbacks.shift();
                            currentCallback();
                        }
                    };
                    var onLoadError = function OSF_OUtil_loadScript$onLoadError() {
                        delete _loadedScripts[url];
                        if (_loadedScriptEntry.timer != null) {
                            clearTimeout(_loadedScriptEntry.timer);
                            delete _loadedScriptEntry.timer;
                        }
                        var pendingCallbackCount = _loadedScriptEntry.pendingCallbacks.length;
                        for (var i = 0; i < pendingCallbackCount; i++) {
                            var currentCallback = _loadedScriptEntry.pendingCallbacks.shift();
                            currentCallback();
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
                    timeoutInMs = timeoutInMs || _defaultScriptLoadingTimeout;
                    _loadedScriptEntry.timer = setTimeout(onLoadError, timeoutInMs);
                    script.setAttribute("crossOrigin", "anonymous");
                    script.src = url;
                    doc.getElementsByTagName("head")[0].appendChild(script);
                }
                else if (_loadedScriptEntry.loaded) {
                    callback();
                }
                else {
                    _loadedScriptEntry.pendingCallbacks.push(callback);
                }
            }
        },
        loadCSS: function OSF_OUtil$loadCSS(url) {
            if (url) {
                var doc = window.document;
                var link = doc.createElement("link");
                link.type = "text/css";
                link.rel = "stylesheet";
                link.href = url;
                doc.getElementsByTagName("head")[0].appendChild(link);
            }
        },
        parseEnum: function OSF_OUtil$parseEnum(str, enumObject) {
            var parsed = enumObject[str.trim()];
            if (typeof (parsed) == 'undefined') {
                OsfMsAjaxFactory.msAjaxDebug.trace("invalid enumeration string:" + str);
                throw OsfMsAjaxFactory.msAjaxError.argument("str");
            }
            return parsed;
        },
        delayExecutionAndCache: function OSF_OUtil$delayExecutionAndCache() {
            var obj = { calc: arguments[0] };
            return function () {
                if (obj.calc) {
                    obj.val = obj.calc.apply(this, arguments);
                    delete obj.calc;
                }
                return obj.val;
            };
        },
        getUniqueId: function OSF_OUtil$getUniqueId() {
            _uniqueId = _uniqueId + 1;
            return _uniqueId.toString();
        },
        formatString: function OSF_OUtil$formatString() {
            var args = arguments;
            var source = args[0];
            return source.replace(/{(\d+)}/gm, function (match, number) {
                var index = parseInt(number, 10) + 1;
                return args[index] === undefined ? '{' + number + '}' : args[index];
            });
        },
        generateConversationId: function OSF_OUtil$generateConversationId() {
            return [_random(), _random(), (new Date()).getTime().toString()].join('_');
        },
        getFrameName: function OSF_OUtil$getFrameName(cacheKey) {
            return _xdmSessionKeyPrefix + cacheKey + this.generateConversationId();
        },
        addXdmInfoAsHash: function OSF_OUtil$addXdmInfoAsHash(url, xdmInfoValue) {
            return OSF.OUtil.addInfoAsHash(url, _xdmInfoKey, xdmInfoValue, false);
        },
        addSerializerVersionAsHash: function OSF_OUtil$addSerializerVersionAsHash(url, serializerVersion) {
            return OSF.OUtil.addInfoAsHash(url, _serializerVersionKey, serializerVersion, true);
        },
        addInfoAsHash: function OSF_OUtil$addInfoAsHash(url, keyName, infoValue, encodeInfo) {
            url = url.trim() || '';
            var urlParts = url.split(_fragmentSeparator);
            var urlWithoutFragment = urlParts.shift();
            var fragment = urlParts.join(_fragmentSeparator);
            var newFragment;
            if (encodeInfo) {
                newFragment = [keyName, encodeURIComponent(infoValue), fragment].join('');
            }
            else {
                newFragment = [fragment, keyName, infoValue].join('');
            }
            return [urlWithoutFragment, _fragmentSeparator, newFragment].join('');
        },
        parseHostInfoFromWindowName: function OSF_OUtil$parseHostInfoFromWindowName(skipSessionStorage, windowName) {
            return OSF.OUtil.parseInfoFromWindowName(skipSessionStorage, windowName, OSF.WindowNameItemKeys.HostInfo);
        },
        parseXdmInfo: function OSF_OUtil$parseXdmInfo(skipSessionStorage) {
            var xdmInfoValue = OSF.OUtil.parseXdmInfoWithGivenFragment(skipSessionStorage, window.location.hash);
            if (!xdmInfoValue) {
                xdmInfoValue = OSF.OUtil.parseXdmInfoFromWindowName(skipSessionStorage, window.name);
            }
            return xdmInfoValue;
        },
        parseXdmInfoFromWindowName: function OSF_OUtil$parseXdmInfoFromWindowName(skipSessionStorage, windowName) {
            return OSF.OUtil.parseInfoFromWindowName(skipSessionStorage, windowName, OSF.WindowNameItemKeys.XdmInfo);
        },
        parseXdmInfoWithGivenFragment: function OSF_OUtil$parseXdmInfoWithGivenFragment(skipSessionStorage, fragment) {
            return OSF.OUtil.parseInfoWithGivenFragment(_xdmInfoKey, _xdmSessionKeyPrefix, false, skipSessionStorage, fragment);
        },
        parseSerializerVersion: function OSF_OUtil$parseSerializerVersion(skipSessionStorage) {
            var serializerVersion = OSF.OUtil.parseSerializerVersionWithGivenFragment(skipSessionStorage, window.location.hash);
            if (isNaN(serializerVersion)) {
                serializerVersion = OSF.OUtil.parseSerializerVersionFromWindowName(skipSessionStorage, window.name);
            }
            return serializerVersion;
        },
        parseSerializerVersionFromWindowName: function OSF_OUtil$parseSerializerVersionFromWindowName(skipSessionStorage, windowName) {
            return parseInt(OSF.OUtil.parseInfoFromWindowName(skipSessionStorage, windowName, OSF.WindowNameItemKeys.SerializerVersion));
        },
        parseSerializerVersionWithGivenFragment: function OSF_OUtil$parseSerializerVersionWithGivenFragment(skipSessionStorage, fragment) {
            return parseInt(OSF.OUtil.parseInfoWithGivenFragment(_serializerVersionKey, _serializerVersionKeyPrefix, true, skipSessionStorage, fragment));
        },
        parseInfoFromWindowName: function OSF_OUtil$parseInfoFromWindowName(skipSessionStorage, windowName, infoKey) {
            try {
                var windowNameObj = JSON.parse(windowName);
                var infoValue = windowNameObj != null ? windowNameObj[infoKey] : null;
                var osfSessionStorage = _getSessionStorage();
                if (!skipSessionStorage && osfSessionStorage && windowNameObj != null) {
                    var sessionKey = windowNameObj[OSF.WindowNameItemKeys.BaseFrameName] + infoKey;
                    if (infoValue) {
                        osfSessionStorage.setItem(sessionKey, infoValue);
                    }
                    else {
                        infoValue = osfSessionStorage.getItem(sessionKey);
                    }
                }
                return infoValue;
            }
            catch (Exception) {
                return null;
            }
        },
        parseInfoWithGivenFragment: function OSF_OUtil$parseInfoWithGivenFragment(infoKey, infoKeyPrefix, decodeInfo, skipSessionStorage, fragment) {
            var fragmentParts = fragment.split(infoKey);
            var infoValue = fragmentParts.length > 1 ? fragmentParts[fragmentParts.length - 1] : null;
            if (decodeInfo && infoValue != null) {
                if (infoValue.indexOf(_fragmentInfoDelimiter) >= 0) {
                    infoValue = infoValue.split(_fragmentInfoDelimiter)[0];
                }
                infoValue = decodeURIComponent(infoValue);
            }
            var osfSessionStorage = _getSessionStorage();
            if (!skipSessionStorage && osfSessionStorage) {
                var sessionKeyStart = window.name.indexOf(infoKeyPrefix);
                if (sessionKeyStart > -1) {
                    var sessionKeyEnd = window.name.indexOf(";", sessionKeyStart);
                    if (sessionKeyEnd == -1) {
                        sessionKeyEnd = window.name.length;
                    }
                    var sessionKey = window.name.substring(sessionKeyStart, sessionKeyEnd);
                    if (infoValue) {
                        osfSessionStorage.setItem(sessionKey, infoValue);
                    }
                    else {
                        infoValue = osfSessionStorage.getItem(sessionKey);
                    }
                }
            }
            return infoValue;
        },
        getConversationId: function OSF_OUtil$getConversationId() {
            var searchString = window.location.search;
            var conversationId = null;
            if (searchString) {
                var index = searchString.indexOf("&");
                conversationId = index > 0 ? searchString.substring(1, index) : searchString.substr(1);
                if (conversationId && conversationId.charAt(conversationId.length - 1) === '=') {
                    conversationId = conversationId.substring(0, conversationId.length - 1);
                    if (conversationId) {
                        conversationId = decodeURIComponent(conversationId);
                    }
                }
            }
            return conversationId;
        },
        getInfoItems: function OSF_OUtil$getInfoItems(strInfo) {
            var items = strInfo.split("$");
            if (typeof items[1] == "undefined") {
                items = strInfo.split("|");
            }
            if (typeof items[1] == "undefined") {
                items = strInfo.split("%7C");
            }
            return items;
        },
        getXdmFieldValue: function OSF_OUtil$getXdmFieldValue(xdmFieldName, skipSessionStorage) {
            var fieldValue = '';
            var xdmInfoValue = OSF.OUtil.parseXdmInfo(skipSessionStorage);
            if (xdmInfoValue) {
                var items = OSF.OUtil.getInfoItems(xdmInfoValue);
                if (items != undefined && items.length >= 3) {
                    switch (xdmFieldName) {
                        case OSF.XdmFieldName.ConversationUrl:
                            fieldValue = items[2];
                            break;
                        case OSF.XdmFieldName.AppId:
                            fieldValue = items[1];
                            break;
                    }
                }
            }
            return fieldValue;
        },
        validateParamObject: function OSF_OUtil$validateParamObject(params, expectedProperties, callback) {
            var e = Function._validateParams(arguments, [{ name: "params", type: Object, mayBeNull: false },
                { name: "expectedProperties", type: Object, mayBeNull: false },
                { name: "callback", type: Function, mayBeNull: true }
            ]);
            if (e)
                throw e;
            for (var p in expectedProperties) {
                e = Function._validateParameter(params[p], expectedProperties[p], p);
                if (e)
                    throw e;
            }
        },
        writeProfilerMark: function OSF_OUtil$writeProfilerMark(text) {
            if (window.msWriteProfilerMark) {
                window.msWriteProfilerMark(text);
                OsfMsAjaxFactory.msAjaxDebug.trace(text);
            }
        },
        outputDebug: function OSF_OUtil$outputDebug(text) {
            if (typeof (OsfMsAjaxFactory) !== 'undefined' && OsfMsAjaxFactory.msAjaxDebug && OsfMsAjaxFactory.msAjaxDebug.trace) {
                OsfMsAjaxFactory.msAjaxDebug.trace(text);
            }
        },
        defineNondefaultProperty: function OSF_OUtil$defineNondefaultProperty(obj, prop, descriptor, attributes) {
            descriptor = descriptor || {};
            for (var nd in attributes) {
                var attribute = attributes[nd];
                if (descriptor[attribute] == undefined) {
                    descriptor[attribute] = true;
                }
            }
            Object.defineProperty(obj, prop, descriptor);
            return obj;
        },
        defineNondefaultProperties: function OSF_OUtil$defineNondefaultProperties(obj, descriptors, attributes) {
            descriptors = descriptors || {};
            for (var prop in descriptors) {
                OSF.OUtil.defineNondefaultProperty(obj, prop, descriptors[prop], attributes);
            }
            return obj;
        },
        defineEnumerableProperty: function OSF_OUtil$defineEnumerableProperty(obj, prop, descriptor) {
            return OSF.OUtil.defineNondefaultProperty(obj, prop, descriptor, ["enumerable"]);
        },
        defineEnumerableProperties: function OSF_OUtil$defineEnumerableProperties(obj, descriptors) {
            return OSF.OUtil.defineNondefaultProperties(obj, descriptors, ["enumerable"]);
        },
        defineMutableProperty: function OSF_OUtil$defineMutableProperty(obj, prop, descriptor) {
            return OSF.OUtil.defineNondefaultProperty(obj, prop, descriptor, ["writable", "enumerable", "configurable"]);
        },
        defineMutableProperties: function OSF_OUtil$defineMutableProperties(obj, descriptors) {
            return OSF.OUtil.defineNondefaultProperties(obj, descriptors, ["writable", "enumerable", "configurable"]);
        },
        finalizeProperties: function OSF_OUtil$finalizeProperties(obj, descriptor) {
            descriptor = descriptor || {};
            var props = Object.getOwnPropertyNames(obj);
            var propsLength = props.length;
            for (var i = 0; i < propsLength; i++) {
                var prop = props[i];
                var desc = Object.getOwnPropertyDescriptor(obj, prop);
                if (!desc.get && !desc.set) {
                    desc.writable = descriptor.writable || false;
                }
                desc.configurable = descriptor.configurable || false;
                desc.enumerable = descriptor.enumerable || true;
                Object.defineProperty(obj, prop, desc);
            }
            return obj;
        },
        mapList: function OSF_OUtil$MapList(list, mapFunction) {
            var ret = [];
            if (list) {
                for (var item in list) {
                    ret.push(mapFunction(list[item]));
                }
            }
            return ret;
        },
        listContainsKey: function OSF_OUtil$listContainsKey(list, key) {
            for (var item in list) {
                if (key == item) {
                    return true;
                }
            }
            return false;
        },
        listContainsValue: function OSF_OUtil$listContainsElement(list, value) {
            for (var item in list) {
                if (value == list[item]) {
                    return true;
                }
            }
            return false;
        },
        augmentList: function OSF_OUtil$augmentList(list, addenda) {
            var add = list.push ? function (key, value) { list.push(value); } : function (key, value) { list[key] = value; };
            for (var key in addenda) {
                add(key, addenda[key]);
            }
        },
        redefineList: function OSF_Outil$redefineList(oldList, newList) {
            for (var key1 in oldList) {
                delete oldList[key1];
            }
            for (var key2 in newList) {
                oldList[key2] = newList[key2];
            }
        },
        isArray: function OSF_OUtil$isArray(obj) {
            return Object.prototype.toString.apply(obj) === "[object Array]";
        },
        isFunction: function OSF_OUtil$isFunction(obj) {
            return Object.prototype.toString.apply(obj) === "[object Function]";
        },
        isDate: function OSF_OUtil$isDate(obj) {
            return Object.prototype.toString.apply(obj) === "[object Date]";
        },
        addEventListener: function OSF_OUtil$addEventListener(element, eventName, listener) {
            if (element.addEventListener) {
                element.addEventListener(eventName, listener, false);
            }
            else if ((Sys.Browser.agent === Sys.Browser.InternetExplorer) && element.attachEvent) {
                element.attachEvent("on" + eventName, listener);
            }
            else {
                element["on" + eventName] = listener;
            }
        },
        removeEventListener: function OSF_OUtil$removeEventListener(element, eventName, listener) {
            if (element.removeEventListener) {
                element.removeEventListener(eventName, listener, false);
            }
            else if ((Sys.Browser.agent === Sys.Browser.InternetExplorer) && element.detachEvent) {
                element.detachEvent("on" + eventName, listener);
            }
            else {
                element["on" + eventName] = null;
            }
        },
        getCookieValue: function OSF_OUtil$getCookieValue(cookieName) {
            var tmpCookieString = RegExp(cookieName + "[^;]+").exec(document.cookie);
            return tmpCookieString.toString().replace(/^[^=]+./, "");
        },
        xhrGet: function OSF_OUtil$xhrGet(url, onSuccess, onError) {
            var xmlhttp;
            try {
                xmlhttp = new XMLHttpRequest();
                xmlhttp.onreadystatechange = function () {
                    if (xmlhttp.readyState == 4) {
                        if (xmlhttp.status == 200) {
                            onSuccess(xmlhttp.responseText);
                        }
                        else {
                            onError(xmlhttp.status);
                        }
                    }
                };
                xmlhttp.open("GET", url, true);
                xmlhttp.send();
            }
            catch (ex) {
                onError(ex);
            }
        },
        xhrGetFull: function OSF_OUtil$xhrGetFull(url, oneDriveFileName, onSuccess, onError) {
            var xmlhttp;
            var requestedFileName = oneDriveFileName;
            try {
                xmlhttp = new XMLHttpRequest();
                xmlhttp.onreadystatechange = function () {
                    if (xmlhttp.readyState == 4) {
                        if (xmlhttp.status == 200) {
                            onSuccess(xmlhttp, requestedFileName);
                        }
                        else {
                            onError(xmlhttp.status);
                        }
                    }
                };
                xmlhttp.open("GET", url, true);
                xmlhttp.send();
            }
            catch (ex) {
                onError(ex);
            }
        },
        encodeBase64: function OSF_Outil$encodeBase64(input) {
            if (!input)
                return input;
            var codex = "ABCDEFGHIJKLMNOP" + "QRSTUVWXYZabcdef" + "ghijklmnopqrstuv" + "wxyz0123456789+/=";
            var output = [];
            var temp = [];
            var index = 0;
            var c1, c2, c3, a, b, c;
            var i;
            var length = input.length;
            do {
                c1 = input.charCodeAt(index++);
                c2 = input.charCodeAt(index++);
                c3 = input.charCodeAt(index++);
                i = 0;
                a = c1 & 255;
                b = c1 >> 8;
                c = c2 & 255;
                temp[i++] = a >> 2;
                temp[i++] = ((a & 3) << 4) | (b >> 4);
                temp[i++] = ((b & 15) << 2) | (c >> 6);
                temp[i++] = c & 63;
                if (!isNaN(c2)) {
                    a = c2 >> 8;
                    b = c3 & 255;
                    c = c3 >> 8;
                    temp[i++] = a >> 2;
                    temp[i++] = ((a & 3) << 4) | (b >> 4);
                    temp[i++] = ((b & 15) << 2) | (c >> 6);
                    temp[i++] = c & 63;
                }
                if (isNaN(c2)) {
                    temp[i - 1] = 64;
                }
                else if (isNaN(c3)) {
                    temp[i - 2] = 64;
                    temp[i - 1] = 64;
                }
                for (var t = 0; t < i; t++) {
                    output.push(codex.charAt(temp[t]));
                }
            } while (index < length);
            return output.join("");
        },
        getSessionStorage: function OSF_Outil$getSessionStorage() {
            return _getSessionStorage();
        },
        getLocalStorage: function OSF_Outil$getLocalStorage() {
            if (!_safeLocalStorage) {
                try {
                    var localStorage = window.localStorage;
                }
                catch (ex) {
                    localStorage = null;
                }
                _safeLocalStorage = new OfficeExt.SafeStorage(localStorage);
            }
            return _safeLocalStorage;
        },
        convertIntToCssHexColor: function OSF_Outil$convertIntToCssHexColor(val) {
            var hex = "#" + (Number(val) + 0x1000000).toString(16).slice(-6);
            return hex;
        },
        attachClickHandler: function OSF_Outil$attachClickHandler(element, handler) {
            element.onclick = function (e) {
                handler();
            };
            element.ontouchend = function (e) {
                handler();
                e.preventDefault();
            };
        },
        getQueryStringParamValue: function OSF_Outil$getQueryStringParamValue(queryString, paramName) {
            var e = Function._validateParams(arguments, [{ name: "queryString", type: String, mayBeNull: false },
                { name: "paramName", type: String, mayBeNull: false }
            ]);
            if (e) {
                OsfMsAjaxFactory.msAjaxDebug.trace("OSF_Outil_getQueryStringParamValue: Parameters cannot be null.");
                return "";
            }
            var queryExp = new RegExp("[\\?&]" + paramName + "=([^&#]*)", "i");
            if (!queryExp.test(queryString)) {
                OsfMsAjaxFactory.msAjaxDebug.trace("OSF_Outil_getQueryStringParamValue: The parameter is not found.");
                return "";
            }
            return queryExp.exec(queryString)[1];
        },
        isiOS: function OSF_Outil$isiOS() {
            return (window.navigator.userAgent.match(/(iPad|iPhone|iPod)/g) ? true : false);
        },
        isChrome: function OSF_Outil$isChrome() {
            return (window.navigator.userAgent.indexOf("Chrome") > 0) && !OSF.OUtil.isEdge();
        },
        isEdge: function OSF_Outil$isEdge() {
            return window.navigator.userAgent.indexOf("Edge") > 0;
        },
        isIE: function OSF_Outil$isIE() {
            return window.navigator.userAgent.indexOf("Trident") > 0;
        },
        isFirefox: function OSF_Outil$isFirefox() {
            return window.navigator.userAgent.indexOf("Firefox") > 0;
        },
        shallowCopy: function OSF_Outil$shallowCopy(sourceObj) {
            if (sourceObj == null) {
                return null;
            }
            else if (!(sourceObj instanceof Object)) {
                return sourceObj;
            }
            else if (Array.isArray(sourceObj)) {
                var copyArr = [];
                for (var i = 0; i < sourceObj.length; i++) {
                    copyArr.push(sourceObj[i]);
                }
                return copyArr;
            }
            else {
                var copyObj = sourceObj.constructor();
                for (var property in sourceObj) {
                    if (sourceObj.hasOwnProperty(property)) {
                        copyObj[property] = sourceObj[property];
                    }
                }
                return copyObj;
            }
        },
        createObject: function OSF_Outil$createObject(properties) {
            var obj = null;
            if (properties) {
                obj = {};
                var len = properties.length;
                for (var i = 0; i < len; i++) {
                    obj[properties[i].name] = properties[i].value;
                }
            }
            return obj;
        },
        addClass: function OSF_OUtil$addClass(elmt, val) {
            if (!OSF.OUtil.hasClass(elmt, val)) {
                var className = elmt.getAttribute(_classN);
                if (className) {
                    elmt.setAttribute(_classN, className + " " + val);
                }
                else {
                    elmt.setAttribute(_classN, val);
                }
            }
        },
        removeClass: function OSF_OUtil$removeClass(elmt, val) {
            if (OSF.OUtil.hasClass(elmt, val)) {
                var className = elmt.getAttribute(_classN);
                var reg = new RegExp('(\\s|^)' + val + '(\\s|$)');
                className = className.replace(reg, '');
                elmt.setAttribute(_classN, className);
            }
        },
        hasClass: function OSF_OUtil$hasClass(elmt, clsName) {
            var className = elmt.getAttribute(_classN);
            return className && className.match(new RegExp('(\\s|^)' + clsName + '(\\s|$)'));
        },
        focusToFirstTabbable: function OSF_OUtil$focusToFirstTabbable(all, backward) {
            var next;
            var focused = false;
            var candidate;
            var setFlag = function (e) {
                focused = true;
            };
            var findNextPos = function (allLen, currPos, backward) {
                if (currPos < 0 || currPos > allLen) {
                    return -1;
                }
                else if (currPos === 0 && backward) {
                    return -1;
                }
                else if (currPos === allLen - 1 && !backward) {
                    return -1;
                }
                if (backward) {
                    return currPos - 1;
                }
                else {
                    return currPos + 1;
                }
            };
            all = _reOrderTabbableElements(all);
            next = backward ? all.length - 1 : 0;
            if (all.length === 0) {
                return null;
            }
            while (!focused && next >= 0 && next < all.length) {
                candidate = all[next];
                window.focus();
                candidate.addEventListener('focus', setFlag);
                candidate.focus();
                candidate.removeEventListener('focus', setFlag);
                next = findNextPos(all.length, next, backward);
                if (!focused && candidate === document.activeElement) {
                    focused = true;
                }
            }
            if (focused) {
                return candidate;
            }
            else {
                return null;
            }
        },
        focusToNextTabbable: function OSF_OUtil$focusToNextTabbable(all, curr, shift) {
            var currPos;
            var next;
            var focused = false;
            var candidate;
            var setFlag = function (e) {
                focused = true;
            };
            var findCurrPos = function (all, curr) {
                var i = 0;
                for (; i < all.length; i++) {
                    if (all[i] === curr) {
                        return i;
                    }
                }
                return -1;
            };
            var findNextPos = function (allLen, currPos, shift) {
                if (currPos < 0 || currPos > allLen) {
                    return -1;
                }
                else if (currPos === 0 && shift) {
                    return -1;
                }
                else if (currPos === allLen - 1 && !shift) {
                    return -1;
                }
                if (shift) {
                    return currPos - 1;
                }
                else {
                    return currPos + 1;
                }
            };
            all = _reOrderTabbableElements(all);
            currPos = findCurrPos(all, curr);
            next = findNextPos(all.length, currPos, shift);
            if (next < 0) {
                return null;
            }
            while (!focused && next >= 0 && next < all.length) {
                candidate = all[next];
                candidate.addEventListener('focus', setFlag);
                candidate.focus();
                candidate.removeEventListener('focus', setFlag);
                next = findNextPos(all.length, next, shift);
                if (!focused && candidate === document.activeElement) {
                    focused = true;
                }
            }
            if (focused) {
                return candidate;
            }
            else {
                return null;
            }
        }
    };
})();
OSF.OUtil.Guid = (function () {
    var hexCode = ["0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "a", "b", "c", "d", "e", "f"];
    return {
        generateNewGuid: function OSF_Outil_Guid$generateNewGuid() {
            var result = "";
            var tick = (new Date()).getTime();
            var index = 0;
            for (; index < 32 && tick > 0; index++) {
                if (index == 8 || index == 12 || index == 16 || index == 20) {
                    result += "-";
                }
                result += hexCode[tick % 16];
                tick = Math.floor(tick / 16);
            }
            for (; index < 32; index++) {
                if (index == 8 || index == 12 || index == 16 || index == 20) {
                    result += "-";
                }
                result += hexCode[Math.floor(Math.random() * 16)];
            }
            return result;
        }
    };
})();
window.OSF = OSF;
OSF.OUtil.setNamespace("OSF", window);
OSF.MessageIDs = {
    "FetchBundleUrl": 0,
    "LoadReactBundle": 1,
    "LoadBundleSuccess": 2,
    "LoadBundleError": 3
};
OSF.AppName = {
    Unsupported: 0,
    Excel: 1,
    Word: 2,
    PowerPoint: 4,
    Outlook: 8,
    ExcelWebApp: 16,
    WordWebApp: 32,
    OutlookWebApp: 64,
    Project: 128,
    AccessWebApp: 256,
    PowerpointWebApp: 512,
    ExcelIOS: 1024,
    Sway: 2048,
    WordIOS: 4096,
    PowerPointIOS: 8192,
    Access: 16384,
    Lync: 32768,
    OutlookIOS: 65536,
    OneNoteWebApp: 131072,
    OneNote: 262144,
    ExcelWinRT: 524288,
    WordWinRT: 1048576,
    PowerpointWinRT: 2097152,
    OutlookAndroid: 4194304,
    OneNoteWinRT: 8388608,
    ExcelAndroid: 8388609,
    VisioWebApp: 8388610,
    OneNoteIOS: 8388611,
    WordAndroid: 8388613,
    PowerpointAndroid: 8388614
};
OSF.InternalPerfMarker = {
    DataCoercionBegin: "Agave.HostCall.CoerceDataStart",
    DataCoercionEnd: "Agave.HostCall.CoerceDataEnd"
};
OSF.HostCallPerfMarker = {
    IssueCall: "Agave.HostCall.IssueCall",
    ReceiveResponse: "Agave.HostCall.ReceiveResponse",
    RuntimeExceptionRaised: "Agave.HostCall.RuntimeExecptionRaised"
};
OSF.AgaveHostAction = {
    "Select": 0,
    "UnSelect": 1,
    "CancelDialog": 2,
    "InsertAgave": 3,
    "CtrlF6In": 4,
    "CtrlF6Exit": 5,
    "CtrlF6ExitShift": 6,
    "SelectWithError": 7,
    "NotifyHostError": 8,
    "RefreshAddinCommands": 9,
    "PageIsReady": 10,
    "TabIn": 11,
    "TabInShift": 12,
    "TabExit": 13,
    "TabExitShift": 14,
    "EscExit": 15,
    "F2Exit": 16,
    "ExitNoFocusable": 17,
    "ExitNoFocusableShift": 18,
    "MouseEnter": 19,
    "MouseLeave": 20,
    "UpdateTargetUrl": 21,
    "InstallCustomFunctions": 22,
    "SendTelemetryEvent": 23
};
OSF.SharedConstants = {
    "NotificationConversationIdSuffix": '_ntf'
};
OSF.DialogMessageType = {
    DialogMessageReceived: 0,
    DialogParentMessageReceived: 1,
    DialogClosed: 12006
};
OSF.OfficeAppContext = function OSF_OfficeAppContext(id, appName, appVersion, appUILocale, dataLocale, docUrl, clientMode, settings, reason, osfControlType, eToken, correlationId, appInstanceId, touchEnabled, commerceAllowed, appMinorVersion, requirementMatrix, hostCustomMessage, hostFullVersion, clientWindowHeight, clientWindowWidth, addinName, appDomains, dialogRequirementMatrix) {
    this._id = id;
    this._appName = appName;
    this._appVersion = appVersion;
    this._appUILocale = appUILocale;
    this._dataLocale = dataLocale;
    this._docUrl = docUrl;
    this._clientMode = clientMode;
    this._settings = settings;
    this._reason = reason;
    this._osfControlType = osfControlType;
    this._eToken = eToken;
    this._correlationId = correlationId;
    this._appInstanceId = appInstanceId;
    this._touchEnabled = touchEnabled;
    this._commerceAllowed = commerceAllowed;
    this._appMinorVersion = appMinorVersion;
    this._requirementMatrix = requirementMatrix;
    this._hostCustomMessage = hostCustomMessage;
    this._hostFullVersion = hostFullVersion;
    this._isDialog = false;
    this._clientWindowHeight = clientWindowHeight;
    this._clientWindowWidth = clientWindowWidth;
    this._addinName = addinName;
    this._appDomains = appDomains;
    this._dialogRequirementMatrix = dialogRequirementMatrix;
    this.get_id = function get_id() { return this._id; };
    this.get_appName = function get_appName() { return this._appName; };
    this.get_appVersion = function get_appVersion() { return this._appVersion; };
    this.get_appUILocale = function get_appUILocale() { return this._appUILocale; };
    this.get_dataLocale = function get_dataLocale() { return this._dataLocale; };
    this.get_docUrl = function get_docUrl() { return this._docUrl; };
    this.get_clientMode = function get_clientMode() { return this._clientMode; };
    this.get_bindings = function get_bindings() { return this._bindings; };
    this.get_settings = function get_settings() { return this._settings; };
    this.get_reason = function get_reason() { return this._reason; };
    this.get_osfControlType = function get_osfControlType() { return this._osfControlType; };
    this.get_eToken = function get_eToken() { return this._eToken; };
    this.get_correlationId = function get_correlationId() { return this._correlationId; };
    this.get_appInstanceId = function get_appInstanceId() { return this._appInstanceId; };
    this.get_touchEnabled = function get_touchEnabled() { return this._touchEnabled; };
    this.get_commerceAllowed = function get_commerceAllowed() { return this._commerceAllowed; };
    this.get_appMinorVersion = function get_appMinorVersion() { return this._appMinorVersion; };
    this.get_requirementMatrix = function get_requirementMatrix() { return this._requirementMatrix; };
    this.get_dialogRequirementMatrix = function get_dialogRequirementMatrix() { return this._dialogRequirementMatrix; };
    this.get_hostCustomMessage = function get_hostCustomMessage() { return this._hostCustomMessage; };
    this.get_hostFullVersion = function get_hostFullVersion() { return this._hostFullVersion; };
    this.get_isDialog = function get_isDialog() { return this._isDialog; };
    this.get_clientWindowHeight = function get_clientWindowHeight() { return this._clientWindowHeight; };
    this.get_clientWindowWidth = function get_clientWindowWidth() { return this._clientWindowWidth; };
    this.get_addinName = function get_addinName() { return this._addinName; };
    this.get_appDomains = function get_appDomains() { return this._appDomains; };
};
OSF.OsfControlType = {
    DocumentLevel: 0,
    ContainerLevel: 1
};
OSF.ClientMode = {
    ReadOnly: 0,
    ReadWrite: 1
};
OSF.OUtil.setNamespace("Microsoft", window);
OSF.OUtil.setNamespace("Office", Microsoft);
OSF.OUtil.setNamespace("Client", Microsoft.Office);
OSF.OUtil.setNamespace("WebExtension", Microsoft.Office);
Microsoft.Office.WebExtension.InitializationReason = {
    Inserted: "inserted",
    DocumentOpened: "documentOpened"
};
Microsoft.Office.WebExtension.ValueFormat = {
    Unformatted: "unformatted",
    Formatted: "formatted"
};
Microsoft.Office.WebExtension.FilterType = {
    All: "all"
};
Microsoft.Office.WebExtension.Parameters = {
    BindingType: "bindingType",
    CoercionType: "coercionType",
    ValueFormat: "valueFormat",
    FilterType: "filterType",
    Columns: "columns",
    SampleData: "sampleData",
    GoToType: "goToType",
    SelectionMode: "selectionMode",
    Id: "id",
    PromptText: "promptText",
    ItemName: "itemName",
    FailOnCollision: "failOnCollision",
    StartRow: "startRow",
    StartColumn: "startColumn",
    RowCount: "rowCount",
    ColumnCount: "columnCount",
    Callback: "callback",
    AsyncContext: "asyncContext",
    Data: "data",
    Rows: "rows",
    OverwriteIfStale: "overwriteIfStale",
    FileType: "fileType",
    EventType: "eventType",
    Handler: "handler",
    SliceSize: "sliceSize",
    SliceIndex: "sliceIndex",
    ActiveView: "activeView",
    Status: "status",
    PlatformType: "platformType",
    HostType: "hostType",
    ForceConsent: "forceConsent",
    ForceAddAccount: "forceAddAccount",
    AuthChallenge: "authChallenge",
    Reserved: "reserved",
    Tcid: "tcid",
    Xml: "xml",
    Namespace: "namespace",
    Prefix: "prefix",
    XPath: "xPath",
    Text: "text",
    ImageLeft: "imageLeft",
    ImageTop: "imageTop",
    ImageWidth: "imageWidth",
    ImageHeight: "imageHeight",
    TaskId: "taskId",
    FieldId: "fieldId",
    FieldValue: "fieldValue",
    ServerUrl: "serverUrl",
    ListName: "listName",
    ResourceId: "resourceId",
    ViewType: "viewType",
    ViewName: "viewName",
    GetRawValue: "getRawValue",
    CellFormat: "cellFormat",
    TableOptions: "tableOptions",
    TaskIndex: "taskIndex",
    ResourceIndex: "resourceIndex",
    CustomFieldId: "customFieldId",
    Url: "url",
    MessageHandler: "messageHandler",
    Width: "width",
    Height: "height",
    RequireHTTPs: "requireHTTPS",
    MessageToParent: "messageToParent",
    DisplayInIframe: "displayInIframe",
    MessageContent: "messageContent",
    HideTitle: "hideTitle",
    UseDeviceIndependentPixels: "useDeviceIndependentPixels",
    PromptBeforeOpen: "promptBeforeOpen",
    AppCommandInvocationCompletedData: "appCommandInvocationCompletedData",
    Base64: "base64",
    FormId: "formId"
};
OSF.OUtil.setNamespace("DDA", OSF);
OSF.DDA.DocumentMode = {
    ReadOnly: 1,
    ReadWrite: 0
};
OSF.DDA.PropertyDescriptors = {
    AsyncResultStatus: "AsyncResultStatus"
};
OSF.DDA.EventDescriptors = {};
OSF.DDA.ListDescriptors = {};
OSF.DDA.UI = {};
OSF.DDA.getXdmEventName = function OSF_DDA$GetXdmEventName(id, eventType) {
    if (eventType == Microsoft.Office.WebExtension.EventType.BindingSelectionChanged ||
        eventType == Microsoft.Office.WebExtension.EventType.BindingDataChanged ||
        eventType == Microsoft.Office.WebExtension.EventType.DataNodeDeleted ||
        eventType == Microsoft.Office.WebExtension.EventType.DataNodeInserted ||
        eventType == Microsoft.Office.WebExtension.EventType.DataNodeReplaced) {
        return id + "_" + eventType;
    }
    else {
        return eventType;
    }
};
OSF.DDA.MethodDispId = {
    dispidMethodMin: 64,
    dispidGetSelectedDataMethod: 64,
    dispidSetSelectedDataMethod: 65,
    dispidAddBindingFromSelectionMethod: 66,
    dispidAddBindingFromPromptMethod: 67,
    dispidGetBindingMethod: 68,
    dispidReleaseBindingMethod: 69,
    dispidGetBindingDataMethod: 70,
    dispidSetBindingDataMethod: 71,
    dispidAddRowsMethod: 72,
    dispidClearAllRowsMethod: 73,
    dispidGetAllBindingsMethod: 74,
    dispidLoadSettingsMethod: 75,
    dispidSaveSettingsMethod: 76,
    dispidGetDocumentCopyMethod: 77,
    dispidAddBindingFromNamedItemMethod: 78,
    dispidAddColumnsMethod: 79,
    dispidGetDocumentCopyChunkMethod: 80,
    dispidReleaseDocumentCopyMethod: 81,
    dispidNavigateToMethod: 82,
    dispidGetActiveViewMethod: 83,
    dispidGetDocumentThemeMethod: 84,
    dispidGetOfficeThemeMethod: 85,
    dispidGetFilePropertiesMethod: 86,
    dispidClearFormatsMethod: 87,
    dispidSetTableOptionsMethod: 88,
    dispidSetFormatsMethod: 89,
    dispidExecuteRichApiRequestMethod: 93,
    dispidAppCommandInvocationCompletedMethod: 94,
    dispidCloseContainerMethod: 97,
    dispidGetAccessTokenMethod: 98,
    dispidOpenBrowserWindow: 102,
    dispidCreateDocumentMethod: 105,
    dispidInsertFormMethod: 106,
    dispidDisplayRibbonCalloutAsyncMethod: 109,
    dispidGetSelectedTaskMethod: 110,
    dispidGetSelectedResourceMethod: 111,
    dispidGetTaskMethod: 112,
    dispidGetResourceFieldMethod: 113,
    dispidGetWSSUrlMethod: 114,
    dispidGetTaskFieldMethod: 115,
    dispidGetProjectFieldMethod: 116,
    dispidGetSelectedViewMethod: 117,
    dispidGetTaskByIndexMethod: 118,
    dispidGetResourceByIndexMethod: 119,
    dispidSetTaskFieldMethod: 120,
    dispidSetResourceFieldMethod: 121,
    dispidGetMaxTaskIndexMethod: 122,
    dispidGetMaxResourceIndexMethod: 123,
    dispidCreateTaskMethod: 124,
    dispidAddDataPartMethod: 128,
    dispidGetDataPartByIdMethod: 129,
    dispidGetDataPartsByNamespaceMethod: 130,
    dispidGetDataPartXmlMethod: 131,
    dispidGetDataPartNodesMethod: 132,
    dispidDeleteDataPartMethod: 133,
    dispidGetDataNodeValueMethod: 134,
    dispidGetDataNodeXmlMethod: 135,
    dispidGetDataNodesMethod: 136,
    dispidSetDataNodeValueMethod: 137,
    dispidSetDataNodeXmlMethod: 138,
    dispidAddDataNamespaceMethod: 139,
    dispidGetDataUriByPrefixMethod: 140,
    dispidGetDataPrefixByUriMethod: 141,
    dispidGetDataNodeTextMethod: 142,
    dispidSetDataNodeTextMethod: 143,
    dispidMessageParentMethod: 144,
    dispidSendMessageMethod: 145,
    dispidExecuteFeature: 146,
    dispidQueryFeature: 147,
    dispidMethodMax: 147
};
OSF.DDA.EventDispId = {
    dispidEventMin: 0,
    dispidInitializeEvent: 0,
    dispidSettingsChangedEvent: 1,
    dispidDocumentSelectionChangedEvent: 2,
    dispidBindingSelectionChangedEvent: 3,
    dispidBindingDataChangedEvent: 4,
    dispidDocumentOpenEvent: 5,
    dispidDocumentCloseEvent: 6,
    dispidActiveViewChangedEvent: 7,
    dispidDocumentThemeChangedEvent: 8,
    dispidOfficeThemeChangedEvent: 9,
    dispidDialogMessageReceivedEvent: 10,
    dispidDialogNotificationShownInAddinEvent: 11,
    dispidDialogParentMessageReceivedEvent: 12,
    dispidObjectDeletedEvent: 13,
    dispidObjectSelectionChangedEvent: 14,
    dispidObjectDataChangedEvent: 15,
    dispidContentControlAddedEvent: 16,
    dispidActivationStatusChangedEvent: 32,
    dispidRichApiMessageEvent: 33,
    dispidAppCommandInvokedEvent: 39,
    dispidOlkItemSelectedChangedEvent: 46,
    dispidOlkRecipientsChangedEvent: 47,
    dispidOlkAppointmentTimeChangedEvent: 48,
    dispidOlkRecurrenceChangedEvent: 49,
    dispidTaskSelectionChangedEvent: 56,
    dispidResourceSelectionChangedEvent: 57,
    dispidViewSelectionChangedEvent: 58,
    dispidDataNodeAddedEvent: 60,
    dispidDataNodeReplacedEvent: 61,
    dispidDataNodeDeletedEvent: 62,
    dispidEventMax: 63
};
OSF.DDA.ErrorCodeManager = (function () {
    var _errorMappings = {};
    return {
        getErrorArgs: function OSF_DDA_ErrorCodeManager$getErrorArgs(errorCode) {
            var errorArgs = _errorMappings[errorCode];
            if (!errorArgs) {
                errorArgs = _errorMappings[this.errorCodes.ooeInternalError];
            }
            else {
                if (!errorArgs.name) {
                    errorArgs.name = _errorMappings[this.errorCodes.ooeInternalError].name;
                }
                if (!errorArgs.message) {
                    errorArgs.message = _errorMappings[this.errorCodes.ooeInternalError].message;
                }
            }
            return errorArgs;
        },
        addErrorMessage: function OSF_DDA_ErrorCodeManager$addErrorMessage(errorCode, errorNameMessage) {
            _errorMappings[errorCode] = errorNameMessage;
        },
        errorCodes: {
            ooeSuccess: 0,
            ooeChunkResult: 1,
            ooeCoercionTypeNotSupported: 1000,
            ooeGetSelectionNotMatchDataType: 1001,
            ooeCoercionTypeNotMatchBinding: 1002,
            ooeInvalidGetRowColumnCounts: 1003,
            ooeSelectionNotSupportCoercionType: 1004,
            ooeInvalidGetStartRowColumn: 1005,
            ooeNonUniformPartialGetNotSupported: 1006,
            ooeGetDataIsTooLarge: 1008,
            ooeFileTypeNotSupported: 1009,
            ooeGetDataParametersConflict: 1010,
            ooeInvalidGetColumns: 1011,
            ooeInvalidGetRows: 1012,
            ooeInvalidReadForBlankRow: 1013,
            ooeUnsupportedDataObject: 2000,
            ooeCannotWriteToSelection: 2001,
            ooeDataNotMatchSelection: 2002,
            ooeOverwriteWorksheetData: 2003,
            ooeDataNotMatchBindingSize: 2004,
            ooeInvalidSetStartRowColumn: 2005,
            ooeInvalidDataFormat: 2006,
            ooeDataNotMatchCoercionType: 2007,
            ooeDataNotMatchBindingType: 2008,
            ooeSetDataIsTooLarge: 2009,
            ooeNonUniformPartialSetNotSupported: 2010,
            ooeInvalidSetColumns: 2011,
            ooeInvalidSetRows: 2012,
            ooeSetDataParametersConflict: 2013,
            ooeCellDataAmountBeyondLimits: 2014,
            ooeSelectionCannotBound: 3000,
            ooeBindingNotExist: 3002,
            ooeBindingToMultipleSelection: 3003,
            ooeInvalidSelectionForBindingType: 3004,
            ooeOperationNotSupportedOnThisBindingType: 3005,
            ooeNamedItemNotFound: 3006,
            ooeMultipleNamedItemFound: 3007,
            ooeInvalidNamedItemForBindingType: 3008,
            ooeUnknownBindingType: 3009,
            ooeOperationNotSupportedOnMatrixData: 3010,
            ooeInvalidColumnsForBinding: 3011,
            ooeSettingNameNotExist: 4000,
            ooeSettingsCannotSave: 4001,
            ooeSettingsAreStale: 4002,
            ooeOperationNotSupported: 5000,
            ooeInternalError: 5001,
            ooeDocumentReadOnly: 5002,
            ooeEventHandlerNotExist: 5003,
            ooeInvalidApiCallInContext: 5004,
            ooeShuttingDown: 5005,
            ooeUnsupportedEnumeration: 5007,
            ooeIndexOutOfRange: 5008,
            ooeBrowserAPINotSupported: 5009,
            ooeInvalidParam: 5010,
            ooeRequestTimeout: 5011,
            ooeInvalidOrTimedOutSession: 5012,
            ooeInvalidApiArguments: 5013,
            ooeOperationCancelled: 5014,
            ooeTooManyIncompleteRequests: 5100,
            ooeRequestTokenUnavailable: 5101,
            ooeActivityLimitReached: 5102,
            ooeCustomXmlNodeNotFound: 6000,
            ooeCustomXmlError: 6100,
            ooeCustomXmlExceedQuota: 6101,
            ooeCustomXmlOutOfDate: 6102,
            ooeNoCapability: 7000,
            ooeCannotNavTo: 7001,
            ooeSpecifiedIdNotExist: 7002,
            ooeNavOutOfBound: 7004,
            ooeElementMissing: 8000,
            ooeProtectedError: 8001,
            ooeInvalidCellsValue: 8010,
            ooeInvalidTableOptionValue: 8011,
            ooeInvalidFormatValue: 8012,
            ooeRowIndexOutOfRange: 8020,
            ooeColIndexOutOfRange: 8021,
            ooeFormatValueOutOfRange: 8022,
            ooeCellFormatAmountBeyondLimits: 8023,
            ooeMemoryFileLimit: 11000,
            ooeNetworkProblemRetrieveFile: 11001,
            ooeInvalidSliceSize: 11002,
            ooeInvalidCallback: 11101,
            ooeInvalidWidth: 12000,
            ooeInvalidHeight: 12001,
            ooeNavigationError: 12002,
            ooeInvalidScheme: 12003,
            ooeAppDomains: 12004,
            ooeRequireHTTPS: 12005,
            ooeWebDialogClosed: 12006,
            ooeDialogAlreadyOpened: 12007,
            ooeEndUserAllow: 12008,
            ooeEndUserIgnore: 12009,
            ooeNotUILessDialog: 12010,
            ooeCrossZone: 12011,
            ooeNotSSOAgave: 13000,
            ooeSSOUserNotSignedIn: 13001,
            ooeSSOUserAborted: 13002,
            ooeSSOUnsupportedUserIdentity: 13003,
            ooeSSOInvalidResourceUrl: 13004,
            ooeSSOInvalidGrant: 13005,
            ooeSSOClientError: 13006,
            ooeSSOServerError: 13007,
            ooeAddinIsAlreadyRequestingToken: 13008,
            ooeSSOUserConsentNotSupportedByCurrentAddinCategory: 13009,
            ooeSSOConnectionLost: 13010,
            ooeResourceNotAllowed: 13011,
            ooeAccessDenied: 13990,
            ooeGeneralException: 13991
        },
        initializeErrorMessages: function OSF_DDA_ErrorCodeManager$initializeErrorMessages(stringNS) {
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeCoercionTypeNotSupported] = { name: stringNS.L_InvalidCoercion, message: stringNS.L_CoercionTypeNotSupported };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeGetSelectionNotMatchDataType] = { name: stringNS.L_DataReadError, message: stringNS.L_GetSelectionNotSupported };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeCoercionTypeNotMatchBinding] = { name: stringNS.L_InvalidCoercion, message: stringNS.L_CoercionTypeNotMatchBinding };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidGetRowColumnCounts] = { name: stringNS.L_DataReadError, message: stringNS.L_InvalidGetRowColumnCounts };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeSelectionNotSupportCoercionType] = { name: stringNS.L_DataReadError, message: stringNS.L_SelectionNotSupportCoercionType };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidGetStartRowColumn] = { name: stringNS.L_DataReadError, message: stringNS.L_InvalidGetStartRowColumn };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeNonUniformPartialGetNotSupported] = { name: stringNS.L_DataReadError, message: stringNS.L_NonUniformPartialGetNotSupported };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeGetDataIsTooLarge] = { name: stringNS.L_DataReadError, message: stringNS.L_GetDataIsTooLarge };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeFileTypeNotSupported] = { name: stringNS.L_DataReadError, message: stringNS.L_FileTypeNotSupported };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeGetDataParametersConflict] = { name: stringNS.L_DataReadError, message: stringNS.L_GetDataParametersConflict };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidGetColumns] = { name: stringNS.L_DataReadError, message: stringNS.L_InvalidGetColumns };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidGetRows] = { name: stringNS.L_DataReadError, message: stringNS.L_InvalidGetRows };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidReadForBlankRow] = { name: stringNS.L_DataReadError, message: stringNS.L_InvalidReadForBlankRow };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeUnsupportedDataObject] = { name: stringNS.L_DataWriteError, message: stringNS.L_UnsupportedDataObject };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeCannotWriteToSelection] = { name: stringNS.L_DataWriteError, message: stringNS.L_CannotWriteToSelection };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeDataNotMatchSelection] = { name: stringNS.L_DataWriteError, message: stringNS.L_DataNotMatchSelection };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeOverwriteWorksheetData] = { name: stringNS.L_DataWriteError, message: stringNS.L_OverwriteWorksheetData };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeDataNotMatchBindingSize] = { name: stringNS.L_DataWriteError, message: stringNS.L_DataNotMatchBindingSize };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidSetStartRowColumn] = { name: stringNS.L_DataWriteError, message: stringNS.L_InvalidSetStartRowColumn };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidDataFormat] = { name: stringNS.L_InvalidFormat, message: stringNS.L_InvalidDataFormat };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeDataNotMatchCoercionType] = { name: stringNS.L_InvalidDataObject, message: stringNS.L_DataNotMatchCoercionType };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeDataNotMatchBindingType] = { name: stringNS.L_InvalidDataObject, message: stringNS.L_DataNotMatchBindingType };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeSetDataIsTooLarge] = { name: stringNS.L_DataWriteError, message: stringNS.L_SetDataIsTooLarge };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeNonUniformPartialSetNotSupported] = { name: stringNS.L_DataWriteError, message: stringNS.L_NonUniformPartialSetNotSupported };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidSetColumns] = { name: stringNS.L_DataWriteError, message: stringNS.L_InvalidSetColumns };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidSetRows] = { name: stringNS.L_DataWriteError, message: stringNS.L_InvalidSetRows };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeSetDataParametersConflict] = { name: stringNS.L_DataWriteError, message: stringNS.L_SetDataParametersConflict };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeSelectionCannotBound] = { name: stringNS.L_BindingCreationError, message: stringNS.L_SelectionCannotBound };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeBindingNotExist] = { name: stringNS.L_InvalidBindingError, message: stringNS.L_BindingNotExist };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeBindingToMultipleSelection] = { name: stringNS.L_BindingCreationError, message: stringNS.L_BindingToMultipleSelection };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidSelectionForBindingType] = { name: stringNS.L_BindingCreationError, message: stringNS.L_InvalidSelectionForBindingType };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeOperationNotSupportedOnThisBindingType] = { name: stringNS.L_InvalidBindingOperation, message: stringNS.L_OperationNotSupportedOnThisBindingType };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeNamedItemNotFound] = { name: stringNS.L_BindingCreationError, message: stringNS.L_NamedItemNotFound };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeMultipleNamedItemFound] = { name: stringNS.L_BindingCreationError, message: stringNS.L_MultipleNamedItemFound };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidNamedItemForBindingType] = { name: stringNS.L_BindingCreationError, message: stringNS.L_InvalidNamedItemForBindingType };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeUnknownBindingType] = { name: stringNS.L_InvalidBinding, message: stringNS.L_UnknownBindingType };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeOperationNotSupportedOnMatrixData] = { name: stringNS.L_InvalidBindingOperation, message: stringNS.L_OperationNotSupportedOnMatrixData };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidColumnsForBinding] = { name: stringNS.L_InvalidBinding, message: stringNS.L_InvalidColumnsForBinding };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeSettingNameNotExist] = { name: stringNS.L_ReadSettingsError, message: stringNS.L_SettingNameNotExist };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeSettingsCannotSave] = { name: stringNS.L_SaveSettingsError, message: stringNS.L_SettingsCannotSave };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeSettingsAreStale] = { name: stringNS.L_SettingsStaleError, message: stringNS.L_SettingsAreStale };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeOperationNotSupported] = { name: stringNS.L_HostError, message: stringNS.L_OperationNotSupported };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInternalError] = { name: stringNS.L_InternalError, message: stringNS.L_InternalErrorDescription };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeDocumentReadOnly] = { name: stringNS.L_PermissionDenied, message: stringNS.L_DocumentReadOnly };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeEventHandlerNotExist] = { name: stringNS.L_EventRegistrationError, message: stringNS.L_EventHandlerNotExist };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidApiCallInContext] = { name: stringNS.L_InvalidAPICall, message: stringNS.L_InvalidApiCallInContext };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeShuttingDown] = { name: stringNS.L_ShuttingDown, message: stringNS.L_ShuttingDown };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeUnsupportedEnumeration] = { name: stringNS.L_UnsupportedEnumeration, message: stringNS.L_UnsupportedEnumerationMessage };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeIndexOutOfRange] = { name: stringNS.L_IndexOutOfRange, message: stringNS.L_IndexOutOfRange };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeBrowserAPINotSupported] = { name: stringNS.L_APINotSupported, message: stringNS.L_BrowserAPINotSupported };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeRequestTimeout] = { name: stringNS.L_APICallFailed, message: stringNS.L_RequestTimeout };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidOrTimedOutSession] = { name: stringNS.L_InvalidOrTimedOutSession, message: stringNS.L_InvalidOrTimedOutSessionMessage };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeTooManyIncompleteRequests] = { name: stringNS.L_APICallFailed, message: stringNS.L_TooManyIncompleteRequests };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeRequestTokenUnavailable] = { name: stringNS.L_APICallFailed, message: stringNS.L_RequestTokenUnavailable };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeActivityLimitReached] = { name: stringNS.L_APICallFailed, message: stringNS.L_ActivityLimitReached };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidApiArguments] = { name: stringNS.L_APICallFailed, message: stringNS.L_InvalidApiArgumentsMessage };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeCustomXmlNodeNotFound] = { name: stringNS.L_InvalidNode, message: stringNS.L_CustomXmlNodeNotFound };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeCustomXmlError] = { name: stringNS.L_CustomXmlError, message: stringNS.L_CustomXmlError };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeCustomXmlExceedQuota] = { name: stringNS.L_CustomXmlExceedQuotaName, message: stringNS.L_CustomXmlExceedQuotaMessage };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeCustomXmlOutOfDate] = { name: stringNS.L_CustomXmlOutOfDateName, message: stringNS.L_CustomXmlOutOfDateMessage };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeNoCapability] = { name: stringNS.L_PermissionDenied, message: stringNS.L_NoCapability };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeCannotNavTo] = { name: stringNS.L_CannotNavigateTo, message: stringNS.L_CannotNavigateTo };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeSpecifiedIdNotExist] = { name: stringNS.L_SpecifiedIdNotExist, message: stringNS.L_SpecifiedIdNotExist };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeNavOutOfBound] = { name: stringNS.L_NavOutOfBound, message: stringNS.L_NavOutOfBound };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeCellDataAmountBeyondLimits] = { name: stringNS.L_DataWriteReminder, message: stringNS.L_CellDataAmountBeyondLimits };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeElementMissing] = { name: stringNS.L_MissingParameter, message: stringNS.L_ElementMissing };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeProtectedError] = { name: stringNS.L_PermissionDenied, message: stringNS.L_NoCapability };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidCellsValue] = { name: stringNS.L_InvalidValue, message: stringNS.L_InvalidCellsValue };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidTableOptionValue] = { name: stringNS.L_InvalidValue, message: stringNS.L_InvalidTableOptionValue };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidFormatValue] = { name: stringNS.L_InvalidValue, message: stringNS.L_InvalidFormatValue };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeRowIndexOutOfRange] = { name: stringNS.L_OutOfRange, message: stringNS.L_RowIndexOutOfRange };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeColIndexOutOfRange] = { name: stringNS.L_OutOfRange, message: stringNS.L_ColIndexOutOfRange };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeFormatValueOutOfRange] = { name: stringNS.L_OutOfRange, message: stringNS.L_FormatValueOutOfRange };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeCellFormatAmountBeyondLimits] = { name: stringNS.L_FormattingReminder, message: stringNS.L_CellFormatAmountBeyondLimits };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeMemoryFileLimit] = { name: stringNS.L_MemoryLimit, message: stringNS.L_CloseFileBeforeRetrieve };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeNetworkProblemRetrieveFile] = { name: stringNS.L_NetworkProblem, message: stringNS.L_NetworkProblemRetrieveFile };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidSliceSize] = { name: stringNS.L_InvalidValue, message: stringNS.L_SliceSizeNotSupported };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeDialogAlreadyOpened] = { name: stringNS.L_DisplayDialogError, message: stringNS.L_DialogAlreadyOpened };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidWidth] = { name: stringNS.L_IndexOutOfRange, message: stringNS.L_IndexOutOfRange };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidHeight] = { name: stringNS.L_IndexOutOfRange, message: stringNS.L_IndexOutOfRange };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeNavigationError] = { name: stringNS.L_DisplayDialogError, message: stringNS.L_NetworkProblem };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidScheme] = { name: stringNS.L_DialogNavigateError, message: stringNS.L_DialogInvalidScheme };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeAppDomains] = { name: stringNS.L_DisplayDialogError, message: stringNS.L_DialogAddressNotTrusted };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeRequireHTTPS] = { name: stringNS.L_DisplayDialogError, message: stringNS.L_DialogRequireHTTPS };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeEndUserIgnore] = { name: stringNS.L_DisplayDialogError, message: stringNS.L_UserClickIgnore };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeCrossZone] = { name: stringNS.L_DisplayDialogError, message: stringNS.L_NewWindowCrossZoneErrorString };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeNotSSOAgave] = { name: stringNS.L_APINotSupported, message: stringNS.L_InvalidSSOAddinMessage };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeSSOUserNotSignedIn] = { name: stringNS.L_UserNotSignedIn, message: stringNS.L_UserNotSignedIn };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeSSOUserAborted] = { name: stringNS.L_UserAborted, message: stringNS.L_UserAbortedMessage };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeSSOUnsupportedUserIdentity] = { name: stringNS.L_UnsupportedUserIdentity, message: stringNS.L_UnsupportedUserIdentityMessage };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeSSOInvalidResourceUrl] = { name: stringNS.L_InvalidResourceUrl, message: stringNS.L_InvalidResourceUrlMessage };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeSSOInvalidGrant] = { name: stringNS.L_InvalidGrant, message: stringNS.L_InvalidGrantMessage };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeSSOClientError] = { name: stringNS.L_SSOClientError, message: stringNS.L_SSOClientErrorMessage };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeSSOServerError] = { name: stringNS.L_SSOServerError, message: stringNS.L_SSOServerErrorMessage };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeAddinIsAlreadyRequestingToken] = { name: stringNS.L_AddinIsAlreadyRequestingToken, message: stringNS.L_AddinIsAlreadyRequestingTokenMessage };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeSSOUserConsentNotSupportedByCurrentAddinCategory] = { name: stringNS.L_SSOUserConsentNotSupportedByCurrentAddinCategory, message: stringNS.L_SSOUserConsentNotSupportedByCurrentAddinCategoryMessage };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeSSOConnectionLost] = { name: stringNS.L_SSOConnectionLostError, message: stringNS.L_SSOConnectionLostErrorMessage };
            _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeOperationCancelled] = { name: stringNS.L_OperationCancelledError, message: stringNS.L_OperationCancelledErrorMessage };
        }
    };
})();
var OfficeExt;
(function (OfficeExt) {
    var Requirement;
    (function (Requirement) {
        var RequirementVersion = (function () {
            function RequirementVersion() {
            }
            return RequirementVersion;
        })();
        Requirement.RequirementVersion = RequirementVersion;
        var RequirementMatrix = (function () {
            function RequirementMatrix(_setMap) {
                this.isSetSupported = function _isSetSupported(name, minVersion) {
                    if (name == undefined) {
                        return false;
                    }
                    if (minVersion == undefined) {
                        minVersion = 0;
                    }
                    var setSupportArray = this._setMap;
                    var sets = setSupportArray._sets;
                    if (sets.hasOwnProperty(name.toLowerCase())) {
                        var setMaxVersion = sets[name.toLowerCase()];
                        try {
                            var setMaxVersionNum = this._getVersion(setMaxVersion);
                            minVersion = minVersion + "";
                            var minVersionNum = this._getVersion(minVersion);
                            if (setMaxVersionNum.major > 0 && setMaxVersionNum.major > minVersionNum.major) {
                                return true;
                            }
                            if (setMaxVersionNum.minor > 0 &&
                                setMaxVersionNum.minor > 0 &&
                                setMaxVersionNum.major == minVersionNum.major &&
                                setMaxVersionNum.minor >= minVersionNum.minor) {
                                return true;
                            }
                        }
                        catch (e) {
                            return false;
                        }
                    }
                    return false;
                };
                this._getVersion = function (version) {
                    var temp = version.split(".");
                    var major = 0;
                    var minor = 0;
                    if (temp.length < 2 && isNaN(Number(version))) {
                        throw "version format incorrect";
                    }
                    else {
                        major = Number(temp[0]);
                        if (temp.length >= 2) {
                            minor = Number(temp[1]);
                        }
                        if (isNaN(major) || isNaN(minor)) {
                            throw "version format incorrect";
                        }
                    }
                    var result = { "minor": minor, "major": major };
                    return result;
                };
                this._setMap = _setMap;
                this.isSetSupported = this.isSetSupported.bind(this);
            }
            return RequirementMatrix;
        })();
        Requirement.RequirementMatrix = RequirementMatrix;
        var DefaultSetRequirement = (function () {
            function DefaultSetRequirement(setMap) {
                this._addSetMap = function DefaultSetRequirement_addSetMap(addedSet) {
                    for (var name in addedSet) {
                        this._sets[name] = addedSet[name];
                    }
                };
                this._sets = setMap;
            }
            return DefaultSetRequirement;
        })();
        Requirement.DefaultSetRequirement = DefaultSetRequirement;
        var DefaultDialogSetRequirement = (function (_super) {
            __extends(DefaultDialogSetRequirement, _super);
            function DefaultDialogSetRequirement() {
                _super.call(this, {
                    "dialogapi": 1.1
                });
            }
            return DefaultDialogSetRequirement;
        })(DefaultSetRequirement);
        Requirement.DefaultDialogSetRequirement = DefaultDialogSetRequirement;
        var ExcelClientDefaultSetRequirement = (function (_super) {
            __extends(ExcelClientDefaultSetRequirement, _super);
            function ExcelClientDefaultSetRequirement() {
                _super.call(this, {
                    "bindingevents": 1.1,
                    "documentevents": 1.1,
                    "excelapi": 1.1,
                    "matrixbindings": 1.1,
                    "matrixcoercion": 1.1,
                    "selection": 1.1,
                    "settings": 1.1,
                    "tablebindings": 1.1,
                    "tablecoercion": 1.1,
                    "textbindings": 1.1,
                    "textcoercion": 1.1
                });
            }
            return ExcelClientDefaultSetRequirement;
        })(DefaultSetRequirement);
        Requirement.ExcelClientDefaultSetRequirement = ExcelClientDefaultSetRequirement;
        var ExcelClientV1DefaultSetRequirement = (function (_super) {
            __extends(ExcelClientV1DefaultSetRequirement, _super);
            function ExcelClientV1DefaultSetRequirement() {
                _super.call(this);
                this._addSetMap({
                    "imagecoercion": 1.1
                });
            }
            return ExcelClientV1DefaultSetRequirement;
        })(ExcelClientDefaultSetRequirement);
        Requirement.ExcelClientV1DefaultSetRequirement = ExcelClientV1DefaultSetRequirement;
        var OutlookClientDefaultSetRequirement = (function (_super) {
            __extends(OutlookClientDefaultSetRequirement, _super);
            function OutlookClientDefaultSetRequirement() {
                _super.call(this, {
                    "mailbox": 1.3
                });
            }
            return OutlookClientDefaultSetRequirement;
        })(DefaultSetRequirement);
        Requirement.OutlookClientDefaultSetRequirement = OutlookClientDefaultSetRequirement;
        var WordClientDefaultSetRequirement = (function (_super) {
            __extends(WordClientDefaultSetRequirement, _super);
            function WordClientDefaultSetRequirement() {
                _super.call(this, {
                    "bindingevents": 1.1,
                    "compressedfile": 1.1,
                    "customxmlparts": 1.1,
                    "documentevents": 1.1,
                    "file": 1.1,
                    "htmlcoercion": 1.1,
                    "matrixbindings": 1.1,
                    "matrixcoercion": 1.1,
                    "ooxmlcoercion": 1.1,
                    "pdffile": 1.1,
                    "selection": 1.1,
                    "settings": 1.1,
                    "tablebindings": 1.1,
                    "tablecoercion": 1.1,
                    "textbindings": 1.1,
                    "textcoercion": 1.1,
                    "textfile": 1.1,
                    "wordapi": 1.1
                });
            }
            return WordClientDefaultSetRequirement;
        })(DefaultSetRequirement);
        Requirement.WordClientDefaultSetRequirement = WordClientDefaultSetRequirement;
        var WordClientV1DefaultSetRequirement = (function (_super) {
            __extends(WordClientV1DefaultSetRequirement, _super);
            function WordClientV1DefaultSetRequirement() {
                _super.call(this);
                this._addSetMap({
                    "customxmlparts": 1.2,
                    "wordapi": 1.2,
                    "imagecoercion": 1.1
                });
            }
            return WordClientV1DefaultSetRequirement;
        })(WordClientDefaultSetRequirement);
        Requirement.WordClientV1DefaultSetRequirement = WordClientV1DefaultSetRequirement;
        var PowerpointClientDefaultSetRequirement = (function (_super) {
            __extends(PowerpointClientDefaultSetRequirement, _super);
            function PowerpointClientDefaultSetRequirement() {
                _super.call(this, {
                    "activeview": 1.1,
                    "compressedfile": 1.1,
                    "documentevents": 1.1,
                    "file": 1.1,
                    "pdffile": 1.1,
                    "selection": 1.1,
                    "settings": 1.1,
                    "textcoercion": 1.1
                });
            }
            return PowerpointClientDefaultSetRequirement;
        })(DefaultSetRequirement);
        Requirement.PowerpointClientDefaultSetRequirement = PowerpointClientDefaultSetRequirement;
        var PowerpointClientV1DefaultSetRequirement = (function (_super) {
            __extends(PowerpointClientV1DefaultSetRequirement, _super);
            function PowerpointClientV1DefaultSetRequirement() {
                _super.call(this);
                this._addSetMap({
                    "imagecoercion": 1.1
                });
            }
            return PowerpointClientV1DefaultSetRequirement;
        })(PowerpointClientDefaultSetRequirement);
        Requirement.PowerpointClientV1DefaultSetRequirement = PowerpointClientV1DefaultSetRequirement;
        var ProjectClientDefaultSetRequirement = (function (_super) {
            __extends(ProjectClientDefaultSetRequirement, _super);
            function ProjectClientDefaultSetRequirement() {
                _super.call(this, {
                    "selection": 1.1,
                    "textcoercion": 1.1
                });
            }
            return ProjectClientDefaultSetRequirement;
        })(DefaultSetRequirement);
        Requirement.ProjectClientDefaultSetRequirement = ProjectClientDefaultSetRequirement;
        var ExcelWebDefaultSetRequirement = (function (_super) {
            __extends(ExcelWebDefaultSetRequirement, _super);
            function ExcelWebDefaultSetRequirement() {
                _super.call(this, {
                    "bindingevents": 1.1,
                    "documentevents": 1.1,
                    "matrixbindings": 1.1,
                    "matrixcoercion": 1.1,
                    "selection": 1.1,
                    "settings": 1.1,
                    "tablebindings": 1.1,
                    "tablecoercion": 1.1,
                    "textbindings": 1.1,
                    "textcoercion": 1.1,
                    "file": 1.1
                });
            }
            return ExcelWebDefaultSetRequirement;
        })(DefaultSetRequirement);
        Requirement.ExcelWebDefaultSetRequirement = ExcelWebDefaultSetRequirement;
        var WordWebDefaultSetRequirement = (function (_super) {
            __extends(WordWebDefaultSetRequirement, _super);
            function WordWebDefaultSetRequirement() {
                _super.call(this, {
                    "compressedfile": 1.1,
                    "documentevents": 1.1,
                    "file": 1.1,
                    "imagecoercion": 1.1,
                    "matrixcoercion": 1.1,
                    "ooxmlcoercion": 1.1,
                    "pdffile": 1.1,
                    "selection": 1.1,
                    "settings": 1.1,
                    "tablecoercion": 1.1,
                    "textcoercion": 1.1,
                    "textfile": 1.1
                });
            }
            return WordWebDefaultSetRequirement;
        })(DefaultSetRequirement);
        Requirement.WordWebDefaultSetRequirement = WordWebDefaultSetRequirement;
        var PowerpointWebDefaultSetRequirement = (function (_super) {
            __extends(PowerpointWebDefaultSetRequirement, _super);
            function PowerpointWebDefaultSetRequirement() {
                _super.call(this, {
                    "activeview": 1.1,
                    "settings": 1.1
                });
            }
            return PowerpointWebDefaultSetRequirement;
        })(DefaultSetRequirement);
        Requirement.PowerpointWebDefaultSetRequirement = PowerpointWebDefaultSetRequirement;
        var OutlookWebDefaultSetRequirement = (function (_super) {
            __extends(OutlookWebDefaultSetRequirement, _super);
            function OutlookWebDefaultSetRequirement() {
                _super.call(this, {
                    "mailbox": 1.3
                });
            }
            return OutlookWebDefaultSetRequirement;
        })(DefaultSetRequirement);
        Requirement.OutlookWebDefaultSetRequirement = OutlookWebDefaultSetRequirement;
        var SwayWebDefaultSetRequirement = (function (_super) {
            __extends(SwayWebDefaultSetRequirement, _super);
            function SwayWebDefaultSetRequirement() {
                _super.call(this, {
                    "activeview": 1.1,
                    "documentevents": 1.1,
                    "selection": 1.1,
                    "settings": 1.1,
                    "textcoercion": 1.1
                });
            }
            return SwayWebDefaultSetRequirement;
        })(DefaultSetRequirement);
        Requirement.SwayWebDefaultSetRequirement = SwayWebDefaultSetRequirement;
        var AccessWebDefaultSetRequirement = (function (_super) {
            __extends(AccessWebDefaultSetRequirement, _super);
            function AccessWebDefaultSetRequirement() {
                _super.call(this, {
                    "bindingevents": 1.1,
                    "partialtablebindings": 1.1,
                    "settings": 1.1,
                    "tablebindings": 1.1,
                    "tablecoercion": 1.1
                });
            }
            return AccessWebDefaultSetRequirement;
        })(DefaultSetRequirement);
        Requirement.AccessWebDefaultSetRequirement = AccessWebDefaultSetRequirement;
        var ExcelIOSDefaultSetRequirement = (function (_super) {
            __extends(ExcelIOSDefaultSetRequirement, _super);
            function ExcelIOSDefaultSetRequirement() {
                _super.call(this, {
                    "bindingevents": 1.1,
                    "documentevents": 1.1,
                    "matrixbindings": 1.1,
                    "matrixcoercion": 1.1,
                    "selection": 1.1,
                    "settings": 1.1,
                    "tablebindings": 1.1,
                    "tablecoercion": 1.1,
                    "textbindings": 1.1,
                    "textcoercion": 1.1
                });
            }
            return ExcelIOSDefaultSetRequirement;
        })(DefaultSetRequirement);
        Requirement.ExcelIOSDefaultSetRequirement = ExcelIOSDefaultSetRequirement;
        var WordIOSDefaultSetRequirement = (function (_super) {
            __extends(WordIOSDefaultSetRequirement, _super);
            function WordIOSDefaultSetRequirement() {
                _super.call(this, {
                    "bindingevents": 1.1,
                    "compressedfile": 1.1,
                    "customxmlparts": 1.1,
                    "documentevents": 1.1,
                    "file": 1.1,
                    "htmlcoercion": 1.1,
                    "matrixbindings": 1.1,
                    "matrixcoercion": 1.1,
                    "ooxmlcoercion": 1.1,
                    "pdffile": 1.1,
                    "selection": 1.1,
                    "settings": 1.1,
                    "tablebindings": 1.1,
                    "tablecoercion": 1.1,
                    "textbindings": 1.1,
                    "textcoercion": 1.1,
                    "textfile": 1.1
                });
            }
            return WordIOSDefaultSetRequirement;
        })(DefaultSetRequirement);
        Requirement.WordIOSDefaultSetRequirement = WordIOSDefaultSetRequirement;
        var WordIOSV1DefaultSetRequirement = (function (_super) {
            __extends(WordIOSV1DefaultSetRequirement, _super);
            function WordIOSV1DefaultSetRequirement() {
                _super.call(this);
                this._addSetMap({
                    "customxmlparts": 1.2,
                    "wordapi": 1.2
                });
            }
            return WordIOSV1DefaultSetRequirement;
        })(WordIOSDefaultSetRequirement);
        Requirement.WordIOSV1DefaultSetRequirement = WordIOSV1DefaultSetRequirement;
        var PowerpointIOSDefaultSetRequirement = (function (_super) {
            __extends(PowerpointIOSDefaultSetRequirement, _super);
            function PowerpointIOSDefaultSetRequirement() {
                _super.call(this, {
                    "activeview": 1.1,
                    "compressedfile": 1.1,
                    "documentevents": 1.1,
                    "file": 1.1,
                    "pdffile": 1.1,
                    "selection": 1.1,
                    "settings": 1.1,
                    "textcoercion": 1.1
                });
            }
            return PowerpointIOSDefaultSetRequirement;
        })(DefaultSetRequirement);
        Requirement.PowerpointIOSDefaultSetRequirement = PowerpointIOSDefaultSetRequirement;
        var OutlookIOSDefaultSetRequirement = (function (_super) {
            __extends(OutlookIOSDefaultSetRequirement, _super);
            function OutlookIOSDefaultSetRequirement() {
                _super.call(this, {
                    "mailbox": 1.1
                });
            }
            return OutlookIOSDefaultSetRequirement;
        })(DefaultSetRequirement);
        Requirement.OutlookIOSDefaultSetRequirement = OutlookIOSDefaultSetRequirement;
        var RequirementsMatrixFactory = (function () {
            function RequirementsMatrixFactory() {
            }
            RequirementsMatrixFactory.initializeOsfDda = function () {
                OSF.OUtil.setNamespace("Requirement", OSF.DDA);
            };
            RequirementsMatrixFactory.getDefaultRequirementMatrix = function (appContext) {
                this.initializeDefaultSetMatrix();
                var defaultRequirementMatrix = undefined;
                var clientRequirement = appContext.get_requirementMatrix();
                if (clientRequirement != undefined && clientRequirement.length > 0 && typeof (JSON) !== "undefined") {
                    var matrixItem = JSON.parse(appContext.get_requirementMatrix().toLowerCase());
                    defaultRequirementMatrix = new RequirementMatrix(new DefaultSetRequirement(matrixItem));
                }
                else {
                    var appLocator = RequirementsMatrixFactory.getClientFullVersionString(appContext);
                    if (RequirementsMatrixFactory.DefaultSetArrayMatrix != undefined && RequirementsMatrixFactory.DefaultSetArrayMatrix[appLocator] != undefined) {
                        defaultRequirementMatrix = new RequirementMatrix(RequirementsMatrixFactory.DefaultSetArrayMatrix[appLocator]);
                    }
                    else {
                        defaultRequirementMatrix = new RequirementMatrix(new DefaultSetRequirement({}));
                    }
                }
                return defaultRequirementMatrix;
            };
            RequirementsMatrixFactory.getDefaultDialogRequirementMatrix = function (appContext) {
                var defaultRequirementMatrix = undefined;
                var clientRequirement = appContext.get_dialogRequirementMatrix();
                if (clientRequirement != undefined && clientRequirement.length > 0 && typeof (JSON) !== "undefined") {
                    var matrixItem = JSON.parse(appContext.get_requirementMatrix().toLowerCase());
                    defaultRequirementMatrix = new RequirementMatrix(new DefaultSetRequirement(matrixItem));
                }
                else {
                    defaultRequirementMatrix = new RequirementMatrix(new DefaultDialogSetRequirement());
                }
                return defaultRequirementMatrix;
            };
            RequirementsMatrixFactory.getClientFullVersionString = function (appContext) {
                var appMinorVersion = appContext.get_appMinorVersion();
                var appMinorVersionString = "";
                var appFullVersion = "";
                var appName = appContext.get_appName();
                var isIOSClient = appName == 1024 ||
                    appName == 4096 ||
                    appName == 8192 ||
                    appName == 65536;
                if (isIOSClient && appContext.get_appVersion() == 1) {
                    if (appName == 4096 && appMinorVersion >= 15) {
                        appFullVersion = "16.00.01";
                    }
                    else {
                        appFullVersion = "16.00";
                    }
                }
                else if (appContext.get_appName() == 64) {
                    appFullVersion = appContext.get_appVersion();
                }
                else {
                    if (appMinorVersion < 10) {
                        appMinorVersionString = "0" + appMinorVersion;
                    }
                    else {
                        appMinorVersionString = "" + appMinorVersion;
                    }
                    appFullVersion = appContext.get_appVersion() + "." + appMinorVersionString;
                }
                return appContext.get_appName() + "-" + appFullVersion;
            };
            RequirementsMatrixFactory.initializeDefaultSetMatrix = function () {
                RequirementsMatrixFactory.DefaultSetArrayMatrix[RequirementsMatrixFactory.Excel_RCLIENT_1600] = new ExcelClientDefaultSetRequirement();
                RequirementsMatrixFactory.DefaultSetArrayMatrix[RequirementsMatrixFactory.Word_RCLIENT_1600] = new WordClientDefaultSetRequirement();
                RequirementsMatrixFactory.DefaultSetArrayMatrix[RequirementsMatrixFactory.PowerPoint_RCLIENT_1600] = new PowerpointClientDefaultSetRequirement();
                RequirementsMatrixFactory.DefaultSetArrayMatrix[RequirementsMatrixFactory.Excel_RCLIENT_1601] = new ExcelClientV1DefaultSetRequirement();
                RequirementsMatrixFactory.DefaultSetArrayMatrix[RequirementsMatrixFactory.Word_RCLIENT_1601] = new WordClientV1DefaultSetRequirement();
                RequirementsMatrixFactory.DefaultSetArrayMatrix[RequirementsMatrixFactory.PowerPoint_RCLIENT_1601] = new PowerpointClientV1DefaultSetRequirement();
                RequirementsMatrixFactory.DefaultSetArrayMatrix[RequirementsMatrixFactory.Outlook_RCLIENT_1600] = new OutlookClientDefaultSetRequirement();
                RequirementsMatrixFactory.DefaultSetArrayMatrix[RequirementsMatrixFactory.Excel_WAC_1600] = new ExcelWebDefaultSetRequirement();
                RequirementsMatrixFactory.DefaultSetArrayMatrix[RequirementsMatrixFactory.Word_WAC_1600] = new WordWebDefaultSetRequirement();
                RequirementsMatrixFactory.DefaultSetArrayMatrix[RequirementsMatrixFactory.Outlook_WAC_1600] = new OutlookWebDefaultSetRequirement();
                RequirementsMatrixFactory.DefaultSetArrayMatrix[RequirementsMatrixFactory.Outlook_WAC_1601] = new OutlookWebDefaultSetRequirement();
                RequirementsMatrixFactory.DefaultSetArrayMatrix[RequirementsMatrixFactory.Project_RCLIENT_1600] = new ProjectClientDefaultSetRequirement();
                RequirementsMatrixFactory.DefaultSetArrayMatrix[RequirementsMatrixFactory.Access_WAC_1600] = new AccessWebDefaultSetRequirement();
                RequirementsMatrixFactory.DefaultSetArrayMatrix[RequirementsMatrixFactory.PowerPoint_WAC_1600] = new PowerpointWebDefaultSetRequirement();
                RequirementsMatrixFactory.DefaultSetArrayMatrix[RequirementsMatrixFactory.Excel_IOS_1600] = new ExcelIOSDefaultSetRequirement();
                RequirementsMatrixFactory.DefaultSetArrayMatrix[RequirementsMatrixFactory.SWAY_WAC_1600] = new SwayWebDefaultSetRequirement();
                RequirementsMatrixFactory.DefaultSetArrayMatrix[RequirementsMatrixFactory.Word_IOS_1600] = new WordIOSDefaultSetRequirement();
                RequirementsMatrixFactory.DefaultSetArrayMatrix[RequirementsMatrixFactory.Word_IOS_16001] = new WordIOSV1DefaultSetRequirement();
                RequirementsMatrixFactory.DefaultSetArrayMatrix[RequirementsMatrixFactory.PowerPoint_IOS_1600] = new PowerpointIOSDefaultSetRequirement();
                RequirementsMatrixFactory.DefaultSetArrayMatrix[RequirementsMatrixFactory.Outlook_IOS_1600] = new OutlookIOSDefaultSetRequirement();
            };
            RequirementsMatrixFactory.Excel_RCLIENT_1600 = "1-16.00";
            RequirementsMatrixFactory.Excel_RCLIENT_1601 = "1-16.01";
            RequirementsMatrixFactory.Word_RCLIENT_1600 = "2-16.00";
            RequirementsMatrixFactory.Word_RCLIENT_1601 = "2-16.01";
            RequirementsMatrixFactory.PowerPoint_RCLIENT_1600 = "4-16.00";
            RequirementsMatrixFactory.PowerPoint_RCLIENT_1601 = "4-16.01";
            RequirementsMatrixFactory.Outlook_RCLIENT_1600 = "8-16.00";
            RequirementsMatrixFactory.Excel_WAC_1600 = "16-16.00";
            RequirementsMatrixFactory.Word_WAC_1600 = "32-16.00";
            RequirementsMatrixFactory.Outlook_WAC_1600 = "64-16.00";
            RequirementsMatrixFactory.Outlook_WAC_1601 = "64-16.01";
            RequirementsMatrixFactory.Project_RCLIENT_1600 = "128-16.00";
            RequirementsMatrixFactory.Access_WAC_1600 = "256-16.00";
            RequirementsMatrixFactory.PowerPoint_WAC_1600 = "512-16.00";
            RequirementsMatrixFactory.Excel_IOS_1600 = "1024-16.00";
            RequirementsMatrixFactory.SWAY_WAC_1600 = "2048-16.00";
            RequirementsMatrixFactory.Word_IOS_1600 = "4096-16.00";
            RequirementsMatrixFactory.Word_IOS_16001 = "4096-16.00.01";
            RequirementsMatrixFactory.PowerPoint_IOS_1600 = "8192-16.00";
            RequirementsMatrixFactory.Outlook_IOS_1600 = "65536-16.00";
            RequirementsMatrixFactory.DefaultSetArrayMatrix = {};
            return RequirementsMatrixFactory;
        })();
        Requirement.RequirementsMatrixFactory = RequirementsMatrixFactory;
    })(Requirement = OfficeExt.Requirement || (OfficeExt.Requirement = {}));
})(OfficeExt || (OfficeExt = {}));
OfficeExt.Requirement.RequirementsMatrixFactory.initializeOsfDda();
Microsoft.Office.WebExtension.ApplicationMode = {
    WebEditor: "webEditor",
    WebViewer: "webViewer",
    Client: "client"
};
Microsoft.Office.WebExtension.DocumentMode = {
    ReadOnly: "readOnly",
    ReadWrite: "readWrite"
};
OSF.NamespaceManager = (function OSF_NamespaceManager() {
    var _userOffice;
    var _useShortcut = false;
    return {
        enableShortcut: function OSF_NamespaceManager$enableShortcut() {
            if (!_useShortcut) {
                if (window.Office) {
                    _userOffice = window.Office;
                }
                else {
                    OSF.OUtil.setNamespace("Office", window);
                }
                window.Office = Microsoft.Office.WebExtension;
                _useShortcut = true;
            }
        },
        disableShortcut: function OSF_NamespaceManager$disableShortcut() {
            if (_useShortcut) {
                if (_userOffice) {
                    window.Office = _userOffice;
                }
                else {
                    OSF.OUtil.unsetNamespace("Office", window);
                }
                _useShortcut = false;
            }
        }
    };
})();
OSF.NamespaceManager.enableShortcut();
Microsoft.Office.WebExtension.useShortNamespace = function Microsoft_Office_WebExtension_useShortNamespace(useShortcut) {
    if (useShortcut) {
        OSF.NamespaceManager.enableShortcut();
    }
    else {
        OSF.NamespaceManager.disableShortcut();
    }
};
Microsoft.Office.WebExtension.select = function Microsoft_Office_WebExtension_select(str, errorCallback) {
    var promise;
    if (str && typeof str == "string") {
        var index = str.indexOf("#");
        if (index != -1) {
            var op = str.substring(0, index);
            var target = str.substring(index + 1);
            switch (op) {
                case "binding":
                case "bindings":
                    if (target) {
                        promise = new OSF.DDA.BindingPromise(target);
                    }
                    break;
            }
        }
    }
    if (!promise) {
        if (errorCallback) {
            var callbackType = typeof errorCallback;
            if (callbackType == "function") {
                var callArgs = {};
                callArgs[Microsoft.Office.WebExtension.Parameters.Callback] = errorCallback;
                OSF.DDA.issueAsyncResult(callArgs, OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidApiCallInContext, OSF.DDA.ErrorCodeManager.getErrorArgs(OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidApiCallInContext));
            }
            else {
                throw OSF.OUtil.formatString(Strings.OfficeOM.L_CallbackNotAFunction, callbackType);
            }
        }
    }
    else {
        promise.onFail = errorCallback;
        return promise;
    }
};
OSF.DDA.Context = function OSF_DDA_Context(officeAppContext, document, license, appOM, getOfficeTheme) {
    OSF.OUtil.defineEnumerableProperties(this, {
        "contentLanguage": {
            value: officeAppContext.get_dataLocale()
        },
        "displayLanguage": {
            value: officeAppContext.get_appUILocale()
        },
        "touchEnabled": {
            value: officeAppContext.get_touchEnabled()
        },
        "commerceAllowed": {
            value: officeAppContext.get_commerceAllowed()
        },
        "host": {
            value: OfficeExt.HostName.Host.getInstance().getHost()
        },
        "platform": {
            value: OfficeExt.HostName.Host.getInstance().getPlatform()
        },
        "diagnostics": {
            value: OfficeExt.HostName.Host.getInstance().getDiagnostics(officeAppContext.get_hostFullVersion())
        }
    });
    if (license) {
        OSF.OUtil.defineEnumerableProperty(this, "license", {
            value: license
        });
    }
    if (officeAppContext.ui) {
        OSF.OUtil.defineEnumerableProperty(this, "ui", {
            value: officeAppContext.ui
        });
    }
    if (officeAppContext.auth) {
        OSF.OUtil.defineEnumerableProperty(this, "auth", {
            value: officeAppContext.auth
        });
    }
    if (officeAppContext.application) {
        OSF.OUtil.defineEnumerableProperty(this, "application", {
            value: officeAppContext.application
        });
    }
    if (officeAppContext.get_isDialog()) {
        var requirements = OfficeExt.Requirement.RequirementsMatrixFactory.getDefaultDialogRequirementMatrix(officeAppContext);
        OSF.OUtil.defineEnumerableProperty(this, "requirements", {
            value: requirements
        });
    }
    else {
        if (document) {
            OSF.OUtil.defineEnumerableProperty(this, "document", {
                value: document
            });
        }
        if (appOM) {
            var displayName = appOM.displayName || "appOM";
            delete appOM.displayName;
            OSF.OUtil.defineEnumerableProperty(this, displayName, {
                value: appOM
            });
        }
        if (getOfficeTheme) {
            OSF.OUtil.defineEnumerableProperty(this, "officeTheme", {
                get: function () {
                    return getOfficeTheme();
                }
            });
        }
        var requirements = OfficeExt.Requirement.RequirementsMatrixFactory.getDefaultRequirementMatrix(officeAppContext);
        OSF.OUtil.defineEnumerableProperty(this, "requirements", {
            value: requirements
        });
    }
};
OSF.DDA.OutlookContext = function OSF_DDA_OutlookContext(appContext, settings, license, appOM, getOfficeTheme) {
    OSF.DDA.OutlookContext.uber.constructor.call(this, appContext, null, license, appOM, getOfficeTheme);
    if (settings) {
        OSF.OUtil.defineEnumerableProperty(this, "roamingSettings", {
            value: settings
        });
    }
};
OSF.OUtil.extend(OSF.DDA.OutlookContext, OSF.DDA.Context);
OSF.DDA.OutlookAppOm = function OSF_DDA_OutlookAppOm(appContext, window, appReady) { };
OSF.DDA.Application = function OSF_DDA_Application(officeAppContext) {
};
OSF.DDA.Document = function OSF_DDA_Document(officeAppContext, settings) {
    var mode;
    switch (officeAppContext.get_clientMode()) {
        case OSF.ClientMode.ReadOnly:
            mode = Microsoft.Office.WebExtension.DocumentMode.ReadOnly;
            break;
        case OSF.ClientMode.ReadWrite:
            mode = Microsoft.Office.WebExtension.DocumentMode.ReadWrite;
            break;
    }
    ;
    if (settings) {
        OSF.OUtil.defineEnumerableProperty(this, "settings", {
            value: settings
        });
    }
    ;
    OSF.OUtil.defineMutableProperties(this, {
        "mode": {
            value: mode
        },
        "url": {
            value: officeAppContext.get_docUrl()
        }
    });
};
OSF.DDA.JsomDocument = function OSF_DDA_JsomDocument(officeAppContext, bindingFacade, settings) {
    OSF.DDA.JsomDocument.uber.constructor.call(this, officeAppContext, settings);
    if (bindingFacade) {
        OSF.OUtil.defineEnumerableProperty(this, "bindings", {
            get: function OSF_DDA_Document$GetBindings() { return bindingFacade; }
        });
    }
    var am = OSF.DDA.AsyncMethodNames;
    OSF.DDA.DispIdHost.addAsyncMethods(this, [
        am.GetSelectedDataAsync,
        am.SetSelectedDataAsync
    ]);
    OSF.DDA.DispIdHost.addEventSupport(this, new OSF.EventDispatch([Microsoft.Office.WebExtension.EventType.DocumentSelectionChanged]));
};
OSF.OUtil.extend(OSF.DDA.JsomDocument, OSF.DDA.Document);
OSF.OUtil.defineEnumerableProperty(Microsoft.Office.WebExtension, "context", {
    get: function Microsoft_Office_WebExtension$GetContext() {
        var context;
        if (OSF && OSF._OfficeAppFactory) {
            context = OSF._OfficeAppFactory.getContext();
        }
        return context;
    }
});
OSF.DDA.License = function OSF_DDA_License(eToken) {
    OSF.OUtil.defineEnumerableProperty(this, "value", {
        value: eToken
    });
};
OSF.DDA.ApiMethodCall = function OSF_DDA_ApiMethodCall(requiredParameters, supportedOptions, privateStateCallbacks, checkCallArgs, displayName) {
    var requiredCount = requiredParameters.length;
    var getInvalidParameterString = OSF.OUtil.delayExecutionAndCache(function () {
        return OSF.OUtil.formatString(Strings.OfficeOM.L_InvalidParameters, displayName);
    });
    this.verifyArguments = function OSF_DDA_ApiMethodCall$VerifyArguments(params, args) {
        for (var name in params) {
            var param = params[name];
            var arg = args[name];
            if (param["enum"]) {
                switch (typeof arg) {
                    case "string":
                        if (OSF.OUtil.listContainsValue(param["enum"], arg)) {
                            break;
                        }
                    case "undefined":
                        throw OSF.DDA.ErrorCodeManager.errorCodes.ooeUnsupportedEnumeration;
                    default:
                        throw getInvalidParameterString();
                }
            }
            if (param["types"]) {
                if (!OSF.OUtil.listContainsValue(param["types"], typeof arg)) {
                    throw getInvalidParameterString();
                }
            }
        }
    };
    this.extractRequiredArguments = function OSF_DDA_ApiMethodCall$ExtractRequiredArguments(userArgs, caller, stateInfo) {
        if (userArgs.length < requiredCount) {
            throw OsfMsAjaxFactory.msAjaxError.parameterCount(Strings.OfficeOM.L_MissingRequiredArguments);
        }
        var requiredArgs = [];
        var index;
        for (index = 0; index < requiredCount; index++) {
            requiredArgs.push(userArgs[index]);
        }
        this.verifyArguments(requiredParameters, requiredArgs);
        var ret = {};
        for (index = 0; index < requiredCount; index++) {
            var param = requiredParameters[index];
            var arg = requiredArgs[index];
            if (param.verify) {
                var isValid = param.verify(arg, caller, stateInfo);
                if (!isValid) {
                    throw getInvalidParameterString();
                }
            }
            ret[param.name] = arg;
        }
        return ret;
    },
        this.fillOptions = function OSF_DDA_ApiMethodCall$FillOptions(options, requiredArgs, caller, stateInfo) {
            options = options || {};
            for (var optionName in supportedOptions) {
                if (!OSF.OUtil.listContainsKey(options, optionName)) {
                    var value = undefined;
                    var option = supportedOptions[optionName];
                    if (option.calculate && requiredArgs) {
                        value = option.calculate(requiredArgs, caller, stateInfo);
                    }
                    if (!value && option.defaultValue !== undefined) {
                        value = option.defaultValue;
                    }
                    options[optionName] = value;
                }
            }
            return options;
        };
    this.constructCallArgs = function OSF_DAA_ApiMethodCall$ConstructCallArgs(required, options, caller, stateInfo) {
        var callArgs = {};
        for (var r in required) {
            callArgs[r] = required[r];
        }
        for (var o in options) {
            callArgs[o] = options[o];
        }
        for (var s in privateStateCallbacks) {
            callArgs[s] = privateStateCallbacks[s](caller, stateInfo);
        }
        if (checkCallArgs) {
            callArgs = checkCallArgs(callArgs, caller, stateInfo);
        }
        return callArgs;
    };
};
OSF.OUtil.setNamespace("AsyncResultEnum", OSF.DDA);
OSF.DDA.AsyncResultEnum.Properties = {
    Context: "Context",
    Value: "Value",
    Status: "Status",
    Error: "Error"
};
Microsoft.Office.WebExtension.AsyncResultStatus = {
    Succeeded: "succeeded",
    Failed: "failed"
};
OSF.DDA.AsyncResultEnum.ErrorCode = {
    Success: 0,
    Failed: 1
};
OSF.DDA.AsyncResultEnum.ErrorProperties = {
    Name: "Name",
    Message: "Message",
    Code: "Code"
};
OSF.DDA.AsyncMethodNames = {};
OSF.DDA.AsyncMethodNames.addNames = function (methodNames) {
    for (var entry in methodNames) {
        var am = {};
        OSF.OUtil.defineEnumerableProperties(am, {
            "id": {
                value: entry
            },
            "displayName": {
                value: methodNames[entry]
            }
        });
        OSF.DDA.AsyncMethodNames[entry] = am;
    }
};
OSF.DDA.AsyncMethodCall = function OSF_DDA_AsyncMethodCall(requiredParameters, supportedOptions, privateStateCallbacks, onSucceeded, onFailed, checkCallArgs, displayName) {
    var requiredCount = requiredParameters.length;
    var apiMethods = new OSF.DDA.ApiMethodCall(requiredParameters, supportedOptions, privateStateCallbacks, checkCallArgs, displayName);
    function OSF_DAA_AsyncMethodCall$ExtractOptions(userArgs, requiredArgs, caller, stateInfo) {
        if (userArgs.length > requiredCount + 2) {
            throw OsfMsAjaxFactory.msAjaxError.parameterCount(Strings.OfficeOM.L_TooManyArguments);
        }
        var options, parameterCallback;
        for (var i = userArgs.length - 1; i >= requiredCount; i--) {
            var argument = userArgs[i];
            switch (typeof argument) {
                case "object":
                    if (options) {
                        throw OsfMsAjaxFactory.msAjaxError.parameterCount(Strings.OfficeOM.L_TooManyOptionalObjects);
                    }
                    else {
                        options = argument;
                    }
                    break;
                case "function":
                    if (parameterCallback) {
                        throw OsfMsAjaxFactory.msAjaxError.parameterCount(Strings.OfficeOM.L_TooManyOptionalFunction);
                    }
                    else {
                        parameterCallback = argument;
                    }
                    break;
                default:
                    throw OsfMsAjaxFactory.msAjaxError.argument(Strings.OfficeOM.L_InValidOptionalArgument);
                    break;
            }
        }
        options = apiMethods.fillOptions(options, requiredArgs, caller, stateInfo);
        if (parameterCallback) {
            if (options[Microsoft.Office.WebExtension.Parameters.Callback]) {
                throw Strings.OfficeOM.L_RedundantCallbackSpecification;
            }
            else {
                options[Microsoft.Office.WebExtension.Parameters.Callback] = parameterCallback;
            }
        }
        apiMethods.verifyArguments(supportedOptions, options);
        return options;
    }
    ;
    this.verifyAndExtractCall = function OSF_DAA_AsyncMethodCall$VerifyAndExtractCall(userArgs, caller, stateInfo) {
        var required = apiMethods.extractRequiredArguments(userArgs, caller, stateInfo);
        var options = OSF_DAA_AsyncMethodCall$ExtractOptions(userArgs, required, caller, stateInfo);
        var callArgs = apiMethods.constructCallArgs(required, options, caller, stateInfo);
        return callArgs;
    };
    this.processResponse = function OSF_DAA_AsyncMethodCall$ProcessResponse(status, response, caller, callArgs) {
        var payload;
        if (status == OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess) {
            if (onSucceeded) {
                payload = onSucceeded(response, caller, callArgs);
            }
            else {
                payload = response;
            }
        }
        else {
            if (onFailed) {
                payload = onFailed(status, response);
            }
            else {
                payload = OSF.DDA.ErrorCodeManager.getErrorArgs(status);
            }
        }
        return payload;
    };
    this.getCallArgs = function (suppliedArgs) {
        var options, parameterCallback;
        for (var i = suppliedArgs.length - 1; i >= requiredCount; i--) {
            var argument = suppliedArgs[i];
            switch (typeof argument) {
                case "object":
                    options = argument;
                    break;
                case "function":
                    parameterCallback = argument;
                    break;
            }
        }
        options = options || {};
        if (parameterCallback) {
            options[Microsoft.Office.WebExtension.Parameters.Callback] = parameterCallback;
        }
        return options;
    };
};
OSF.DDA.AsyncMethodCallFactory = (function () {
    return {
        manufacture: function (params) {
            var supportedOptions = params.supportedOptions ? OSF.OUtil.createObject(params.supportedOptions) : [];
            var privateStateCallbacks = params.privateStateCallbacks ? OSF.OUtil.createObject(params.privateStateCallbacks) : [];
            return new OSF.DDA.AsyncMethodCall(params.requiredArguments || [], supportedOptions, privateStateCallbacks, params.onSucceeded, params.onFailed, params.checkCallArgs, params.method.displayName);
        }
    };
})();
OSF.DDA.AsyncMethodCalls = {};
OSF.DDA.AsyncMethodCalls.define = function (callDefinition) {
    OSF.DDA.AsyncMethodCalls[callDefinition.method.id] = OSF.DDA.AsyncMethodCallFactory.manufacture(callDefinition);
};
OSF.DDA.Error = function OSF_DDA_Error(name, message, code) {
    OSF.OUtil.defineEnumerableProperties(this, {
        "name": {
            value: name
        },
        "message": {
            value: message
        },
        "code": {
            value: code
        }
    });
};
OSF.DDA.AsyncResult = function OSF_DDA_AsyncResult(initArgs, errorArgs) {
    OSF.OUtil.defineEnumerableProperties(this, {
        "value": {
            value: initArgs[OSF.DDA.AsyncResultEnum.Properties.Value]
        },
        "status": {
            value: errorArgs ? Microsoft.Office.WebExtension.AsyncResultStatus.Failed : Microsoft.Office.WebExtension.AsyncResultStatus.Succeeded
        }
    });
    if (initArgs[OSF.DDA.AsyncResultEnum.Properties.Context]) {
        OSF.OUtil.defineEnumerableProperty(this, "asyncContext", {
            value: initArgs[OSF.DDA.AsyncResultEnum.Properties.Context]
        });
    }
    if (errorArgs) {
        OSF.OUtil.defineEnumerableProperty(this, "error", {
            value: new OSF.DDA.Error(errorArgs[OSF.DDA.AsyncResultEnum.ErrorProperties.Name], errorArgs[OSF.DDA.AsyncResultEnum.ErrorProperties.Message], errorArgs[OSF.DDA.AsyncResultEnum.ErrorProperties.Code])
        });
    }
};
OSF.DDA.issueAsyncResult = function OSF_DDA$IssueAsyncResult(callArgs, status, payload) {
    var callback = callArgs[Microsoft.Office.WebExtension.Parameters.Callback];
    if (callback) {
        var asyncInitArgs = {};
        asyncInitArgs[OSF.DDA.AsyncResultEnum.Properties.Context] = callArgs[Microsoft.Office.WebExtension.Parameters.AsyncContext];
        var errorArgs;
        if (status == OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess) {
            asyncInitArgs[OSF.DDA.AsyncResultEnum.Properties.Value] = payload;
        }
        else {
            errorArgs = {};
            payload = payload || OSF.DDA.ErrorCodeManager.getErrorArgs(OSF.DDA.ErrorCodeManager.errorCodes.ooeInternalError);
            errorArgs[OSF.DDA.AsyncResultEnum.ErrorProperties.Code] = status || OSF.DDA.ErrorCodeManager.errorCodes.ooeInternalError;
            errorArgs[OSF.DDA.AsyncResultEnum.ErrorProperties.Name] = payload.name || payload;
            errorArgs[OSF.DDA.AsyncResultEnum.ErrorProperties.Message] = payload.message || payload;
        }
        callback(new OSF.DDA.AsyncResult(asyncInitArgs, errorArgs));
    }
};
OSF.DDA.SyncMethodNames = {};
OSF.DDA.SyncMethodNames.addNames = function (methodNames) {
    for (var entry in methodNames) {
        var am = {};
        OSF.OUtil.defineEnumerableProperties(am, {
            "id": {
                value: entry
            },
            "displayName": {
                value: methodNames[entry]
            }
        });
        OSF.DDA.SyncMethodNames[entry] = am;
    }
};
OSF.DDA.SyncMethodCall = function OSF_DDA_SyncMethodCall(requiredParameters, supportedOptions, privateStateCallbacks, checkCallArgs, displayName) {
    var requiredCount = requiredParameters.length;
    var apiMethods = new OSF.DDA.ApiMethodCall(requiredParameters, supportedOptions, privateStateCallbacks, checkCallArgs, displayName);
    function OSF_DAA_SyncMethodCall$ExtractOptions(userArgs, requiredArgs, caller, stateInfo) {
        if (userArgs.length > requiredCount + 1) {
            throw OsfMsAjaxFactory.msAjaxError.parameterCount(Strings.OfficeOM.L_TooManyArguments);
        }
        var options, parameterCallback;
        for (var i = userArgs.length - 1; i >= requiredCount; i--) {
            var argument = userArgs[i];
            switch (typeof argument) {
                case "object":
                    if (options) {
                        throw OsfMsAjaxFactory.msAjaxError.parameterCount(Strings.OfficeOM.L_TooManyOptionalObjects);
                    }
                    else {
                        options = argument;
                    }
                    break;
                default:
                    throw OsfMsAjaxFactory.msAjaxError.argument(Strings.OfficeOM.L_InValidOptionalArgument);
                    break;
            }
        }
        options = apiMethods.fillOptions(options, requiredArgs, caller, stateInfo);
        apiMethods.verifyArguments(supportedOptions, options);
        return options;
    }
    ;
    this.verifyAndExtractCall = function OSF_DAA_AsyncMethodCall$VerifyAndExtractCall(userArgs, caller, stateInfo) {
        var required = apiMethods.extractRequiredArguments(userArgs, caller, stateInfo);
        var options = OSF_DAA_SyncMethodCall$ExtractOptions(userArgs, required, caller, stateInfo);
        var callArgs = apiMethods.constructCallArgs(required, options, caller, stateInfo);
        return callArgs;
    };
};
OSF.DDA.SyncMethodCallFactory = (function () {
    return {
        manufacture: function (params) {
            var supportedOptions = params.supportedOptions ? OSF.OUtil.createObject(params.supportedOptions) : [];
            return new OSF.DDA.SyncMethodCall(params.requiredArguments || [], supportedOptions, params.privateStateCallbacks, params.checkCallArgs, params.method.displayName);
        }
    };
})();
OSF.DDA.SyncMethodCalls = {};
OSF.DDA.SyncMethodCalls.define = function (callDefinition) {
    OSF.DDA.SyncMethodCalls[callDefinition.method.id] = OSF.DDA.SyncMethodCallFactory.manufacture(callDefinition);
};
OSF.DDA.ListType = (function () {
    var listTypes = {};
    return {
        setListType: function OSF_DDA_ListType$AddListType(t, prop) { listTypes[t] = prop; },
        isListType: function OSF_DDA_ListType$IsListType(t) { return OSF.OUtil.listContainsKey(listTypes, t); },
        getDescriptor: function OSF_DDA_ListType$getDescriptor(t) { return listTypes[t]; }
    };
})();
OSF.DDA.HostParameterMap = function (specialProcessor, mappings) {
    var toHostMap = "toHost";
    var fromHostMap = "fromHost";
    var sourceData = "sourceData";
    var self = "self";
    var dynamicTypes = {};
    dynamicTypes[Microsoft.Office.WebExtension.Parameters.Data] = {
        toHost: function (data) {
            if (data != null && data.rows !== undefined) {
                var tableData = {};
                tableData[OSF.DDA.TableDataProperties.TableRows] = data.rows;
                tableData[OSF.DDA.TableDataProperties.TableHeaders] = data.headers;
                data = tableData;
            }
            return data;
        },
        fromHost: function (args) {
            return args;
        }
    };
    dynamicTypes[Microsoft.Office.WebExtension.Parameters.SampleData] = dynamicTypes[Microsoft.Office.WebExtension.Parameters.Data];
    function mapValues(preimageSet, mapping) {
        var ret = preimageSet ? {} : undefined;
        for (var entry in preimageSet) {
            var preimage = preimageSet[entry];
            var image;
            if (OSF.DDA.ListType.isListType(entry)) {
                image = [];
                for (var subEntry in preimage) {
                    image.push(mapValues(preimage[subEntry], mapping));
                }
            }
            else if (OSF.OUtil.listContainsKey(dynamicTypes, entry)) {
                image = dynamicTypes[entry][mapping](preimage);
            }
            else if (mapping == fromHostMap && specialProcessor.preserveNesting(entry)) {
                image = mapValues(preimage, mapping);
            }
            else {
                var maps = mappings[entry];
                if (maps) {
                    var map = maps[mapping];
                    if (map) {
                        image = map[preimage];
                        if (image === undefined) {
                            image = preimage;
                        }
                    }
                }
                else {
                    image = preimage;
                }
            }
            ret[entry] = image;
        }
        return ret;
    }
    ;
    function generateArguments(imageSet, parameters) {
        var ret;
        for (var param in parameters) {
            var arg;
            if (specialProcessor.isComplexType(param)) {
                arg = generateArguments(imageSet, mappings[param][toHostMap]);
            }
            else {
                arg = imageSet[param];
            }
            if (arg != undefined) {
                if (!ret) {
                    ret = {};
                }
                var index = parameters[param];
                if (index == self) {
                    index = param;
                }
                ret[index] = specialProcessor.pack(param, arg);
            }
        }
        return ret;
    }
    ;
    function extractArguments(source, parameters, extracted) {
        if (!extracted) {
            extracted = {};
        }
        for (var param in parameters) {
            var index = parameters[param];
            var value;
            if (index == self) {
                value = source;
            }
            else if (index == sourceData) {
                extracted[param] = source.toArray();
                continue;
            }
            else {
                value = source[index];
            }
            if (value === null || value === undefined) {
                extracted[param] = undefined;
            }
            else {
                value = specialProcessor.unpack(param, value);
                var map;
                if (specialProcessor.isComplexType(param)) {
                    map = mappings[param][fromHostMap];
                    if (specialProcessor.preserveNesting(param)) {
                        extracted[param] = extractArguments(value, map);
                    }
                    else {
                        extractArguments(value, map, extracted);
                    }
                }
                else {
                    if (OSF.DDA.ListType.isListType(param)) {
                        map = {};
                        var entryDescriptor = OSF.DDA.ListType.getDescriptor(param);
                        map[entryDescriptor] = self;
                        var extractedValues = new Array(value.length);
                        for (var item in value) {
                            extractedValues[item] = extractArguments(value[item], map);
                        }
                        extracted[param] = extractedValues;
                    }
                    else {
                        extracted[param] = value;
                    }
                }
            }
        }
        return extracted;
    }
    ;
    function applyMap(mapName, preimage, mapping) {
        var parameters = mappings[mapName][mapping];
        var image;
        if (mapping == "toHost") {
            var imageSet = mapValues(preimage, mapping);
            image = generateArguments(imageSet, parameters);
        }
        else if (mapping == "fromHost") {
            var argumentSet = extractArguments(preimage, parameters);
            image = mapValues(argumentSet, mapping);
        }
        return image;
    }
    ;
    if (!mappings) {
        mappings = {};
    }
    this.addMapping = function (mapName, description) {
        var toHost, fromHost;
        if (description.map) {
            toHost = description.map;
            fromHost = {};
            for (var preimage in toHost) {
                var image = toHost[preimage];
                if (image == self) {
                    image = preimage;
                }
                fromHost[image] = preimage;
            }
        }
        else {
            toHost = description.toHost;
            fromHost = description.fromHost;
        }
        var pair = mappings[mapName];
        if (pair) {
            var currMap = pair[toHostMap];
            for (var th in currMap)
                toHost[th] = currMap[th];
            currMap = pair[fromHostMap];
            for (var fh in currMap)
                fromHost[fh] = currMap[fh];
        }
        else {
            pair = mappings[mapName] = {};
        }
        pair[toHostMap] = toHost;
        pair[fromHostMap] = fromHost;
    };
    this.toHost = function (mapName, preimage) { return applyMap(mapName, preimage, toHostMap); };
    this.fromHost = function (mapName, image) { return applyMap(mapName, image, fromHostMap); };
    this.self = self;
    this.sourceData = sourceData;
    this.addComplexType = function (ct) { specialProcessor.addComplexType(ct); };
    this.getDynamicType = function (dt) { return specialProcessor.getDynamicType(dt); };
    this.setDynamicType = function (dt, handler) { specialProcessor.setDynamicType(dt, handler); };
    this.dynamicTypes = dynamicTypes;
    this.doMapValues = function (preimageSet, mapping) { return mapValues(preimageSet, mapping); };
};
OSF.DDA.SpecialProcessor = function (complexTypes, dynamicTypes) {
    this.addComplexType = function OSF_DDA_SpecialProcessor$addComplexType(ct) {
        complexTypes.push(ct);
    };
    this.getDynamicType = function OSF_DDA_SpecialProcessor$getDynamicType(dt) {
        return dynamicTypes[dt];
    };
    this.setDynamicType = function OSF_DDA_SpecialProcessor$setDynamicType(dt, handler) {
        dynamicTypes[dt] = handler;
    };
    this.isComplexType = function OSF_DDA_SpecialProcessor$isComplexType(t) {
        return OSF.OUtil.listContainsValue(complexTypes, t);
    };
    this.isDynamicType = function OSF_DDA_SpecialProcessor$isDynamicType(p) {
        return OSF.OUtil.listContainsKey(dynamicTypes, p);
    };
    this.preserveNesting = function OSF_DDA_SpecialProcessor$preserveNesting(p) {
        var pn = [];
        if (OSF.DDA.PropertyDescriptors)
            pn.push(OSF.DDA.PropertyDescriptors.Subset);
        if (OSF.DDA.DataNodeEventProperties) {
            pn = pn.concat([
                OSF.DDA.DataNodeEventProperties.OldNode,
                OSF.DDA.DataNodeEventProperties.NewNode,
                OSF.DDA.DataNodeEventProperties.NextSiblingNode
            ]);
        }
        return OSF.OUtil.listContainsValue(pn, p);
    };
    this.pack = function OSF_DDA_SpecialProcessor$pack(param, arg) {
        var value;
        if (this.isDynamicType(param)) {
            value = dynamicTypes[param].toHost(arg);
        }
        else {
            value = arg;
        }
        return value;
    };
    this.unpack = function OSF_DDA_SpecialProcessor$unpack(param, arg) {
        var value;
        if (this.isDynamicType(param)) {
            value = dynamicTypes[param].fromHost(arg);
        }
        else {
            value = arg;
        }
        return value;
    };
};
OSF.DDA.getDecoratedParameterMap = function (specialProcessor, initialDefs) {
    var parameterMap = new OSF.DDA.HostParameterMap(specialProcessor);
    var self = parameterMap.self;
    function createObject(properties) {
        var obj = null;
        if (properties) {
            obj = {};
            var len = properties.length;
            for (var i = 0; i < len; i++) {
                obj[properties[i].name] = properties[i].value;
            }
        }
        return obj;
    }
    parameterMap.define = function define(definition) {
        var args = {};
        var toHost = createObject(definition.toHost);
        if (definition.invertible) {
            args.map = toHost;
        }
        else if (definition.canonical) {
            args.toHost = args.fromHost = toHost;
        }
        else {
            args.toHost = toHost;
            args.fromHost = createObject(definition.fromHost);
        }
        parameterMap.addMapping(definition.type, args);
        if (definition.isComplexType)
            parameterMap.addComplexType(definition.type);
    };
    for (var id in initialDefs)
        parameterMap.define(initialDefs[id]);
    return parameterMap;
};
OSF.OUtil.setNamespace("DispIdHost", OSF.DDA);
OSF.DDA.DispIdHost.Methods = {
    InvokeMethod: "invokeMethod",
    AddEventHandler: "addEventHandler",
    RemoveEventHandler: "removeEventHandler",
    OpenDialog: "openDialog",
    CloseDialog: "closeDialog",
    MessageParent: "messageParent",
    SendMessage: "sendMessage"
};
OSF.DDA.DispIdHost.Delegates = {
    ExecuteAsync: "executeAsync",
    RegisterEventAsync: "registerEventAsync",
    UnregisterEventAsync: "unregisterEventAsync",
    ParameterMap: "parameterMap",
    OpenDialog: "openDialog",
    CloseDialog: "closeDialog",
    MessageParent: "messageParent",
    SendMessage: "sendMessage"
};
OSF.DDA.DispIdHost.Facade = function OSF_DDA_DispIdHost_Facade(getDelegateMethods, parameterMap) {
    var dispIdMap = {};
    var jsom = OSF.DDA.AsyncMethodNames;
    var did = OSF.DDA.MethodDispId;
    var methodMap = {
        "GoToByIdAsync": did.dispidNavigateToMethod,
        "GetSelectedDataAsync": did.dispidGetSelectedDataMethod,
        "SetSelectedDataAsync": did.dispidSetSelectedDataMethod,
        "GetDocumentCopyChunkAsync": did.dispidGetDocumentCopyChunkMethod,
        "ReleaseDocumentCopyAsync": did.dispidReleaseDocumentCopyMethod,
        "GetDocumentCopyAsync": did.dispidGetDocumentCopyMethod,
        "AddFromSelectionAsync": did.dispidAddBindingFromSelectionMethod,
        "AddFromPromptAsync": did.dispidAddBindingFromPromptMethod,
        "AddFromNamedItemAsync": did.dispidAddBindingFromNamedItemMethod,
        "GetAllAsync": did.dispidGetAllBindingsMethod,
        "GetByIdAsync": did.dispidGetBindingMethod,
        "ReleaseByIdAsync": did.dispidReleaseBindingMethod,
        "GetDataAsync": did.dispidGetBindingDataMethod,
        "SetDataAsync": did.dispidSetBindingDataMethod,
        "AddRowsAsync": did.dispidAddRowsMethod,
        "AddColumnsAsync": did.dispidAddColumnsMethod,
        "DeleteAllDataValuesAsync": did.dispidClearAllRowsMethod,
        "RefreshAsync": did.dispidLoadSettingsMethod,
        "SaveAsync": did.dispidSaveSettingsMethod,
        "GetActiveViewAsync": did.dispidGetActiveViewMethod,
        "GetFilePropertiesAsync": did.dispidGetFilePropertiesMethod,
        "GetOfficeThemeAsync": did.dispidGetOfficeThemeMethod,
        "GetDocumentThemeAsync": did.dispidGetDocumentThemeMethod,
        "ClearFormatsAsync": did.dispidClearFormatsMethod,
        "SetTableOptionsAsync": did.dispidSetTableOptionsMethod,
        "SetFormatsAsync": did.dispidSetFormatsMethod,
        "GetAccessTokenAsync": did.dispidGetAccessTokenMethod,
        "ExecuteRichApiRequestAsync": did.dispidExecuteRichApiRequestMethod,
        "AppCommandInvocationCompletedAsync": did.dispidAppCommandInvocationCompletedMethod,
        "CloseContainerAsync": did.dispidCloseContainerMethod,
        "OpenBrowserWindow": did.dispidOpenBrowserWindow,
        "CreateDocumentAsync": did.dispidCreateDocumentMethod,
        "InsertFormAsync": did.dispidInsertFormMethod,
        "ExecuteFeature": did.dispidExecuteFeature,
        "QueryFeature": did.dispidQueryFeature,
        "AddDataPartAsync": did.dispidAddDataPartMethod,
        "GetDataPartByIdAsync": did.dispidGetDataPartByIdMethod,
        "GetDataPartsByNameSpaceAsync": did.dispidGetDataPartsByNamespaceMethod,
        "GetPartXmlAsync": did.dispidGetDataPartXmlMethod,
        "GetPartNodesAsync": did.dispidGetDataPartNodesMethod,
        "DeleteDataPartAsync": did.dispidDeleteDataPartMethod,
        "GetNodeValueAsync": did.dispidGetDataNodeValueMethod,
        "GetNodeXmlAsync": did.dispidGetDataNodeXmlMethod,
        "GetRelativeNodesAsync": did.dispidGetDataNodesMethod,
        "SetNodeValueAsync": did.dispidSetDataNodeValueMethod,
        "SetNodeXmlAsync": did.dispidSetDataNodeXmlMethod,
        "AddDataPartNamespaceAsync": did.dispidAddDataNamespaceMethod,
        "GetDataPartNamespaceAsync": did.dispidGetDataUriByPrefixMethod,
        "GetDataPartPrefixAsync": did.dispidGetDataPrefixByUriMethod,
        "GetNodeTextAsync": did.dispidGetDataNodeTextMethod,
        "SetNodeTextAsync": did.dispidSetDataNodeTextMethod,
        "GetSelectedTask": did.dispidGetSelectedTaskMethod,
        "GetTask": did.dispidGetTaskMethod,
        "GetWSSUrl": did.dispidGetWSSUrlMethod,
        "GetTaskField": did.dispidGetTaskFieldMethod,
        "GetSelectedResource": did.dispidGetSelectedResourceMethod,
        "GetResourceField": did.dispidGetResourceFieldMethod,
        "GetProjectField": did.dispidGetProjectFieldMethod,
        "GetSelectedView": did.dispidGetSelectedViewMethod,
        "GetTaskByIndex": did.dispidGetTaskByIndexMethod,
        "GetResourceByIndex": did.dispidGetResourceByIndexMethod,
        "SetTaskField": did.dispidSetTaskFieldMethod,
        "SetResourceField": did.dispidSetResourceFieldMethod,
        "GetMaxTaskIndex": did.dispidGetMaxTaskIndexMethod,
        "GetMaxResourceIndex": did.dispidGetMaxResourceIndexMethod,
        "CreateTask": did.dispidCreateTaskMethod
    };
    for (var method in methodMap) {
        if (jsom[method]) {
            dispIdMap[jsom[method].id] = methodMap[method];
        }
    }
    jsom = OSF.DDA.SyncMethodNames;
    did = OSF.DDA.MethodDispId;
    var syncMethodMap = {
        "MessageParent": did.dispidMessageParentMethod,
        "SendMessage": did.dispidSendMessageMethod
    };
    for (var method in syncMethodMap) {
        if (jsom[method]) {
            dispIdMap[jsom[method].id] = syncMethodMap[method];
        }
    }
    jsom = Microsoft.Office.WebExtension.EventType;
    did = OSF.DDA.EventDispId;
    var eventMap = {
        "SettingsChanged": did.dispidSettingsChangedEvent,
        "DocumentSelectionChanged": did.dispidDocumentSelectionChangedEvent,
        "BindingSelectionChanged": did.dispidBindingSelectionChangedEvent,
        "BindingDataChanged": did.dispidBindingDataChangedEvent,
        "ActiveViewChanged": did.dispidActiveViewChangedEvent,
        "OfficeThemeChanged": did.dispidOfficeThemeChangedEvent,
        "DocumentThemeChanged": did.dispidDocumentThemeChangedEvent,
        "AppCommandInvoked": did.dispidAppCommandInvokedEvent,
        "DialogMessageReceived": did.dispidDialogMessageReceivedEvent,
        "DialogParentMessageReceived": did.dispidDialogParentMessageReceivedEvent,
        "ObjectDeleted": did.dispidObjectDeletedEvent,
        "ObjectSelectionChanged": did.dispidObjectSelectionChangedEvent,
        "ObjectDataChanged": did.dispidObjectDataChangedEvent,
        "ContentControlAdded": did.dispidContentControlAddedEvent,
        "RichApiMessage": did.dispidRichApiMessageEvent,
        "ItemChanged": did.dispidOlkItemSelectedChangedEvent,
        "RecipientsChanged": did.dispidOlkRecipientsChangedEvent,
        "AppointmentTimeChanged": did.dispidOlkAppointmentTimeChangedEvent,
        "RecurrenceChanged": did.dispidOlkRecurrenceChangedEvent,
        "TaskSelectionChanged": did.dispidTaskSelectionChangedEvent,
        "ResourceSelectionChanged": did.dispidResourceSelectionChangedEvent,
        "ViewSelectionChanged": did.dispidViewSelectionChangedEvent,
        "DataNodeInserted": did.dispidDataNodeAddedEvent,
        "DataNodeReplaced": did.dispidDataNodeReplacedEvent,
        "DataNodeDeleted": did.dispidDataNodeDeletedEvent
    };
    for (var event in eventMap) {
        if (jsom[event]) {
            dispIdMap[jsom[event]] = eventMap[event];
        }
    }
    function IsObjectEvent(dispId) {
        return (dispId == OSF.DDA.EventDispId.dispidObjectDeletedEvent ||
            dispId == OSF.DDA.EventDispId.dispidObjectSelectionChangedEvent ||
            dispId == OSF.DDA.EventDispId.dispidObjectDataChangedEvent ||
            dispId == OSF.DDA.EventDispId.dispidContentControlAddedEvent);
    }
    function onException(ex, asyncMethodCall, suppliedArgs, callArgs) {
        if (typeof ex == "number") {
            if (!callArgs) {
                callArgs = asyncMethodCall.getCallArgs(suppliedArgs);
            }
            OSF.DDA.issueAsyncResult(callArgs, ex, OSF.DDA.ErrorCodeManager.getErrorArgs(ex));
        }
        else {
            throw ex;
        }
    }
    ;
    this[OSF.DDA.DispIdHost.Methods.InvokeMethod] = function OSF_DDA_DispIdHost_Facade$InvokeMethod(method, suppliedArguments, caller, privateState) {
        var callArgs;
        try {
            var methodName = method.id;
            var asyncMethodCall = OSF.DDA.AsyncMethodCalls[methodName];
            callArgs = asyncMethodCall.verifyAndExtractCall(suppliedArguments, caller, privateState);
            var dispId = dispIdMap[methodName];
            var delegate = getDelegateMethods(methodName);
            var richApiInExcelMethodSubstitution = null;
            if (window.Excel && window.Office.context.requirements.isSetSupported("RedirectV1Api")) {
                window.Excel._RedirectV1APIs = true;
            }
            if (window.Excel && window.Excel._RedirectV1APIs && (richApiInExcelMethodSubstitution = window.Excel._V1APIMap[methodName])) {
                var preprocessedCallArgs = OSF.OUtil.shallowCopy(callArgs);
                delete preprocessedCallArgs[Microsoft.Office.WebExtension.Parameters.AsyncContext];
                if (richApiInExcelMethodSubstitution.preprocess) {
                    preprocessedCallArgs = richApiInExcelMethodSubstitution.preprocess(preprocessedCallArgs);
                }
                var ctx = new window.Excel.RequestContext();
                var result = richApiInExcelMethodSubstitution.call(ctx, preprocessedCallArgs);
                ctx.sync()
                    .then(function () {
                    var response = result.value;
                    var status = response.status;
                    delete response["status"];
                    delete response["@odata.type"];
                    if (richApiInExcelMethodSubstitution.postprocess) {
                        response = richApiInExcelMethodSubstitution.postprocess(response, preprocessedCallArgs);
                    }
                    if (status != 0) {
                        response = OSF.DDA.ErrorCodeManager.getErrorArgs(status);
                    }
                    OSF.DDA.issueAsyncResult(callArgs, status, response);
                })["catch"](function (error) {
                    OSF.DDA.issueAsyncResult(callArgs, OSF.DDA.ErrorCodeManager.errorCodes.ooeFailure, null);
                });
            }
            else {
                var hostCallArgs;
                if (parameterMap.toHost) {
                    hostCallArgs = parameterMap.toHost(dispId, callArgs);
                }
                else {
                    hostCallArgs = callArgs;
                }
                var startTime = (new Date()).getTime();
                delegate[OSF.DDA.DispIdHost.Delegates.ExecuteAsync]({
                    "dispId": dispId,
                    "hostCallArgs": hostCallArgs,
                    "onCalling": function OSF_DDA_DispIdFacade$Execute_onCalling() { },
                    "onReceiving": function OSF_DDA_DispIdFacade$Execute_onReceiving() { },
                    "onComplete": function (status, hostResponseArgs) {
                        var responseArgs;
                        if (status == OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess) {
                            if (parameterMap.fromHost) {
                                responseArgs = parameterMap.fromHost(dispId, hostResponseArgs);
                            }
                            else {
                                responseArgs = hostResponseArgs;
                            }
                        }
                        else {
                            responseArgs = hostResponseArgs;
                        }
                        var payload = asyncMethodCall.processResponse(status, responseArgs, caller, callArgs);
                        OSF.DDA.issueAsyncResult(callArgs, status, payload);
                        if (OSF.AppTelemetry) {
                            OSF.AppTelemetry.onMethodDone(dispId, hostCallArgs, Math.abs((new Date()).getTime() - startTime), status);
                        }
                    }
                });
            }
        }
        catch (ex) {
            onException(ex, asyncMethodCall, suppliedArguments, callArgs);
        }
    };
    this[OSF.DDA.DispIdHost.Methods.AddEventHandler] = function OSF_DDA_DispIdHost_Facade$AddEventHandler(suppliedArguments, eventDispatch, caller, isPopupWindow) {
        var callArgs;
        var eventType, handler;
        var isObjectEvent = false;
        function onEnsureRegistration(status) {
            if (status == OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess) {
                var added = !isObjectEvent ? eventDispatch.addEventHandler(eventType, handler) :
                    eventDispatch.addObjectEventHandler(eventType, callArgs[Microsoft.Office.WebExtension.Parameters.Id], handler);
                if (!added) {
                    status = OSF.DDA.ErrorCodeManager.errorCodes.ooeEventHandlerAdditionFailed;
                }
            }
            var error;
            if (status != OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess) {
                error = OSF.DDA.ErrorCodeManager.getErrorArgs(status);
            }
            OSF.DDA.issueAsyncResult(callArgs, status, error);
        }
        try {
            var asyncMethodCall = OSF.DDA.AsyncMethodCalls[OSF.DDA.AsyncMethodNames.AddHandlerAsync.id];
            callArgs = asyncMethodCall.verifyAndExtractCall(suppliedArguments, caller, eventDispatch);
            eventType = callArgs[Microsoft.Office.WebExtension.Parameters.EventType];
            handler = callArgs[Microsoft.Office.WebExtension.Parameters.Handler];
            if (isPopupWindow) {
                onEnsureRegistration(OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess);
                return;
            }
            var dispId = dispIdMap[eventType];
            isObjectEvent = IsObjectEvent(dispId);
            var targetId = (isObjectEvent ? callArgs[Microsoft.Office.WebExtension.Parameters.Id] : (caller.id || ""));
            var count = isObjectEvent ? eventDispatch.getObjectEventHandlerCount(eventType, targetId) : eventDispatch.getEventHandlerCount(eventType);
            if (count == 0) {
                var invoker = getDelegateMethods(eventType)[OSF.DDA.DispIdHost.Delegates.RegisterEventAsync];
                invoker({
                    "eventType": eventType,
                    "dispId": dispId,
                    "targetId": targetId,
                    "onCalling": function OSF_DDA_DispIdFacade$Execute_onCalling() { OSF.OUtil.writeProfilerMark(OSF.HostCallPerfMarker.IssueCall); },
                    "onReceiving": function OSF_DDA_DispIdFacade$Execute_onReceiving() { OSF.OUtil.writeProfilerMark(OSF.HostCallPerfMarker.ReceiveResponse); },
                    "onComplete": onEnsureRegistration,
                    "onEvent": function handleEvent(hostArgs) {
                        var args = parameterMap.fromHost(dispId, hostArgs);
                        if (!isObjectEvent)
                            eventDispatch.fireEvent(OSF.DDA.OMFactory.manufactureEventArgs(eventType, caller, args));
                        else
                            eventDispatch.fireObjectEvent(targetId, OSF.DDA.OMFactory.manufactureEventArgs(eventType, targetId, args));
                    }
                });
            }
            else {
                onEnsureRegistration(OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess);
            }
        }
        catch (ex) {
            onException(ex, asyncMethodCall, suppliedArguments, callArgs);
        }
    };
    this[OSF.DDA.DispIdHost.Methods.RemoveEventHandler] = function OSF_DDA_DispIdHost_Facade$RemoveEventHandler(suppliedArguments, eventDispatch, caller) {
        var callArgs;
        var eventType, handler;
        var isObjectEvent = false;
        function onEnsureRegistration(status) {
            var error;
            if (status != OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess) {
                error = OSF.DDA.ErrorCodeManager.getErrorArgs(status);
            }
            OSF.DDA.issueAsyncResult(callArgs, status, error);
        }
        try {
            var asyncMethodCall = OSF.DDA.AsyncMethodCalls[OSF.DDA.AsyncMethodNames.RemoveHandlerAsync.id];
            callArgs = asyncMethodCall.verifyAndExtractCall(suppliedArguments, caller, eventDispatch);
            eventType = callArgs[Microsoft.Office.WebExtension.Parameters.EventType];
            handler = callArgs[Microsoft.Office.WebExtension.Parameters.Handler];
            var dispId = dispIdMap[eventType];
            isObjectEvent = IsObjectEvent(dispId);
            var targetId = (isObjectEvent ? callArgs[Microsoft.Office.WebExtension.Parameters.Id] : (caller.id || ""));
            var status, removeSuccess;
            if (handler === null) {
                removeSuccess = isObjectEvent ? eventDispatch.clearObjectEventHandlers(eventType, targetId) : eventDispatch.clearEventHandlers(eventType);
                status = OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess;
            }
            else {
                removeSuccess = isObjectEvent ? eventDispatch.removeObjectEventHandler(eventType, targetId, handler) : eventDispatch.removeEventHandler(eventType, handler);
                status = removeSuccess ? OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess : OSF.DDA.ErrorCodeManager.errorCodes.ooeEventHandlerNotExist;
            }
            var count = isObjectEvent ? eventDispatch.getObjectEventHandlerCount(eventType, targetId) : eventDispatch.getEventHandlerCount(eventType);
            if (removeSuccess && count == 0) {
                var invoker = getDelegateMethods(eventType)[OSF.DDA.DispIdHost.Delegates.UnregisterEventAsync];
                invoker({
                    "eventType": eventType,
                    "dispId": dispId,
                    "targetId": targetId,
                    "onCalling": function OSF_DDA_DispIdFacade$Execute_onCalling() { OSF.OUtil.writeProfilerMark(OSF.HostCallPerfMarker.IssueCall); },
                    "onReceiving": function OSF_DDA_DispIdFacade$Execute_onReceiving() { OSF.OUtil.writeProfilerMark(OSF.HostCallPerfMarker.ReceiveResponse); },
                    "onComplete": onEnsureRegistration
                });
            }
            else {
                onEnsureRegistration(status);
            }
        }
        catch (ex) {
            onException(ex, asyncMethodCall, suppliedArguments, callArgs);
        }
    };
    this[OSF.DDA.DispIdHost.Methods.OpenDialog] = function OSF_DDA_DispIdHost_Facade$OpenDialog(suppliedArguments, eventDispatch, caller) {
        var callArgs;
        var targetId;
        var dialogMessageEvent = Microsoft.Office.WebExtension.EventType.DialogMessageReceived;
        var dialogOtherEvent = Microsoft.Office.WebExtension.EventType.DialogEventReceived;
        function onEnsureRegistration(status) {
            var payload;
            if (status != OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess) {
                payload = OSF.DDA.ErrorCodeManager.getErrorArgs(status);
            }
            else {
                var onSucceedArgs = {};
                onSucceedArgs[Microsoft.Office.WebExtension.Parameters.Id] = targetId;
                onSucceedArgs[Microsoft.Office.WebExtension.Parameters.Data] = eventDispatch;
                var payload = asyncMethodCall.processResponse(status, onSucceedArgs, caller, callArgs);
                OSF.DialogShownStatus.hasDialogShown = true;
                eventDispatch.clearEventHandlers(dialogMessageEvent);
                eventDispatch.clearEventHandlers(dialogOtherEvent);
            }
            OSF.DDA.issueAsyncResult(callArgs, status, payload);
        }
        try {
            if (dialogMessageEvent == undefined || dialogOtherEvent == undefined) {
                onEnsureRegistration(OSF.DDA.ErrorCodeManager.ooeOperationNotSupported);
            }
            if (OSF.DDA.AsyncMethodNames.DisplayDialogAsync == null) {
                onEnsureRegistration(OSF.DDA.ErrorCodeManager.errorCodes.ooeInternalError);
                return;
            }
            var asyncMethodCall = OSF.DDA.AsyncMethodCalls[OSF.DDA.AsyncMethodNames.DisplayDialogAsync.id];
            callArgs = asyncMethodCall.verifyAndExtractCall(suppliedArguments, caller, eventDispatch);
            var dispId = dispIdMap[dialogMessageEvent];
            var delegateMethods = getDelegateMethods(dialogMessageEvent);
            var invoker = delegateMethods[OSF.DDA.DispIdHost.Delegates.OpenDialog] != undefined ?
                delegateMethods[OSF.DDA.DispIdHost.Delegates.OpenDialog] :
                delegateMethods[OSF.DDA.DispIdHost.Delegates.RegisterEventAsync];
            targetId = JSON.stringify(callArgs);
            if (!OSF.DialogShownStatus.hasDialogShown) {
                eventDispatch.clearQueuedEvent(dialogMessageEvent);
                eventDispatch.clearQueuedEvent(dialogOtherEvent);
                eventDispatch.clearQueuedEvent(Microsoft.Office.WebExtension.EventType.DialogParentMessageReceived);
            }
            invoker({
                "eventType": dialogMessageEvent,
                "dispId": dispId,
                "targetId": targetId,
                "onCalling": function OSF_DDA_DispIdFacade$Execute_onCalling() { OSF.OUtil.writeProfilerMark(OSF.HostCallPerfMarker.IssueCall); },
                "onReceiving": function OSF_DDA_DispIdFacade$Execute_onReceiving() { OSF.OUtil.writeProfilerMark(OSF.HostCallPerfMarker.ReceiveResponse); },
                "onComplete": onEnsureRegistration,
                "onEvent": function handleEvent(hostArgs) {
                    var args = parameterMap.fromHost(dispId, hostArgs);
                    var event = OSF.DDA.OMFactory.manufactureEventArgs(dialogMessageEvent, caller, args);
                    if (event.type == dialogOtherEvent) {
                        var payload = OSF.DDA.ErrorCodeManager.getErrorArgs(event.error);
                        var errorArgs = {};
                        errorArgs[OSF.DDA.AsyncResultEnum.ErrorProperties.Code] = status || OSF.DDA.ErrorCodeManager.errorCodes.ooeInternalError;
                        errorArgs[OSF.DDA.AsyncResultEnum.ErrorProperties.Name] = payload.name || payload;
                        errorArgs[OSF.DDA.AsyncResultEnum.ErrorProperties.Message] = payload.message || payload;
                        event.error = new OSF.DDA.Error(errorArgs[OSF.DDA.AsyncResultEnum.ErrorProperties.Name], errorArgs[OSF.DDA.AsyncResultEnum.ErrorProperties.Message], errorArgs[OSF.DDA.AsyncResultEnum.ErrorProperties.Code]);
                    }
                    eventDispatch.fireOrQueueEvent(event);
                    if (args[OSF.DDA.PropertyDescriptors.MessageType] == OSF.DialogMessageType.DialogClosed) {
                        eventDispatch.clearEventHandlers(dialogMessageEvent);
                        eventDispatch.clearEventHandlers(dialogOtherEvent);
                        eventDispatch.clearEventHandlers(Microsoft.Office.WebExtension.EventType.DialogParentMessageReceived);
                        OSF.DialogShownStatus.hasDialogShown = false;
                    }
                }
            });
        }
        catch (ex) {
            onException(ex, asyncMethodCall, suppliedArguments, callArgs);
        }
    };
    this[OSF.DDA.DispIdHost.Methods.CloseDialog] = function OSF_DDA_DispIdHost_Facade$CloseDialog(suppliedArguments, targetId, eventDispatch, caller) {
        var callArgs;
        var dialogMessageEvent, dialogOtherEvent;
        var closeStatus = OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess;
        function closeCallback(status) {
            closeStatus = status;
            OSF.DialogShownStatus.hasDialogShown = false;
        }
        try {
            var asyncMethodCall = OSF.DDA.AsyncMethodCalls[OSF.DDA.AsyncMethodNames.CloseAsync.id];
            callArgs = asyncMethodCall.verifyAndExtractCall(suppliedArguments, caller, eventDispatch);
            dialogMessageEvent = Microsoft.Office.WebExtension.EventType.DialogMessageReceived;
            dialogOtherEvent = Microsoft.Office.WebExtension.EventType.DialogEventReceived;
            eventDispatch.clearEventHandlers(dialogMessageEvent);
            eventDispatch.clearEventHandlers(dialogOtherEvent);
            var dispId = dispIdMap[dialogMessageEvent];
            var delegateMethods = getDelegateMethods(dialogMessageEvent);
            var invoker = delegateMethods[OSF.DDA.DispIdHost.Delegates.CloseDialog] != undefined ?
                delegateMethods[OSF.DDA.DispIdHost.Delegates.CloseDialog] :
                delegateMethods[OSF.DDA.DispIdHost.Delegates.UnregisterEventAsync];
            invoker({
                "eventType": dialogMessageEvent,
                "dispId": dispId,
                "targetId": targetId,
                "onCalling": function OSF_DDA_DispIdFacade$Execute_onCalling() { OSF.OUtil.writeProfilerMark(OSF.HostCallPerfMarker.IssueCall); },
                "onReceiving": function OSF_DDA_DispIdFacade$Execute_onReceiving() { OSF.OUtil.writeProfilerMark(OSF.HostCallPerfMarker.ReceiveResponse); },
                "onComplete": closeCallback
            });
        }
        catch (ex) {
            onException(ex, asyncMethodCall, suppliedArguments, callArgs);
        }
        if (closeStatus != OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess) {
            throw OSF.OUtil.formatString(Strings.OfficeOM.L_FunctionCallFailed, OSF.DDA.AsyncMethodNames.CloseAsync.displayName, closeStatus);
        }
    };
    this[OSF.DDA.DispIdHost.Methods.MessageParent] = function OSF_DDA_DispIdHost_Facade$MessageParent(suppliedArguments, caller) {
        var stateInfo = {};
        var syncMethodCall = OSF.DDA.SyncMethodCalls[OSF.DDA.SyncMethodNames.MessageParent.id];
        var callArgs = syncMethodCall.verifyAndExtractCall(suppliedArguments, caller, stateInfo);
        var delegate = getDelegateMethods(OSF.DDA.SyncMethodNames.MessageParent.id);
        var invoker = delegate[OSF.DDA.DispIdHost.Delegates.MessageParent];
        var dispId = dispIdMap[OSF.DDA.SyncMethodNames.MessageParent.id];
        return invoker({
            "dispId": dispId,
            "hostCallArgs": callArgs,
            "onCalling": function OSF_DDA_DispIdFacade$Execute_onCalling() { OSF.OUtil.writeProfilerMark(OSF.HostCallPerfMarker.IssueCall); },
            "onReceiving": function OSF_DDA_DispIdFacade$Execute_onReceiving() { OSF.OUtil.writeProfilerMark(OSF.HostCallPerfMarker.ReceiveResponse); }
        });
    };
    this[OSF.DDA.DispIdHost.Methods.SendMessage] = function OSF_DDA_DispIdHost_Facade$SendMessage(suppliedArguments, eventDispatch, caller) {
        var stateInfo = {};
        var syncMethodCall = OSF.DDA.SyncMethodCalls[OSF.DDA.SyncMethodNames.SendMessage.id];
        var callArgs = syncMethodCall.verifyAndExtractCall(suppliedArguments, caller, stateInfo);
        var delegate = getDelegateMethods(OSF.DDA.SyncMethodNames.SendMessage.id);
        var invoker = delegate[OSF.DDA.DispIdHost.Delegates.SendMessage];
        var dispId = dispIdMap[OSF.DDA.SyncMethodNames.SendMessage.id];
        return invoker({
            "dispId": dispId,
            "hostCallArgs": callArgs,
            "onCalling": function OSF_DDA_DispIdFacade$Execute_onCalling() { OSF.OUtil.writeProfilerMark(OSF.HostCallPerfMarker.IssueCall); },
            "onReceiving": function OSF_DDA_DispIdFacade$Execute_onReceiving() { OSF.OUtil.writeProfilerMark(OSF.HostCallPerfMarker.ReceiveResponse); }
        });
    };
};
OSF.DDA.DispIdHost.addAsyncMethods = function OSF_DDA_DispIdHost$AddAsyncMethods(target, asyncMethodNames, privateState) {
    for (var entry in asyncMethodNames) {
        var method = asyncMethodNames[entry];
        var name = method.displayName;
        if (!target[name]) {
            OSF.OUtil.defineEnumerableProperty(target, name, {
                value: (function (asyncMethod) {
                    return function () {
                        var invokeMethod = OSF._OfficeAppFactory.getHostFacade()[OSF.DDA.DispIdHost.Methods.InvokeMethod];
                        invokeMethod(asyncMethod, arguments, target, privateState);
                    };
                })(method)
            });
        }
    }
};
OSF.DDA.DispIdHost.addEventSupport = function OSF_DDA_DispIdHost$AddEventSupport(target, eventDispatch, isPopupWindow) {
    var add = OSF.DDA.AsyncMethodNames.AddHandlerAsync.displayName;
    var remove = OSF.DDA.AsyncMethodNames.RemoveHandlerAsync.displayName;
    if (!target[add]) {
        OSF.OUtil.defineEnumerableProperty(target, add, {
            value: function () {
                var addEventHandler = OSF._OfficeAppFactory.getHostFacade()[OSF.DDA.DispIdHost.Methods.AddEventHandler];
                addEventHandler(arguments, eventDispatch, target, isPopupWindow);
            }
        });
    }
    if (!target[remove]) {
        OSF.OUtil.defineEnumerableProperty(target, remove, {
            value: function () {
                var removeEventHandler = OSF._OfficeAppFactory.getHostFacade()[OSF.DDA.DispIdHost.Methods.RemoveEventHandler];
                removeEventHandler(arguments, eventDispatch, target);
            }
        });
    }
};
var OfficeExt;
(function (OfficeExt) {
    var MsAjaxTypeHelper = (function () {
        function MsAjaxTypeHelper() {
        }
        MsAjaxTypeHelper.isInstanceOfType = function (type, instance) {
            if (typeof (instance) === "undefined" || instance === null)
                return false;
            if (instance instanceof type)
                return true;
            var instanceType = instance.constructor;
            if (!instanceType || (typeof (instanceType) !== "function") || !instanceType.__typeName || instanceType.__typeName === 'Object') {
                instanceType = Object;
            }
            return !!(instanceType === type) ||
                (instanceType.__typeName && type.__typeName && instanceType.__typeName === type.__typeName);
        };
        return MsAjaxTypeHelper;
    })();
    OfficeExt.MsAjaxTypeHelper = MsAjaxTypeHelper;
    var MsAjaxError = (function () {
        function MsAjaxError() {
        }
        MsAjaxError.create = function (message, errorInfo) {
            var err = new Error(message);
            err.message = message;
            if (errorInfo) {
                for (var v in errorInfo) {
                    err[v] = errorInfo[v];
                }
            }
            err.popStackFrame();
            return err;
        };
        MsAjaxError.parameterCount = function (message) {
            var displayMessage = "Sys.ParameterCountException: " + (message ? message : "Parameter count mismatch.");
            var err = MsAjaxError.create(displayMessage, { name: 'Sys.ParameterCountException' });
            err.popStackFrame();
            return err;
        };
        MsAjaxError.argument = function (paramName, message) {
            var displayMessage = "Sys.ArgumentException: " + (message ? message : "Value does not fall within the expected range.");
            if (paramName) {
                displayMessage += "\n" + MsAjaxString.format("Parameter name: {0}", paramName);
            }
            var err = MsAjaxError.create(displayMessage, { name: "Sys.ArgumentException", paramName: paramName });
            err.popStackFrame();
            return err;
        };
        MsAjaxError.argumentNull = function (paramName, message) {
            var displayMessage = "Sys.ArgumentNullException: " + (message ? message : "Value cannot be null.");
            if (paramName) {
                displayMessage += "\n" + MsAjaxString.format("Parameter name: {0}", paramName);
            }
            var err = MsAjaxError.create(displayMessage, { name: "Sys.ArgumentNullException", paramName: paramName });
            err.popStackFrame();
            return err;
        };
        MsAjaxError.argumentOutOfRange = function (paramName, actualValue, message) {
            var displayMessage = "Sys.ArgumentOutOfRangeException: " + (message ? message : "Specified argument was out of the range of valid values.");
            if (paramName) {
                displayMessage += "\n" + MsAjaxString.format("Parameter name: {0}", paramName);
            }
            if (typeof (actualValue) !== "undefined" && actualValue !== null) {
                displayMessage += "\n" + MsAjaxString.format("Actual value was {0}.", actualValue);
            }
            var err = MsAjaxError.create(displayMessage, {
                name: "Sys.ArgumentOutOfRangeException",
                paramName: paramName,
                actualValue: actualValue
            });
            err.popStackFrame();
            return err;
        };
        MsAjaxError.argumentType = function (paramName, actualType, expectedType, message) {
            var displayMessage = "Sys.ArgumentTypeException: ";
            if (message) {
                displayMessage += message;
            }
            else if (actualType && expectedType) {
                displayMessage += MsAjaxString.format("Object of type '{0}' cannot be converted to type '{1}'.", actualType.getName ? actualType.getName() : actualType, expectedType.getName ? expectedType.getName() : expectedType);
            }
            else {
                displayMessage += "Object cannot be converted to the required type.";
            }
            if (paramName) {
                displayMessage += "\n" + MsAjaxString.format("Parameter name: {0}", paramName);
            }
            var err = MsAjaxError.create(displayMessage, {
                name: "Sys.ArgumentTypeException",
                paramName: paramName,
                actualType: actualType,
                expectedType: expectedType
            });
            err.popStackFrame();
            return err;
        };
        MsAjaxError.argumentUndefined = function (paramName, message) {
            var displayMessage = "Sys.ArgumentUndefinedException: " + (message ? message : "Value cannot be undefined.");
            if (paramName) {
                displayMessage += "\n" + MsAjaxString.format("Parameter name: {0}", paramName);
            }
            var err = MsAjaxError.create(displayMessage, { name: "Sys.ArgumentUndefinedException", paramName: paramName });
            err.popStackFrame();
            return err;
        };
        MsAjaxError.invalidOperation = function (message) {
            var displayMessage = "Sys.InvalidOperationException: " + (message ? message : "Operation is not valid due to the current state of the object.");
            var err = MsAjaxError.create(displayMessage, { name: 'Sys.InvalidOperationException' });
            err.popStackFrame();
            return err;
        };
        return MsAjaxError;
    })();
    OfficeExt.MsAjaxError = MsAjaxError;
    var MsAjaxString = (function () {
        function MsAjaxString() {
        }
        MsAjaxString.format = function (format) {
            var args = [];
            for (var _i = 1; _i < arguments.length; _i++) {
                args[_i - 1] = arguments[_i];
            }
            var source = format;
            return source.replace(/{(\d+)}/gm, function (match, number) {
                var index = parseInt(number, 10);
                return args[index] === undefined ? '{' + number + '}' : args[index];
            });
        };
        MsAjaxString.startsWith = function (str, prefix) {
            return (str.substr(0, prefix.length) === prefix);
        };
        return MsAjaxString;
    })();
    OfficeExt.MsAjaxString = MsAjaxString;
    var MsAjaxDebug = (function () {
        function MsAjaxDebug() {
        }
        MsAjaxDebug.trace = function (text) {
            if (typeof Debug !== "undefined" && Debug.writeln)
                Debug.writeln(text);
            if (window.console && window.console.log)
                window.console.log(text);
            if (window.opera && window.opera.postError)
                window.opera.postError(text);
            if (window.debugService && window.debugService.trace)
                window.debugService.trace(text);
            var a = document.getElementById("TraceConsole");
            if (a && a.tagName.toUpperCase() === "TEXTAREA") {
                a.innerHTML += text + "\n";
            }
        };
        return MsAjaxDebug;
    })();
    OfficeExt.MsAjaxDebug = MsAjaxDebug;
    if (!OsfMsAjaxFactory.isMsAjaxLoaded()) {
        var registerTypeInternal = function registerTypeInternal(type, name, isClass) {
            if (type.__typeName === undefined || type.__typeName === null) {
                type.__typeName = name;
            }
            if (type.__class === undefined || type.__class === null) {
                type.__class = isClass;
            }
        };
        registerTypeInternal(Function, "Function", true);
        registerTypeInternal(Error, "Error", true);
        registerTypeInternal(Object, "Object", true);
        registerTypeInternal(String, "String", true);
        registerTypeInternal(Boolean, "Boolean", true);
        registerTypeInternal(Date, "Date", true);
        registerTypeInternal(Number, "Number", true);
        registerTypeInternal(RegExp, "RegExp", true);
        registerTypeInternal(Array, "Array", true);
        if (!Function.createCallback) {
            Function.createCallback = function Function$createCallback(method, context) {
                var e = Function._validateParams(arguments, [
                    { name: "method", type: Function },
                    { name: "context", mayBeNull: true }
                ]);
                if (e)
                    throw e;
                return function () {
                    var l = arguments.length;
                    if (l > 0) {
                        var args = [];
                        for (var i = 0; i < l; i++) {
                            args[i] = arguments[i];
                        }
                        args[l] = context;
                        return method.apply(this, args);
                    }
                    return method.call(this, context);
                };
            };
        }
        if (!Function.createDelegate) {
            Function.createDelegate = function Function$createDelegate(instance, method) {
                var e = Function._validateParams(arguments, [
                    { name: "instance", mayBeNull: true },
                    { name: "method", type: Function }
                ]);
                if (e)
                    throw e;
                return function () {
                    return method.apply(instance, arguments);
                };
            };
        }
        if (!Function._validateParams) {
            Function._validateParams = function (params, expectedParams, validateParameterCount) {
                var e, expectedLength = expectedParams.length;
                validateParameterCount = validateParameterCount || (typeof (validateParameterCount) === "undefined");
                e = Function._validateParameterCount(params, expectedParams, validateParameterCount);
                if (e) {
                    e.popStackFrame();
                    return e;
                }
                for (var i = 0, l = params.length; i < l; i++) {
                    var expectedParam = expectedParams[Math.min(i, expectedLength - 1)], paramName = expectedParam.name;
                    if (expectedParam.parameterArray) {
                        paramName += "[" + (i - expectedLength + 1) + "]";
                    }
                    else if (!validateParameterCount && (i >= expectedLength)) {
                        break;
                    }
                    e = Function._validateParameter(params[i], expectedParam, paramName);
                    if (e) {
                        e.popStackFrame();
                        return e;
                    }
                }
                return null;
            };
        }
        if (!Function._validateParameterCount) {
            Function._validateParameterCount = function (params, expectedParams, validateParameterCount) {
                var i, error, expectedLen = expectedParams.length, actualLen = params.length;
                if (actualLen < expectedLen) {
                    var minParams = expectedLen;
                    for (i = 0; i < expectedLen; i++) {
                        var param = expectedParams[i];
                        if (param.optional || param.parameterArray) {
                            minParams--;
                        }
                    }
                    if (actualLen < minParams) {
                        error = true;
                    }
                }
                else if (validateParameterCount && (actualLen > expectedLen)) {
                    error = true;
                    for (i = 0; i < expectedLen; i++) {
                        if (expectedParams[i].parameterArray) {
                            error = false;
                            break;
                        }
                    }
                }
                if (error) {
                    var e = MsAjaxError.parameterCount();
                    e.popStackFrame();
                    return e;
                }
                return null;
            };
        }
        if (!Function._validateParameter) {
            Function._validateParameter = function (param, expectedParam, paramName) {
                var e, expectedType = expectedParam.type, expectedInteger = !!expectedParam.integer, expectedDomElement = !!expectedParam.domElement, mayBeNull = !!expectedParam.mayBeNull;
                e = Function._validateParameterType(param, expectedType, expectedInteger, expectedDomElement, mayBeNull, paramName);
                if (e) {
                    e.popStackFrame();
                    return e;
                }
                var expectedElementType = expectedParam.elementType, elementMayBeNull = !!expectedParam.elementMayBeNull;
                if (expectedType === Array && typeof (param) !== "undefined" && param !== null &&
                    (expectedElementType || !elementMayBeNull)) {
                    var expectedElementInteger = !!expectedParam.elementInteger, expectedElementDomElement = !!expectedParam.elementDomElement;
                    for (var i = 0; i < param.length; i++) {
                        var elem = param[i];
                        e = Function._validateParameterType(elem, expectedElementType, expectedElementInteger, expectedElementDomElement, elementMayBeNull, paramName + "[" + i + "]");
                        if (e) {
                            e.popStackFrame();
                            return e;
                        }
                    }
                }
                return null;
            };
        }
        if (!Function._validateParameterType) {
            Function._validateParameterType = function (param, expectedType, expectedInteger, expectedDomElement, mayBeNull, paramName) {
                var e, i;
                if (typeof (param) === "undefined") {
                    if (mayBeNull) {
                        return null;
                    }
                    else {
                        e = OfficeExt.MsAjaxError.argumentUndefined(paramName);
                        e.popStackFrame();
                        return e;
                    }
                }
                if (param === null) {
                    if (mayBeNull) {
                        return null;
                    }
                    else {
                        e = OfficeExt.MsAjaxError.argumentNull(paramName);
                        e.popStackFrame();
                        return e;
                    }
                }
                if (expectedType && !OfficeExt.MsAjaxTypeHelper.isInstanceOfType(expectedType, param)) {
                    e = OfficeExt.MsAjaxError.argumentType(paramName, typeof (param), expectedType);
                    e.popStackFrame();
                    return e;
                }
                return null;
            };
        }
        if (!window.Type) {
            window.Type = Function;
        }
        if (!Type.registerNamespace) {
            Type.registerNamespace = function (ns) {
                var namespaceParts = ns.split('.');
                var currentNamespace = window;
                for (var i = 0; i < namespaceParts.length; i++) {
                    currentNamespace[namespaceParts[i]] = currentNamespace[namespaceParts[i]] || {};
                    currentNamespace = currentNamespace[namespaceParts[i]];
                }
            };
        }
        if (!Type.prototype.registerClass) {
            Type.prototype.registerClass = function (cls) { cls = {}; };
        }
        if (typeof (Sys) === "undefined") {
            Type.registerNamespace('Sys');
        }
        if (!Error.prototype.popStackFrame) {
            Error.prototype.popStackFrame = function () {
                if (arguments.length !== 0)
                    throw MsAjaxError.parameterCount();
                if (typeof (this.stack) === "undefined" || this.stack === null ||
                    typeof (this.fileName) === "undefined" || this.fileName === null ||
                    typeof (this.lineNumber) === "undefined" || this.lineNumber === null) {
                    return;
                }
                var stackFrames = this.stack.split("\n");
                var currentFrame = stackFrames[0];
                var pattern = this.fileName + ":" + this.lineNumber;
                while (typeof (currentFrame) !== "undefined" &&
                    currentFrame !== null &&
                    currentFrame.indexOf(pattern) === -1) {
                    stackFrames.shift();
                    currentFrame = stackFrames[0];
                }
                var nextFrame = stackFrames[1];
                if (typeof (nextFrame) === "undefined" || nextFrame === null) {
                    return;
                }
                var nextFrameParts = nextFrame.match(/@(.*):(\d+)$/);
                if (typeof (nextFrameParts) === "undefined" || nextFrameParts === null) {
                    return;
                }
                this.fileName = nextFrameParts[1];
                this.lineNumber = parseInt(nextFrameParts[2]);
                stackFrames.shift();
                this.stack = stackFrames.join("\n");
            };
        }
        OsfMsAjaxFactory.msAjaxError = MsAjaxError;
        OsfMsAjaxFactory.msAjaxString = MsAjaxString;
        OsfMsAjaxFactory.msAjaxDebug = MsAjaxDebug;
    }
})(OfficeExt || (OfficeExt = {}));
OSF.OUtil.setNamespace("SafeArray", OSF.DDA);
OSF.DDA.SafeArray.Response = {
    Status: 0,
    Payload: 1
};
OSF.DDA.SafeArray.UniqueArguments = {
    Offset: "offset",
    Run: "run",
    BindingSpecificData: "bindingSpecificData",
    MergedCellGuid: "{66e7831f-81b2-42e2-823c-89e872d541b3}"
};
OSF.OUtil.setNamespace("Delegate", OSF.DDA.SafeArray);
OSF.DDA.SafeArray.Delegate._onException = function OSF_DDA_SafeArray_Delegate$OnException(ex, args) {
    var status;
    var statusNumber = ex.number;
    if (statusNumber) {
        switch (statusNumber) {
            case -2146828218:
                status = OSF.DDA.ErrorCodeManager.errorCodes.ooeNoCapability;
                break;
            case -2147467259:
                status = OSF.DDA.ErrorCodeManager.errorCodes.ooeDialogAlreadyOpened;
                break;
            case -2146828283:
                status = OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidParam;
                break;
            case -2147209089:
                status = OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidParam;
                break;
            case -2147208704:
                status = OSF.DDA.ErrorCodeManager.errorCodes.ooeTooManyIncompleteRequests;
                break;
            case -2146827850:
            default:
                status = OSF.DDA.ErrorCodeManager.errorCodes.ooeInternalError;
                break;
        }
    }
    if (args.onComplete) {
        args.onComplete(status || OSF.DDA.ErrorCodeManager.errorCodes.ooeInternalError);
    }
};
OSF.DDA.SafeArray.Delegate._onExceptionSyncMethod = function OSF_DDA_SafeArray_Delegate$OnExceptionSyncMethod(ex, args) {
    var status;
    var number = ex.number;
    if (number) {
        switch (number) {
            case -2146828218:
                status = OSF.DDA.ErrorCodeManager.errorCodes.ooeNoCapability;
                break;
            case -2146827850:
            default:
                status = OSF.DDA.ErrorCodeManager.errorCodes.ooeInternalError;
                break;
        }
    }
    return status || OSF.DDA.ErrorCodeManager.errorCodes.ooeInternalError;
};
OSF.DDA.SafeArray.Delegate.SpecialProcessor = function OSF_DDA_SafeArray_Delegate_SpecialProcessor() {
    function _2DVBArrayToJaggedArray(vbArr) {
        var ret;
        try {
            var rows = vbArr.ubound(1);
            var cols = vbArr.ubound(2);
            vbArr = vbArr.toArray();
            if (rows == 1 && cols == 1) {
                ret = [vbArr];
            }
            else {
                ret = [];
                for (var row = 0; row < rows; row++) {
                    var rowArr = [];
                    for (var col = 0; col < cols; col++) {
                        var datum = vbArr[row * cols + col];
                        if (datum != OSF.DDA.SafeArray.UniqueArguments.MergedCellGuid) {
                            rowArr.push(datum);
                        }
                    }
                    if (rowArr.length > 0) {
                        ret.push(rowArr);
                    }
                }
            }
        }
        catch (ex) {
        }
        return ret;
    }
    var complexTypes = [];
    var dynamicTypes = {};
    dynamicTypes[Microsoft.Office.WebExtension.Parameters.Data] = (function () {
        var tableRows = 0;
        var tableHeaders = 1;
        return {
            toHost: function OSF_DDA_SafeArray_Delegate_SpecialProcessor_Data$toHost(data) {
                if (OSF.DDA.TableDataProperties && typeof data != "string" && data[OSF.DDA.TableDataProperties.TableRows] !== undefined) {
                    var tableData = [];
                    tableData[tableRows] = data[OSF.DDA.TableDataProperties.TableRows];
                    tableData[tableHeaders] = data[OSF.DDA.TableDataProperties.TableHeaders];
                    data = tableData;
                }
                return data;
            },
            fromHost: function OSF_DDA_SafeArray_Delegate_SpecialProcessor_Data$fromHost(hostArgs) {
                var ret;
                if (hostArgs.toArray) {
                    var dimensions = hostArgs.dimensions();
                    if (dimensions === 2) {
                        ret = _2DVBArrayToJaggedArray(hostArgs);
                    }
                    else {
                        var array = hostArgs.toArray();
                        if (array.length === 2 && ((array[0] != null && array[0].toArray) || (array[1] != null && array[1].toArray))) {
                            ret = {};
                            ret[OSF.DDA.TableDataProperties.TableRows] = _2DVBArrayToJaggedArray(array[tableRows]);
                            ret[OSF.DDA.TableDataProperties.TableHeaders] = _2DVBArrayToJaggedArray(array[tableHeaders]);
                        }
                        else {
                            ret = array;
                        }
                    }
                }
                else {
                    ret = hostArgs;
                }
                return ret;
            }
        };
    })();
    OSF.DDA.SafeArray.Delegate.SpecialProcessor.uber.constructor.call(this, complexTypes, dynamicTypes);
    this.unpack = function OSF_DDA_SafeArray_Delegate_SpecialProcessor$unpack(param, arg) {
        var value;
        if (this.isComplexType(param) || OSF.DDA.ListType.isListType(param)) {
            var toArraySupported = (arg || typeof arg === "unknown") && arg.toArray;
            value = toArraySupported ? arg.toArray() : arg || {};
        }
        else if (this.isDynamicType(param)) {
            value = dynamicTypes[param].fromHost(arg);
        }
        else {
            value = arg;
        }
        return value;
    };
};
OSF.OUtil.extend(OSF.DDA.SafeArray.Delegate.SpecialProcessor, OSF.DDA.SpecialProcessor);
OSF.DDA.SafeArray.Delegate.ParameterMap = OSF.DDA.getDecoratedParameterMap(new OSF.DDA.SafeArray.Delegate.SpecialProcessor(), [
    {
        type: Microsoft.Office.WebExtension.Parameters.ValueFormat,
        toHost: [
            { name: Microsoft.Office.WebExtension.ValueFormat.Unformatted, value: 0 },
            { name: Microsoft.Office.WebExtension.ValueFormat.Formatted, value: 1 }
        ]
    },
    {
        type: Microsoft.Office.WebExtension.Parameters.FilterType,
        toHost: [
            { name: Microsoft.Office.WebExtension.FilterType.All, value: 0 }
        ]
    }
]);
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
    type: OSF.DDA.PropertyDescriptors.AsyncResultStatus,
    fromHost: [
        { name: Microsoft.Office.WebExtension.AsyncResultStatus.Succeeded, value: 0 },
        { name: Microsoft.Office.WebExtension.AsyncResultStatus.Failed, value: 1 }
    ]
});
OSF.DDA.SafeArray.Delegate.executeAsync = function OSF_DDA_SafeArray_Delegate$ExecuteAsync(args) {
    function toArray(args) {
        var arrArgs = args;
        if (OSF.OUtil.isArray(args)) {
            var len = arrArgs.length;
            for (var i = 0; i < len; i++) {
                arrArgs[i] = toArray(arrArgs[i]);
            }
        }
        else if (OSF.OUtil.isDate(args)) {
            arrArgs = args.getVarDate();
        }
        else if (typeof args === "object" && !OSF.OUtil.isArray(args)) {
            arrArgs = [];
            for (var index in args) {
                if (!OSF.OUtil.isFunction(args[index])) {
                    arrArgs[index] = toArray(args[index]);
                }
            }
        }
        return arrArgs;
    }
    function fromSafeArray(value) {
        var ret = value;
        if (value != null && value.toArray) {
            var arrayResult = value.toArray();
            ret = new Array(arrayResult.length);
            for (var i = 0; i < arrayResult.length; i++) {
                ret[i] = fromSafeArray(arrayResult[i]);
            }
        }
        return ret;
    }
    try {
        if (args.onCalling) {
            args.onCalling();
        }
        OSF.ClientHostController.execute(args.dispId, toArray(args.hostCallArgs), function OSF_DDA_SafeArrayFacade$Execute_OnResponse(hostResponseArgs, resultCode) {
            var result = hostResponseArgs.toArray();
            var status = result[OSF.DDA.SafeArray.Response.Status];
            if (status == OSF.DDA.ErrorCodeManager.errorCodes.ooeChunkResult) {
                var payload = result[OSF.DDA.SafeArray.Response.Payload];
                payload = fromSafeArray(payload);
                if (payload != null) {
                    if (!args._chunkResultData) {
                        args._chunkResultData = new Array();
                    }
                    args._chunkResultData[payload[0]] = payload[1];
                }
                return false;
            }
            if (args.onReceiving) {
                args.onReceiving();
            }
            if (args.onComplete) {
                var payload;
                if (status == OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess) {
                    if (result.length > 2) {
                        payload = [];
                        for (var i = 1; i < result.length; i++)
                            payload[i - 1] = result[i];
                    }
                    else {
                        payload = result[OSF.DDA.SafeArray.Response.Payload];
                    }
                    if (args._chunkResultData) {
                        payload = fromSafeArray(payload);
                        if (payload != null) {
                            var expectedChunkCount = payload[payload.length - 1];
                            if (args._chunkResultData.length == expectedChunkCount) {
                                payload[payload.length - 1] = args._chunkResultData;
                            }
                            else {
                                status = OSF.DDA.ErrorCodeManager.errorCodes.ooeInternalError;
                            }
                        }
                    }
                }
                else {
                    payload = result[OSF.DDA.SafeArray.Response.Payload];
                }
                args.onComplete(status, payload);
            }
            return true;
        });
    }
    catch (ex) {
        OSF.DDA.SafeArray.Delegate._onException(ex, args);
    }
};
OSF.DDA.SafeArray.Delegate._getOnAfterRegisterEvent = function OSF_DDA_SafeArrayDelegate$GetOnAfterRegisterEvent(register, args) {
    var startTime = (new Date()).getTime();
    return function OSF_DDA_SafeArrayDelegate$OnAfterRegisterEvent(hostResponseArgs) {
        if (args.onReceiving) {
            args.onReceiving();
        }
        var status = hostResponseArgs.toArray ? hostResponseArgs.toArray()[OSF.DDA.SafeArray.Response.Status] : hostResponseArgs;
        if (args.onComplete) {
            args.onComplete(status);
        }
        if (OSF.AppTelemetry) {
            OSF.AppTelemetry.onRegisterDone(register, args.dispId, Math.abs((new Date()).getTime() - startTime), status);
        }
    };
};
OSF.DDA.SafeArray.Delegate.registerEventAsync = function OSF_DDA_SafeArray_Delegate$RegisterEventAsync(args) {
    if (args.onCalling) {
        args.onCalling();
    }
    var callback = OSF.DDA.SafeArray.Delegate._getOnAfterRegisterEvent(true, args);
    try {
        OSF.ClientHostController.registerEvent(args.dispId, args.targetId, function OSF_DDA_SafeArrayDelegate$RegisterEventAsync_OnEvent(eventDispId, payload) {
            if (args.onEvent) {
                args.onEvent(payload);
            }
            if (OSF.AppTelemetry) {
                OSF.AppTelemetry.onEventDone(args.dispId);
            }
        }, callback);
    }
    catch (ex) {
        OSF.DDA.SafeArray.Delegate._onException(ex, args);
    }
};
OSF.DDA.SafeArray.Delegate.unregisterEventAsync = function OSF_DDA_SafeArray_Delegate$UnregisterEventAsync(args) {
    if (args.onCalling) {
        args.onCalling();
    }
    var callback = OSF.DDA.SafeArray.Delegate._getOnAfterRegisterEvent(false, args);
    try {
        OSF.ClientHostController.unregisterEvent(args.dispId, args.targetId, callback);
    }
    catch (ex) {
        OSF.DDA.SafeArray.Delegate._onException(ex, args);
    }
};
OSF.ClientMode = {
    ReadWrite: 0,
    ReadOnly: 1
};
OSF.DDA.RichInitializationReason = {
    1: Microsoft.Office.WebExtension.InitializationReason.Inserted,
    2: Microsoft.Office.WebExtension.InitializationReason.DocumentOpened
};
OSF.InitializationHelper = function OSF_InitializationHelper(hostInfo, webAppState, context, settings, hostFacade) {
    this._hostInfo = hostInfo;
    this._webAppState = webAppState;
    this._context = context;
    this._settings = settings;
    this._hostFacade = hostFacade;
    this._initializeSettings = this.initializeSettings;
};
OSF.InitializationHelper.prototype.deserializeSettings = function OSF_InitializationHelper$deserializeSettings(serializedSettings, refreshSupported) {
    var settings;
    var osfSessionStorage = OSF.OUtil.getSessionStorage();
    if (osfSessionStorage) {
        var storageSettings = osfSessionStorage.getItem(OSF._OfficeAppFactory.getCachedSessionSettingsKey());
        if (storageSettings) {
            serializedSettings = JSON.parse(storageSettings);
        }
        else {
            storageSettings = JSON.stringify(serializedSettings);
            osfSessionStorage.setItem(OSF._OfficeAppFactory.getCachedSessionSettingsKey(), storageSettings);
        }
    }
    var deserializedSettings = OSF.DDA.SettingsManager.deserializeSettings(serializedSettings);
    if (refreshSupported) {
        settings = new OSF.DDA.RefreshableSettings(deserializedSettings);
    }
    else {
        settings = new OSF.DDA.Settings(deserializedSettings);
    }
    return settings;
};
OSF.InitializationHelper.prototype.saveAndSetDialogInfo = function OSF_InitializationHelper$saveAndSetDialogInfo(hostInfoValue) {
};
OSF.InitializationHelper.prototype.setAgaveHostCommunication = function OSF_InitializationHelper$setAgaveHostCommunication() {
};
OSF.InitializationHelper.prototype.prepareRightBeforeWebExtensionInitialize = function OSF_InitializationHelper$prepareRightBeforeWebExtensionInitialize(appContext) {
    this.prepareApiSurface(appContext);
    Microsoft.Office.WebExtension.initialize(this.getInitializationReason(appContext));
};
OSF.InitializationHelper.prototype.prepareApiSurface = function OSF_InitializationHelper$prepareApiSurfaceAndInitialize(appContext) {
    var license = new OSF.DDA.License(appContext.get_eToken());
    var getOfficeThemeHandler = (OSF.DDA.OfficeTheme && OSF.DDA.OfficeTheme.getOfficeTheme) ? OSF.DDA.OfficeTheme.getOfficeTheme : null;
    if (appContext.get_isDialog()) {
        if (OSF.DDA.UI.ChildUI) {
            appContext.ui = new OSF.DDA.UI.ChildUI();
        }
    }
    else {
        if (OSF.DDA.UI.ParentUI) {
            appContext.ui = new OSF.DDA.UI.ParentUI();
            if (OfficeExt.Container) {
                OSF.DDA.DispIdHost.addAsyncMethods(appContext.ui, [OSF.DDA.AsyncMethodNames.CloseContainerAsync]);
            }
        }
    }
    if (OSF.DDA.OpenBrowser) {
        OSF.DDA.DispIdHost.addAsyncMethods(appContext.ui, [OSF.DDA.AsyncMethodNames.OpenBrowserWindow]);
    }
    if (OSF.DDA.ExecuteFeature) {
        OSF.DDA.DispIdHost.addAsyncMethods(appContext.ui, [OSF.DDA.AsyncMethodNames.ExecuteFeature]);
    }
    if (OSF.DDA.QueryFeature) {
        OSF.DDA.DispIdHost.addAsyncMethods(appContext.ui, [OSF.DDA.AsyncMethodNames.QueryFeature]);
    }
    if (OSF.DDA.Auth) {
        appContext.auth = new OSF.DDA.Auth();
        OSF.DDA.DispIdHost.addAsyncMethods(appContext.auth, [OSF.DDA.AsyncMethodNames.GetAccessTokenAsync]);
    }
    OSF._OfficeAppFactory.setContext(new OSF.DDA.Context(appContext, appContext.doc, license, null, getOfficeThemeHandler));
    var getDelegateMethods, parameterMap;
    getDelegateMethods = OSF.DDA.DispIdHost.getClientDelegateMethods;
    parameterMap = OSF.DDA.SafeArray.Delegate.ParameterMap;
    OSF._OfficeAppFactory.setHostFacade(new OSF.DDA.DispIdHost.Facade(getDelegateMethods, parameterMap));
};
OSF.InitializationHelper.prototype.getInitializationReason = function (appContext) { return OSF.DDA.RichInitializationReason[appContext.get_reason()]; };
OSF.DDA.DispIdHost.getClientDelegateMethods = function (actionId) {
    var delegateMethods = {};
    delegateMethods[OSF.DDA.DispIdHost.Delegates.ExecuteAsync] = OSF.DDA.SafeArray.Delegate.executeAsync;
    delegateMethods[OSF.DDA.DispIdHost.Delegates.RegisterEventAsync] = OSF.DDA.SafeArray.Delegate.registerEventAsync;
    delegateMethods[OSF.DDA.DispIdHost.Delegates.UnregisterEventAsync] = OSF.DDA.SafeArray.Delegate.unregisterEventAsync;
    delegateMethods[OSF.DDA.DispIdHost.Delegates.OpenDialog] = OSF.DDA.SafeArray.Delegate.openDialog;
    delegateMethods[OSF.DDA.DispIdHost.Delegates.CloseDialog] = OSF.DDA.SafeArray.Delegate.closeDialog;
    delegateMethods[OSF.DDA.DispIdHost.Delegates.MessageParent] = OSF.DDA.SafeArray.Delegate.messageParent;
    delegateMethods[OSF.DDA.DispIdHost.Delegates.SendMessage] = OSF.DDA.SafeArray.Delegate.sendMessage;
    if (OSF.DDA.AsyncMethodNames.RefreshAsync && actionId == OSF.DDA.AsyncMethodNames.RefreshAsync.id) {
        var readSerializedSettings = function (hostCallArgs, onCalling, onReceiving) {
            return OSF.DDA.ClientSettingsManager.read(onCalling, onReceiving);
        };
        delegateMethods[OSF.DDA.DispIdHost.Delegates.ExecuteAsync] = OSF.DDA.ClientSettingsManager.getSettingsExecuteMethod(readSerializedSettings);
    }
    if (OSF.DDA.AsyncMethodNames.SaveAsync && actionId == OSF.DDA.AsyncMethodNames.SaveAsync.id) {
        var writeSerializedSettings = function (hostCallArgs, onCalling, onReceiving) {
            return OSF.DDA.ClientSettingsManager.write(hostCallArgs[OSF.DDA.SettingsManager.SerializedSettings], hostCallArgs[Microsoft.Office.WebExtension.Parameters.OverwriteIfStale], onCalling, onReceiving);
        };
        delegateMethods[OSF.DDA.DispIdHost.Delegates.ExecuteAsync] = OSF.DDA.ClientSettingsManager.getSettingsExecuteMethod(writeSerializedSettings);
    }
    return delegateMethods;
};
var OfficeExt;
(function (OfficeExt) {
    var RichClientHostController = (function () {
        function RichClientHostController() {
        }
        RichClientHostController.prototype.execute = function (id, params, callback) {
            window.external.Execute(id, params, callback);
        };
        RichClientHostController.prototype.registerEvent = function (id, targetId, handler, callback) {
            window.external.RegisterEvent(id, targetId, handler, callback);
        };
        RichClientHostController.prototype.unregisterEvent = function (id, targetId, callback) {
            window.external.UnregisterEvent(id, targetId, callback);
        };
        return RichClientHostController;
    })();
    OfficeExt.RichClientHostController = RichClientHostController;
})(OfficeExt || (OfficeExt = {}));
var OfficeExt;
(function (OfficeExt) {
    var Win32RichClientHostController = (function (_super) {
        __extends(Win32RichClientHostController, _super);
        function Win32RichClientHostController() {
            _super.apply(this, arguments);
        }
        Win32RichClientHostController.prototype.messageParent = function (params) {
            var message = params[Microsoft.Office.WebExtension.Parameters.MessageToParent];
            window.external.MessageParent(message);
        };
        Win32RichClientHostController.prototype.openDialog = function (id, targetId, handler, callback) {
            this.registerEvent(id, targetId, handler, callback);
        };
        Win32RichClientHostController.prototype.closeDialog = function (id, targetId, callback) {
            this.unregisterEvent(id, targetId, callback);
        };
        Win32RichClientHostController.prototype.sendMessage = function (params) {
        };
        return Win32RichClientHostController;
    })(OfficeExt.RichClientHostController);
    OfficeExt.Win32RichClientHostController = Win32RichClientHostController;
})(OfficeExt || (OfficeExt = {}));
OSF.ClientHostController = new OfficeExt.Win32RichClientHostController();
var OfficeExt;
(function (OfficeExt) {
    var OfficeTheme;
    (function (OfficeTheme) {
        var OfficeThemeManager = (function () {
            function OfficeThemeManager() {
                this._osfOfficeTheme = null;
                this._osfOfficeThemeTimeStamp = null;
            }
            OfficeThemeManager.prototype.getOfficeTheme = function () {
                if (OSF.DDA._OsfControlContext) {
                    if (this._osfOfficeTheme && this._osfOfficeThemeTimeStamp && ((new Date()).getTime() - this._osfOfficeThemeTimeStamp < OfficeThemeManager._osfOfficeThemeCacheValidPeriod)) {
                        if (OSF.AppTelemetry) {
                            OSF.AppTelemetry.onPropertyDone("GetOfficeThemeInfo", 0);
                        }
                    }
                    else {
                        var startTime = (new Date()).getTime();
                        var osfOfficeTheme = OSF.DDA._OsfControlContext.GetOfficeThemeInfo();
                        var endTime = (new Date()).getTime();
                        if (OSF.AppTelemetry) {
                            OSF.AppTelemetry.onPropertyDone("GetOfficeThemeInfo", Math.abs(endTime - startTime));
                        }
                        this._osfOfficeTheme = JSON.parse(osfOfficeTheme);
                        for (var color in this._osfOfficeTheme) {
                            this._osfOfficeTheme[color] = OSF.OUtil.convertIntToCssHexColor(this._osfOfficeTheme[color]);
                        }
                        this._osfOfficeThemeTimeStamp = endTime;
                    }
                    return this._osfOfficeTheme;
                }
            };
            OfficeThemeManager.instance = function () {
                if (OfficeThemeManager._instance == null) {
                    OfficeThemeManager._instance = new OfficeThemeManager();
                }
                return OfficeThemeManager._instance;
            };
            OfficeThemeManager._osfOfficeThemeCacheValidPeriod = 5000;
            OfficeThemeManager._instance = null;
            return OfficeThemeManager;
        })();
        OfficeTheme.OfficeThemeManager = OfficeThemeManager;
        OSF.OUtil.setNamespace("OfficeTheme", OSF.DDA);
        OSF.DDA.OfficeTheme.getOfficeTheme = OfficeExt.OfficeTheme.OfficeThemeManager.instance().getOfficeTheme;
    })(OfficeTheme = OfficeExt.OfficeTheme || (OfficeExt.OfficeTheme = {}));
})(OfficeExt || (OfficeExt = {}));
OSF.DDA.ClientSettingsManager = {
    getSettingsExecuteMethod: function OSF_DDA_ClientSettingsManager$getSettingsExecuteMethod(hostDelegateMethod) {
        return function (args) {
            var status, response;
            try {
                response = hostDelegateMethod(args.hostCallArgs, args.onCalling, args.onReceiving);
                status = OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess;
            }
            catch (ex) {
                status = OSF.DDA.ErrorCodeManager.errorCodes.ooeInternalError;
                response = { name: Strings.OfficeOM.L_InternalError, message: ex };
            }
            if (args.onComplete) {
                args.onComplete(status, response);
            }
        };
    },
    read: function OSF_DDA_ClientSettingsManager$read(onCalling, onReceiving) {
        var keys = [];
        var values = [];
        if (onCalling) {
            onCalling();
        }
        OSF.DDA._OsfControlContext.GetSettings().Read(keys, values);
        if (onReceiving) {
            onReceiving();
        }
        var serializedSettings = {};
        for (var index = 0; index < keys.length; index++) {
            serializedSettings[keys[index]] = values[index];
        }
        return serializedSettings;
    },
    write: function OSF_DDA_ClientSettingsManager$write(serializedSettings, overwriteIfStale, onCalling, onReceiving) {
        var keys = [];
        var values = [];
        for (var key in serializedSettings) {
            keys.push(key);
            values.push(serializedSettings[key]);
        }
        if (onCalling) {
            onCalling();
        }
        OSF.DDA._OsfControlContext.GetSettings().Write(keys, values);
        if (onReceiving) {
            onReceiving();
        }
    }
};
OSF.InitializationHelper.prototype.initializeSettings = function OSF_InitializationHelper$initializeSettings(refreshSupported) {
    var serializedSettings = OSF.DDA.ClientSettingsManager.read();
    var settings = this.deserializeSettings(serializedSettings, refreshSupported);
    return settings;
};
OSF.InitializationHelper.prototype.getAppContext = function OSF_InitializationHelper$getAppContext(wnd, gotAppContext) {
    var returnedContext;
    var context;
    var warningText = "Warning: Office.js is loaded outside of Office client";
    try {
        if (window.external && typeof window.external.GetContext !== 'undefined') {
            context = OSF.DDA._OsfControlContext = window.external.GetContext();
        }
        else {
            OsfMsAjaxFactory.msAjaxDebug.trace(warningText);
            return;
        }
    }
    catch (e) {
        OsfMsAjaxFactory.msAjaxDebug.trace(warningText);
        return;
    }
    var appType = context.GetAppType();
    var id = context.GetSolutionRef();
    var version = context.GetAppVersionMajor();
    var minorVersion = context.GetAppVersionMinor();
    var UILocale = context.GetAppUILocale();
    var dataLocale = context.GetAppDataLocale();
    var docUrl = context.GetDocUrl();
    var clientMode = context.GetAppCapabilities();
    var reason = context.GetActivationMode();
    var osfControlType = context.GetControlIntegrationLevel();
    var settings = [];
    var eToken;
    try {
        eToken = context.GetSolutionToken();
    }
    catch (ex) {
    }
    var correlationId;
    if (typeof context.GetCorrelationId !== "undefined") {
        correlationId = context.GetCorrelationId();
    }
    var appInstanceId;
    if (typeof context.GetInstanceId !== "undefined") {
        appInstanceId = context.GetInstanceId();
    }
    var touchEnabled;
    if (typeof context.GetTouchEnabled !== "undefined") {
        touchEnabled = context.GetTouchEnabled();
    }
    var commerceAllowed;
    if (typeof context.GetCommerceAllowed !== "undefined") {
        commerceAllowed = context.GetCommerceAllowed();
    }
    var requirementMatrix;
    if (typeof context.GetSupportedMatrix !== "undefined") {
        requirementMatrix = context.GetSupportedMatrix();
    }
    var hostCustomMessage;
    if (typeof context.GetHostCustomMessage !== "undefined") {
        hostCustomMessage = context.GetHostCustomMessage();
    }
    var hostFullVersion;
    if (typeof context.GetHostFullVersion !== "undefined") {
        hostFullVersion = context.GetHostFullVersion();
    }
    var dialogRequirementMatrix;
    if (typeof context.GetDialogRequirementMatrix != "undefined") {
        dialogRequirementMatrix = context.GetDialogRequirementMatrix();
    }
    eToken = eToken ? eToken.toString() : "";
    returnedContext = new OSF.OfficeAppContext(id, appType, version, UILocale, dataLocale, docUrl, clientMode, settings, reason, osfControlType, eToken, correlationId, appInstanceId, touchEnabled, commerceAllowed, minorVersion, requirementMatrix, hostCustomMessage, hostFullVersion, undefined, undefined, undefined, dialogRequirementMatrix);
    if (OSF.AppTelemetry) {
        OSF.AppTelemetry.initialize(returnedContext);
    }
    gotAppContext(returnedContext);
};
var OSFLog;
(function (OSFLog) {
    var BaseUsageData = (function () {
        function BaseUsageData(table) {
            this._table = table;
            this._fields = {};
        }
        Object.defineProperty(BaseUsageData.prototype, "Fields", {
            get: function () {
                return this._fields;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(BaseUsageData.prototype, "Table", {
            get: function () {
                return this._table;
            },
            enumerable: true,
            configurable: true
        });
        BaseUsageData.prototype.SerializeFields = function () {
        };
        BaseUsageData.prototype.SetSerializedField = function (key, value) {
            if (typeof (value) !== "undefined" && value !== null) {
                this._serializedFields[key] = value.toString();
            }
        };
        BaseUsageData.prototype.SerializeRow = function () {
            this._serializedFields = {};
            this.SetSerializedField("Table", this._table);
            this.SerializeFields();
            return JSON.stringify(this._serializedFields);
        };
        return BaseUsageData;
    })();
    OSFLog.BaseUsageData = BaseUsageData;
    var AppActivatedUsageData = (function (_super) {
        __extends(AppActivatedUsageData, _super);
        function AppActivatedUsageData() {
            _super.call(this, "AppActivated");
        }
        Object.defineProperty(AppActivatedUsageData.prototype, "CorrelationId", {
            get: function () { return this.Fields["CorrelationId"]; },
            set: function (value) { this.Fields["CorrelationId"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(AppActivatedUsageData.prototype, "SessionId", {
            get: function () { return this.Fields["SessionId"]; },
            set: function (value) { this.Fields["SessionId"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(AppActivatedUsageData.prototype, "AppId", {
            get: function () { return this.Fields["AppId"]; },
            set: function (value) { this.Fields["AppId"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(AppActivatedUsageData.prototype, "AppInstanceId", {
            get: function () { return this.Fields["AppInstanceId"]; },
            set: function (value) { this.Fields["AppInstanceId"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(AppActivatedUsageData.prototype, "AppURL", {
            get: function () { return this.Fields["AppURL"]; },
            set: function (value) { this.Fields["AppURL"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(AppActivatedUsageData.prototype, "AssetId", {
            get: function () { return this.Fields["AssetId"]; },
            set: function (value) { this.Fields["AssetId"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(AppActivatedUsageData.prototype, "Browser", {
            get: function () { return this.Fields["Browser"]; },
            set: function (value) { this.Fields["Browser"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(AppActivatedUsageData.prototype, "UserId", {
            get: function () { return this.Fields["UserId"]; },
            set: function (value) { this.Fields["UserId"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(AppActivatedUsageData.prototype, "Host", {
            get: function () { return this.Fields["Host"]; },
            set: function (value) { this.Fields["Host"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(AppActivatedUsageData.prototype, "HostVersion", {
            get: function () { return this.Fields["HostVersion"]; },
            set: function (value) { this.Fields["HostVersion"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(AppActivatedUsageData.prototype, "ClientId", {
            get: function () { return this.Fields["ClientId"]; },
            set: function (value) { this.Fields["ClientId"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(AppActivatedUsageData.prototype, "AppSizeWidth", {
            get: function () { return this.Fields["AppSizeWidth"]; },
            set: function (value) { this.Fields["AppSizeWidth"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(AppActivatedUsageData.prototype, "AppSizeHeight", {
            get: function () { return this.Fields["AppSizeHeight"]; },
            set: function (value) { this.Fields["AppSizeHeight"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(AppActivatedUsageData.prototype, "Message", {
            get: function () { return this.Fields["Message"]; },
            set: function (value) { this.Fields["Message"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(AppActivatedUsageData.prototype, "DocUrl", {
            get: function () { return this.Fields["DocUrl"]; },
            set: function (value) { this.Fields["DocUrl"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(AppActivatedUsageData.prototype, "OfficeJSVersion", {
            get: function () { return this.Fields["OfficeJSVersion"]; },
            set: function (value) { this.Fields["OfficeJSVersion"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(AppActivatedUsageData.prototype, "HostJSVersion", {
            get: function () { return this.Fields["HostJSVersion"]; },
            set: function (value) { this.Fields["HostJSVersion"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(AppActivatedUsageData.prototype, "WacHostEnvironment", {
            get: function () { return this.Fields["WacHostEnvironment"]; },
            set: function (value) { this.Fields["WacHostEnvironment"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(AppActivatedUsageData.prototype, "IsFromWacAutomation", {
            get: function () { return this.Fields["IsFromWacAutomation"]; },
            set: function (value) { this.Fields["IsFromWacAutomation"] = value; },
            enumerable: true,
            configurable: true
        });
        AppActivatedUsageData.prototype.SerializeFields = function () {
            this.SetSerializedField("CorrelationId", this.CorrelationId);
            this.SetSerializedField("SessionId", this.SessionId);
            this.SetSerializedField("AppId", this.AppId);
            this.SetSerializedField("AppInstanceId", this.AppInstanceId);
            this.SetSerializedField("AppURL", this.AppURL);
            this.SetSerializedField("AssetId", this.AssetId);
            this.SetSerializedField("Browser", this.Browser);
            this.SetSerializedField("UserId", this.UserId);
            this.SetSerializedField("Host", this.Host);
            this.SetSerializedField("HostVersion", this.HostVersion);
            this.SetSerializedField("ClientId", this.ClientId);
            this.SetSerializedField("AppSizeWidth", this.AppSizeWidth);
            this.SetSerializedField("AppSizeHeight", this.AppSizeHeight);
            this.SetSerializedField("Message", this.Message);
            this.SetSerializedField("DocUrl", this.DocUrl);
            this.SetSerializedField("OfficeJSVersion", this.OfficeJSVersion);
            this.SetSerializedField("HostJSVersion", this.HostJSVersion);
            this.SetSerializedField("WacHostEnvironment", this.WacHostEnvironment);
            this.SetSerializedField("IsFromWacAutomation", this.IsFromWacAutomation);
        };
        return AppActivatedUsageData;
    })(BaseUsageData);
    OSFLog.AppActivatedUsageData = AppActivatedUsageData;
    var ScriptLoadUsageData = (function (_super) {
        __extends(ScriptLoadUsageData, _super);
        function ScriptLoadUsageData() {
            _super.call(this, "ScriptLoad");
        }
        Object.defineProperty(ScriptLoadUsageData.prototype, "CorrelationId", {
            get: function () { return this.Fields["CorrelationId"]; },
            set: function (value) { this.Fields["CorrelationId"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ScriptLoadUsageData.prototype, "SessionId", {
            get: function () { return this.Fields["SessionId"]; },
            set: function (value) { this.Fields["SessionId"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ScriptLoadUsageData.prototype, "ScriptId", {
            get: function () { return this.Fields["ScriptId"]; },
            set: function (value) { this.Fields["ScriptId"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ScriptLoadUsageData.prototype, "StartTime", {
            get: function () { return this.Fields["StartTime"]; },
            set: function (value) { this.Fields["StartTime"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ScriptLoadUsageData.prototype, "ResponseTime", {
            get: function () { return this.Fields["ResponseTime"]; },
            set: function (value) { this.Fields["ResponseTime"] = value; },
            enumerable: true,
            configurable: true
        });
        ScriptLoadUsageData.prototype.SerializeFields = function () {
            this.SetSerializedField("CorrelationId", this.CorrelationId);
            this.SetSerializedField("SessionId", this.SessionId);
            this.SetSerializedField("ScriptId", this.ScriptId);
            this.SetSerializedField("StartTime", this.StartTime);
            this.SetSerializedField("ResponseTime", this.ResponseTime);
        };
        return ScriptLoadUsageData;
    })(BaseUsageData);
    OSFLog.ScriptLoadUsageData = ScriptLoadUsageData;
    var AppClosedUsageData = (function (_super) {
        __extends(AppClosedUsageData, _super);
        function AppClosedUsageData() {
            _super.call(this, "AppClosed");
        }
        Object.defineProperty(AppClosedUsageData.prototype, "CorrelationId", {
            get: function () { return this.Fields["CorrelationId"]; },
            set: function (value) { this.Fields["CorrelationId"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(AppClosedUsageData.prototype, "SessionId", {
            get: function () { return this.Fields["SessionId"]; },
            set: function (value) { this.Fields["SessionId"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(AppClosedUsageData.prototype, "FocusTime", {
            get: function () { return this.Fields["FocusTime"]; },
            set: function (value) { this.Fields["FocusTime"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(AppClosedUsageData.prototype, "AppSizeFinalWidth", {
            get: function () { return this.Fields["AppSizeFinalWidth"]; },
            set: function (value) { this.Fields["AppSizeFinalWidth"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(AppClosedUsageData.prototype, "AppSizeFinalHeight", {
            get: function () { return this.Fields["AppSizeFinalHeight"]; },
            set: function (value) { this.Fields["AppSizeFinalHeight"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(AppClosedUsageData.prototype, "OpenTime", {
            get: function () { return this.Fields["OpenTime"]; },
            set: function (value) { this.Fields["OpenTime"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(AppClosedUsageData.prototype, "CloseMethod", {
            get: function () { return this.Fields["CloseMethod"]; },
            set: function (value) { this.Fields["CloseMethod"] = value; },
            enumerable: true,
            configurable: true
        });
        AppClosedUsageData.prototype.SerializeFields = function () {
            this.SetSerializedField("CorrelationId", this.CorrelationId);
            this.SetSerializedField("SessionId", this.SessionId);
            this.SetSerializedField("FocusTime", this.FocusTime);
            this.SetSerializedField("AppSizeFinalWidth", this.AppSizeFinalWidth);
            this.SetSerializedField("AppSizeFinalHeight", this.AppSizeFinalHeight);
            this.SetSerializedField("OpenTime", this.OpenTime);
            this.SetSerializedField("CloseMethod", this.CloseMethod);
        };
        return AppClosedUsageData;
    })(BaseUsageData);
    OSFLog.AppClosedUsageData = AppClosedUsageData;
    var APIUsageUsageData = (function (_super) {
        __extends(APIUsageUsageData, _super);
        function APIUsageUsageData() {
            _super.call(this, "APIUsage");
        }
        Object.defineProperty(APIUsageUsageData.prototype, "CorrelationId", {
            get: function () { return this.Fields["CorrelationId"]; },
            set: function (value) { this.Fields["CorrelationId"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(APIUsageUsageData.prototype, "SessionId", {
            get: function () { return this.Fields["SessionId"]; },
            set: function (value) { this.Fields["SessionId"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(APIUsageUsageData.prototype, "APIType", {
            get: function () { return this.Fields["APIType"]; },
            set: function (value) { this.Fields["APIType"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(APIUsageUsageData.prototype, "APIID", {
            get: function () { return this.Fields["APIID"]; },
            set: function (value) { this.Fields["APIID"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(APIUsageUsageData.prototype, "Parameters", {
            get: function () { return this.Fields["Parameters"]; },
            set: function (value) { this.Fields["Parameters"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(APIUsageUsageData.prototype, "ResponseTime", {
            get: function () { return this.Fields["ResponseTime"]; },
            set: function (value) { this.Fields["ResponseTime"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(APIUsageUsageData.prototype, "ErrorType", {
            get: function () { return this.Fields["ErrorType"]; },
            set: function (value) { this.Fields["ErrorType"] = value; },
            enumerable: true,
            configurable: true
        });
        APIUsageUsageData.prototype.SerializeFields = function () {
            this.SetSerializedField("CorrelationId", this.CorrelationId);
            this.SetSerializedField("SessionId", this.SessionId);
            this.SetSerializedField("APIType", this.APIType);
            this.SetSerializedField("APIID", this.APIID);
            this.SetSerializedField("Parameters", this.Parameters);
            this.SetSerializedField("ResponseTime", this.ResponseTime);
            this.SetSerializedField("ErrorType", this.ErrorType);
        };
        return APIUsageUsageData;
    })(BaseUsageData);
    OSFLog.APIUsageUsageData = APIUsageUsageData;
    var AppInitializationUsageData = (function (_super) {
        __extends(AppInitializationUsageData, _super);
        function AppInitializationUsageData() {
            _super.call(this, "AppInitialization");
        }
        Object.defineProperty(AppInitializationUsageData.prototype, "CorrelationId", {
            get: function () { return this.Fields["CorrelationId"]; },
            set: function (value) { this.Fields["CorrelationId"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(AppInitializationUsageData.prototype, "SessionId", {
            get: function () { return this.Fields["SessionId"]; },
            set: function (value) { this.Fields["SessionId"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(AppInitializationUsageData.prototype, "SuccessCode", {
            get: function () { return this.Fields["SuccessCode"]; },
            set: function (value) { this.Fields["SuccessCode"] = value; },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(AppInitializationUsageData.prototype, "Message", {
            get: function () { return this.Fields["Message"]; },
            set: function (value) { this.Fields["Message"] = value; },
            enumerable: true,
            configurable: true
        });
        AppInitializationUsageData.prototype.SerializeFields = function () {
            this.SetSerializedField("CorrelationId", this.CorrelationId);
            this.SetSerializedField("SessionId", this.SessionId);
            this.SetSerializedField("SuccessCode", this.SuccessCode);
            this.SetSerializedField("Message", this.Message);
        };
        return AppInitializationUsageData;
    })(BaseUsageData);
    OSFLog.AppInitializationUsageData = AppInitializationUsageData;
})(OSFLog || (OSFLog = {}));
var Logger;
(function (Logger) {
    "use strict";
    (function (TraceLevel) {
        TraceLevel[TraceLevel["info"] = 0] = "info";
        TraceLevel[TraceLevel["warning"] = 1] = "warning";
        TraceLevel[TraceLevel["error"] = 2] = "error";
    })(Logger.TraceLevel || (Logger.TraceLevel = {}));
    var TraceLevel = Logger.TraceLevel;
    (function (SendFlag) {
        SendFlag[SendFlag["none"] = 0] = "none";
        SendFlag[SendFlag["flush"] = 1] = "flush";
    })(Logger.SendFlag || (Logger.SendFlag = {}));
    var SendFlag = Logger.SendFlag;
    function allowUploadingData() {
    }
    Logger.allowUploadingData = allowUploadingData;
    function sendLog(traceLevel, message, flag) {
    }
    Logger.sendLog = sendLog;
    function creatULSEndpoint() {
        try {
            return new ULSEndpointProxy();
        }
        catch (e) {
            return null;
        }
    }
    var ULSEndpointProxy = (function () {
        function ULSEndpointProxy() {
        }
        ULSEndpointProxy.prototype.writeLog = function (log) {
        };
        ULSEndpointProxy.prototype.loadProxyFrame = function () {
        };
        return ULSEndpointProxy;
    })();
    if (!OSF.Logger) {
        OSF.Logger = Logger;
    }
    Logger.ulsEndpoint = creatULSEndpoint();
})(Logger || (Logger = {}));
var OSFAriaLogger;
(function (OSFAriaLogger) {
    var AriaLogger = (function () {
        function AriaLogger() {
        }
        AriaLogger.prototype.getAriaCDNLocation = function () {
            return (OSF._OfficeAppFactory.getLoadScriptHelper().getOfficeJsBasePath() + "/ariatelemetry/aria-web-telemetry.js");
        };
        AriaLogger.getInstance = function () {
            if (AriaLogger.AriaLoggerObj === undefined) {
                AriaLogger.AriaLoggerObj = new AriaLogger();
            }
            return AriaLogger.AriaLoggerObj;
        };
        AriaLogger.prototype.isIUsageData = function (arg) {
            return arg["Fields"] !== undefined;
        };
        AriaLogger.prototype.loadAriaScriptAndLog = function (tableName, telemetryData) {
            var startAfterMs = 1000;
            OSF.OUtil.loadScript(this.getAriaCDNLocation(), function () {
                try {
                    if (!this.ALogger) {
                        var OfficeExtensibilityTenantID = "db334b301e7b474db5e0f02f07c51a47-a1b5bc36-1bbe-482f-a64a-c2d9cb606706-7439";
                        this.ALogger = AWTLogManager.initialize(OfficeExtensibilityTenantID);
                    }
                    var eventProperties = new AWTEventProperties();
                    eventProperties.setName("Office.Extensibility.OfficeJS." + tableName);
                    for (var key in telemetryData) {
                        if (key.toLowerCase() !== "table") {
                            eventProperties.setProperty(key, telemetryData[key]);
                        }
                    }
                    var today = new Date();
                    eventProperties.setProperty("Date", today.toISOString());
                    this.ALogger.logEvent(eventProperties);
                }
                catch (e) {
                }
            }, startAfterMs);
        };
        AriaLogger.prototype.logData = function (data) {
            if (this.isIUsageData(data)) {
                this.loadAriaScriptAndLog(data["Table"], data["Fields"]);
            }
            else {
                this.loadAriaScriptAndLog(data["Table"], data);
            }
        };
        return AriaLogger;
    })();
    OSFAriaLogger.AriaLogger = AriaLogger;
})(OSFAriaLogger || (OSFAriaLogger = {}));
var OSFAppTelemetry;
(function (OSFAppTelemetry) {
    "use strict";
    var appInfo;
    var sessionId = OSF.OUtil.Guid.generateNewGuid();
    var osfControlAppCorrelationId = "";
    var omexDomainRegex = new RegExp("^https?://store\\.office(ppe|-int)?\\.com/", "i");
    OSFAppTelemetry.enableTelemetry = true;
    ;
    var AppInfo = (function () {
        function AppInfo() {
        }
        return AppInfo;
    })();
    var Event = (function () {
        function Event(name, handler) {
            this.name = name;
            this.handler = handler;
        }
        return Event;
    })();
    var AppStorage = (function () {
        function AppStorage() {
            this.clientIDKey = "Office API client";
            this.logIdSetKey = "Office App Log Id Set";
        }
        AppStorage.prototype.getClientId = function () {
            var clientId = this.getValue(this.clientIDKey);
            if (!clientId || clientId.length <= 0 || clientId.length > 40) {
                clientId = OSF.OUtil.Guid.generateNewGuid();
                this.setValue(this.clientIDKey, clientId);
            }
            return clientId;
        };
        AppStorage.prototype.saveLog = function (logId, log) {
            var logIdSet = this.getValue(this.logIdSetKey);
            logIdSet = ((logIdSet && logIdSet.length > 0) ? (logIdSet + ";") : "") + logId;
            this.setValue(this.logIdSetKey, logIdSet);
            this.setValue(logId, log);
        };
        AppStorage.prototype.enumerateLog = function (callback, clean) {
            var logIdSet = this.getValue(this.logIdSetKey);
            if (logIdSet) {
                var ids = logIdSet.split(";");
                for (var id in ids) {
                    var logId = ids[id];
                    var log = this.getValue(logId);
                    if (log) {
                        if (callback) {
                            callback(logId, log);
                        }
                        if (clean) {
                            this.remove(logId);
                        }
                    }
                }
                if (clean) {
                    this.remove(this.logIdSetKey);
                }
            }
        };
        AppStorage.prototype.getValue = function (key) {
            var osfLocalStorage = OSF.OUtil.getLocalStorage();
            var value = "";
            if (osfLocalStorage) {
                value = osfLocalStorage.getItem(key);
            }
            return value;
        };
        AppStorage.prototype.setValue = function (key, value) {
            var osfLocalStorage = OSF.OUtil.getLocalStorage();
            if (osfLocalStorage) {
                osfLocalStorage.setItem(key, value);
            }
        };
        AppStorage.prototype.remove = function (key) {
            var osfLocalStorage = OSF.OUtil.getLocalStorage();
            if (osfLocalStorage) {
                try {
                    osfLocalStorage.removeItem(key);
                }
                catch (ex) {
                }
            }
        };
        return AppStorage;
    })();
    var AppLogger = (function () {
        function AppLogger() {
        }
        AppLogger.prototype.LogData = function (data) {
            if (!OSFAppTelemetry.enableTelemetry) {
                return;
            }
            try {
                OSFAriaLogger.AriaLogger.getInstance().logData(data);
            }
            catch (e) {
            }
        };
        AppLogger.prototype.LogRawData = function (log) {
            if (!OSFAppTelemetry.enableTelemetry) {
                return;
            }
            try {
                OSFAriaLogger.AriaLogger.getInstance().logData(JSON.parse(log));
            }
            catch (e) {
            }
        };
        return AppLogger;
    })();
    function trimStringToLowerCase(input) {
        if (input) {
            input = input.replace(/[{}]/g, "").toLowerCase();
        }
        return (input || "");
    }
    function initialize(context) {
        if (!OSFAppTelemetry.enableTelemetry) {
            return;
        }
        if (appInfo) {
            return;
        }
        appInfo = new AppInfo();
        if (context.get_hostFullVersion()) {
            appInfo.hostVersion = context.get_hostFullVersion();
        }
        else {
            appInfo.hostVersion = context.get_appVersion();
        }
        appInfo.appId = context.get_id();
        appInfo.host = context.get_appName();
        appInfo.browser = window.navigator.userAgent;
        appInfo.correlationId = trimStringToLowerCase(context.get_correlationId());
        appInfo.clientId = (new AppStorage()).getClientId();
        appInfo.appInstanceId = context.get_appInstanceId();
        if (appInfo.appInstanceId) {
            appInfo.appInstanceId = appInfo.appInstanceId.replace(/[{}]/g, "").toLowerCase();
        }
        appInfo.message = context.get_hostCustomMessage();
        appInfo.officeJSVersion = OSF.ConstantNames.FileVersion;
        appInfo.hostJSVersion = "16.0.10726.30000";
        if (context._wacHostEnvironment) {
            appInfo.wacHostEnvironment = context._wacHostEnvironment;
        }
        if (context._isFromWacAutomation !== undefined && context._isFromWacAutomation !== null) {
            appInfo.isFromWacAutomation = context._isFromWacAutomation.toString().toLowerCase();
        }
        var docUrl = context.get_docUrl();
        appInfo.docUrl = omexDomainRegex.test(docUrl) ? docUrl : "";
        var url = location.href;
        if (url) {
            url = url.split("?")[0].split("#")[0];
        }
        appInfo.appURL = url;
        (function getUserIdAndAssetIdFromToken(token, appInfo) {
            var xmlContent;
            var parser;
            var xmlDoc;
            appInfo.assetId = "";
            appInfo.userId = "";
            try {
                xmlContent = decodeURIComponent(token);
                parser = new DOMParser();
                xmlDoc = parser.parseFromString(xmlContent, "text/xml");
                var cidNode = xmlDoc.getElementsByTagName("t")[0].attributes.getNamedItem("cid");
                var oidNode = xmlDoc.getElementsByTagName("t")[0].attributes.getNamedItem("oid");
                if (cidNode && cidNode.nodeValue) {
                    appInfo.userId = cidNode.nodeValue;
                }
                else if (oidNode && oidNode.nodeValue) {
                    appInfo.userId = oidNode.nodeValue;
                }
                appInfo.assetId = xmlDoc.getElementsByTagName("t")[0].attributes.getNamedItem("aid").nodeValue;
            }
            catch (e) {
            }
            finally {
                xmlContent = null;
                xmlDoc = null;
                parser = null;
            }
        })(context.get_eToken(), appInfo);
        (function handleLifecycle() {
            var startTime = new Date();
            var lastFocus = null;
            var focusTime = 0;
            var finished = false;
            var adjustFocusTime = function () {
                if (document.hasFocus()) {
                    if (lastFocus == null) {
                        lastFocus = new Date();
                    }
                }
                else if (lastFocus) {
                    focusTime += Math.abs((new Date()).getTime() - lastFocus.getTime());
                    lastFocus = null;
                }
            };
            var eventList = [];
            eventList.push(new Event("focus", adjustFocusTime));
            eventList.push(new Event("blur", adjustFocusTime));
            eventList.push(new Event("focusout", adjustFocusTime));
            eventList.push(new Event("focusin", adjustFocusTime));
            var exitFunction = function () {
                for (var i = 0; i < eventList.length; i++) {
                    OSF.OUtil.removeEventListener(window, eventList[i].name, eventList[i].handler);
                }
                eventList.length = 0;
                if (!finished) {
                    if (document.hasFocus() && lastFocus) {
                        focusTime += Math.abs((new Date()).getTime() - lastFocus.getTime());
                        lastFocus = null;
                    }
                    OSFAppTelemetry.onAppClosed(Math.abs((new Date()).getTime() - startTime.getTime()), focusTime);
                    finished = true;
                }
            };
            eventList.push(new Event("beforeunload", exitFunction));
            eventList.push(new Event("unload", exitFunction));
            for (var i = 0; i < eventList.length; i++) {
                OSF.OUtil.addEventListener(window, eventList[i].name, eventList[i].handler);
            }
            adjustFocusTime();
        })();
        OSFAppTelemetry.onAppActivated();
    }
    OSFAppTelemetry.initialize = initialize;
    function onAppActivated() {
        if (!appInfo) {
            return;
        }
        (new AppStorage()).enumerateLog(function (id, log) { return (new AppLogger()).LogRawData(log); }, true);
        var data = new OSFLog.AppActivatedUsageData();
        data.SessionId = sessionId;
        data.AppId = appInfo.appId;
        data.AssetId = appInfo.assetId;
        data.AppURL = appInfo.appURL;
        data.UserId = "";
        data.ClientId = appInfo.clientId;
        data.Browser = appInfo.browser;
        data.Host = appInfo.host;
        data.HostVersion = appInfo.hostVersion;
        data.CorrelationId = trimStringToLowerCase(appInfo.correlationId);
        data.AppSizeWidth = window.innerWidth;
        data.AppSizeHeight = window.innerHeight;
        data.AppInstanceId = appInfo.appInstanceId;
        data.Message = appInfo.message;
        data.DocUrl = appInfo.docUrl;
        data.OfficeJSVersion = appInfo.officeJSVersion;
        data.HostJSVersion = appInfo.hostJSVersion;
        if (appInfo.wacHostEnvironment) {
            data.WacHostEnvironment = appInfo.wacHostEnvironment;
        }
        if (appInfo.isFromWacAutomation !== undefined && appInfo.isFromWacAutomation !== null) {
            data.IsFromWacAutomation = appInfo.isFromWacAutomation;
        }
        (new AppLogger()).LogData(data);
    }
    OSFAppTelemetry.onAppActivated = onAppActivated;
    function onScriptDone(scriptId, msStartTime, msResponseTime, appCorrelationId) {
        var data = new OSFLog.ScriptLoadUsageData();
        data.CorrelationId = trimStringToLowerCase(appCorrelationId);
        data.SessionId = sessionId;
        data.ScriptId = scriptId;
        data.StartTime = msStartTime;
        data.ResponseTime = msResponseTime;
        (new AppLogger()).LogData(data);
    }
    OSFAppTelemetry.onScriptDone = onScriptDone;
    function onCallDone(apiType, id, parameters, msResponseTime, errorType) {
        if (!appInfo) {
            return;
        }
        var data = new OSFLog.APIUsageUsageData();
        data.CorrelationId = trimStringToLowerCase(osfControlAppCorrelationId);
        data.SessionId = sessionId;
        data.APIType = apiType;
        data.APIID = id;
        data.Parameters = parameters;
        data.ResponseTime = msResponseTime;
        data.ErrorType = errorType;
        (new AppLogger()).LogData(data);
    }
    OSFAppTelemetry.onCallDone = onCallDone;
    ;
    function onMethodDone(id, args, msResponseTime, errorType) {
        var parameters = null;
        if (args) {
            if (typeof args == "number") {
                parameters = String(args);
            }
            else if (typeof args === "object") {
                for (var index in args) {
                    if (parameters !== null) {
                        parameters += ",";
                    }
                    else {
                        parameters = "";
                    }
                    if (typeof args[index] == "number") {
                        parameters += String(args[index]);
                    }
                }
            }
            else {
                parameters = "";
            }
        }
        OSF.AppTelemetry.onCallDone("method", id, parameters, msResponseTime, errorType);
    }
    OSFAppTelemetry.onMethodDone = onMethodDone;
    function onPropertyDone(propertyName, msResponseTime) {
        OSF.AppTelemetry.onCallDone("property", -1, propertyName, msResponseTime);
    }
    OSFAppTelemetry.onPropertyDone = onPropertyDone;
    function onEventDone(id, errorType) {
        OSF.AppTelemetry.onCallDone("event", id, null, 0, errorType);
    }
    OSFAppTelemetry.onEventDone = onEventDone;
    function onRegisterDone(register, id, msResponseTime, errorType) {
        OSF.AppTelemetry.onCallDone(register ? "registerevent" : "unregisterevent", id, null, msResponseTime, errorType);
    }
    OSFAppTelemetry.onRegisterDone = onRegisterDone;
    function onAppClosed(openTime, focusTime) {
        if (!appInfo) {
            return;
        }
        var data = new OSFLog.AppClosedUsageData();
        data.CorrelationId = trimStringToLowerCase(osfControlAppCorrelationId);
        data.SessionId = sessionId;
        data.FocusTime = focusTime;
        data.OpenTime = openTime;
        data.AppSizeFinalWidth = window.innerWidth;
        data.AppSizeFinalHeight = window.innerHeight;
        (new AppStorage()).saveLog(sessionId, data.SerializeRow());
    }
    OSFAppTelemetry.onAppClosed = onAppClosed;
    function setOsfControlAppCorrelationId(correlationId) {
        osfControlAppCorrelationId = trimStringToLowerCase(correlationId);
    }
    OSFAppTelemetry.setOsfControlAppCorrelationId = setOsfControlAppCorrelationId;
    function doAppInitializationLogging(isException, message) {
        var data = new OSFLog.AppInitializationUsageData();
        data.CorrelationId = trimStringToLowerCase(osfControlAppCorrelationId);
        data.SessionId = sessionId;
        data.SuccessCode = isException ? 1 : 0;
        data.Message = message;
        (new AppLogger()).LogData(data);
    }
    OSFAppTelemetry.doAppInitializationLogging = doAppInitializationLogging;
    function logAppCommonMessage(message) {
        doAppInitializationLogging(false, message);
    }
    OSFAppTelemetry.logAppCommonMessage = logAppCommonMessage;
    function logAppException(errorMessage) {
        doAppInitializationLogging(true, errorMessage);
    }
    OSFAppTelemetry.logAppException = logAppException;
    OSF.AppTelemetry = OSFAppTelemetry;
})(OSFAppTelemetry || (OSFAppTelemetry = {}));
OSF.InitializationHelper.prototype.loadAppSpecificScriptAndCreateOM = function OSF_InitializationHelper$loadAppSpecificScriptAndCreateOM(appContext, appReady, basePath) {
    OSF.DDA.ErrorCodeManager.initializeErrorMessages(Strings.OfficeOM);
    OSF.DDA.DispIdHost.addAsyncMethods(OSF.DDA.RichApi, [OSF.DDA.AsyncMethodNames.ExecuteRichApiRequestAsync]);
    OSF.DDA.RichApi.richApiMessageManager = new OfficeExt.RichApiMessageManager();
    appReady();
};
OSF.DDA.AsyncMethodNames.addNames({
    ExecuteRichApiRequestAsync: "executeRichApiRequestAsync"
});
OSF.DDA.AsyncMethodCalls.define({
    method: OSF.DDA.AsyncMethodNames.ExecuteRichApiRequestAsync,
    requiredArguments: [
        {
            name: Microsoft.Office.WebExtension.Parameters.Data,
            types: ["object"]
        }
    ],
    supportedOptions: []
});
OSF.OUtil.setNamespace("RichApi", OSF.DDA);
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
    type: OSF.DDA.MethodDispId.dispidExecuteRichApiRequestMethod,
    toHost: [
        { name: Microsoft.Office.WebExtension.Parameters.Data, value: 0 }
    ],
    fromHost: [
        { name: Microsoft.Office.WebExtension.Parameters.Data, value: OSF.DDA.SafeArray.Delegate.ParameterMap.self }
    ]
});
Microsoft.Office.WebExtension.EventType = {};
OSF.EventDispatch = function OSF_EventDispatch(eventTypes) {
    this._eventHandlers = {};
    this._objectEventHandlers = {};
    this._queuedEventsArgs = {};
    if (eventTypes != null) {
        for (var i = 0; i < eventTypes.length; i++) {
            var eventType = eventTypes[i];
            var isObjectEvent = (eventType == "objectDeleted" || eventType == "objectSelectionChanged" || eventType == "objectDataChanged" || eventType == "contentControlAdded");
            if (!isObjectEvent)
                this._eventHandlers[eventType] = [];
            else
                this._objectEventHandlers[eventType] = {};
            this._queuedEventsArgs[eventType] = [];
        }
    }
};
OSF.EventDispatch.prototype = {
    getSupportedEvents: function OSF_EventDispatch$getSupportedEvents() {
        var events = [];
        for (var eventName in this._eventHandlers)
            events.push(eventName);
        for (var eventName in this._objectEventHandlers)
            events.push(eventName);
        return events;
    },
    supportsEvent: function OSF_EventDispatch$supportsEvent(event) {
        for (var eventName in this._eventHandlers) {
            if (event == eventName)
                return true;
        }
        for (var eventName in this._objectEventHandlers) {
            if (event == eventName)
                return true;
        }
        return false;
    },
    hasEventHandler: function OSF_EventDispatch$hasEventHandler(eventType, handler) {
        var handlers = this._eventHandlers[eventType];
        if (handlers && handlers.length > 0) {
            for (var i = 0; i < handlers.length; i++) {
                if (handlers[i] === handler)
                    return true;
            }
        }
        return false;
    },
    hasObjectEventHandler: function OSF_EventDispatch$hasObjectEventHandler(eventType, objectId, handler) {
        var handlers = this._objectEventHandlers[eventType];
        if (handlers != null) {
            var _handlers = handlers[objectId];
            for (var i = 0; _handlers != null && i < _handlers.length; i++) {
                if (_handlers[i] === handler)
                    return true;
            }
        }
        return false;
    },
    addEventHandler: function OSF_EventDispatch$addEventHandler(eventType, handler) {
        if (typeof handler != "function") {
            return false;
        }
        var handlers = this._eventHandlers[eventType];
        if (handlers && !this.hasEventHandler(eventType, handler)) {
            handlers.push(handler);
            return true;
        }
        else {
            return false;
        }
    },
    addObjectEventHandler: function OSF_EventDispatch$addObjectEventHandler(eventType, objectId, handler) {
        if (typeof handler != "function") {
            return false;
        }
        var handlers = this._objectEventHandlers[eventType];
        if (handlers && !this.hasObjectEventHandler(eventType, objectId, handler)) {
            if (handlers[objectId] == null)
                handlers[objectId] = [];
            handlers[objectId].push(handler);
            return true;
        }
        return false;
    },
    addEventHandlerAndFireQueuedEvent: function OSF_EventDispatch$addEventHandlerAndFireQueuedEvent(eventType, handler) {
        var handlers = this._eventHandlers[eventType];
        var isFirstHandler = handlers.length == 0;
        var succeed = this.addEventHandler(eventType, handler);
        if (isFirstHandler && succeed) {
            this.fireQueuedEvent(eventType);
        }
        return succeed;
    },
    removeEventHandler: function OSF_EventDispatch$removeEventHandler(eventType, handler) {
        var handlers = this._eventHandlers[eventType];
        if (handlers && handlers.length > 0) {
            for (var index = 0; index < handlers.length; index++) {
                if (handlers[index] === handler) {
                    handlers.splice(index, 1);
                    return true;
                }
            }
        }
        return false;
    },
    removeObjectEventHandler: function OSF_EventDispatch$removeObjectEventHandler(eventType, objectId, handler) {
        var handlers = this._objectEventHandlers[eventType];
        if (handlers != null) {
            var _handlers = handlers[objectId];
            for (var i = 0; _handlers != null && i < _handlers.length; i++) {
                if (_handlers[i] === handler) {
                    _handlers.splice(i, 1);
                    return true;
                }
            }
        }
        return false;
    },
    clearEventHandlers: function OSF_EventDispatch$clearEventHandlers(eventType) {
        if (typeof this._eventHandlers[eventType] != "undefined" && this._eventHandlers[eventType].length > 0) {
            this._eventHandlers[eventType] = [];
            return true;
        }
        return false;
    },
    clearObjectEventHandlers: function OSF_EventDispatch$clearObjectEventHandlers(eventType, objectId) {
        if (this._objectEventHandlers[eventType] != null && this._objectEventHandlers[eventType][objectId] != null) {
            this._objectEventHandlers[eventType][objectId] = [];
            return true;
        }
        return false;
    },
    getEventHandlerCount: function OSF_EventDispatch$getEventHandlerCount(eventType) {
        return this._eventHandlers[eventType] != undefined ? this._eventHandlers[eventType].length : -1;
    },
    getObjectEventHandlerCount: function OSF_EventDispatch$getObjectEventHandlerCount(eventType, objectId) {
        if (this._objectEventHandlers[eventType] == null || this._objectEventHandlers[eventType][objectId] == null)
            return 0;
        return this._objectEventHandlers[eventType][objectId].length;
    },
    fireEvent: function OSF_EventDispatch$fireEvent(eventArgs) {
        if (eventArgs.type == undefined)
            return false;
        var eventType = eventArgs.type;
        if (eventType && this._eventHandlers[eventType]) {
            var eventHandlers = this._eventHandlers[eventType];
            for (var i = 0; i < eventHandlers.length; i++) {
                eventHandlers[i](eventArgs);
            }
            return true;
        }
        else {
            return false;
        }
    },
    fireObjectEvent: function OSF_EventDispatch$fireObjectEvent(objectId, eventArgs) {
        if (eventArgs.type == undefined)
            return false;
        var eventType = eventArgs.type;
        if (eventType && this._objectEventHandlers[eventType]) {
            var eventHandlers = this._objectEventHandlers[eventType];
            var _handlers = eventHandlers[objectId];
            if (_handlers != null) {
                for (var i = 0; i < _handlers.length; i++)
                    _handlers[i](eventArgs);
                return true;
            }
        }
        return false;
    },
    fireOrQueueEvent: function OSF_EventDispatch$fireOrQueueEvent(eventArgs) {
        var eventType = eventArgs.type;
        if (eventType && this._eventHandlers[eventType]) {
            var eventHandlers = this._eventHandlers[eventType];
            var queuedEvents = this._queuedEventsArgs[eventType];
            if (eventHandlers.length == 0) {
                queuedEvents.push(eventArgs);
            }
            else {
                this.fireEvent(eventArgs);
            }
            return true;
        }
        else {
            return false;
        }
    },
    fireQueuedEvent: function OSF_EventDispatch$queueEvent(eventType) {
        if (eventType && this._eventHandlers[eventType]) {
            var eventHandlers = this._eventHandlers[eventType];
            var queuedEvents = this._queuedEventsArgs[eventType];
            if (eventHandlers.length > 0) {
                var eventHandler = eventHandlers[0];
                while (queuedEvents.length > 0) {
                    var eventArgs = queuedEvents.shift();
                    eventHandler(eventArgs);
                }
                return true;
            }
        }
        return false;
    },
    clearQueuedEvent: function OSF_EventDispatch$clearQueuedEvent(eventType) {
        if (eventType && this._eventHandlers[eventType]) {
            var queuedEvents = this._queuedEventsArgs[eventType];
            if (queuedEvents) {
                this._queuedEventsArgs[eventType] = [];
            }
        }
    }
};
OSF.DDA.OMFactory = OSF.DDA.OMFactory || {};
OSF.DDA.OMFactory.manufactureEventArgs = function OSF_DDA_OMFactory$manufactureEventArgs(eventType, target, eventProperties) {
    var args;
    switch (eventType) {
        case Microsoft.Office.WebExtension.EventType.DocumentSelectionChanged:
            args = new OSF.DDA.DocumentSelectionChangedEventArgs(target);
            break;
        case Microsoft.Office.WebExtension.EventType.BindingSelectionChanged:
            args = new OSF.DDA.BindingSelectionChangedEventArgs(this.manufactureBinding(eventProperties, target.document), eventProperties[OSF.DDA.PropertyDescriptors.Subset]);
            break;
        case Microsoft.Office.WebExtension.EventType.BindingDataChanged:
            args = new OSF.DDA.BindingDataChangedEventArgs(this.manufactureBinding(eventProperties, target.document));
            break;
        case Microsoft.Office.WebExtension.EventType.SettingsChanged:
            args = new OSF.DDA.SettingsChangedEventArgs(target);
            break;
        case Microsoft.Office.WebExtension.EventType.ActiveViewChanged:
            args = new OSF.DDA.ActiveViewChangedEventArgs(eventProperties);
            break;
        case Microsoft.Office.WebExtension.EventType.OfficeThemeChanged:
            args = new OSF.DDA.Theming.OfficeThemeChangedEventArgs(eventProperties);
            break;
        case Microsoft.Office.WebExtension.EventType.DocumentThemeChanged:
            args = new OSF.DDA.Theming.DocumentThemeChangedEventArgs(eventProperties);
            break;
        case Microsoft.Office.WebExtension.EventType.AppCommandInvoked:
            args = OSF.DDA.AppCommand.AppCommandInvokedEventArgs.create(eventProperties);
            break;
        case Microsoft.Office.WebExtension.EventType.ObjectDeleted:
        case Microsoft.Office.WebExtension.EventType.ObjectSelectionChanged:
        case Microsoft.Office.WebExtension.EventType.ObjectDataChanged:
        case Microsoft.Office.WebExtension.EventType.ContentControlAdded:
            args = new OSF.DDA.ObjectEventArgs(eventType, eventProperties[Microsoft.Office.WebExtension.Parameters.Id]);
            break;
        case Microsoft.Office.WebExtension.EventType.RichApiMessage:
            args = new OSF.DDA.RichApiMessageEventArgs(eventType, eventProperties);
            break;
        case Microsoft.Office.WebExtension.EventType.DataNodeInserted:
            args = new OSF.DDA.NodeInsertedEventArgs(this.manufactureDataNode(eventProperties[OSF.DDA.DataNodeEventProperties.NewNode]), eventProperties[OSF.DDA.DataNodeEventProperties.InUndoRedo]);
            break;
        case Microsoft.Office.WebExtension.EventType.DataNodeReplaced:
            args = new OSF.DDA.NodeReplacedEventArgs(this.manufactureDataNode(eventProperties[OSF.DDA.DataNodeEventProperties.OldNode]), this.manufactureDataNode(eventProperties[OSF.DDA.DataNodeEventProperties.NewNode]), eventProperties[OSF.DDA.DataNodeEventProperties.InUndoRedo]);
            break;
        case Microsoft.Office.WebExtension.EventType.DataNodeDeleted:
            args = new OSF.DDA.NodeDeletedEventArgs(this.manufactureDataNode(eventProperties[OSF.DDA.DataNodeEventProperties.OldNode]), this.manufactureDataNode(eventProperties[OSF.DDA.DataNodeEventProperties.NextSiblingNode]), eventProperties[OSF.DDA.DataNodeEventProperties.InUndoRedo]);
            break;
        case Microsoft.Office.WebExtension.EventType.TaskSelectionChanged:
            args = new OSF.DDA.TaskSelectionChangedEventArgs(target);
            break;
        case Microsoft.Office.WebExtension.EventType.ResourceSelectionChanged:
            args = new OSF.DDA.ResourceSelectionChangedEventArgs(target);
            break;
        case Microsoft.Office.WebExtension.EventType.ViewSelectionChanged:
            args = new OSF.DDA.ViewSelectionChangedEventArgs(target);
            break;
        case Microsoft.Office.WebExtension.EventType.DialogMessageReceived:
            args = new OSF.DDA.DialogEventArgs(eventProperties);
            break;
        case Microsoft.Office.WebExtension.EventType.DialogParentMessageReceived:
            args = new OSF.DDA.DialogParentEventArgs(eventProperties);
            break;
        case Microsoft.Office.WebExtension.EventType.ItemChanged:
            if (OSF._OfficeAppFactory.getHostInfo()["hostType"] == "outlook") {
                args = new OSF.DDA.OlkItemSelectedChangedEventArgs(eventProperties);
                target.initialize(args["initialData"]);
                if (OSF._OfficeAppFactory.getHostInfo()["hostPlatform"] == "win32" || OSF._OfficeAppFactory.getHostInfo()["hostPlatform"] == "mac") {
                    target.setCurrentItemNumber(args["itemNumber"].itemNumber);
                }
            }
            else {
                throw OsfMsAjaxFactory.msAjaxError.argument(Microsoft.Office.WebExtension.Parameters.EventType, OSF.OUtil.formatString(Strings.OfficeOM.L_NotSupportedEventType, eventType));
            }
            break;
        case Microsoft.Office.WebExtension.EventType.RecipientsChanged:
            if (OSF._OfficeAppFactory.getHostInfo()["hostType"] == "outlook") {
                args = new OSF.DDA.OlkRecipientsChangedEventArgs(eventProperties);
            }
            else {
                throw OsfMsAjaxFactory.msAjaxError.argument(Microsoft.Office.WebExtension.Parameters.EventType, OSF.OUtil.formatString(Strings.OfficeOM.L_NotSupportedEventType, eventType));
            }
            break;
        case Microsoft.Office.WebExtension.EventType.AppointmentTimeChanged:
            if (OSF._OfficeAppFactory.getHostInfo()["hostType"] == "outlook") {
                args = new OSF.DDA.OlkAppointmentTimeChangedEventArgs(eventProperties);
            }
            else {
                throw OsfMsAjaxFactory.msAjaxError.argument(Microsoft.Office.WebExtension.Parameters.EventType, OSF.OUtil.formatString(Strings.OfficeOM.L_NotSupportedEventType, eventType));
            }
            break;
        case Microsoft.Office.WebExtension.EventType.RecurrenceChanged:
            if (OSF._OfficeAppFactory.getHostInfo()["hostType"] == "outlook") {
                args = new OSF.DDA.OlkRecurrenceChangedEventArgs(eventProperties);
            }
            else {
                throw OsfMsAjaxFactory.msAjaxError.argument(Microsoft.Office.WebExtension.Parameters.EventType, OSF.OUtil.formatString(Strings.OfficeOM.L_NotSupportedEventType, eventType));
            }
            break;
        default:
            throw OsfMsAjaxFactory.msAjaxError.argument(Microsoft.Office.WebExtension.Parameters.EventType, OSF.OUtil.formatString(Strings.OfficeOM.L_NotSupportedEventType, eventType));
    }
    return args;
};
OSF.DDA.AsyncMethodNames.addNames({
    AddHandlerAsync: "addHandlerAsync",
    RemoveHandlerAsync: "removeHandlerAsync"
});
OSF.DDA.AsyncMethodCalls.define({
    method: OSF.DDA.AsyncMethodNames.AddHandlerAsync,
    requiredArguments: [{
            "name": Microsoft.Office.WebExtension.Parameters.EventType,
            "enum": Microsoft.Office.WebExtension.EventType,
            "verify": function (eventType, caller, eventDispatch) { return eventDispatch.supportsEvent(eventType); }
        },
        {
            "name": Microsoft.Office.WebExtension.Parameters.Handler,
            "types": ["function"]
        }
    ],
    supportedOptions: [],
    privateStateCallbacks: []
});
OSF.DDA.AsyncMethodCalls.define({
    method: OSF.DDA.AsyncMethodNames.RemoveHandlerAsync,
    requiredArguments: [
        {
            "name": Microsoft.Office.WebExtension.Parameters.EventType,
            "enum": Microsoft.Office.WebExtension.EventType,
            "verify": function (eventType, caller, eventDispatch) { return eventDispatch.supportsEvent(eventType); }
        }
    ],
    supportedOptions: [
        {
            name: Microsoft.Office.WebExtension.Parameters.Handler,
            value: {
                "types": ["function", "object"],
                "defaultValue": null
            }
        }
    ],
    privateStateCallbacks: []
});
OSF.OUtil.augmentList(Microsoft.Office.WebExtension.EventType, { RichApiMessage: "richApiMessage" });
OSF.DDA.RichApiMessageEventArgs = function OSF_DDA_RichApiMessageEventArgs(eventType, eventProperties) {
    var entryArray = eventProperties[Microsoft.Office.WebExtension.Parameters.Data];
    var entries = [];
    if (entryArray) {
        for (var i = 0; i < entryArray.length; i++) {
            var elem = entryArray[i];
            if (elem.toArray) {
                elem = elem.toArray();
            }
            entries.push({
                messageCategory: elem[0],
                messageType: elem[1],
                targetId: elem[2],
                message: elem[3],
                id: elem[4],
                isRemoteOverride: elem[5]
            });
        }
    }
    OSF.OUtil.defineEnumerableProperties(this, {
        "type": { value: Microsoft.Office.WebExtension.EventType.RichApiMessage },
        "entries": { value: entries }
    });
};
var OfficeExt;
(function (OfficeExt) {
    var RichApiMessageManager = (function () {
        function RichApiMessageManager() {
            this._eventDispatch = null;
            this._eventDispatch = new OSF.EventDispatch([
                Microsoft.Office.WebExtension.EventType.RichApiMessage,
            ]);
            OSF.DDA.DispIdHost.addEventSupport(this, this._eventDispatch);
        }
        return RichApiMessageManager;
    })();
    OfficeExt.RichApiMessageManager = RichApiMessageManager;
})(OfficeExt || (OfficeExt = {}));
OSF.DDA.SafeArray.Delegate.ParameterMap.define({
    type: OSF.DDA.EventDispId.dispidRichApiMessageEvent,
    toHost: [
        { name: Microsoft.Office.WebExtension.Parameters.Data, value: 0 }
    ],
    fromHost: [
        { name: Microsoft.Office.WebExtension.Parameters.Data, value: OSF.DDA.SafeArray.Delegate.ParameterMap.sourceData }
    ]
});




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
var DialogApi = __webpack_require__(4);
var AsyncStorage = __webpack_require__(6);
window._OfficeRuntimeNative = {
    displayWebDialog: DialogApi.displayWebDialog,
    AsyncStorage: AsyncStorage.AsyncStorage
};


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
/* Begin_PlaceHolder_ModuleHeader */
/* End_PlaceHolder_ModuleHeader */
var _hostName = 'Office';
var _defaultApiSetName = 'OfficeSharedApi';
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
var _typeDialogService = 'DialogService';
var Dialog = /** @class */ (function () {
    function Dialog(_dialogService) {
        this._dialogService = _dialogService;
    }
    Dialog.prototype.close = function () {
        this._dialogService.close();
        return this._dialogService.context.sync();
    };
    return Dialog;
}());
exports.Dialog = Dialog;
function displayWebDialog(url, options) {
    return new OfficeExtension.CoreUtility.Promise(function (resolve, reject) {
        // Check parameters
        if (options.width && options.height && (!isInt(options.width) || !isInt(options.height))) {
            // null dimensions are valid
            throw new OfficeExtension.Error({ code: 'InvalidArgument', message: 'Dimensions must be % or number.' });
        }
        var ctx = new OfficeExtension.ClientRequestContext();
        var dialogService = DialogService.newObject(ctx);
        var dialog = new Dialog(dialogService);
        var eventResult = dialogService.onDialogMessage.add(function (args) {
            // Check if message recieved is an error message
            OfficeExtension.Utility.log('dialogMessageHandler:' + JSON.stringify(args));
            // Determine what kind of event to fire
            switch (args.type) {
                case 17 /* dialogInitializationDoneEvent */:
                    if (args.error) {
                        reject(args.error);
                    }
                    else {
                        resolve(dialog);
                    }
                    break;
                case 12 /* dialogParentMessageReceivedEvent */:
                    // Message parent case
                    if (options.onMessage) {
                        options.onMessage(args.message, dialog);
                    }
                    break;
                case 10 /* dialogMessageEvent */:
                default:
                    if (args.originalErrorCode === 12006 /* dialogClosedError */) {
                        if (eventResult) {
                            eventResult.remove();
                            ctx.sync(); // A best-effort to remove
                        }
                        if (options.onClose) {
                            options.onClose();
                        }
                    }
                    else {
                        if (options.onRuntimeError) {
                            options.onRuntimeError(args.error, dialog);
                        }
                    }
            }
            return OfficeExtension.CoreUtility.Promise.resolve();
        });
        return ctx
            .sync()
            .then(function () {
            var dialogOptions = {
                width: options.width ? parseInt(options.width) : 50,
                height: options.height ? parseInt(options.height) : 50,
                displayInIFrame: options.displayInIFrame,
                hideTitle: options.hideTitle
            };
            dialogService.displayDialog(url, dialogOptions);
            return ctx.sync();
            // Note: actual resolving will happen once you get the dialogInitializationDone event
        })
            .catch(function (e) {
            reject(e);
        });
    });
    function isInt(value) {
        // (/^(\-|\+)?([0-9]+)$/.test
        // Ensures the entire string is a number
        return /^(\-|\+)?([0-9]+)%?$/.test(value);
    }
}
exports.displayWebDialog = displayWebDialog;
/* End_PlaceHolder_DialogService_BeforeDeclaration */
/**
 *
 * Represents interface for opening a dialog
 *
 * [Api set: Dialog 1.2]
 */
var DialogService = /** @class */ (function (_super) {
    __extends(DialogService, _super);
    function DialogService() {
        return _super !== null && _super.apply(this, arguments) || this;
    }
    Object.defineProperty(DialogService.prototype, "_className", {
        get: function () {
            return 'DialogService';
        },
        enumerable: true,
        configurable: true
    });
    /* Begin_PlaceHolder_DialogService_Custom_Members */
    /* End_PlaceHolder_DialogService_Custom_Members */
    // SET method is absent, because no settable properties on self or children.
    DialogService.prototype.close = function () {
        /* Begin_PlaceHolder_DialogService_Close */
        /* End_PlaceHolder_DialogService_Close */
        _invokeMethod(this, 'Close', 1 /* Read */, [], 4 /* concurrent */, 0 /* none */);
    };
    DialogService.prototype.displayDialog = function (url, options) {
        /* Begin_PlaceHolder_DialogService_DisplayDialog */
        /* End_PlaceHolder_DialogService_DisplayDialog */
        _invokeMethod(this, 'DisplayDialog', 1 /* Read */, [url, options], 4 /* concurrent */, 0 /* none */);
    };
    /** Handle results returned from the document
     * @private
     */
    DialogService.prototype._handleResult = function (value) {
        _super.prototype._handleResult.call(this, value);
        if (_isNullOrUndefined(value))
            return;
        var obj = value;
        _fixObjectPathIfNecessary(this, obj);
        /* Begin_PlaceHolder_DialogService_HandleResult */
        /* End_PlaceHolder_DialogService_HandleResult */
    };
    /** Handle retrieve results
     * @private
     */
    DialogService.prototype._handleRetrieveResult = function (value, result) {
        _super.prototype._handleRetrieveResult.call(this, value, result);
        /* Begin_PlaceHolder_DialogService_HandleRetrieveResult */
        /* End_PlaceHolder_DialogService_HandleRetrieveResult */
        _processRetrieveResult(this, value, result);
    };
    /**
     * Create a new instance of DialogService object
     */
    DialogService.newObject = function (context) {
        return _createTopLevelServiceObject(DialogService, context, 'Microsoft.Dialog.DialogService', false /*isCollection*/, 4 /* concurrent */);
    };
    Object.defineProperty(DialogService.prototype, "onDialogMessage", {
        /**
         *
         * Occurs when the Dialog sends a message
         *
         * [Api set: Dialog 1.2]
         */
        get: function () {
            /* Begin_PlaceHolder_DialogService_DialogMessage_get */
            /* End_PlaceHolder_DialogService_DialogMessage_get */
            if (!this.m_dialogMessage) {
                /* Begin_PlaceHolder_DialogService_DialogMessage_get_PreInit */
                /* End_PlaceHolder_DialogService_DialogMessage_get_PreInit */
                /* Begin_PlaceHolder_DialogService_DialogMessage_get_EventHandlers */
                this.m_dialogMessage = new OfficeExtension.GenericEventHandlers(
                /* End_PlaceHolder_DialogService_DialogMessage_get_EventHandlers */
                this.context, this, 'DialogMessage', 
                // Please add eventInfo between the placeholders
                /* Begin_PlaceHolder_DialogService_DialogMessage_Constructor_Parameters */
                {
                    eventType: 65536 /* dialogRichApiMessageEvent */,
                    registerFunc: function () { return void {}; },
                    unregisterFunc: function () { return void {}; },
                    getTargetIdFunc: function () {
                        return null;
                    },
                    eventArgsTransformFunc: function (args) {
                        var transformedArgs;
                        try {
                            var parsedMessage = JSON.parse(args.message);
                            var error = parsedMessage.errorCode
                                ? new OfficeExtension.Error(lookupErrorCodeAndMessage(parsedMessage.errorCode))
                                : null;
                            transformedArgs = {
                                originalErrorCode: parsedMessage.errorCode,
                                type: parsedMessage.type,
                                error: error,
                                message: parsedMessage.message
                            };
                        }
                        catch (e) {
                            transformedArgs = {
                                originalErrorCode: null,
                                type: 17 /* dialogInitializationDoneEvent */,
                                error: new OfficeExtension.Error({ code: 'GenericException', message: 'Unknown error' }),
                                message: e.message
                            };
                        }
                        return OfficeExtension.Utility._createPromiseFromResult(transformedArgs);
                        // Helper
                        function lookupErrorCodeAndMessage(internalCode) {
                            var table = (_a = {},
                                _a[12002 /* dialogInvalidURLError */] = {
                                    code: 'InvalidUrl',
                                    message: 'Cannot load URL, no such page or bad URL syntax.'
                                },
                                _a[12003 /* dialogHttpsRequiredError */] = { code: 'InvalidUrl', message: 'HTTPS is required.' },
                                _a[12004 /* dialogUntrustedDomainError */] = { code: 'Untrusted', message: 'Domain is not trusted.' },
                                _a[12005 /* dialogHttpsRequiredErrorInitialization */] = {
                                    code: 'InvalidUrl',
                                    message: 'HTTPS is required.'
                                },
                                _a[12007 /* dialogAlreadyOpenedError */] = {
                                    code: 'FailedToOpen',
                                    message: 'Another dialog is already opened.'
                                },
                                _a);
                            if (table[internalCode]) {
                                return table[internalCode];
                            }
                            else {
                                return { code: 'Unknown', message: 'An unknown error has occured' };
                            }
                            var _a;
                        }
                    }
                }
                /* End_PlaceHolder_DialogService_DialogMessage_Constructor_Parameters */
                );
                /* Begin_PlaceHolder_DialogService_DialogMessage_get_AfterInit */
                /* End_PlaceHolder_DialogService_DialogMessage_get_AfterInit */
            }
            return this.m_dialogMessage;
        },
        enumerable: true,
        configurable: true
    });
    DialogService.prototype.toJSON = function () {
        return _toJson(this, /* scalarProperties: */ {}, /* navigationProperties: */ {});
    };
    return DialogService;
}(OfficeExtension.ClientObject));
exports.DialogService = DialogService;
// Keep non-const so we can check membership
var DialogEventType;
(function (DialogEventType) {
    DialogEventType[DialogEventType["dialogMessageReceived"] = 0] = "dialogMessageReceived";
    DialogEventType[DialogEventType["dialogEventReceived"] = 1] = "dialogEventReceived";
})(DialogEventType || (DialogEventType = {}));
/* Begin_PlaceHolder_ErrorCodesTypeName */
var DialogErrorCodes;
(function (DialogErrorCodes) {
    /* End_PlaceHolder_ErrorCodesTypeName */
    DialogErrorCodes["generalException"] = "GeneralException";
    /* Begin_PlaceHolder_ErrorCodesAdditional */
    /* End_PlaceHolder_ErrorCodesAdditional */
})(DialogErrorCodes = exports.DialogErrorCodes || (exports.DialogErrorCodes = {}));


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
/* Begin_PlaceHolder_ModuleHeader */
/* End_PlaceHolder_ModuleHeader */
var _hostName = 'Office';
var _defaultApiSetName = 'OfficeSharedApi';
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
var _typePersistentKvStorageService = 'PersistentKvStorageService';
/* Begin_PlaceHolder_PersistentKvStorageService_BeforeDeclaration */
exports.AsyncStorage = {
    getItem: function (key, callback) {
        return callStorageManager(function (storage, invokeId) { return storage.multiGet(invokeId, JSON.stringify([key])); }, function (result) {
            var parsedResult = JSON.parse(result);
            return parsedResult && parsedResult[0] && parsedResult[0][1] ? parsedResult[0][1] : null;
        }, callback);
    },
    setItem: function (key, value, callback) {
        return callStorageManager(function (storage, invokeId) { return storage.multiSet(invokeId, JSON.stringify([[key, value]])); }, function () { return null; }, callback);
    },
    removeItem: function (key, callback) {
        return callStorageManager(function (storage, invokeId) { return storage.multiRemove(invokeId, JSON.stringify([key])); }, function () { return null; }, callback);
    },
    multiGet: function (keys, callback) {
        return callStorageManager(function (storage, invokeId) { return storage.multiGet(invokeId, JSON.stringify(keys)); }, function (result) { return JSON.parse(result); }, callback);
    },
    multiSet: function (keyValuePairs, callback) {
        return callStorageManager(function (storage, invokeId) { return storage.multiSet(invokeId, JSON.stringify(keyValuePairs)); }, function () { return null; }, callback);
    },
    multiRemove: function (keys, callback) {
        return callStorageManager(function (storage, invokeId) { return storage.multiRemove(invokeId, JSON.stringify(keys)); }, function () { return null; }, callback);
    },
    getAllKeys: function (callback) {
        return callStorageManager(function (storage, invokeId) { return storage.getAllKeys(invokeId); }, function (result) { return JSON.parse(result); }, callback);
    },
    clear: function (callback) {
        return callStorageManager(function (storage, invokeId) { return storage.clear(invokeId); }, function () { return null; }, callback);
    }
};
function callStorageManager(nativeCall, getValueOnSuccess, callback) {
    return new OfficeExtension.CoreUtility.Promise(function (resolve, reject) {
        var storageManager = PersistentKvStorageManager.getInstance();
        var invokeId = storageManager.setCallBack(function (result, error) {
            if (error) {
                if (callback) {
                    callback(error);
                }
                reject(error);
                return;
            }
            // Otherwise:
            var value = getValueOnSuccess(result);
            if (callback) {
                callback(null, value);
            }
            resolve(value);
        });
        storageManager.ctx
            .sync()
            .then(function () {
            var storageService = storageManager.getPersistentKvStorageService();
            nativeCall(storageService, invokeId);
            return storageManager.ctx.sync();
        })
            .catch(function (e) {
            reject(e);
        });
    });
}
var PersistentKvStorageManager = /** @class */ (function () {
    function PersistentKvStorageManager() {
        var _this = this;
        this._invokeId = 0;
        this._callDict = {};
        this.ctx = new OfficeExtension.ClientRequestContext();
        this._perkvstorService = PersistentKvStorageService.newObject(this.ctx);
        this._eventResult = this._perkvstorService.onPersistentStorageMessage.add(function (args) {
            OfficeExtension.Utility.log('persistentKvStoragegMessageHandler:' + JSON.stringify(args));
            var callback = _this._callDict[args.invokeId];
            if (callback) {
                callback(args.message, args.error);
                delete _this._callDict[args.invokeId];
            }
        });
    }
    PersistentKvStorageManager.getInstance = function () {
        if (PersistentKvStorageManager.instance === undefined) {
            PersistentKvStorageManager.instance = new PersistentKvStorageManager();
        }
        else {
            PersistentKvStorageManager.instance._perkvstorService = PersistentKvStorageService.newObject(PersistentKvStorageManager.instance.ctx);
        }
        return PersistentKvStorageManager.instance;
    };
    PersistentKvStorageManager.prototype.getPersistentKvStorageService = function () {
        return this._perkvstorService;
    };
    PersistentKvStorageManager.prototype.getCallBack = function (callId) {
        return this._callDict[callId];
    };
    PersistentKvStorageManager.prototype.setCallBack = function (callback) {
        var id = this._invokeId;
        this._callDict[this._invokeId++] = callback;
        return id;
    };
    return PersistentKvStorageManager;
}());
/* End_PlaceHolder_PersistentKvStorageService_BeforeDeclaration */
/**
 *
 * Represents interface for PersistentStorageService
 *
 * [Api set: PersistentKvStorage 1.9]
 */
var PersistentKvStorageService = /** @class */ (function (_super) {
    __extends(PersistentKvStorageService, _super);
    function PersistentKvStorageService() {
        return _super !== null && _super.apply(this, arguments) || this;
    }
    Object.defineProperty(PersistentKvStorageService.prototype, "_className", {
        get: function () {
            return 'PersistentKvStorageService';
        },
        enumerable: true,
        configurable: true
    });
    /* Begin_PlaceHolder_PersistentKvStorageService_Custom_Members */
    /* End_PlaceHolder_PersistentKvStorageService_Custom_Members */
    // SET method is absent, because no settable properties on self or children.
    PersistentKvStorageService.prototype.clear = function (id) {
        /* Begin_PlaceHolder_PersistentKvStorageService_Clear */
        /* End_PlaceHolder_PersistentKvStorageService_Clear */
        _invokeMethod(this, 'Clear', 1 /* Read */, [id], 4 /* concurrent */, 0 /* none */);
    };
    PersistentKvStorageService.prototype.getAllKeys = function (id) {
        /* Begin_PlaceHolder_PersistentKvStorageService_GetAllKeys */
        /* End_PlaceHolder_PersistentKvStorageService_GetAllKeys */
        _invokeMethod(this, 'GetAllKeys', 1 /* Read */, [id], 4 /* concurrent */, 0 /* none */);
    };
    PersistentKvStorageService.prototype.multiGet = function (id, jsonKeys) {
        /* Begin_PlaceHolder_PersistentKvStorageService_MultiGet */
        /* End_PlaceHolder_PersistentKvStorageService_MultiGet */
        _invokeMethod(this, 'MultiGet', 1 /* Read */, [id, jsonKeys], 4 /* concurrent */, 0 /* none */);
    };
    PersistentKvStorageService.prototype.multiRemove = function (id, jsonKeys) {
        /* Begin_PlaceHolder_PersistentKvStorageService_MultiRemove */
        /* End_PlaceHolder_PersistentKvStorageService_MultiRemove */
        _invokeMethod(this, 'MultiRemove', 1 /* Read */, [id, jsonKeys], 4 /* concurrent */, 0 /* none */);
    };
    PersistentKvStorageService.prototype.multiSet = function (id, jsonKeyValue) {
        /* Begin_PlaceHolder_PersistentKvStorageService_MultiSet */
        /* End_PlaceHolder_PersistentKvStorageService_MultiSet */
        _invokeMethod(this, 'MultiSet', 1 /* Read */, [id, jsonKeyValue], 4 /* concurrent */, 0 /* none */);
    };
    /** Handle results returned from the document
     * @private
     */
    PersistentKvStorageService.prototype._handleResult = function (value) {
        _super.prototype._handleResult.call(this, value);
        if (_isNullOrUndefined(value))
            return;
        var obj = value;
        _fixObjectPathIfNecessary(this, obj);
        /* Begin_PlaceHolder_PersistentKvStorageService_HandleResult */
        /* End_PlaceHolder_PersistentKvStorageService_HandleResult */
    };
    /** Handle retrieve results
     * @private
     */
    PersistentKvStorageService.prototype._handleRetrieveResult = function (value, result) {
        _super.prototype._handleRetrieveResult.call(this, value, result);
        /* Begin_PlaceHolder_PersistentKvStorageService_HandleRetrieveResult */
        /* End_PlaceHolder_PersistentKvStorageService_HandleRetrieveResult */
        _processRetrieveResult(this, value, result);
    };
    /**
     * Create a new instance of PersistentKvStorageService object
     */
    PersistentKvStorageService.newObject = function (context) {
        return _createTopLevelServiceObject(PersistentKvStorageService, context, 'Microsoft.PersistentKvStorage.PersistentKvStorageService', false /*isCollection*/, 4 /* concurrent */);
    };
    Object.defineProperty(PersistentKvStorageService.prototype, "onPersistentStorageMessage", {
        /**
         *
         * Occurs when the Result is sent
         *
         * [Api set: PersistentKvStorage 1.9]
         */
        get: function () {
            /* Begin_PlaceHolder_PersistentKvStorageService_PersistentStorageMessage_get */
            /* End_PlaceHolder_PersistentKvStorageService_PersistentStorageMessage_get */
            if (!this.m_persistentStorageMessage) {
                /* Begin_PlaceHolder_PersistentKvStorageService_PersistentStorageMessage_get_PreInit */
                /* End_PlaceHolder_PersistentKvStorageService_PersistentStorageMessage_get_PreInit */
                /* Begin_PlaceHolder_PersistentKvStorageService_PersistentStorageMessage_get_EventHandlers */
                this.m_persistentStorageMessage = new OfficeExtension.GenericEventHandlers(
                /* End_PlaceHolder_PersistentKvStorageService_PersistentStorageMessage_get_EventHandlers */
                this.context, this, 'PersistentStorageMessage', 
                // Please add eventInfo between the placeholders
                /* Begin_PlaceHolder_PersistentKvStorageService_PersistentStorageMessage_Constructor_Parameters */
                {
                    eventType: 65537 /* perkvstorRichApiMessageEvent */,
                    registerFunc: function () { return void {}; },
                    unregisterFunc: function () { return void {}; },
                    getTargetIdFunc: function () {
                        return null;
                    },
                    eventArgsTransformFunc: function (args) {
                        var perkvstorArgs;
                        try {
                            var parsedMessage = JSON.parse(args.message);
                            var hr = parseInt(parsedMessage.errorCode);
                            var error = hr != 0 ? new OfficeExtension.Error(getErrorCodeAndMessage(hr)) : null;
                            perkvstorArgs = {
                                invokeId: parsedMessage.invokeId,
                                message: parsedMessage.message,
                                error: error
                            };
                        }
                        catch (e) {
                            perkvstorArgs = {
                                invokeId: -1,
                                message: e.message,
                                error: new OfficeExtension.Error({ code: 'GenericException', message: 'Unknown error' })
                            };
                        }
                        return OfficeExtension.Utility._createPromiseFromResult(perkvstorArgs);
                        // Helper
                        function getErrorCodeAndMessage(internalCode) {
                            var table = (_a = {},
                                _a[16389 /* persistentKvStorageEFail */] = {
                                    code: 'GenericException',
                                    message: 'Unknown error.'
                                },
                                _a[65535 /* persistentKvStorageEUnexpected */] = {
                                    code: 'Unexcepted',
                                    message: 'Catastrophic failure.'
                                },
                                _a[14 /* persistentKvStorageEOutOfMemory */] = {
                                    code: 'OutOfMemory',
                                    message: 'Ran out of memory.'
                                },
                                _a[87 /* persistentKvStorageEInvalidArg */] = {
                                    code: 'InvalidArg',
                                    message: 'One or more arguments are invalid.'
                                },
                                _a[16385 /* persistentKvStorageENotImpl */] = {
                                    code: 'NotImplemented',
                                    message: 'Not implemented.'
                                },
                                _a[6 /* persistentKvStorageEHandle */] = {
                                    code: 'BadHandle',
                                    message: 'File Handle is not Set.'
                                },
                                _a[5 /* persistentKvStorageEAccessDenied */] = {
                                    code: 'AccessDenied',
                                    message: "Can't read the AsyncStorage File."
                                },
                                _a);
                            if (table[internalCode]) {
                                return table[internalCode];
                            }
                            else {
                                return { code: 'Unknown', message: 'An unknown error has occured' };
                            }
                            var _a;
                        }
                    }
                }
                /* End_PlaceHolder_PersistentKvStorageService_PersistentStorageMessage_Constructor_Parameters */
                );
                /* Begin_PlaceHolder_PersistentKvStorageService_PersistentStorageMessage_get_AfterInit */
                /* End_PlaceHolder_PersistentKvStorageService_PersistentStorageMessage_get_AfterInit */
            }
            return this.m_persistentStorageMessage;
        },
        enumerable: true,
        configurable: true
    });
    PersistentKvStorageService.prototype.toJSON = function () {
        return _toJson(this, /* scalarProperties: */ {}, /* navigationProperties: */ {});
    };
    return PersistentKvStorageService;
}(OfficeExtension.ClientObject));
exports.PersistentKvStorageService = PersistentKvStorageService;
/* Begin_PlaceHolder_ErrorCodesTypeName */
var ErrorCodes;
(function (ErrorCodes) {
    /* End_PlaceHolder_ErrorCodesTypeName */
    ErrorCodes["generalException"] = "GeneralException";
    /* Begin_PlaceHolder_ErrorCodesAdditional */
    /* End_PlaceHolder_ErrorCodesAdditional */
})(ErrorCodes = exports.ErrorCodes || (exports.ErrorCodes = {}));


/***/ })
/******/ ]);
//# sourceMappingURL=officeruntimenative.g.js.map