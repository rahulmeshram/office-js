/*
	Copyright (c) Microsoft Corporation.  All rights reserved.
*/

/*
	Your use of this file is governed by the Microsoft Services Agreement http://go.microsoft.com/fwlink/?LinkId=266419.
*/

/*
* @overview es6-promise - a tiny implementation of Promises/A+.
* @copyright Copyright (c) 2014 Yehuda Katz, Tom Dale, Stefan Penner and contributors (Conversion to ES6 API by Jake Archibald)
* @license   Licensed under MIT license
*            See https://raw.githubusercontent.com/jakearchibald/es6-promise/master/LICENSE
* @version   2.3.0
*/


// Sources:
// osfweb: 16.0\11004.10000
// runtime: 16.0.11001.30000
// core: 16.0\11006.10000
// host: none



window.OfficeExtensionBatch = window.OfficeExtension;



!function(e){var n={};function t(r){if(n[r])return n[r].exports;var o=n[r]={i:r,l:!1,exports:{}};return e[r].call(o.exports,o,o.exports,t),o.l=!0,o.exports}t.m=e,t.c=n,t.d=function(e,n,r){t.o(e,n)||Object.defineProperty(e,n,{enumerable:!0,get:r})},t.r=function(e){"undefined"!=typeof Symbol&&Symbol.toStringTag&&Object.defineProperty(e,Symbol.toStringTag,{value:"Module"}),Object.defineProperty(e,"__esModule",{value:!0})},t.t=function(e,n){if(1&n&&(e=t(e)),8&n)return e;if(4&n&&"object"==typeof e&&e&&e.__esModule)return e;var r=Object.create(null);if(t.r(r),Object.defineProperty(r,"default",{enumerable:!0,value:e}),2&n&&"string"!=typeof e)for(var o in e)t.d(r,o,function(n){return e[n]}.bind(null,o));return r},t.n=function(e){var n=e&&e.__esModule?function(){return e.default}:function(){return e};return t.d(n,"a",n),n},t.o=function(e,n){return Object.prototype.hasOwnProperty.call(e,n)},t.p="",t(t.s=0)}([function(e,n,t){"use strict";Object.defineProperty(n,"__esModule",{value:!0});var r=t(1),o=t(2);window._OfficeRuntimeWeb={displayWebDialog:o.displayWebDialog,AsyncStorage:r}},function(e,n,t){"use strict";Object.defineProperty(n,"__esModule",{value:!0});var r="_Office_AsyncStorage_",o=r+"|_unusedKey_";function i(){window.localStorage.setItem(o,null),window.localStorage.removeItem(o)}function u(e,n){return void 0===n&&(n=function(){}),new Promise(function(t,r){try{i(),e(),n(null),t()}catch(e){n(e),r(e)}})}function c(e,n){return void 0===n&&(n=function(){}),new Promise(function(t,r){try{i();var o=e();n(null,o),t(o)}catch(e){n(e,null),r(e)}})}function a(e,n,t){return void 0===t&&(t=function(){}),new Promise(function(r,o){var u=[];try{i()}catch(e){u.push(e)}e.forEach(function(e){try{n(e)}catch(e){u.push(e)}}),t(u),u.length>0?o(u):r()})}n.getItem=function(e,n){return c(function(){return window.localStorage.getItem(r+e)},n)},n.setItem=function(e,n,t){return u(function(){return window.localStorage.setItem(r+e,n)},t)},n.removeItem=function(e,n){return u(function(){return window.localStorage.removeItem(r+e)},n)},n.clear=function(e){return u(function(){Object.keys(window.localStorage).filter(function(e){return 0===e.indexOf(r)}).forEach(function(e){return window.localStorage.removeItem(e)})},e)},n.getAllKeys=function(e){return c(function(){return Object.keys(window.localStorage).filter(function(e){return 0===e.indexOf(r)}).map(function(e){return e.substr(r.length)})},e)},n.multiSet=function(e,n){return a(e,function(e){var n=e[0],t=e[1];return window.localStorage.setItem(r+n,t)},n)},n.multiRemove=function(e,n){return a(e,function(e){return window.localStorage.removeItem(r+e)},n)},n.multiGet=function(e,n){return new Promise(function(t,o){n||(n=function(){});var i=[],u=e.map(function(e){try{return[e,window.localStorage.getItem(r+e)]}catch(e){i.push(e)}}).filter(function(e){return e});i.length>0?(n(i,u),o(i)):(n(null,u),t(u))})}},function(e,n,t){"use strict";Object.defineProperty(n,"__esModule",{value:!0});var r=t(3),o=function(){function e(e){this._dialog=e}return e.prototype.close=function(){return this._dialog.close(),r.CoreUtility.Promise.resolve()},e}();n.Dialog=o,n.displayWebDialog=function(e,n){return new r.CoreUtility.Promise(function(t,i){if(n.width&&n.height&&(!f(n.width)||!f(n.height)))throw new r.Error({code:"InvalidArgument",message:'Dimensions must be "number%" or number.'});var u,c={width:n.width?parseInt(n.width,10):50,height:n.height?parseInt(n.height,10):50,displayInIframe:n.displayInIFrame||!1};function a(e){n.onMessage&&n.onMessage(e.message,u)}function l(e){12006===e.error?n.onClose&&n.onClose():n.onRuntimeError&&n.onRuntimeError(new r.Error(s(e.error)),u)}function f(e){return/^(\-|\+)?([0-9]+)%?$/.test(e)}function s(e){var n,t=((n={})[12002]={code:"InvalidUrl",message:"Cannot load URL, no such page or bad URL syntax."},n[12003]={code:"InvalidUrl",message:"HTTPS is required."},n[12004]={code:"Untrusted",message:"Domain is not trusted."},n[12005]={code:"InvalidUrl",message:"HTTPS is required."},n[12007]={code:"FailedToOpen",message:"Another dialog is already opened."},n);return t[e]?t[e]:{code:"Unknown",message:"An unknown error has occured"}}Office.context.ui.displayDialogAsync(e,c,function(e){"failed"===e.status?i(new r.Error(s(e.error.code))):((u=e.value).addEventHandler(Office.EventType.DialogMessageReceived,a),u.addEventHandler(Office.EventType.DialogEventReceived,l),t(new o(u)))})})}},function(e,n){e.exports=OfficeExtensionBatch}]);



window.OfficeRuntime = window._OfficeRuntimeWeb;