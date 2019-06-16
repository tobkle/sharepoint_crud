/**
 * @name Sharepoint CRUD
 * @description Provides Create, Read, Update, Delete Opertions on Sharepoint Lists
 * @author Tobias Klemmer <tobias.klemmer@wuerth-it.com>
 * @date 2019-06-16
 */

(function (global, factory) {
    typeof exports === 'object' && typeof module !== 'undefined' ? factory(exports) :
        typeof define === 'function' && define.amd ? define(['exports'], factory) :
            (factory((global.SPCRUD = {})));
}(this, (function (exports) {

    // Encode Ajax Url Parameters
    function param(object) {
        var encodedString = '?';
        for (var prop in object) {
            if (object.hasOwnProperty(prop)) {
                if (encodedString.length > 0) {
                    encodedString += '&';
                }
                encodedString += encodeURI(prop + '=' + object[prop]);
            }
        }
        return encodedString;
    };

    // Template for Error Handling
    function Error(message) {
        this.message = message;
    };

    // Perform an AJAX call
    function ajax(parameters) {
        // for Internet Explorer
        var url     = parameters.url     ? parameters.url : null;
        var type    = parameters.type    ? parameters.type : null; 
        var headers = parameters.headers ? parameters.headers : null;
        var params  = parameters.params  ? parameters.params : null; 
        var data    = parameters.data    ? parameters.data : {};
        // Deliver a Promise for a non-blocking asynchronous call
        return new Promise(function (resolve, reject) {
            var urlWithParams = params ? url + param(params) : url;
            var xhr = new XMLHttpRequest();
            xhr.open(type, urlWithParams);
            xhr.withCredentials = true;
            if (headers) {
                if (headers.length) {
                    headers.forEach(function (header) {
                        Object.keys(header).forEach(function (key) {
                            xhr.setRequestHeader(key, header[key]);
                        });
                    });
                } else {
                    Object.keys(headers).forEach(function (key) {
                        xhr.setRequestHeader(key, headers[key]);
                    });
                }
            }
            xhr.onload = function () {
                if (xhr.status >= 200 && xhr.status < 300) {
                    resolve(xhr.responseText);
                } else {
                    reject({ code: xhr.status, message: xhr.responseText });
                }
            };
            xhr.send(data);
        });
    };

    // Try to get List Details
    function getLists(siteUrl) {
        return ajax({
            url: siteUrl + "/_api/Web/Lists",
            type: "GET",
            headers: {
                "accept": "application/json;odata=verbose",
                "content-Type": "application/json;odata=verbose"
            },
            params: {},
            data: JSON.stringify({})
        });
    };

    // Sign later Calls with this Form Digest, otherwise 403 FORBIDDEN response
    function getFormDigest(siteUrl) {
        return ajax({
            url: siteUrl + "/_api/contextinfo",
            type: 'POST',
            headers: {
                "Accept": "application/json; odata=verbose"
            },
            data: JSON.stringify({})
        });
    };

    // Try to reach the Sharepoint Site and create a Library Object therefore,
    // register the CRUD operations for this Library for easier handling
    this.connect = function (siteUrl) {
        return new Promise(function (resolve, reject) {
            if (!siteUrl) reject('connect: Please provide a valid Sharepoint Site URL.');
            var newSite = {};
            return getLists(siteUrl)
            .then(function(response){
                var Lists = JSON.parse(response).d.results;
                if (Lists && Lists.length){
                    Lists.forEach(function(List){
                        newSite[List.Title] = List;
                        newSite[List.Title].siteUrl = siteUrl;
                        newSite[List.Title].createListItem = this.createListItem.bind(List);
                        newSite[List.Title].readListItem = this.readListItem.bind(List);
                        newSite[List.Title].readListItems = this.readListItems.bind(List);
                        newSite[List.Title].updateListItem = this.updateListItem.bind(List);
                        newSite[List.Title].deleteListItem = this.deleteListItem.bind(List);
                        newSite[List.Title].getListFields = this.getListFields.bind(List);
                        exports[List.Title] = newSite[List.Title];
                    });
                }
                resolve(newSite);
            }).catch(function (error) {
                reject(error);
            });
        });
    };
    
    // Create a new list item
    this.createListItem = function (newItem) {
        var siteUrl = this.siteUrl;
        var Id = this.Id;
        if (!newItem) throw new Error("createListItem: Please provide a new item.")
        if (!newItem["__metadata"]) {
            newItem["__metadata"] = {
                type: this.ListItemEntityTypeFullName
            };
        }
        return getFormDigest(siteUrl)
            .then(function (response) {
                return JSON.parse(response);
            })
            .then(function (data) {
                return ajax({
                    url: siteUrl + "/_api/Web/Lists(guid'" + Id + "')/items",
                    type: "POST",
                    headers: {
                        "accept": "application/json;odata=verbose",
                        "X-RequestDigest": data.d.GetContextWebInformation.FormDigestValue,
                        "content-Type": "application/json;odata=verbose"
                    },
                    data: JSON.stringify(newItem)
                });
            });
    };
        
    // Read list items
    this.readListItems = function (numberOfRecords) {
        var siteUrl = this.siteUrl;
        var Id = this.Id;
        var top = numberOfRecords ? parseInt(numberOfRecords) : 9999; // 9999 is maximum
        return ajax({
            url: siteUrl + "/_api/Web/Lists(guid'" + Id + "')/items",
            type: "GET",
            headers: {
                "accept": "application/json;odata=verbose",
                "content-Type": "application/json;odata=verbose"
            },
            params: {
                top: top
            },
            data: JSON.stringify({})
        });
    };
        
    // Read a specific list item by Id
    this.readListItem = function (itemId) {
        var siteUrl = this.siteUrl;
        var Id = this.Id;
        if (!itemId) throw new Error('readListItem: Please provide an item id.');
        return ajax({
            url: siteUrl + "/_api/Web/Lists(guid'" + Id + "')/GetItemById('" + itemId + "')",
            type: "GET",
            headers: {
                "accept": "application/json;odata=verbose",
                "content-Type": "application/json;odata=verbose"
            },
            data: JSON.stringify({})
        });
    };
        
    // Update a specific list item by Id
    this.updateListItem = function (itemId, updateProperties) {
        var siteUrl = this.siteUrl;
        var Id = this.Id;
        if (!itemId) throw new Error('updateListItem: Please provide an item id.');
        if (!updateProperties) console.error('updateListItem: Please provide update properties.');
        var oldItem = null;
        return this.readListItem(itemId)
        .then(function (oldResponse) {
            oldItem = JSON.parse(oldResponse).d;
            updateProperties["__metadata"] = oldItem.__metadata;
            return getFormDigest(siteUrl);
        })
        .then(function (response) {
            return JSON.parse(response);
        })
        .then(function (data) {
            return ajax({
                url: siteUrl + "/_api/Web/Lists(guid'" + Id + "')/GetItemById('" + itemId + "')",
                type: "PATCH",
                headers: {
                    "Accept": "application/json;odata=verbose",
                    "Content-Type": "application/json;odata=verbose",
                    "X-RequestDigest": data.d.GetContextWebInformation.FormDigestValue,
                    "X-Http-Method": "PATCH",
                    "If-Match": oldItem && oldItem.__metadata && oldItem.__metadata.etag ? oldItem.__metadata.etag : "*"
                },
                data: JSON.stringify(updateProperties)
            });
        });
    };
        
    // Delete a specific list item by Id
    this.deleteListItem = function (itemId) {
        var siteUrl = this.siteUrl;
        var Id = this.Id;
        if (!itemId) throw new Error('updateListItem: Please provide an item id.');
        var oldItem = null;
        return this.readListItem(itemId)
        .then(function (oldResponse) {
            oldItem = JSON.parse(oldResponse).d;
            return getFormDigest(siteUrl);
        })
        .then(function (response) {
            return JSON.parse(response);
        })
        .then(function (data) {
            return ajax({
                url: siteUrl + "/_api/Web/Lists(guid'" + Id + "')/GetItemById('" + itemId + "')",
                type: "DELETE",
                headers: {
                    "Accept": "application/json;odata=verbose",
                    "Content-Type": "application/json;odata=verbose",
                    "X-RequestDigest": data.d.GetContextWebInformation.FormDigestValue,
                    "X-Http-Method": "DELETE",
                    "If-Match": oldItem && oldItem.__metadata && oldItem.__metadata.etag ? oldItem.__metadata.etag : "*"
                },
                data: JSON.stringify({})
            });
        });
    };

    // Try to get List Fields Details
    this.getListFields = function () {
        var siteUrl = this.Fields.__deferred.uri;
        return ajax({
            url: siteUrl,
            type: "GET",
            headers: {
                "accept": "application/json;odata=verbose",
                "content-Type": "application/json;odata=verbose"
            },
            params: {},
            data: JSON.stringify({})
        });
    };
        

    exports.connect = this.connect;

    Object.defineProperty(exports, '__esModule', { value: true });
})));

// Promise Polyfill for Internet Explorer
(function (global, factory) {
    typeof exports === 'object' && typeof module !== 'undefined' ? factory() :
        typeof define === 'function' && define.amd ? define(factory) :
            (factory());
}(this, (function () {
    'use strict';

    /**
     * @this {Promise}
     */
    function finallyConstructor(callback) {
        var constructor = this.constructor;
        return this.then(
            function (value) {
                // @ts-ignore
                return constructor.resolve(callback()).then(function () {
                    return value;
                });
            },
            function (reason) {
                // @ts-ignore
                return constructor.resolve(callback()).then(function () {
                    // @ts-ignore
                    return constructor.reject(reason);
                });
            }
        );
    }

    // Store setTimeout reference so promise-polyfill will be unaffected by
    // other code modifying setTimeout (like sinon.useFakeTimers())
    var setTimeoutFunc = setTimeout;

    function isArray(x) {
        return Boolean(x && x.length);
    }

    function noop() { }

    // Polyfill for Function.prototype.bind
    function bind(fn, thisArg) {
        return function () {
            fn.apply(thisArg, arguments);
        };
    }

    /**
     * @constructor
     * @param {Function} fn
     */
    function Promise(fn) {
        if (!(this instanceof Promise))
            throw new TypeError('Promises must be constructed via new');
        if (typeof fn !== 'function') throw new TypeError('not a function');
        /** @type {!number} */
        this._state = 0;
        /** @type {!boolean} */
        this._handled = false;
        /** @type {Promise|undefined} */
        this._value = undefined;
        /** @type {!Array<!Function>} */
        this._deferreds = [];

        doResolve(fn, this);
    }

    function handle(self, deferred) {
        while (self._state === 3) {
            self = self._value;
        }
        if (self._state === 0) {
            self._deferreds.push(deferred);
            return;
        }
        self._handled = true;
        Promise._immediateFn(function () {
            var cb = self._state === 1 ? deferred.onFulfilled : deferred.onRejected;
            if (cb === null) {
                (self._state === 1 ? resolve : reject)(deferred.promise, self._value);
                return;
            }
            var ret;
            try {
                ret = cb(self._value);
            } catch (e) {
                reject(deferred.promise, e);
                return;
            }
            resolve(deferred.promise, ret);
        });
    }

    function resolve(self, newValue) {
        try {
            // Promise Resolution Procedure: https://github.com/promises-aplus/promises-spec#the-promise-resolution-procedure
            if (newValue === self)
                throw new TypeError('A promise cannot be resolved with itself.');
            if (
                newValue &&
                (typeof newValue === 'object' || typeof newValue === 'function')
            ) {
                var then = newValue.then;
                if (newValue instanceof Promise) {
                    self._state = 3;
                    self._value = newValue;
                    finale(self);
                    return;
                } else if (typeof then === 'function') {
                    doResolve(bind(then, newValue), self);
                    return;
                }
            }
            self._state = 1;
            self._value = newValue;
            finale(self);
        } catch (e) {
            reject(self, e);
        }
    }

    function reject(self, newValue) {
        self._state = 2;
        self._value = newValue;
        finale(self);
    }

    function finale(self) {
        if (self._state === 2 && self._deferreds.length === 0) {
            Promise._immediateFn(function () {
                if (!self._handled) {
                    Promise._unhandledRejectionFn(self._value);
                }
            });
        }

        for (var i = 0, len = self._deferreds.length; i < len; i++) {
            handle(self, self._deferreds[i]);
        }
        self._deferreds = null;
    }

    /**
     * @constructor
     */
    function Handler(onFulfilled, onRejected, promise) {
        this.onFulfilled = typeof onFulfilled === 'function' ? onFulfilled : null;
        this.onRejected = typeof onRejected === 'function' ? onRejected : null;
        this.promise = promise;
    }

    /**
     * Take a potentially misbehaving resolver function and make sure
     * onFulfilled and onRejected are only called once.
     *
     * Makes no guarantees about asynchrony.
     */
    function doResolve(fn, self) {
        var done = false;
        try {
            fn(
                function (value) {
                    if (done) return;
                    done = true;
                    resolve(self, value);
                },
                function (reason) {
                    if (done) return;
                    done = true;
                    reject(self, reason);
                }
            );
        } catch (ex) {
            if (done) return;
            done = true;
            reject(self, ex);
        }
    }

    Promise.prototype['catch'] = function (onRejected) {
        return this.then(null, onRejected);
    };

    Promise.prototype.then = function (onFulfilled, onRejected) {
        // @ts-ignore
        var prom = new this.constructor(noop);

        handle(this, new Handler(onFulfilled, onRejected, prom));
        return prom;
    };

    Promise.prototype['finally'] = finallyConstructor;

    Promise.all = function (arr) {
        return new Promise(function (resolve, reject) {
            if (!isArray(arr)) {
                return reject(new TypeError('Promise.all accepts an array'));
            }

            var args = Array.prototype.slice.call(arr);
            if (args.length === 0) return resolve([]);
            var remaining = args.length;

            function res(i, val) {
                try {
                    if (val && (typeof val === 'object' || typeof val === 'function')) {
                        var then = val.then;
                        if (typeof then === 'function') {
                            then.call(
                                val,
                                function (val) {
                                    res(i, val);
                                },
                                reject
                            );
                            return;
                        }
                    }
                    args[i] = val;
                    if (--remaining === 0) {
                        resolve(args);
                    }
                } catch (ex) {
                    reject(ex);
                }
            }

            for (var i = 0; i < args.length; i++) {
                res(i, args[i]);
            }
        });
    };

    Promise.resolve = function (value) {
        if (value && typeof value === 'object' && value.constructor === Promise) {
            return value;
        }

        return new Promise(function (resolve) {
            resolve(value);
        });
    };

    Promise.reject = function (value) {
        return new Promise(function (resolve, reject) {
            reject(value);
        });
    };

    Promise.race = function (arr) {
        return new Promise(function (resolve, reject) {
            if (!isArray(arr)) {
                return reject(new TypeError('Promise.race accepts an array'));
            }

            for (var i = 0, len = arr.length; i < len; i++) {
                Promise.resolve(arr[i]).then(resolve, reject);
            }
        });
    };

    // Use polyfill for setImmediate for performance gains
    Promise._immediateFn =
        // @ts-ignore
        (typeof setImmediate === 'function' &&
            function (fn) {
                // @ts-ignore
                setImmediate(fn);
            }) ||
        function (fn) {
            setTimeoutFunc(fn, 0);
        };

    Promise._unhandledRejectionFn = function _unhandledRejectionFn(err) {
        if (typeof console !== 'undefined' && console) {
            console.warn('Possible Unhandled Promise Rejection:', err); // eslint-disable-line no-console
        }
    };

    /** @suppress {undefinedVars} */
    var globalNS = (function () {
        // the only reliable means to get the global object is
        // `Function('return this')()`
        // However, this causes CSP violations in Chrome apps.
        if (typeof self !== 'undefined') {
            return self;
        }
        if (typeof window !== 'undefined') {
            return window;
        }
        if (typeof global !== 'undefined') {
            return global;
        }
        throw new Error('unable to locate global object');
    })();

    if (!('Promise' in globalNS)) {
        globalNS['Promise'] = Promise;
    } else if (!globalNS.Promise.prototype['finally']) {
        globalNS.Promise.prototype['finally'] = finallyConstructor;
    }

})));
