/**
 * @fileOverview spread sheet converter
 * @name index.js
 * @author Yuhei Aihara <aihara_yuhei@cyberagent.co.jp>
 * https://github.com/yuhei-a/spread2json
 */
var _ = require('lodash');
var async = require('async');

var api = require('./api');
var logger = require('./logger');

function Spread2Json() {
    this.opts = {
        // Cell with a custom sheet option.
        option_cell: 'A1',
        // Line wutg a data atribute.
        attr_line: 2,
        // Line with a data.
        data_line: 4,
        // ref key
        ref_key: '_id',
        // Custom logger.
        logger: undefined,
        // Googgle API options.
        api: {
            client_id: undefined,
            client_secret: undefined,
            redirect_url: 'http://localhost',
            token_path: './dist/token.json'
        }
    };
    this.logger = logger;
}

module.exports = new Spread2Json();


var parser = {
    number: function(d) {
        return Number(d);
    },
    boolean: function(d) {
        return !!d && d.toLowerCase() !== 'false';
    },
    date: function(d) {
        return new Date(d).getTime();
    }
};

/**
 * API function
 */
Spread2Json.prototype.generateAuthUrl = function() { return api.generateAuthUrl.apply(api, arguments); };
Spread2Json.prototype.getAccessToken = function() { return api.getAccessToken.apply(api, arguments); };
Spread2Json.prototype.refreshAccessToken = function() { return api.refreshAccessToken.apply(api, arguments); };

/**
 * setup
 * @param {Object} options
 * @example
 * var options = {
 *     option_cell: 'A1',
 *     attr_line: 2,
 *     data_line: 4,
 *     ref_key: '_id',
 *     api: {
 *         client_id: 'YOUR CLIENT ID HERE',
 *         client_secret: 'YOUR CLIENT SECRET HERE',
 *         redirect_url: 'http://localhost',
 *         token_path: './dist/token.json'
 *     },
 *     logger: CustomLogger
 * }
 */
Spread2Json.prototype.setup = function(options) {
    this.logger.info('setup');
    _.extend(this.opts, options);
    if (this.opts.logger) {
        this.logger = this.opts.logger;
    }
    api.setup.call(api, this.opts.api);
};

/**
 * get worksheet info in spreadsheet
 * @param {String} key spreadsheetKey
 * @param {Function} callback
 * @example
 * > url = 'https://docs.google.com/spreadsheets/d/1YXVzaaxqkPKsr-excIOXScnTQC7y_DKrUKs0ukzSIgo/edit#gid=0'
 * key = '1YXVzaaxqkPKsr-excIOXScnTQC7y_DKrUKs0ukzSIgo'
 */
Spread2Json.prototype.getWorksheet = function(key, callback) {
    api.getWorksheet(key, function(err, result) {
        if (err) {
            return callback(err);
        }
        var entry = result && result.feed && result.feed.entry;
        if (!entry) {
            return callback(result);
        }

        var list = [];
        for (var i = 0; i < entry.length; i++) {
            var worksheet = entry[i];
            list.push({
                id: worksheet.id && worksheet.id.$t && worksheet.id.$t.replace(/.+\/([^\/]+)$/, '$1'),
                updated: worksheet.updated && worksheet.updated.$t,
                title: worksheet.title && worksheet.title.$t
            });
        }
        callback(null, list);
    });
};

/**
 * get worksheets data
 * @param {String} key spreadsheetKey
 * @param {Array} worksheetIds
 * @param {Function} callback
 * @see #getWorksheet arguments.key and returns list.id
 */
Spread2Json.prototype.getWorksheetDatas = function(key, worksheetIds, callback) {
    var _this = this;
    async.map(worksheetIds, function(worksheetId, next) {
        api.getCells(key, worksheetId, function(err, result) {
            if (err) {
                return next(err);
            }

            var sheetName = result && result.feed.title.$t;
            var entry = result && result.feed.entry;
            if (!entry) {
                _this.logger.error('not found entry.', worksheetId);
                return next();
            }
            
            var opts, list = [], idx = -1, originRow;
            _.extend(opts = {}, {
                attr_line: _this.opts.attr_line,
                data_line: _this.opts.data_line,
                ref_key: _this.opts.ref_key
            });
            entry.forEach(function(cell) {
                var title = cell.title.$t.match(/^(\D+)(\d+)$/);
                var column = title[1];
                var row = Number(title[2]);
                var content = cell.content.$t;
                if (cell.title.$t === _this.opts.option_cell) {
                    var _opts;
                    try {
                        _opts = JSON.parse(content) || {};
                    } catch (e) {
                        _opts = {};
                    }
                    _.extend(opts, _opts);
                    return;
                }

                if (row === opts.attr_line) {
                    var type = content.match(/:(\w+)$/);
                    var keys = content.replace(/:\w+$/, '').split('.');
                    opts.format = opts.format || {};
                    opts.format[column] = {
                        type: type && type[1],
                        keys: keys
                    };
                    return;
                }

                var format = opts.format && opts.format[column];
                if (row < opts.data_line || !format) {
                    return;
                }

                if (column === 'A') {
                    idx++;
                    originRow = row;
                    list[idx] = {};
                }

                var _idx = row - originRow;
                var data = list[idx];
                format.keys.forEach(function(_key, i) {
                    var isArray = /^#/.test(_key);
                    var isSplitArray = /^\$/.test(_key);
                    if (isArray) {
                        _key = _key.replace(/^#/, '');
                        data[_key] = data[_key] || [];
                    }
                    if (isSplitArray) {
                        _key = _key.replace(/^\$/, '');
                    }

                    if (i + 1 !== format.keys.length) {
                        if (isArray) {
                            data = data[_key][_idx] = data[_key][_idx] || {};
                            return;
                        }
                        data = data[_key] = data[_key] || {};
                        return;
                    }

                    if (isArray) {
                        var __key = data[_key].length;
                        data = data[_key];
                        _key = __key;
                    }

                    var type = format.type && format.type.toLowerCase();
                    if (type === 'number' || type === 'num') {
                        data[_key] = isSplitArray ? content.split(',').map(parser.number) : parser.number(content);
                    } else if (type === 'boolean' || type === 'bool') {
                        data[_key] = isSplitArray ? content.split(',').map(parser.boolean) : parser.boolean(content);
                    } else if (type === 'date') {
                        data[_key] = isSplitArray ? content.split(',').map(parser.date) : parser.date(content);
                        data[_key] = new Date(content).getTime();
                    } else {
                        data[_key] = isSplitArray ? content.split(',') : content;
                    }
                });
            });

            next(null, {
                name: sheetName,
                opts: opts,
                list: list
            });
        });
    }, callback);
};

/**
 *
 * @param dataMap
 * @param opts
 * @param data
 * @private
 */
Spread2Json.prototype._findOrigin = function(dataMap, opts, data) {
    var origin = dataMap[data.__ref];
    if (!origin || !opts.key) {
        this.logger.error('not found origin.', JSON.stringify(data));
        return;
    }

    var keys = opts.key.split('.');
    var __in = data.__in ? data.__in.split('.') : [];
    for (var i = 0; i < keys.length; i++) {
        if (/^#/.test(keys[i])) {
            var key = keys[i].replace(/^#/, '');
            var index = __in[i] && __in[i].replace(/^#.+:(\d+)$/, '$1');
            if (!index) {
                this.logger.error('not found index.', JSON.stringify(data));
                return;
            }
            origin[key] = origin[key] || [];
            origin = origin[key];
            origin[index] = origin[index] || {};
            origin = origin[index];
        } else if (keys[i] === '$') {
            origin = origin[__in[i]];
        } else if (i + 1 === keys.length) {
            origin[keys[i]] = origin[keys[i]] || (opts.type === 'array' ? [] : {});
            origin = origin[keys[i]];
        } else {
            origin = origin[keys[i]];
        }
        if (!origin) {
            this.logger.error('not found origin parts.', JSON.stringify(data));
            return;
        }
    }

    if (opts.type === 'array') {
        if (!Array.isArray(origin)) {
            this.logger.error('is not Array.', JSON.stringify(data));
            return;
        }
        origin.push({});
        origin = origin[origin.length - 1];
    } else if (opts.type === 'map') {
        if (!data.__key) {
            this.logger.error('not found __key.', JSON.stringify(data));
            return;
        }
        origin = origin[data.__key] = {};
    } else {
        this.logger.error(opts);
        return;
    }
    return origin;
};

/**
 * sheetDatas to json
 * @param {Array} sheetDatas
 * @param {Function} callback
 */
Spread2Json.prototype.toJson = function(sheetDatas, callback) {
    var collectionMap = {};
    var errors = {};
    for (var i = 0; i < sheetDatas.length; i++) {
        var sheetData = sheetDatas[i];
        var name = sheetData.opts.name || sheetData.name;
        var refKey = sheetData.opts.ref_key;
        var dataMap = collectionMap[name] = collectionMap[name] || {};
        for (var j = 0; j < sheetData.list.length; j++) {
            var data = sheetData.list[j];
            if (!sheetData.opts.type || sheetData.opts.type === 'origin') {
                dataMap[data[refKey]] = data;
            } else {
                var origin = this._findOrigin(dataMap, sheetData.opts, data);
                if (origin) {
                    delete data.__ref;
                    delete data.__in;
                    delete data.__key;
                    _.extend(origin, data);
                } else {
                    errors[name] = errors[name] || [];
                    errors[name].push(data);
                }
            }
        }
    }

    callback(Object.keys(errors).length ? errors : null, collectionMap);
};
