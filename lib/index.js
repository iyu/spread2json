/**
 * @fileOverview spread sheet converter
 * @name index.js
 * @author Yuhei Aihara <yu.e.yu.4119@gmail.com>
 * https://github.com/iyu/spread2json
 */
var _ = require('lodash');
var async = require('async');

var api = require('./api');
var logging = require('./logging');

function Spread2Json() {
  this.opts = {
    // Cell with a custom sheet option.
    option_cell: 'A1',
    // Line with a data attribute.
    attr_line: 2,
    // Line with a data.
    data_line: 4,
    // ref key
    ref_key: '_id',
    // Custom logger.
    logger: undefined,
    // Google API options.
    api: {
      client_id: undefined,
      client_secret: undefined,
      redirect_url: 'http://localhost',
      token_file: {
        use: true,
        path: './dist/token.json'
      }
    }
  };
}

module.exports = new Spread2Json();

/**
 * API function
 */
Spread2Json.prototype.generateAuthUrl = function() { return api.generateAuthUrl.apply(api, arguments); };
Spread2Json.prototype.getAccessToken = function() { return api.getAccessToken.apply(api, arguments); };
Spread2Json.prototype.refreshAccessToken = function() { return api.refreshAccessToken.apply(api, arguments); };

/**
 * setup
 * @param {Object} options
 * @param {string} [options.option_cell='A1']
 * @param {number} [options.attr_line=2]
 * @param {number} [options.data_line=4]
 * @param {string} [options.ref_key='_id']
 * @param {Object} options.api
 * @param {string} options.api.client_id
 * @param {string} options.api.client_secret
 * @param {string} [options.api.redirect_url='http://localhost']
 * @param {Object} [options.api.token_file]
 * @param {boolean} [options.api.token_file.use=true]
 * @param {string} [options.api.token_file.path='./dist/token.json']
 * @param {Logger} [logger]
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
 *         token_file: {
 *           use: true
 *           path: './dist/token.json'
 *         }
 *     },
 *     logger: CustomLogger
 * }
 */
Spread2Json.prototype.setup = function(options) {
  _.extend(this.opts, options);

  if (this.opts.logger) {
    logging.logger = this.opts.logger;
  }

  api.setup.call(api, this.opts.api);
  logging.logger.info('spread2json setup');
};

/**
 * get spreadsheets info
 * @param {Credentials} [credentials] - oAuth2.getToken result
 * @param {Function} callback
 */
Spread2Json.prototype.getSpreadsheet = function(credentials, callback) {
  if (arguments.length === 1) {
    callback = credentials;
    credentials = undefined;
  }
  api.getSpreadsheet(credentials, function(err, result) {
    if (err) {
      return callback(err);
    }
    var entry = result && result.feed && result.feed.entry;
    if (!entry) {
      return callback(result);
    }

    var list = [];
    for (var i = 0; i < entry.length; i++) {
      var spreadsheet = entry[i];
      list.push({
        id: spreadsheet.id && spreadsheet.id.$t && spreadsheet.id.$t.replace(/.+\/([^\/]+)$/, '$1'),
        updated: spreadsheet.updated && spreadsheet.updated.$t,
        title: spreadsheet.title && spreadsheet.title.$t,
        link: spreadsheet.link && spreadsheet.link[0] && spreadsheet.link[0].href
      });
    }
    callback(null, list);
  });
};

/**
 * get worksheet info in spreadsheet
 * @param {Credentials} [credentials] - oAuth2.getToken result
 * @param {string} key - spreadsheetKey
 * @param {Function} callback
 * @example
 * > url = 'https://docs.google.com/spreadsheets/d/1YXVzaaxqkPKsr-excIOXScnTQC7y_DKrUKs0ukzSIgo/edit#gid=0'
 * key = '1YXVzaaxqkPKsr-excIOXScnTQC7y_DKrUKs0ukzSIgo'
 */
Spread2Json.prototype.getWorksheet = function(credentials, key, callback) {
  if (arguments.length === 2) {
    callback = key;
    key = credentials;
    credentials = undefined;
  }
  api.getWorksheet(credentials, key, function(err, result) {
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
 * parser
 * @private
 */
Spread2Json.prototype._parser = {
  number: function(d) {
    return Number(d);
  },
  num: function(d) {
    return this.number(d);
  },
  boolean: function(d) {
    return !!d && d.toLowerCase() !== 'false' && d !== '0';
  },
  bool: function(d) {
    return this.boolean(d);
  },
  date: function(d) {
    return new Date(d).getTime();
  },
  auto: function(d) {
    return isNaN(d) ? d : this.number(d);
  }
};

/**
 * format
 * @param {Array} cells
 * @private
 * @example
 * var cells = [
 *     { cell: 'A1', value: '{}' }, { cell: 'A4', value: '_id' },,,
 * ]
 */
Spread2Json.prototype._format = function(cells) {
  var self = this;
  var opts = {};
  var beforeRow;
  var idx = {};
  var list = [];

  _.extend(opts, {
    attr_line: this.opts.attr_line,
    data_line: this.opts.data_line,
    ref_key: this.opts.ref_key
  });

  _.forEach(cells, function(cell) {
    if (beforeRow !== cell.row) {
      _.each(idx, function(i) {
        if (i.type !== 'format') {
          i.value += 1;
        }
      });
      beforeRow = cell.row;
    }
    if (cell.cell === self.opts.option_cell) {
      var _opts = JSON.parse(cell.value);
      _.extend(opts, _opts);
      return;
    }

    if (cell.row === opts.attr_line) {
      var type = cell.value.match(/:(\w+)$/);
      var key = cell.value.replace(/:\w+$/, '');
      var keys = key.split('.');

      opts.format = opts.format || {};
      opts.format[cell.column] = {
        type: type && type[1],
        key: key,
        keys: keys
      };
      return;
    }

    var format = opts.format && opts.format[cell.column];
    var data;
    var _idx;

    if (cell.row < opts.data_line || !format) {
      return;
    }
    if (format.type && format.type.toLowerCase() === 'index') {
      _idx = parseInt(cell.value, 10);
      if (!idx[format.key] || idx[format.key].value !== _idx) {
        idx[format.key] = {
          type: 'format',
          value: _idx
        };
        _.each(idx, function(i, key) {
          if (new RegExp('^' + format.key + '.+$').test(key)) {
            idx[key].value = 0;
          }
        });
      }
      return;
    }
    if (format.key === opts.ref_key || format.key === '__ref') {
      idx = {};
      list.push({});
    }

    data = _.last(list);
    _.each(format.keys, function(_key, i) {
      var isArray = /^#/.test(_key);
      var isSplitArray = /^\$/.test(_key);
      var __key;
      if (isArray) {
        _key = _key.replace(/^#/, '');
        data[_key] = data[_key] || [];
      }
      if (isSplitArray) {
        _key = _key.replace(/^\$/, '');
      }

      if (i + 1 !== format.keys.length) {
        if (isArray) {
          __key = format.keys.slice(0, i + 1).join('.');
          _idx = idx[__key];
          if (!_idx) {
            _idx = idx[__key] = {
              type: 'normal',
              value: data[_key].length ? data[_key].length - 1 : 0
            };
          }
          data = data[_key][_idx.value] = data[_key][_idx.value] || {};
          return;
        }
        data = data[_key] = data[_key] || {};
        return;
      }

      if (isArray) {
        __key = format.keys.slice(0, i + 1).join('.');
        _idx = idx[__key];
        if (!_idx) {
          _idx = idx[__key] = {
            type: 'normal',
            value: data[_key].length ? data[_key].length - 1 : 0
          };
        }
        data = data[_key];
        _key = _idx.value;
      }

      if (data[_key]) {
        return;
      }

      var type = format.type && format.type.toLowerCase();
      if (self._parser[type]) {
        data[_key] = isSplitArray ? cell.value.split(',').map(self._parser[type].bind(self._parser)) : self._parser[type](cell.value);
      } else {
        data[_key] = isSplitArray ? cell.value.split(',') : cell.value;
      }
    });
  });

  return {
    opts: opts,
    list: list
  };
};

/**
 * find origin data
 * @param {Object} dataMap
 * @param {Object} opts
 * @param {Object} data
 * @private
 */
Spread2Json.prototype._findOrigin = function(dataMap, opts, data) {
  var origin = dataMap[data.__ref];
  if (!origin || !opts.key) {
    logging.logger.error('not found origin.', JSON.stringify(data));
    return;
  }

  var keys = opts.key.split('.');
  var __in = data.__in ? data.__in.split('.') : [];
  for (var i = 0; i < keys.length; i++) {
    if (/^#/.test(keys[i])) {
      var key = keys[i].replace(/^#/, '');
      var index = __in[i] && __in[i].replace(/^#.+:(\d+)$/, '$1');
      if (!index) {
        logging.logger.error('not found index.', JSON.stringify(data));
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
      origin = origin[keys[i]] = origin[keys[i]] || {};
    }
    if (!origin) {
      logging.logger.error('not found origin parts.', JSON.stringify(data));
      return;
    }
  }

  if (opts.type === 'array') {
    if (!Array.isArray(origin)) {
      logging.logger.error('is not Array.', JSON.stringify(data));
      return;
    }
    origin.push({});
    origin = origin[origin.length - 1];
  } else if (opts.type === 'map') {
    if (!data.__key) {
      logging.logger.error('not found __key.', JSON.stringify(data));
      return;
    }
    origin = origin[data.__key] = {};
  } else {
    logging.logger.error(opts);
    return;
  }
  return origin;
};

/**
 * get worksheets data
 * @param {Credentials} [credentials] - oAuth2.getToken result
 * @param {string} key - spreadsheetKey
 * @param {Array} worksheetIds
 * @param {Function} callback
 * @see #getWorksheet arguments.key and returns list.id
 */
Spread2Json.prototype.getWorksheetDatas = function(credentials, key, worksheetIds, callback) {
  if (arguments.length === 3) {
    callback = worksheetIds;
    worksheetIds = key;
    key = credentials;
    credentials = undefined;
  }
  var self = this;
  var errList;
  async.map(worksheetIds, function(worksheetId, next) {
    api.getCells(credentials, key, worksheetId, function(err, result) {
      if (err) {
        return next(err);
      }

      var sheetName = result && result.feed && result.feed.title.$t;
      var entry = result && result.feed && result.feed.entry;
      if (!entry) {
        logging.logger.error('not found entry.', worksheetId);
        return next();
      }

      var cells = _.map(entry, function(cell) {
        var title = cell.title.$t.match(/^(\D+)(\d+)$/);
        return {
          cell: cell.title.$t,
          column: title[1],
          row: parseInt(title[2], 10),
          value: cell.content.$t
        };
      });

      var _result;
      try {
        _result = self._format(cells);
      } catch (e) {
        logging.logger.error('invalid sheet format.', sheetName);
        errList = errList || [];
        errList.push({
          name: sheetName,
          error: e
        });
        return next();
      }

      _result.name = sheetName;

      next(null, _result);
    });
  }, function(err, result) {
    if (err) {
      return callback(err);
    }

    callback(null, _.compact(result), errList);
  });
};

/**
 * sheetDatas to json
 * @param {Array} sheetDatas
 * @param {Function} callback
 */
Spread2Json.prototype.toJson = function(sheetDatas, callback) {
  var collectionMap = {};
  var optionMap = {};
  var errors = {};
  for (var i = 0; i < sheetDatas.length; i++) {
    var sheetData = sheetDatas[i];
    var opts = sheetData.opts;
    var name = opts.name || sheetData.name;
    var refKey = opts.ref_key;
    var dataMap = collectionMap[name] = collectionMap[name] || {};
    if (!optionMap[name]) {
      optionMap[name] = opts;
    } else {
      optionMap[name] = _.extend({}, opts, optionMap[name]);
      _.extend(optionMap[name].format, opts.format);
    }
    for (var j = 0; j < sheetData.list.length; j++) {
      var data = sheetData.list[j];
      if (!opts.type || opts.type === 'origin') {
        dataMap[data[refKey]] = data;
      } else {
        var origin = this._findOrigin(dataMap, opts, data);
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

  callback(Object.keys(errors).length ? errors : null, collectionMap, optionMap);
};

/**
 * get worksheets data
 * @param {Credentials} [credentials] - oAuth2.getToken result
 * @param {string} key - spreadsheetKey
 * @param {string} worksheetId
 * @param {Object[]} datas
 * @param {number|string} datas.row
 * @param {number|string} datas.col
 * @param {number|string} datas.value
 * @param {Function} callback
 * @see #getWorksheet arguments.key and returns list.id
 */
Spread2Json.prototype.updateWorksheetDatas = function(credentials, key, worksheetId, datas, callback) {
  if (arguments.length === 4) {
    callback = datas;
    datas = worksheetId;
    worksheetId = key;
    key = credentials;
    credentials = undefined;
  }

  api.batchCells(credentials, key, worksheetId, datas, callback);
};
