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
    // Line with a attribute description.
    desc_line: 3,
    // Line with a data.
    data_line: 4,
    // ref key
    ref_keys: ['_id'],
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
 * util
 */
/**
 * column key map 'A','B','C',,,'AA','AB',,,'AAA','AAB',,,'ZZZ'
 */
var COLUMN_KEYMAP = (function() {
  var map = [0];

  /**
   * @param {number} digit
   * @param {string} value
   */
  function genValue(digit, value) {
    var i, _value;
    for (i = 65; i < 91; i++) {
      _value = value + String.fromCharCode(i);
      if (digit > 1) {
        genValue(digit - 1, _value);
      } else {
        map.push(_value);
      }
    }
  }

  // A~Z
  genValue(1, '');
  // AA~ZZ
  genValue(2, '');
  // AAA~ZZZ
  genValue(3, '');
  return map;
}());

/**
 * column to number
 * @param {string} column
 */
function columnToNumber(column) {
  return COLUMN_KEYMAP.indexOf(column);
}

/**
 * separate cell
 * @param {string} cell
 * @returns {Object}
 * @example
 * console.log(separateCell('A1'));
 * > { cell: 'A1', column: 'A', row: 1 }
 */
function separateCell(cell) {
  var match = cell.match(/^(\D+)(\d+)$/);
  return {
    cell: cell,
    column: match[1],
    col: columnToNumber(match[1]),
    row: parseInt(match[2], 10)
  };
}

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
 * @param {string} [options.ref_key]
 * @param {string[]} [options.ref_keys=['_id']]
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
 *     ref_keys: ['_id'],
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

  if (this.opts.ref_key) {
    this.opts.ref_keys = options.ref_keys || [this.opts.ref_key];
    delete this.opts.ref_key;
  }

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
 * add worksheet
 * @param {Credentials} [credentials] - oAuth2.getToken result
 * @param {string} key - spreadsheetKey
 * @param {string} [title='new sheet'] - worksheet title
 * @param {number} [rowCount=50]
 * @param {number} [colCount=10]
 * @param {Function} callback
 * @example
 * > url = 'https://docs.google.com/spreadsheets/d/1YXVzaaxqkPKsr-excIOXScnTQC7y_DKrUKs0ukzSIgo/edit#gid=0'
 * key = '1YXVzaaxqkPKsr-excIOXScnTQC7y_DKrUKs0ukzSIgo'
 */
Spread2Json.prototype.addWorksheet = function(credentials, key, title, rowCount, colCount, callback) {
  if (arguments.length === 2) {
    callback = colCount;
    colCount = rowCount;
    rowCount = title;
    title = key;
    key = credentials;
    credentials = undefined;
  }
  title = title || 'new sheet';
  rowCount = _.isNumber(rowCount) ? rowCount : 50;
  colCount = _.isNumber(colCount) ? colCount : 10;
  api.addWorksheet(credentials, key, title, rowCount, colCount, function(err, result) {
    if (err) {
      return callback(err);
    }

    var entry = result && result.entry || {};
    callback(null, {
      id: entry.id && entry.id.$t && entry.id.$t.replace(/.+\/([^\/]+)$/, '$1'),
      updated: entry.updated && entry.updated.$t,
      title: entry.title && entry.title.$t
    });
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
 * stringify
 * @private
 */
Spread2Json.prototype._stringify = {
  number: function(d) {
    return d;
  },
  num: function(d) {
    return this.number(d);
  },
  boolean: function(d) {
    return d;
  },
  bool: function(d) {
    return this.boolean(d);
  },
  date: function(d) {
    var date = new Date(d);
    return date.getFullYear() + '/' +
      ('0' + (date.getMonth() + 1)).slice(-2) + '/' +
      ('0' + date.getDate()).slice(-2) + ' ' +
      ('0' + date.getHours()).slice(-2) + ':' +
      ('0' + date.getMinutes()).slice(-2) + ':' +
      ('0' + date.getSeconds()).slice(-2);
  },
  auto: function(d) {
    return d;
  }
};

/**
 * format
 * @param {Object[]} cells
 * @param {string} cells.cell
 * @param {string} cells.column
 * @param {number} cells.row
 * @param {string} cells.value
 * @private
 * @example
 * var cells = [
 *     { cell: 'A1', column: 'A', row: 1, value: '{}' },,,,
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
    ref_keys: this.opts.ref_keys,
    format: {}
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
      if (_opts.ref_key && !_opts.ref_keys) {
        _opts.ref_keys = [_opts.ref_key];
        delete _opts.ref_key;
      }
      _.extend(opts, _opts);
      return;
    }

    if (cell.row === opts.attr_line) {
      var type = cell.value.match(/:(\w+)$/);
      var key = cell.value.replace(/:\w+$/, '');
      var keys = key.split('.');

      opts.format[cell.column] = {
        type: type && type[1] && type[1].toLowerCase(),
        key: key,
        keys: keys
      };
      return;
    }

    var format = opts.format[cell.column];
    var data;
    var _idx;

    if (cell.row < opts.data_line || !format) {
      return;
    }
    if (format.type === 'index') {
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
    if ((!opts.type || opts.type === 'origin') && format.key === opts.ref_keys[0] ||
       format.key === '__ref' || format.key === '__ref_0') {
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

      var type = format.type;
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
  var origin;
  if (data.__ref) {
    origin = dataMap[data.__ref];
  } else {
    var dataKeys = '';
    _.forEach(opts.ref_keys, function(refKey, i) {
      dataKeys += data['__ref_' + i];
      if (i < opts.ref_keys.length - 1) {
        dataKeys += '.';
      }
    });
    origin = dataMap[dataKeys];
  }

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
        var cr = separateCell(cell.title.$t);
        return {
          cell: cr.cell,
          column: cr.column,
          row: cr.row,
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
 * @param {Object[]} sheetDatas
 * @param {string} sheetDatas.name - work sheet name
 * @param {Object} sheetDatas.opts - sheet option
 * @param {Object[]} sheetDatas.list - sheet data list
 * @param {Function} callback
 */
Spread2Json.prototype.toJson = function(sheetDatas, callback) {
  var self = this;
  var collectionMap = {};
  var optionMap = {};
  var errors = {};
  _.forEach(sheetDatas, function(sheetData) {
    var opts = sheetData.opts;
    var name = opts.name || sheetData.name;
    var refKeys = opts.ref_keys;
    var dataMap = collectionMap[name] = collectionMap[name] || {};
    if (!optionMap[name]) {
      optionMap[name] = opts;
    } else {
      optionMap[name] = _.extend({}, opts, optionMap[name]);
      _.extend(optionMap[name].format, opts.format);
    }
    _.forEach(sheetData.list, function(data) {
      if (!opts.type || opts.type === 'origin') {
        var keys = '';
        _.forEach(refKeys, function(refKey, i) {
          keys += data[refKey];
          if (i < refKeys.length - 1) {
            keys += '.';
          }
        });
        dataMap[keys] = data;
      } else {
        var origin = self._findOrigin(dataMap, opts, data);
        if (origin) {
          _.forEach(data, function(d, k) {
            if (/^__ref/.test(k)) {
              delete data[k];
            }
          });
          delete data.__in;
          delete data.__key;
          _.extend(origin, data);
        } else {
          errors[name] = errors[name] || [];
          errors[name].push(data);
        }
      }
    });
  });

  callback(Object.keys(errors).length ? errors : null, collectionMap, optionMap);
};

/**
 * sheetDatas to cells
 * FIXME array of array is unsupported
 * @param {Object} worksheetData
 * @param {Object} [worksheetData.opts={}] - sheet options
 * @param {string[]} worksheetData.format - data attribute
 * @param {string[]} worksheetData.description - attribute description
 * @param {Object[]} worksheetData.datas - datas
 */
Spread2Json.prototype.toCells = function(worksheetData) {
  var self = this;
  var cells = [];
  var opts = worksheetData.opts || {};
  var maxCol = worksheetData.format.length;
  var maxRow = 0;

  // opts
  var cr = separateCell(this.opts.option_cell);
  cells.push({
    col: cr.col,
    row: cr.row,
    value: JSON.stringify(opts)
  });

  // format
  _.forEach(worksheetData.format, function(attr, i) {
    cells.push({
      col: i + 1,
      row: self.opts.attr_line,
      value: attr
    });
  });

  // description
  _.forEach(worksheetData.description, function(desc, i) {
    cells.push({
      col: i + 1,
      row: self.opts.desc_line,
      value: desc
    });
  });

  // datas
  maxRow = this.opts.data_line;
  var row = maxRow;
  /**
   * add cell data to cells
   * @private
   */
  function addCell(data) {
    _.forEach(worksheetData.format, function(attr, i) {
      try {
        var last = _.last(attr.split('.'));
        var hasArray = /#/.test(attr);
        var isSplitArray = /^\$/.test(last);
        var type = attr.replace(/^.+:(.+)$/, '$1');
        var searchKey = attr.replace(/:.+$/, '').replace('$', '');
        var value;
        if (hasArray) {
          var keys = attr.replace(/#([^\.]+).*$/, '$1');
          var arr = _.get(data, keys);
          _.forEach(arr, function(d, j) {
            value = _.get(data, searchKey.replace(/#([^.]+)/, '$1[' + j + ']'));
            if (value !== undefined) {
              cells.push({
                col: i + 1,
                row: row + j,
                value: self._stringify[type] ? self._stringify[type](value) : value
              });
              maxRow = Math.max(maxRow, row + j);
            }
          });
        } else {
          value = _.get(data, searchKey);
          if (value !== undefined) {
            cells.push({
              col: i + 1,
              row: row,
              value: isSplitArray ? value.join() : (self._stringify[type] ? self._stringify[type](value) : value)
            });
          }
        }
      } catch (e) {
        logging.logger.error('parse error.', attr, e.message);
      }
    });
  }

  _.forEach(worksheetData.datas, function(data) {
    if (opts.type === 'array') {
      var origin = _.transform(opts.ref_keys || self.opts.ref_keys, function(result, refKey, i) {
        result['__ref_' + i] = data[refKey];
      }, {});
      var datas = _.get(data, opts.key);
      _.forEach(datas, function(data) {
        data = _.assign({}, origin, data);
        addCell(data);
        row = maxRow = maxRow + 1;
      });
    } else {
      addCell(data);
      row = maxRow = maxRow + 1;
    }
  });

  return {
    maxCol: maxCol,
    maxRow: maxRow,
    cells: cells
  };
};

/**
 * update worksheets data
 * @param {Credentials} [credentials] - oAuth2.getToken result
 * @param {string} key - spreadsheetKey
 * @param {string} worksheetId
 * @param {Object[]} cells
 * @param {number} cells.col
 * @param {number} cells.row
 * @param {string} cells.value
 * @param {Function} callback
 * @see #getWorksheet arguments.key and returns list.id
 */
Spread2Json.prototype.updateWorksheetDatas = function(credentials, key, worksheetId, cells, callback) {
  if (arguments.length === 4) {
    callback = cells;
    cells = worksheetId;
    worksheetId = key;
    key = credentials;
    credentials = undefined;
  }

  var self = this;

  api.batchCells(credentials, key, worksheetId, cells, function(err, result) {
    if (err) {
      return callback(err);
    }

    var entry = result && result.feed && result.feed.entry || [];
    var errorCells = _.transform(entry, function(errors, data) {
      if (data['batch:status'] && data['batch:status'].code >= 400) {
        errors.push({
          id: data['batch:id'],
          message: data.content
        });
      }
    });
    callback(null, errorCells);
  });
};
