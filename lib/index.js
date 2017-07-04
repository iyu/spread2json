/**
 * @fileOverview spread sheet converter
 * @name index.js
 * @author Yuhei Aihara <yu.e.yu.4119@gmail.com>
 * https://github.com/iyu/spread2json
 */

'use strict';

const _ = require('lodash');
const async = require('neo-async');

const api = require('./api');
const logging = require('./logging');


/**
 * util
 */
/**
 * column key map 'A','B','C',,,'AA','AB',,,'AAA','AAB',,,'ZZZ'
 */
const COLUMN_KEYMAP = (() => {
  const map = [];

  /**
   * @param {number} digit
   * @param {string} value
   */
  function genValue(digit, value) {
    let _value;
    for (let i = 65; i < 91; i++) {
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
})();

/**
 * column to number
 * @param {string} column
 */
function columnToNumber(column) {
  return COLUMN_KEYMAP.indexOf(column) + 1;
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
  const match = cell.match(/^(\D+)(\d+)$/);
  return {
    cell,
    column: match[1],
    col: columnToNumber(match[1]),
    row: parseInt(match[2], 10),
  };
}

const genWorksheetLink = _.template('https://docs.google.com/spreadsheets/d/<%= spreadsheetId %>/edit#gid=<%= sheetId %>');

class Spread2Json {
  constructor() {
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
          path: './dist/token.json',
        },
      },
    };

    this._parser = {
      number: (d) => {
        return Number(d);
      },
      num: (d) => {
        return this._parser.number(d);
      },
      boolean: (d) => {
        if (typeof d === 'boolean') {
          return d;
        }
        if (typeof d === 'string') {
          return d.toLowerCase() !== 'false' && d !== '0';
        }
        return !!d;
      },
      bool: (d) => {
        return this._parser.boolean(d);
      },
      date: (d) => {
        return new Date(d).getTime();
      },
      auto: (d) => {
        return isNaN(d) ? d : this._parser.number(d);
      },
    };

    this._stringify = {
      number: (d) => {
        return d;
      },
      num: (d) => {
        return this._stringify.number(d);
      },
      boolean: (d) => {
        return d;
      },
      bool: (d) => {
        return this._stringify.boolean(d);
      },
      date: (d) => {
        const date = new Date(d);
        const yyyy = date.getFullYear();
        const mm = `${`0${date.getMonth() + 1}`.slice(-2)}`;
        const dd = `${`0${date.getDate()}`.slice(-2)}`;
        const h = `${`0${date.getHours()}`.slice(-2)}`;
        const m = `${`0${date.getMinutes()}`.slice(-2)}`;
        const s = `${`0${date.getSeconds()}`.slice(-2)}`;
        return `${yyyy}/${mm}/${dd} ${h}:${m}:${s}`;
      },
      auto: (d) => {
        return d;
      },
    };
  }

  /**
   * API function
   */
  generateAuthUrl() { return api.generateAuthUrl.apply(api, arguments); }
  getAccessToken() { return api.getAccessToken.apply(api, arguments); }
  refreshAccessToken() { return api.refreshAccessToken.apply(api, arguments); }

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
  setup(options) {
    _.assign(this.opts, options);

    if (this.opts.ref_key) {
      this.opts.ref_keys = options.ref_keys || [this.opts.ref_key];
      delete this.opts.ref_key;
    }

    if (this.opts.logger) {
      logging.logger = this.opts.logger;
    }

    api.setup.call(api, this.opts.api);
    logging.logger.info('spread2json setup');
  }

  /**
   * get spreadsheets info
   * @param {Credentials} [credentials] - oAuth2.getToken result
   * @param {Object} [opts] - drive options @see https://developers.google.com/drive/v3/reference/files/list
   * @param {string} [opts.corpus]
   * @param {string} [opts.orderBy]
   * @param {integer} [opts.pageSize]
   * @param {string} [opts.pageToken]
   * @param {string} [opts.q]
   * @param {string} [opts.spaces]
   * @param {Function} callback
   * @property {Object[]} response
   * @property {string} response[].kind - 'drive#file'
   * @property {string} response[].id - spreadsheetId
   * @property {string} response[].name - spreadsheet name
   * @property {string} response[].mimeTpe - 'application/vnd.google-apps.spreadsheet'
   * @property {string} response[].link
   */
  getSpreadsheet(credentials, opts, callback) {
    if (arguments.length === 1) {
      callback = credentials;
      opts = undefined;
      credentials = undefined;
    }
    if (arguments.length === 2) {
      if (_.has(credentials, ['access_token'])) {
        callback = opts;
        opts = undefined;
      } else {
        callback = opts;
        opts = credentials;
        credentials = undefined;
      }
    }
    api.getSpreadsheet(credentials, opts, (err, result) => {
      if (err) {
        return callback(err);
      }

      _.forEach(result, (data) => {
        data.link = genWorksheetLink({
          spreadsheetId: data.id,
          sheetId: 0,
        });
      });
      callback(null, result);
    });
  }

  /**
   * get worksheet info in spreadsheet
   * @param {Credentials} [credentials] - oAuth2.getToken result
   * @param {string} spreadsheetId
   * @param {Function} callback
   * @property {Object[]} response
   * @property {string} response[].title
   * @property {integer} response[].sheetId
   * @property {string} response[].link
   * @example
   * > url = 'https://docs.google.com/spreadsheets/d/1YXVzaaxqkPKsr-excIOXScnTQC7y_DKrUKs0ukzSIgo/edit#gid=0'
   * spreadsheetId = '1YXVzaaxqkPKsr-excIOXScnTQC7y_DKrUKs0ukzSIgo'
   * sheetId = 0
   */
  getWorksheet(credentials, spreadsheetId, callback) {
    if (arguments.length === 2) {
      callback = spreadsheetId;
      spreadsheetId = credentials;
      credentials = undefined;
    }
    api.getWorksheet(credentials, spreadsheetId, (err, result) => {
      if (err) {
        return callback(err);
      }

      const list = _.map(_.get(result, 'sheets'), (sheet) => {
        const data = {
          title: _.get(sheet, ['properties', 'title'], ''),
          sheetId: _.get(sheet, ['properties', 'sheetId'], 0),
        };
        data.link = genWorksheetLink({
          spreadsheetId,
          sheetId: data.sheetId,
        });
        return data;
      });
      callback(null, list);
    });
  }

  /**
   * add worksheet
   * @param {Credentials} [credentials] - oAuth2.getToken result
   * @param {string} spreadsheetId
   * @param {string} [title] - worksheet title
   * @param {number} [rowCount=50]
   * @param {number} [colCount=10]
   * @param {Function} callback
   * @property {Object} response
   * @property {string} response.title
   * @property {integer} response.sheetId
   * @property {string} response.link
   * @example
   * > url = 'https://docs.google.com/spreadsheets/d/1YXVzaaxqkPKsr-excIOXScnTQC7y_DKrUKs0ukzSIgo/edit#gid=0'
   * spreadsheetId = '1YXVzaaxqkPKsr-excIOXScnTQC7y_DKrUKs0ukzSIgo'
   */
  addWorksheet(credentials, spreadsheetId, title, rowCount, colCount, callback) {
    if (arguments.length === 2) {
      callback = colCount;
      colCount = rowCount;
      rowCount = title;
      title = spreadsheetId;
      spreadsheetId = credentials;
      credentials = undefined;
    }
    rowCount = _.isNumber(rowCount) ? rowCount : 50;
    colCount = _.isNumber(colCount) ? colCount : 10;
    api.addWorksheet(credentials, spreadsheetId, title, rowCount, colCount, (err, result) => {
      if (err) {
        return callback(err);
      }

      const data = {
        title: _.get(result, ['replies', 0, 'addSheet', 'properties', 'title'], ''),
        sheetId: _.get(result, ['replies', 0, 'addSheet', 'properties', 'sheetId'], 0),
      };
      data.link = genWorksheetLink({
        spreadsheetId,
        sheetId: data.sheetId,
      });
      callback(null, data);
    });
  }


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
  _format(cells) {
    const opts = {};
    let beforeRow;
    let idx = {};
    const list = [];

    _.assign(opts, {
      attr_line: this.opts.attr_line,
      data_line: this.opts.data_line,
      ref_keys: this.opts.ref_keys,
      format: {},
    });

    _.forEach(cells, (cell) => {
      if (beforeRow !== cell.row) {
        _.forEach(idx, (i) => {
          if (i.type !== 'format') {
            i.value += 1;
          }
        });
        beforeRow = cell.row;
      }
      if (cell.cell === this.opts.option_cell) {
        const _opts = JSON.parse(cell.value);
        if (_opts.ref_key && !_opts.ref_keys) {
          _opts.ref_keys = [_opts.ref_key];
          delete _opts.ref_key;
        }
        _.assign(opts, _opts);
        return;
      }

      if (cell.row === opts.attr_line) {
        const type = cell.value.match(/:(\w+)$/);
        const key = cell.value.replace(/:\w+$/, '');
        const keys = key.split('.');

        opts.format[cell.column] = {
          type: type && type[1] && type[1].toLowerCase(),
          key,
          keys,
        };
        return;
      }

      const format = opts.format[cell.column];
      let data;
      let _idx;

      if (cell.row < opts.data_line || !format) {
        return;
      }
      if (format.type === 'index') {
        _idx = parseInt(cell.value, 10);
        if (!idx[format.key] || idx[format.key].value !== _idx) {
          idx[format.key] = {
            type: 'format',
            value: _idx,
          };
          _.forEach(idx, (i, key) => {
            if (new RegExp(`^${format.key}.+$`).test(key)) {
              idx[key].value = 0;
            }
          });
        }
        return;
      }
      if (((!opts.type || opts.type === 'origin') && format.key === opts.ref_keys[0]) ||
        format.key === '__ref' || format.key === '__ref_0') {
        idx = {};
        list.push({});
      }

      data = _.last(list);
      _.forEach(format.keys, (_key, i) => {
        const isArray = /^#/.test(_key);
        const isSplitArray = /^\$/.test(_key);
        let __key;
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
              idx[__key] = {
                type: 'normal',
                value: data[_key].length ? data[_key].length - 1 : 0,
              };
              _idx = idx[__key];
            }
            data[_key][_idx.value] = data[_key][_idx.value] || {};
            data = data[_key][_idx.value];
            return;
          }
          data[_key] = data[_key] || {};
          data = data[_key];
          return;
        }

        if (isArray) {
          __key = format.keys.slice(0, i + 1).join('.');
          _idx = idx[__key];
          if (!_idx) {
            idx[__key] = {
              type: 'normal',
              value: data[_key].length ? data[_key].length - 1 : 0,
            };
            _idx = idx[__key];
          }
          data = data[_key];
          _key = _idx.value;
        }

        if (data[_key]) {
          return;
        }

        const type = format.type;
        if (this._parser[type]) {
          data[_key] = isSplitArray ? cell.value.split(',').map(this._parser[type]) : this._parser[type](cell.value);
        } else {
          data[_key] = isSplitArray ? cell.value.split(',') : cell.value;
        }
      });
    });

    return {
      opts,
      list,
    };
  }

  /**
   * find origin data
   * @param {Object} dataMap
   * @param {Object} opts
   * @param {Object} data
   * @private
   */
  _findOrigin(dataMap, opts, data) {
    let origin;
    if (data.__ref) {
      origin = dataMap[data.__ref];
    } else {
      let dataKeys = '';
      _.forEach(opts.ref_keys, (refKey, i) => {
        dataKeys += data[`__ref_${i}`];
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

    const keys = opts.key.split('.');
    const __in = data.__in ? data.__in.split('.') : [];
    for (let i = 0; i < keys.length; i++) {
      if (/^#/.test(keys[i])) {
        const key = keys[i].replace(/^#/, '');
        const index = __in[i] && __in[i].replace(/^#.+:(\d+)$/, '$1');
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
        origin[keys[i]] = origin[keys[i]] || {};
        origin = origin[keys[i]];
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
      origin[data.__key] = {};
      origin = origin[data.__key];
    } else {
      logging.logger.error(opts);
      return;
    }
    return origin;
  }

  /**
   * get worksheets data
   * @param {Credentials} [credentials] - oAuth2.getToken result
   * @param {string} spreadsheetId
   * @param {string[]} worksheetNames
   * @param {Function} callback
   * @see #getWorksheet arguments.key and returns list.id
   */
  getWorksheetDatas(credentials, spreadsheetId, worksheetNames, callback) {
    if (arguments.length === 3) {
      callback = worksheetNames;
      worksheetNames = spreadsheetId;
      spreadsheetId = credentials;
      credentials = undefined;
    }
    let errList;
    async.map(worksheetNames, (sheetName, next) => {
      api.getList(credentials, spreadsheetId, sheetName, (err, result) => {
        if (err) {
          return next(err);
        }

        const cells = _.transform(result.values, (ret, columns, i) => {
          _.forEach(columns, (value, j) => {
            if (value === '') {
              return;
            }
            const rowName = i + 1;
            const columnName = COLUMN_KEYMAP[j];
            ret.push({
              cell: columnName + rowName,
              column: columnName,
              row: rowName,
              value,
            });
          });
        }, []);

        let _result;
        try {
          _result = this._format(cells);
        } catch (e) {
          logging.logger.error('invalid sheet format.', sheetName);
          errList = errList || [];
          errList.push({
            name: sheetName,
            error: e,
          });
          return next();
        }

        _result.name = sheetName;

        next(null, _result);
      });
    }, (err, result) => {
      if (err) {
        return callback(err);
      }

      callback(null, _.compact(result), errList);
    });
  }

  /**
   * sheetDatas to json
   * @param {Object[]} sheetDatas
   * @param {string} sheetDatas.name - work sheet name
   * @param {Object} sheetDatas.opts - sheet option
   * @param {Object[]} sheetDatas.list - sheet data list
   * @param {Function} callback
   */
  toJson(sheetDatas, callback) {
    const collectionMap = {};
    const optionMap = {};
    const errors = {};
    _.forEach(sheetDatas, (sheetData) => {
      const opts = sheetData.opts;
      const name = opts.name || sheetData.name;
      const refKeys = opts.ref_keys;
      collectionMap[name] = collectionMap[name] || {};
      const dataMap = collectionMap[name];
      if (!optionMap[name]) {
        optionMap[name] = opts;
      } else {
        optionMap[name] = _.assign({}, opts, optionMap[name]);
        _.assign(optionMap[name].format, opts.format);
      }
      _.forEach(sheetData.list, (data) => {
        if (!opts.type || opts.type === 'origin') {
          let keys = '';
          _.forEach(refKeys, (refKey, i) => {
            keys += data[refKey];
            if (i < refKeys.length - 1) {
              keys += '.';
            }
          });
          dataMap[keys] = data;
        } else {
          const origin = this._findOrigin(dataMap, opts, data);
          if (origin) {
            _.forEach(data, (d, k) => {
              if (/^__ref/.test(k)) {
                delete data[k];
              }
            });
            delete data.__in;
            delete data.__key;
            _.assign(origin, data);
          } else {
            errors[name] = errors[name] || [];
            errors[name].push(data);
          }
        }
      });
    });

    callback(_.ieEmpty(errors) ? null : errors, collectionMap, optionMap);
  }

  /**
   * sheetDatas to cells
   * FIXME array of array is unsupported
   * @param {Object} worksheetData
   * @param {Object} [worksheetData.opts={}] - sheet options
   * @param {string[]} worksheetData.format - data attribute
   * @param {string[]} worksheetData.description - attribute description
   * @param {Object[]} worksheetData.datas - datas
   */
  toCells(worksheetData) {
    const self = this;
    const cells = [];
    const opts = worksheetData.opts || {};
    const maxCol = worksheetData.format.length;
    let maxRow = 0;

    // opts
    const cr = separateCell(this.opts.option_cell);
    cells.push({
      col: cr.col,
      row: cr.row,
      value: JSON.stringify(opts),
    });

    // format
    _.forEach(worksheetData.format, (attr, i) => {
      cells.push({
        col: i + 1,
        row: this.opts.attr_line,
        value: attr,
      });
    });

    // description
    _.forEach(worksheetData.description, (desc, i) => {
      cells.push({
        col: i + 1,
        row: this.opts.desc_line,
        value: desc,
      });
    });

    // datas
    maxRow = this.opts.data_line;
    let row = maxRow;
    /**
     * add cell data to cells
     * @private
     */
    function addCell(data) {
      _.forEach(worksheetData.format, (attr, i) => {
        try {
          const last = _.last(attr.split('.'));
          const hasArray = /#/.test(attr);
          const isSplitArray = /^\$/.test(last);
          const type = attr.replace(/^.+:(.+)$/, '$1');
          const stringify = self._stringify[type] || self._stringify.auto;
          const searchKey = attr.replace(/:.+$/, '').replace('$', '');
          let value;
          if (hasArray) {
            const keys = attr.replace(/#([^.]+).*$/, '$1');
            const arr = _.get(data, keys);
            _.forEach(arr, (d, j) => {
              value = _.get(data, searchKey.replace(/#([^.]+)/, `$1[${j}]`));
              if (value !== undefined) {
                if (isSplitArray) {
                  value = _.map(value, (v) => {
                    return stringify(v);
                  }).join();
                } else {
                  value = stringify(value);
                }
                cells.push({
                  col: i + 1,
                  row: row + j,
                  value,
                });
                maxRow = Math.max(maxRow, row + j);
              }
            });
          } else {
            value = _.get(data, searchKey);
            if (value !== undefined) {
              if (isSplitArray) {
                value = _.map(value, (v) => {
                  return stringify(v);
                }).join();
              } else {
                value = stringify(value);
              }
              cells.push({
                col: i + 1,
                row,
                value,
              });
            }
          }
        } catch (e) {
          logging.logger.error('parse error.', attr, e.message);
        }
      });
    }

    _.forEach(worksheetData.datas, (data) => {
      if (opts.type === 'array') {
        const origin = _.transform(opts.ref_keys || this.opts.ref_keys, (result, refKey, i) => {
          result[`__ref_${i}`] = data[refKey];
        }, {});
        const datas = _.get(data, opts.key);
        _.forEach(datas, (_data) => {
          _data = _.assign({}, origin, _data);
          addCell(_data);
          maxRow += 1;
          row = maxRow;
        });
      } else {
        addCell(data);
        maxRow += 1;
        row = maxRow;
      }
    });

    return {
      maxCol,
      maxRow,
      cells,
    };
  }

  /**
   * update worksheets data
   * @param {Credentials} [credentials] - oAuth2.getToken result
   * @param {string} spreadsheetId
   * @param {string} worksheetName
   * @param {Object[]} cells
   * @param {number} cells.col
   * @param {number} cells.row
   * @param {string} cells.value
   * @param {Function} callback
   * @see #getWorksheet arguments.key and returns list.id
   */
  updateWorksheetDatas(credentials, spreadsheetId, worksheetName, cells, callback) {
    if (arguments.length === 4) {
      callback = cells;
      cells = worksheetName;
      worksheetName = spreadsheetId;
      spreadsheetId = credentials;
      credentials = undefined;
    }

    const entry = [];
    _.forEach(cells, (cell) => {
      entry[cell.row] = entry[cell.row] || [];
      entry[cell.row][cell.col - 1] = cell.value;
    });
    api.batchCells(credentials, spreadsheetId, worksheetName, entry, callback);
  }
}

module.exports = new Spread2Json();

