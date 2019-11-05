/**
 * @fileOverview spread sheet converter
 * @name index
 * @author Yuhei Aihara <yu.e.yu.4119@gmail.com>
 * https://github.com/iyu/spread2json
 */

import _ from 'lodash';
import { Credentials } from 'google-auth-library';
import { drive_v3, sheets_v4 } from 'googleapis';

import api from './api';
import logging from './logging';

interface Spreadsheet extends drive_v3.Schema$File {
  link: string;
}
interface Worksheet {
  title: string;
  sheetId: string;
  link: string;
}
interface Opts {
  attr_line: number;
  data_line: number;
  ref_keys: string[];
  format: { [key: string]: { type: string; key: string; keys: string[]; } };
  key?: string;
  type?: string;
  name?: string;
}
interface JsonSchema {
  type: string;
  items?: JsonSchema,
  properties: { [key: string]: JsonSchema },
  allOf: JsonSchema[],
  oneOf: JsonSchema[],
  format?: string;
}
type Callback<T> = (err: Error | null, result?: T) => T | undefined;
/**
 * util
 */
/**
 * column key map 'A','B','C',,,'AA','AB',,,'AAA','AAB',,,'ZZZ'
 */
const COLUMN_KEYMAP = (() => {
  const map: string[] = [];

  /**
   * @param {number} digit
   * @param {string} value
   */
  function genValue(digit: number, value: string) {
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
 */
const columnToNumber = (column: string): number => {
  return COLUMN_KEYMAP.indexOf(column) + 1;
};

/**
 * separate cell
 * @param {string} cell
 * @returns {Object}
 * @example
 * console.log(separateCell('A1'));
 * > { cell: 'A1', column: 'A', row: 1 }
 */
const separateCell = (cell: string): {
  cell: string;
  column: string;
  col: number;
  row: number;
} => {
  const match = cell.match(/^(\D+)(\d+)$/) || [];
  return {
    cell,
    column: match[1],
    col: columnToNumber(match[1]),
    row: parseInt(match[2], 10),
  };
};

const genWorksheetLink = _.template('https://docs.google.com/spreadsheets/d/<%= spreadsheetId %>/edit#gid=<%= sheetId %>');

class Spread2Json {
  private opts = {
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
  private _parser: { [key: string]: (d: any) => number|boolean|string|any; } = {
    string: (d: any) => {
      return String(d);
    },
    number: (d: any) => {
      return Number(d);
    },
    num: (d: any) => {
      return this._parser.number(d);
    },
    boolean: (d: any) => {
      if (typeof d === 'boolean') {
        return d;
      }
      if (typeof d === 'string') {
        return d.toLowerCase() !== 'false' && d !== '0';
      }
      return !!d;
    },
    bool: (d: any) => {
      return this._parser.boolean(d);
    },
    date: (d: any) => {
      return new Date(d).getTime();
    },
    auto: (d: any) => {
      return Number.isNaN(Number(d)) ? d : this._parser.number(d);
    },
  };

  private _stringify: { [key: string]: (d: any) => string; } = {
    string: (d: any) => {
      return d;
    },
    number: (d: any) => {
      return d;
    },
    num: (d: any) => {
      return this._stringify.number(d);
    },
    boolean: (d: any) => {
      return d;
    },
    bool: (d: any) => {
      return this._stringify.boolean(d);
    },
    date: (d: any) => {
      const date = new Date(d);
      const yyyy = date.getFullYear();
      const mm = `${`0${date.getMonth() + 1}`.slice(-2)}`;
      const dd = `${`0${date.getDate()}`.slice(-2)}`;
      const h = `${`0${date.getHours()}`.slice(-2)}`;
      const m = `${`0${date.getMinutes()}`.slice(-2)}`;
      const s = `${`0${date.getSeconds()}`.slice(-2)}`;
      return `${yyyy}/${mm}/${dd} ${h}:${m}:${s}`;
    },
    auto: (d: any) => {
      return d;
    },
  };

  /**
   * API function
   */
  generateAuthUrl(opts: { scope: string[]; }) { return api.generateAuthUrl(opts); }
  async getAccessToken(code: string, callback?: Callback<Credentials | null | undefined>) {
    if (!callback) {
      callback = (err, result) => {
        if (err) {
          throw err;
        }
        return result;
      };
    }

    try {
      const result = await api.getAccessToken(code);
      return callback(null, result);
    } catch (e) {
      return callback(e);
    }
  }
  async refreshAccessToken(
    credentials?: Credentials,
    callback?: Callback<Credentials>,
  ) {
    if (!callback) {
      callback = (err, result) => {
        if (err) {
          throw err;
        }
        return result;
      };
    }
    try {
      const result = await api.refreshAccessToken(credentials);
      return callback(null, result);
    } catch (e) {
      return callback(e);
    }
  }

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
  setup(options: {
    option_cell?: string;
    attr_line?: number;
    data_line?: number;
    ref_key?: string;
    ref_keys?: string[];
    api: {
      client_id: string;
      client_secret: string;
      redirect_url?: string;
      token_file?: { use: boolean; path: string; };
    };
    logger: any;
  }) {
    const opts = _.assign(this.opts, options);

    if (opts.ref_key) {
      opts.ref_keys = options.ref_keys || [opts.ref_key];
      delete opts.ref_key;
    }

    if (opts.logger) {
      logging.logger = opts.logger;
    }

    api.setup.call(api, opts.api);
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
   * @param {Function} [callback]
   * @property {Object[]} response
   * @property {string} response[].kind - 'drive#file'
   * @property {string} response[].id - spreadsheetId
   * @property {string} response[].name - spreadsheet name
   * @property {string} response[].mimeTpe - 'application/vnd.google-apps.spreadsheet'
   * @property {string} response[].link
   */
  async getSpreadsheet(): Promise<Spreadsheet[]>;
  async getSpreadsheet(credentials: Credentials): Promise<Spreadsheet[]>;
  async getSpreadsheet(opts: drive_v3.Params$Resource$Files$List): Promise<Spreadsheet[]>;
  async getSpreadsheet(callback: Callback<Spreadsheet[]>): Promise<Spreadsheet[]>;
  async getSpreadsheet(
    credentials: Credentials,
    opts: drive_v3.Params$Resource$Files$List,
  ): Promise<Spreadsheet[]>;
  async getSpreadsheet(
    credentials: Credentials,
    callback: Callback<Spreadsheet[]>,
  ): Promise<Spreadsheet[]>;
  async getSpreadsheet(
    opts: drive_v3.Params$Resource$Files$List,
    callback: Callback<Spreadsheet[]>,
  ): Promise<Spreadsheet[]>;
  async getSpreadsheet(
    credentials: Credentials,
    opts: drive_v3.Params$Resource$Files$List,
    callback: Callback<Spreadsheet[]>,
  ): Promise<Spreadsheet[]>;
  async getSpreadsheet(
    credentials?: Credentials | Callback<Spreadsheet[]> | drive_v3.Params$Resource$Files$List,
    opts?: drive_v3.Params$Resource$Files$List | Callback<Spreadsheet[]>,
    callback?: Callback<Spreadsheet[]>,
  ): Promise<Spreadsheet[] | undefined> {
    if (arguments.length === 1) {
      if (typeof credentials === 'function') {
        callback = credentials as Callback<Spreadsheet[]>;
        opts = undefined;
        credentials = undefined;
      } else if (!_.has(credentials, ['access_token'])) {
        opts = credentials as drive_v3.Params$Resource$Files$List;
        credentials = undefined;
      }
    } else if (arguments.length === 2) {
      if (typeof opts === 'function') {
        callback = opts as Callback<Spreadsheet[]>;
        if (!_.has(credentials, ['access_token'])) {
          opts = credentials as drive_v3.Params$Resource$Files$List;
          credentials = undefined;
        } else {
          opts = undefined;
        }
      }
    }
    if (!callback) {
      callback = (err: Error | null, result?: Spreadsheet[]) => {
        if (err) {
          throw err;
        }
        return result;
      };
    }

    let result;
    try {
      result = await api.getSpreadsheet(
        credentials as Credentials,
        opts as drive_v3.Params$Resource$Files$List,
      );
    } catch (e) {
      return callback(e);
    }

    const list: Spreadsheet[] = _.map(result, (sheet) => {
      const data = _.assign({}, sheet, {
        link: genWorksheetLink({
          spreadsheetId: sheet.id,
          sheetId: 0,
        }),
      });
      return data;
    });
    return callback(null, list);
  }

  /**
   * get worksheet info in spreadsheet
   * @param {Credentials} [credentials] - oAuth2.getToken result
   * @param {string} spreadsheetId
   * @param {Function} [callback]
   * @property {Object[]} response
   * @property {string} response[].title
   * @property {integer} response[].sheetId
   * @property {string} response[].link
   * @example
   * > url = 'https://docs.google.com/spreadsheets/d/1YXVzaaxqkPKsr-excIOXScnTQC7y_DKrUKs0ukzSIgo/edit#gid=0'
   * spreadsheetId = '1YXVzaaxqkPKsr-excIOXScnTQC7y_DKrUKs0ukzSIgo'
   * sheetId = 0
   */
  async getWorksheet(spreadsheetId: string): Promise<Worksheet[] | undefined>;
  async getWorksheet(
    spreadsheetId: string,
    callback: Callback<Worksheet[]>,
  ): Promise<Worksheet[] | undefined>;
  async getWorksheet(
    credentials: Credentials,
    spreadsheetId: string,
  ): Promise<Worksheet[] | undefined>;
  async getWorksheet(
    credentials: Credentials,
    spreadsheetId: string,
    callback: Callback<Worksheet[]>,
  ): Promise<Worksheet[] | undefined>;
  async getWorksheet(
    credentials?: Credentials | string,
    spreadsheetId?: string | Callback<Worksheet[]>,
    callback?: Callback<Worksheet[]>,
  ): Promise<Worksheet[] | undefined> {
    if (!_.has(credentials, 'access_token')) {
      callback = spreadsheetId as Callback<Worksheet[]>;
      spreadsheetId = credentials as string;
      credentials = undefined;
    }
    if (!callback) {
      callback = (err: Error | null, result?: Worksheet[]) => {
        if (err) {
          throw err;
        }
        return result;
      };
    }

    let result;
    try {
      result = await api.getWorksheet(credentials as Credentials, spreadsheetId as string);
    } catch (e) {
      return callback(e);
    }
    const list: Worksheet[] = _.map(_.get(result, 'sheets'), (sheet) => {
      const sheetId = _.get(sheet, ['properties', 'sheetId'], 0);
      const data = {
        title: _.get(sheet, ['properties', 'title'], ''),
        sheetId,
        link: genWorksheetLink({
          spreadsheetId,
          sheetId,
        }),
      };
      return data;
    });

    return callback(null, list);
  }

  /**
   * add worksheet
   * @param {Credentials} [credentials] - oAuth2.getToken result
   * @param {string} spreadsheetId
   * @param {string} [title] - worksheet title
   * @param {number} [rowCount=50]
   * @param {number} [colCount=10]
   * @param {Function} [callback]
   * @property {Object} response
   * @property {string} response.title
   * @property {integer} response.sheetId
   * @property {string} response.link
   * @example
   * > url = 'https://docs.google.com/spreadsheets/d/1YXVzaaxqkPKsr-excIOXScnTQC7y_DKrUKs0ukzSIgo/edit#gid=0'
   * spreadsheetId = '1YXVzaaxqkPKsr-excIOXScnTQC7y_DKrUKs0ukzSIgo'
   */
  async addWorksheet(
    spreadsheetId: string,
    title?: string,
    rowCount?: number,
    colCount?: number,
    callback?: Callback<Worksheet>,
  ): Promise<Worksheet | undefined>;
  async addWorksheet(
    credentials: Credentials,
    spreadsheetId: string,
    title?: string,
    rowCount?: number,
    colCount?: number,
    callback?: Callback<Worksheet>,
  ): Promise<Worksheet | undefined>;
  async addWorksheet(
    credentials?: Credentials | string,
    spreadsheetId?: string,
    title?: string | number,
    rowCount?: number,
    colCount?: number | Callback<Worksheet>,
    callback?: Callback<Worksheet>,
  ): Promise<Worksheet | undefined> {
    if (!_.has(credentials, ['access_token'])) {
      callback = colCount as Callback<Worksheet>;
      colCount = rowCount;
      rowCount = title as number;
      title = spreadsheetId;
      spreadsheetId = credentials as string;
      credentials = undefined;
    }
    rowCount = _.isNumber(rowCount) ? rowCount : 50;
    colCount = _.isNumber(colCount) ? colCount : 10;
    if (!callback) {
      callback = (err: Error | null, result?: Worksheet): Worksheet | undefined => {
        if (err) {
          throw err;
        }
        return result;
      };
    }

    let result;
    try {
      result = await api.addWorksheet(
        credentials as Credentials,
        spreadsheetId as string,
        title as string,
        rowCount,
        colCount,
      );
    } catch (e) {
      return callback(e);
    }

    const sheetId = _.get(result, ['replies', 0, 'addSheet', 'properties', 'sheetId'], 0);
    const data = {
      title: _.get(result, ['replies', 0, 'addSheet', 'properties', 'title'], ''),
      sheetId,
      link: genWorksheetLink({
        spreadsheetId,
        sheetId,
      }),
    };
    return callback(null, data);
  }


  /**
   * format
   * @param {Object[]} cells
   * @param {string} cells[].cell
   * @param {string} cells[].column
   * @param {number} cells[].row
   * @param {string} cells[].value
   * @private
   * @example
   * var cells = [
   *     { cell: 'A1', column: 'A', row: 1, value: '{}' },,,,
   * ]
   */
  _format(cells: { cell: string; column: string; row: number; value: string; }[]) {
    const opts: Opts = {
      attr_line: this.opts.attr_line,
      data_line: this.opts.data_line,
      ref_keys: this.opts.ref_keys,
      format: {},
    };
    let beforeRow: number;
    let idx: { [key: string]: { type: string; value: number; } } = {};
    const list: { [key: string]: any[]; }[] = [];

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
          type: (type && type[1] && type[1].toLowerCase()) || '',
          key,
          keys,
        };
        return;
      }

      const format = opts.format[cell.column];

      if (cell.row < opts.data_line || !format) {
        return;
      }
      if (format.type === 'index') {
        const _idx = parseInt(cell.value, 10);
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

      let data = _.last(list) as { [key:string]: any[] } | any;
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
            let _idx = idx[__key];
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
          let _idx = idx[__key];
          if (!_idx) {
            idx[__key] = {
              type: 'normal',
              value: data[_key].length ? data[_key].length - 1 : 0,
            };
            _idx = idx[__key];
          }
          data = data[_key];
          _key = String(_idx.value);
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
  _findOrigin(
    dataMap: { [key: string]: any },
    opts: { key?: string; type: string; ref_keys: string[] },
    data: any,
  ) {
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
  async getWorksheetDatas(spreadsheetId: string, worksheetNames: string[]): Promise<any>;
  async getWorksheetDatas(
    spreadsheetId: string,
    worksheetNames: string[],
    callback: (err: Error|null, result?: any[], errList?: { name: string; error: Error }[]) => any,
  ): Promise<any>;
  async getWorksheetDatas(
    credentials: Credentials,
    spreadsheetId: string,
    worksheetNames: string[],
  ): Promise<any>;
  async getWorksheetDatas(
    credentials: Credentials,
    spreadsheetId: string,
    worksheetNames: string[],
    callback: (err: Error|null, result?: any[], errList?: { name: string, error: Error }[]) => any,
  ): Promise<any>;
  async getWorksheetDatas(
    credentials?: Credentials | string,
    spreadsheetId?: string | string[],
    worksheetNames?:
      string[] |
      ((err: Error | null, result?: any[], errList?: { name: string, error: Error }[]) => any),
    callback?: (err: Error|null, result?: any[], errList?: { name: string, error: Error }[]) => any,
  ): Promise<any> {
    if (!_.has(credentials, 'access_token')) {
      callback = worksheetNames as
        (err: Error | null, result?: any[], errList?: { name: string; error: Error }[]) => any;
      worksheetNames = spreadsheetId as string[];
      spreadsheetId = credentials as string;
      credentials = undefined;
    }
    if (!callback) {
      callback =
        (err: Error | null, result?: any[], errList?: { name: string, error: Error; }[]) => {
          if (err) {
            throw err;
          }
          return [result, errList];
        };
    }

    let errList: any[] | undefined;
    let results;
    try {
      results = await Promise.all(_.map(worksheetNames as string[], async (sheetName: string) => {
        const result = await api.getList(
          credentials as Credentials,
          spreadsheetId as string,
          sheetName,
        );
        const cells = _.transform(
          result.values as any[][],
          (ret: { cell: string; column: string; row: number; value: string; }[], columns, i) => {
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
          },
          [],
        );

        let _result;
        try {
          _result = _.assign({ name: '' }, this._format(cells));
        } catch (e) {
          logging.logger.error('invalid sheet format.', sheetName);
          errList = errList || [];
          errList.push({
            name: sheetName,
            error: e,
          });
          return;
        }

        _result.name = sheetName;

        return _result;
      }));
    } catch (e) {
      return callback(e);
    }

    return callback(null, _.compact(results), errList);
  }

  /**
   * sheetDatas to json
   * @param {Object[]} sheetDatas
   * @param {string} sheetDatas.name - work sheet name
   * @param {Object} sheetDatas.opts - sheet option
   * @param {Object[]} sheetDatas.list - sheet data list
   * @param {Function} [callback]
   */
  toJson(
    sheetDatas: { name: string; opts: Opts; list: any[] }[],
    callback?: (err: any | null, collectionMap: any, optionMap: any) => any,
  ) {
    const collectionMap: { [key: string]: any } = {};
    const optionMap: { [key: string]: Opts } = {};
    const errors: { [key: string]: any[] } = {};
    if (!callback) {
      callback = (err: any | null, _collectionMap: any, _optionMap: any) => {
        if (err) {
          throw err;
        }
        return [_collectionMap, _optionMap];
      };
    }
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
          const origin = this._findOrigin(dataMap, opts as { type: string } & Opts, data);
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

    return callback(_.isEmpty(errors) ? null : errors, collectionMap, optionMap);
  }

  /**
   * jsonschema to spread2json format
   * @param {Object} schema - jsonschema
   * @param {string[]} refkeys - ref keys
   * @param {string} [path=''] - data path
   */
  convertSchema2Format(schema: JsonSchema, refKeys: string[], path?: string): string[] {
    path = path || '';
    let result: string[] = [];
    switch (schema.type) {
      case 'string':
      case 'boolean':
        result.push(`${path}:${schema.type}`);
        break;
      case 'number':
      case 'integer':
        if (schema.format === 'date-time') {
          result.push(`${path}:date`);
        } else {
          result.push(`${path}:number`);
        }
        break;
      case 'object':
        _.forEach(schema.properties, (sc, key) => {
          let nextKey = key;
          if (sc.type === 'array' && sc.items) {
            if (sc.items.type === 'object') {
              nextKey = `#${key}`;
            } else if (sc.items.type !== 'array') {
              nextKey = `$${key}`;
            }
          }
          const nextPath = path ? `${path}.${nextKey}` : nextKey;
          const format = this.convertSchema2Format(sc, refKeys, nextPath);
          result.push.apply(result, format);
        });
        _.forEach(schema.allOf, (sc) => {
          const format = this.convertSchema2Format(sc, refKeys, path);
          result.push.apply(result, format);
          result = _.uniq(result);
        });
        _.forEach(schema.oneOf, (sc) => {
          const format = this.convertSchema2Format(sc, refKeys, path);
          result.push.apply(result, format);
          result = _.uniq(result);
        });
        break;
      case 'array': {
        if (schema.items) {
          const arrayMark = schema.items.type === 'object' ? '#' : '$';
          const format = this.convertSchema2Format(schema.items, refKeys, path || arrayMark);
          result.push.apply(result, format);
        } else {
          throw new Error(`invalid schema type. ${schema.type}`);
        }
        break;
      }
      default:
        throw new Error(`invalid schema type. ${schema.type}`);
    }

    return _.sortBy(result, (n, index) => {
      const i = _.indexOf(refKeys, n.replace(/:\w+$/, ''));
      return i === -1 ? index : (refKeys.length - i) * -1;
    });
  }

  /**
   * multiple sheet
   * @param {string} name
   * @param {string[]} refKeys
   * @param {string[]} format
   * @param {Object[]} datas
   */
  splitSheetData(name: string, refKeys: string[], format: string[], datas: any[]) {
    const map: {
      [key: string]: {
        name: string,
        opts: { name: string, ref_keys: string[]; attr_line: number; desc_line: number; },
        format: string[],
        description: string[],
        datas: any[],
      }
    } = {
      [name]: {
        name,
        opts: _.defaults({
          name,
          ref_keys: refKeys,
        }, {
          attr_line: this.opts.attr_line,
          desc_line: this.opts.desc_line,
          data_line: this.opts.data_line,
          ref_keys: this.opts.ref_keys,
        }),
        format: [],
        description: [],
        datas,
      },
    };

    const matches: (RegExpMatchArray | null)[] = [];
    const checks: RegExp[] = [];
    _.forEach(format, (keys) => {
      const checked = _.some(checks, (regexp) => {
        return regexp.test(keys);
      });
      if (checked) {
        return;
      }

      let match = keys.match(/#/g);
      if (match && match.length >= 2) {
        // keys:'obj.#arr.obj2.#arr2.str' -> match:[keys,'obj.','arr']
        match = keys.match(/^([^#]+)?#([^#.]+)\..+$/);
        if (match) {
          matches.push(match);
          // ^obj.#arr\.
          checks.push(new RegExp(`^${(match[1] ? match[1] : '')}#${match[2]}\\.`));
        }
      }
    });

    _.forEach(format, (keys) => {
      let match: string[] = [];
      let check: RegExp | string = '';
      const checked = _.some(checks, (regexp, i) => {
        match = matches[i] || [];
        check = regexp;
        return regexp.test(keys);
      });
      if (!checked) {
        map[name].format.push(keys);
        map[name].description.push(_.last(keys.replace(/#|\$/g, '').replace(/:.+$/, '').split('.')) || '');
        return;
      }
      const key = (match[1] ? match[1] : '') + match[2];
      const _sheetName = `${name}.${key}`;
      const _keys = keys.replace(check, '');
      if (!map[_sheetName]) {
        map[_sheetName] = _.clone(map[name]);
        map[_sheetName].name = _sheetName;
        map[_sheetName].opts = _.assign({ type: 'array', key }, map[name].opts);
        map[_sheetName].format = _.map(refKeys, (refKey, i) => {
          return `__ref_${i}`;
        });
        map[_sheetName].description = _.clone(refKeys);
      }
      map[_sheetName].format.push(_keys);
      map[_sheetName].description.push(_.last(_keys.replace(/#|\$/g, '.').replace(/:.+$/, '').split('.')) || '');
    });

    return _.values(map);
  }

  /**
   * sheetDatas to cells
   * FIXME array of array is unsupported
   * @param {Object} worksheetData
   * @param {Object} [worksheetData.opts={}] - sheet options
   * @param {string[]} worksheetData.format - data attribute
   * @param {string[]} worksheetData.description - attribute description
   * @param {Object[]} worksheetData.datas - datas
   * @param {string} [worksheetData.name] - sheet name
   */
  toCells(worksheetData: { opts?: any; format: string[]; description: string[]; datas: any[]; name?: string }) { // eslint-disable-line max-len
    const self = this;
    const cells = [];
    const opts = worksheetData.opts || {};
    const maxCol = worksheetData.format.length;
    const sheetName = worksheetData.name;
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
    function addCell(data: any) {
      _.forEach(worksheetData.format, (attr, i) => {
        try {
          const last = _.last(attr.split('.')) || '';
          const hasArray = /#/.test(attr);
          const isSplitArray = /^\$/.test(last);
          const type = attr.replace(/^.+:(.+)$/, '$1');
          const stringify = self._stringify[type] || self._stringify.auto;
          const searchKey = attr.replace(/:.+$/, '').replace('$', '');
          let value;
          if (hasArray) {
            const keys = attr.replace(/#(\w+).*$/, '$1');
            const arr = _.get(data, keys);
            _.forEach(arr, (d, j: number) => {
              value = _.get(data, searchKey.replace(/#(\w+)/, `$1[${j}]`));
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
        const origin = _.transform(
          opts.ref_keys || this.opts.ref_keys,
          (result: { [key: string]: any }, refKey: string, i) => {
            result[`__ref_${i}`] = data[refKey];
          },
          {},
        );
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
      cells: _.sortBy(cells, ['row', 'col']),
      name: sheetName,
      opts,
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
  async updateWorksheetDatas(
    spreadsheetId: string,
    worksheetName: string,
    cells: { col: number; row: number; value: string }[],
  ): Promise<sheets_v4.Schema$BatchUpdateValuesResponse>;
  async updateWorksheetDatas(
    spreadsheetId: string,
    worksheetName: string,
    cells: { col: number; row: number; value: string }[],
    callback: Callback<sheets_v4.Schema$BatchUpdateValuesResponse>,
  ): Promise<sheets_v4.Schema$BatchUpdateValuesResponse>;
  async updateWorksheetDatas(
    credentials: Credentials,
    spreadsheetId: string,
    worksheetName: string,
    cells: { col: number; row: number; value: string }[],
  ): Promise<sheets_v4.Schema$BatchUpdateValuesResponse>;
  async updateWorksheetDatas(
    credentials: Credentials,
    spreadsheetId: string,
    worksheetName: string,
    cells: { col: number; row: number; value: string }[],
    callback: Callback<sheets_v4.Schema$BatchUpdateValuesResponse>,
  ): Promise<sheets_v4.Schema$BatchUpdateValuesResponse>;
  async updateWorksheetDatas(
    credentials?: Credentials | string,
    spreadsheetId?: string,
    worksheetName?: string | { col: number; row: number; value: string }[],
    cells?: { col: number; row: number; value: string }[] |
      Callback<sheets_v4.Schema$BatchUpdateValuesResponse>,
    callback?: Callback<sheets_v4.Schema$BatchUpdateValuesResponse>,
  ) {
    if (arguments.length === 3) {
      cells = worksheetName as { col: number; row: number; value: string; }[];
      worksheetName = spreadsheetId;
      spreadsheetId = credentials as string;
    } else if (arguments.length === 4) {
      if (typeof cells === 'function') {
        callback = cells;
        cells = worksheetName as { col: number; row: number; value: string; }[];
        worksheetName = spreadsheetId;
        spreadsheetId = credentials as string;
        credentials = undefined;
      }
    }
    if (!callback) {
      callback = (err: Error | null, result?: sheets_v4.Schema$BatchUpdateValuesResponse) => {
        if (err) {
          throw err;
        }
        return result;
      };
    }

    const entry: any[][] = [];
    _.forEach(cells as { col: number; row: number; value: string; }[], (cell) => {
      entry[cell.row] = entry[cell.row] || [];
      entry[cell.row][cell.col - 1] = cell.value;
    });

    try {
      const result = await api.batchCells(
        credentials as Credentials,
        spreadsheetId as string,
        worksheetName as string,
        entry,
      );
      return callback(null, result);
    } catch (e) {
      return callback(e);
    }
  }

  static get COLUMN_KEYMAP() {
    return COLUMN_KEYMAP;
  }
}

export default new Spread2Json();

