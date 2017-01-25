/**
 * @fileOverview spread sheet api
 * @name api.js
 * @author Yuhei Aihara <yu.e.yu.4119@gmail.com>
 * https://github.com/iyu/spread2json
 */
'use strict';

var fs = require('fs');
var path = require('path');
var querystring = require('querystring');

var _ = require('lodash');
var google = require('googleapis');

/**
 * @see https://developers.google.com/sheets/guides/authorizing
 */
var SCOPE = [
  'https://www.googleapis.com/auth/spreadsheets',
  'https://www.googleapis.com/auth/drive'
];

function SpreadSheetApi() {
  this.opts = {
    client_id: undefined,
    client_secret: undefined,
    redirect_url: 'http://localhost',
    token_file: {
      use: true,
      path: './dist/token.json'
    }
  };
  this.oAuth2 = undefined;
}

module.exports = new SpreadSheetApi();

/**
 * setup
 * @param {Object} options
 * @param {string} options.client_id
 * @param {string} options.client_secret
 * @param {string} [options.redirect_url='http://localhost']
 * @param {Object} [options.token_file]
 * @param {boolean} [options.token_file.use=true]
 * @param {string} [options.token_file.path='./dist/token.json']
 * @example
 * options = {
 *   client_id: 'xxx',
 *   client_secret: 'xxx',
 *   redirect_url: 'http://localhost',
 *   token_file: {
 *     use: true
 *     path: './token.json'
 *   }
 * }
 */
SpreadSheetApi.prototype.setup = function(options) {
  _.assign(this.opts, options);
  this.oAuth2 = new google.auth.OAuth2(this.opts.client_id, this.opts.client_secret, this.opts.redirect_url);
  if (!this.opts.token_file || !this.opts.token_file.path) {
    this.opts.token_file = {
      use: false
    };
  }
  if (!this.opts.token_file.use) {
    return this;
  }
  var tokenDir = path.dirname(this.opts.token_file.path);
  if (!fs.existsSync(tokenDir)) {
    fs.mkdirSync(tokenDir);
  }
  if (fs.existsSync(this.opts.token_file.path)) {
    var token = JSON.parse(fs.readFileSync(this.opts.token_file.path, 'utf8'));
    this.oAuth2.setCredentials(token);
  }

  return this;
};

/**
 * generate auth url
 * @see https://github.com/google/google-api-nodejs-client/#generating-an-authentication-url
 * @param {Object} [opts={}]
 * @param {string|Array} [opts.scope=SCOPE]
 * @return {string|Error} - URL to consent page.
 */
SpreadSheetApi.prototype.generateAuthUrl = function(opts) {
  if (!this.oAuth2) {
    return new Error('not been setup yet.');
  }
  opts = opts || {};
  opts.scope = opts.scope || SCOPE;

  return this.oAuth2.generateAuthUrl.call(this.oAuth2, opts);
};

/**
 * get access token
 * @param {string} code
 * @param {Function} callback
 */
SpreadSheetApi.prototype.getAccessToken = function(code, callback) {
  var self = this;
  if (!this.oAuth2) {
    return callback(new Error('not been setup yet.'));
  }

  return this.oAuth2.getToken(code, function(err, result) {
    if (err) {
      err.status = 401;
      return callback(err);
    }

    if (self.opts.token_file.use) {
      self.oAuth2.setCredentials(result);
      fs.writeFileSync(self.opts.token_file.path, JSON.stringify(result));
    }
    callback(null, result);
  });
};

/**
 * refresh access token
 * @param {Credentials} [credentials] - oAuth2.getToken result
 * @param {Function} callback
 */
SpreadSheetApi.prototype.refreshAccessToken = function(credentials, callback) {
  var self = this;
  if (!this.oAuth2) {
    return callback(new Error('not been setup yet.'));
  }
  if (arguments.length === 1) {
    callback = credentials;
    credentials = undefined;
  }
  if (credentials) {
    this.oAuth2.setCredentials(credentials);
  }

  return this.oAuth2.refreshAccessToken(function(err, result) {
    if (err) {
      err.status = 401;
      return callback(err);
    }

    if (self.opts.token_file.use) {
      self.oAuth2.setCredentials(result);
      fs.writeFileSync(self.opts.token_file.path, JSON.stringify(result));
    }
    callback(null, result);
  });
};

/**
 * get spreadsheet
 * @param {Credentials} [credentials] - oAuth2.getToken result
 * @param {Object} [opts] - drive options @see https://developers.google.com/drive/v3/reference/files/list
 * @param {string} [opts.corpus]
 * @param {string} [opts.orderBy]
 * @param {integer} [opts.pageSize]
 * @param {string} [opts.pageToken]
 * @param {string} [opts.q]
 * @param {string} [opts.spaces]
 * @param {Function} callback
 */
SpreadSheetApi.prototype.getSpreadsheet = function(credentials, opts, callback) {
  if (!this.oAuth2) {
    return callback(new Error('not been setup yet.'));
  }
  if (credentials) {
    this.oAuth2.setCredentials(credentials);
  }

  var drive = google.drive('v3');
  drive.files.list(_.assign({
    auth: this.oAuth2,
    q: 'mimeType=\'application/vnd.google-apps.spreadsheet\''
  }, opts), function(err, result) {
    if (err) {
      return callback(err);
    }

    callback(null, result.files);
  });
};

/**
 * get worksheet
 * @param {Credentials} [credentials] - oAuth2.getToken result
 * @param {string} spreadsheetId
 * @param {Function} callback
 */
SpreadSheetApi.prototype.getWorksheet = function(credentials, spreadsheetId, callback) {
  if (!this.oAuth2) {
    return callback(new Error('not been setup yet.'));
  }
  if (credentials) {
    this.oAuth2.setCredentials(credentials);
  }

  var sheets = google.sheets('v4');
  sheets.spreadsheets.get({
    auth: this.oAuth2,
    spreadsheetId: spreadsheetId
  }, function(err, result) {
    if (err) {
      return callback(err);
    }

    callback(null, result);
  });
};

/**
 * add worksheet
 * @param {Credentials} [credentials] - oAuth2.getToken result
 * @param {string} spreadsheetId
 * @param {string} title - worksheet title
 * @param {number} rowCount
 * @param {number} columnCount
 * @param {Function} callback
 */
SpreadSheetApi.prototype.addWorksheet = function(credentials, spreadsheetId, title, rowCount, columnCount, callback) {
  if (!this.oAuth2) {
    return callback(new Error('not been setup yet.'));
  }
  if (credentials) {
    this.oAuth2.setCredentials(credentials);
  }

  var sheets = google.sheets('v4');
  sheets.spreadsheets.batchUpdate({
    auth: this.oAuth2,
    spreadsheetId: spreadsheetId,
    resource: {
      requests: [
        {
          addSheet: {
            properties: {
              title: title,
              gridProperties: {
                rowCount: rowCount,
                columnCount: columnCount,
                frozenRowCount: 3,
                frozenColumnCount: 1
              },
              tabColor: {
                red: 1,
                green: 0,
                blue: 0,
                alpha: 1
              }
            }
          }
        }
      ]
    }
  }, function(err, result) {
    if (err) {
      return callback(err);
    }

    callback(null, result);
  });
};

/**
 * get list raw
 * @param {Credentials} [credentials] - oAuth2.getToken result
 * @param {string} spreadsheetId
 * @param {string} sheetName
 * @param {Function} callback
 */
SpreadSheetApi.prototype.getList = function(credentials, spreadsheetId, sheetName, callback) {
  if (!this.oAuth2) {
    return callback(new Error('not been setup yet.'));
  }
  if (credentials) {
    this.oAuth2.setCredentials(credentials);
  }

  var sheets = google.sheets('v4');
  sheets.spreadsheets.values.get({
    auth: this.oAuth2,
    spreadsheetId: spreadsheetId,
    range: querystring.escape(sheetName),
    majorDimension: 'ROWS',
    valueRenderOption: 'FORMATTED_VALUE',
    dateTimeRenderOption: 'FORMATTED_STRING'
  }, function(err, result) {
    if (err) {
      return callback(err);
    }

    callback(null, result);
  });
};

/**
 * get cells raw
 * @param {Credentials} [credentials] - oAuth2.getToken result
 * @param {string} spreadsheetId
 * @param {string} sheetName
 * @param {Array[]} entry
 * @param {Function} callback
 */
SpreadSheetApi.prototype.batchCells = function(credentials, spreadsheetId, sheetName, entry, callback) {
  if (!this.oAuth2) {
    return callback(new Error('not been setup yet.'));
  }
  if (credentials) {
    this.oAuth2.setCredentials(credentials);
  }

  var sheets = google.sheets('v4');
  sheets.spreadsheets.values.batchUpdate({
    auth: this.oAuth2,
    spreadsheetId: spreadsheetId,
    resource: {
      valueInputOption: 'RAW',
      data: [
        {
          range: sheetName,
          values: entry
        }
      ]
    }
  }, function(err, result) {
    if (err) {
      return callback(err);
    }

    callback(null, result);
  });
};
