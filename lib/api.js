/**
 * @fileOverview spread sheet api
 * @name api.js
 * @author Yuhei Aihara <aihara_yuhei@cyberagent.co.jp>
 * https://github.com/yuhei-a/spread2json
 */
var fs = require('fs');
var path = require('path');

var _ = require('lodash');
var googleapis = require('googleapis');

var logger = require('./logger');

/**
 * @see https://developers.google.com/google-apps/spreadsheets/index
 */
var SCOPE = 'https://spreadsheets.google.com/feeds';
var URLS = {
  SPREADSHEETS: function() {
    return SCOPE + '/spreadsheets/private/full?alt=json';
  },
  WORKSHEETS: function(key) {
    return SCOPE + '/worksheets/' + key + '/private/basic?alt=json';
  },
  LIST: function(key, worksheetId) {
    return SCOPE + '/list/' + key + '/' + worksheetId + '/private/basic?alt=json&start-index=2';
  },
  CELLS: function(key, worksheetId) {
    return SCOPE + '/cells/' + key + '/' + worksheetId + '/private/basic?alt=json';
  }
};

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
 * request callback
 * @param {Function} callback
 */
function requestCallback(callback) {
  return function(err, data, res) {
    if (err) {
      if (err.message === 'No access or refresh token is set.' ||
          (res && res.body && res.body.error_description === 'Missing required parameter: refresh_token')) {
        err.status = 401;
      } else {
        logger.error(err.stack || err);
        if (res && res.body) {
          logger.error(res.body);
        }
        if (data) {
          logger.error(data);
        }
      }
      return callback(err);
    }
    logger.debug(res.req._header);
    if (res.statusCode !== 200) {
      logger.error(data);
      var _err = new Error(data);
      _err.status = res.statusCode;
      return callback(_err);
    }

    return callback(null, data);
  };
}

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
  _.extend(this.opts, options);
  this.oAuth2 = new googleapis.auth.OAuth2(this.opts.client_id, this.opts.client_secret, this.opts.redirect_url);
  if (!this.opts.token_file || !this.opts.token_file.path) {
    this.opts.token_file = {
      use: false
    };
  }
  var tokenDir = path.dirname(this.opts.token_file.path);
  if (this.opts.token_file.use && !fs.existsSync(tokenDir)) {
    fs.mkdirSync(tokenDir);
  }
  if (this.opts.token_file.use && fs.existsSync(this.opts.token_file.path)) {
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
 * Sheets APIs
 */

/**
 * get spreadsheet
 * @param {Credentials} [credentials] - oAuth2.getToken result
 * @param {Function} callback
 */
SpreadSheetApi.prototype.getSpreadsheet = function(credentials, callback) {
  if (!this.oAuth2) {
    return callback(new Error('not been setup yet.'));
  }
  if (credentials) {
    this.oAuth2.setCredentials(credentials);
  }

  var opts = {
    url: URLS.SPREADSHEETS()
  };
  return this.oAuth2.request(opts, requestCallback(callback));
};

/**
 * get worksheet
 * @param {Credentials} [credentials] - oAuth2.getToken result
 * @param {string} key
 * @param {Function} callback
 */
SpreadSheetApi.prototype.getWorksheet = function(credentials, key, callback) {
  if (!this.oAuth2) {
    return callback(new Error('not been setup yet.'));
  }
  if (credentials) {
    this.oAuth2.setCredentials(credentials);
  }

  var opts = {
    url: URLS.WORKSHEETS(key)
  };
  return this.oAuth2.request(opts, requestCallback(callback));
};

/**
 * get list raw
 * @param {Credentials} [credentials] - oAuth2.getToken result
 * @param {string} key
 * @param {string} worksheetId
 * @param {Function} callback
 */
SpreadSheetApi.prototype.getList = function(credentials, key, worksheetId, callback) {
  if (!this.oAuth2) {
    return callback(new Error('not been setup yet.'));
  }
  if (credentials) {
    this.oAuth2.setCredentials(credentials);
  }

  var opts = {
    url: URLS.LIST(key, worksheetId)
  };
  return this.oAuth2.request(opts, requestCallback(callback));
};

/**
 * get cells raw
 * @param {Credentials} [credentials] - oAuth2.getToken result
 * @param {string} key
 * @param {string} worksheetId
 * @param {Function} callback
 */
SpreadSheetApi.prototype.getCells = function(credentials, key, worksheetId, callback) {
  if (!this.oAuth2) {
    return callback(new Error('not been setup yet.'));
  }
  if (credentials) {
    this.oAuth2.setCredentials(credentials);
  }

  var opts = {
    url: URLS.LIST(key, worksheetId)
  };
  return this.oAuth2.request(opts, requestCallback(callback));
};
