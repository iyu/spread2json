/**
 * @fileOverview spread sheet api
 * @name api.js
 * @author Yuhei Aihara <yu.e.yu.4119@gmail.com>
 * https://github.com/iyu/spread2json
 */
var fs = require('fs');
var path = require('path');

var _ = require('lodash');
var async = require('async');
var googleapis = require('googleapis');
var xml2json = require('xml2json');

var logging = require('./logging');

/**
 * @see https://developers.google.com/google-apps/spreadsheets/index
 */
var SCOPE = 'https://spreadsheets.google.com/feeds';
var URLS = {
  SPREADSHEETS: function() {
    return SCOPE + '/spreadsheets/private/full?alt=json';
  },
  WORKSHEETS: function(key) {
    return SCOPE + '/worksheets/' + key + '/private/full?alt=json';
  },
  LIST: function(key, worksheetId) {
    return SCOPE + '/list/' + key + '/' + worksheetId + '/private/full?alt=json';
  },
  CELLS: function(key, worksheetId) {
    return SCOPE + '/cells/' + key + '/' + worksheetId + '/private/full?alt=json';
  },
  BATCH: function(key, worksheetId) {
    return SCOPE + '/cells/' + key + '/' + worksheetId + '/private/full/batch';
  }
};

var ESCAPE_CHARCODES = [0, 1, 2, 3, 4, 5, 6, 7, 8, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31];

var ADD_TEMPLATE = _.template(
  '<entry xmlns="http://www.w3.org/2005/Atom"\n' +
  ' xmlns:gs="http://schemas.google.com/spreadsheets/2006">\n' +
  ' <title><%= title %></title>\n' +
  ' <gs:rowCount><%= rowCount %></gs:rowCount>\n' +
  ' <gs:colCount><%= colCount %></gs:colCount>\n' +
  '</entry>'
);

var BATCH_TEMPLATE = {
  BODY: _.template(
    '<feed xmlns="http://www.w3.org/2005/Atom"\n' +
    '      xmlns:batch="http://schemas.google.com/gdata/batch"\n' +
    '      xmlns:gs="http://schemas.google.com/spreadsheets/2006">\n' +
    '  <id>https://spreadsheets.google.com/feeds/cells/<%= key %>/<%= worksheetId %>/private/full</id>\n' +
    '<%= entry %>' +
    '</feed>'),
  ENTRY: _.template(
    '  <entry>\n' +
    '    <batch:id>batchId_R<%= row %>C<%= col %></batch:id>\n' +
    '    <batch:operation type="update"/>\n' +
    '    <id>https://spreadsheets.google.com/feeds/cells/<%= key %>/<%= worksheetId %>/private/full/R<%= row %>C<%= col %></id>\n' +
    '    <link rel="edit" type="application/atom+xml"\n' +
    '      href="https://spreadsheets.google.com/feeds/cells/<%= key %>/<%= worksheetId %>/private/full/R<%= row %>C<%= col %>/version"/>\n' +
    '    <gs:cell row="<%= row %>" col="<%= col %>" inputValue="<%= value %>"/>\n' +
    '  </entry>\n')
};

var BATCH_LIMIT = 40000;

/**
 * xml escape
 * @param {string} str
 * @returns {string}
 */
function xmlEscape(str) {
  if (!_.isString(str)) {
    return str;
  }
  str = _.escape(str);
  return _.map(str, function(c) {
    return _.includes(ESCAPE_CHARCODES, c.charCodeAt()) ? '' : c;
  }).join('');
}

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
 * request
 * @param {Object} opts
 * @param {Function} callback
 */
SpreadSheetApi.prototype.request = function(opts, callback) {
  opts.headers = opts.headers || {};
  opts.headers['GData-Version'] = '3.0';

  return this.oAuth2.request(opts, function(err, body, res) {
    if (err) {
      if (err.message === 'No access or refresh token is set.' ||
          res && res.body && res.body.error_description === 'Missing required parameter: refresh_token') {
        err.status = 401;
      } else {
        logging.logger.error(err.stack || err);
        if (body) {
          logging.logger.error(body);
        }
      }
      return callback(err);
    }

    logging.logger.debug(res.req.method, res.req.path);
    var _err;
    if (/application\/atom\+xml/.test(res.headers['content-type'])) {
      try {
        body = JSON.parse(xml2json.toJson(body));
        if (body.feed && body.feed.entry && body.feed.entry['batch:status']) {
          res.statusCode = body.feed.entry['batch:status'].code;
          if (res.statusCode >= 400) {
            _err = new Error(body.feed.entry.content);
            _err.status = res.statusCode;
            logging.logger.error(_err);
            return callback(_err);
          }
        }
        if (body.feed && body.feed.entry && body.feed.entry['batch:interrupted']) {
          _err = new Error(body.feed.entry.content);
          logging.logger.error(JSON.stringify(body.feed.entry['batch:interrupted']));
          return callback(_err);
        }
      } catch (e) { /* no op */ }
    }
    if (res.statusCode >= 400) {
      _err = new Error(body);
      _err.status = res.statusCode;
      logging.logger.error(_err);
      return callback(_err);
    }

    return callback(null, body, res);
  });
};

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
    uri: URLS.SPREADSHEETS()
  };
  return this.request(opts, callback);
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
    uri: URLS.WORKSHEETS(key)
  };
  return this.request(opts, callback);
};

/**
 * add worksheet
 * @param {Credentials} [credentials] - oAuth2.getToken result
 * @param {string} key
 * @param {string} title - worksheet title
 * @param {number} rowCount
 * @param {number} colCount
 * @param {Function} callback
 */
SpreadSheetApi.prototype.addWorksheet = function(credentials, key, title, rowCount, colCount, callback) {
  if (!this.oAuth2) {
    return callback(new Error('not been setup yet.'));
  }
  if (credentials) {
    this.oAuth2.setCredentials(credentials);
  }

  var opts = {
    uri: URLS.WORKSHEETS(key),
    method: 'POST',
    headers: {
      'Content-Type': 'application/atom+xml'
    },
    body: ADD_TEMPLATE({
      title: xmlEscape(title),
      rowCount: rowCount,
      colCount: colCount
    })
  };
  return this.request(opts, callback);
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
    uri: URLS.LIST(key, worksheetId)
  };
  return this.request(opts, callback);
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
    uri: URLS.CELLS(key, worksheetId)
  };
  return this.request(opts, callback);
};

/**
 * get cells raw
 * @param {Credentials} [credentials] - oAuth2.getToken result
 * @param {string} key
 * @param {string} worksheetId
 * @param {Object[]} entry
 * @param {number|string} entry.row
 * @param {number|string} entry.col
 * @param {number|string} entry.value
 * @param {Function} callback
 */
SpreadSheetApi.prototype.batchCells = function(credentials, key, worksheetId, entry, callback) {
  if (!this.oAuth2) {
    return callback(new Error('not been setup yet.'));
  }
  if (credentials) {
    this.oAuth2.setCredentials(credentials);
  }

  var entryXmlList = _.map(entry, function(v) {
    return BATCH_TEMPLATE.ENTRY({
      key: key,
      worksheetId: worksheetId,
      row: v.row,
      col: v.col,
      value: xmlEscape(v.value)
    });
  });

  if (entryXmlList.length < BATCH_LIMIT) {
    var opts = {
      uri: URLS.BATCH(key, worksheetId),
      method: 'POST',
      headers: {
        'Content-Type': 'application/atom+xml',
        'If-Match': '*'
      },
      body: BATCH_TEMPLATE.BODY({ key: key, worksheetId: worksheetId, entry: entryXmlList.join('') })
    };
    return this.request(opts, callback);
  }

  var self = this;
  var unit = Math.ceil(entryXmlList.length / BATCH_LIMIT);
  var splitEntryXmlList = _.times(unit, function() {
    return entryXmlList.splice(0, BATCH_LIMIT);
  });
  async.eachSeries(splitEntryXmlList, function(list, next) {
    var opts = {
      uri: URLS.BATCH(key, worksheetId),
      method: 'POST',
      headers: {
        'Content-Type': 'application/atom+xml',
        'If-Match': '*'
      },
      body: BATCH_TEMPLATE.BODY({ key: key, worksheetId: worksheetId, entry: list.join('') })
    };
    self.request(opts, next);
  }, callback);
};
