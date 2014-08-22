/**
 * @fileOverview spread sheet api
 * @name api.js
 * @author Yuhei Aihara <aihara_yuhei@cyberagent.co.jp>
 * https://github.com/yuhei-a/spread2json
 */
var fs = require('fs');

var _ = require('lodash');
var google = require('googleapis');

var SCOPE = 'https://spreadsheets.google.com/feeds';

function SpreadSheetApi() {
    this.opts = {
        client_id: undefined,
        client_secret: undefined,
        redirect_url: 'http://localhost',
        token_path: './dist/token.json'
    };
    this.oAuth2 = undefined;
}

module.exports = new SpreadSheetApi();

/**
 * setup
 * @param {Object} options
 * @example
 * options = {
 *     client_id: 'xxx',
 *     client_secret: 'xxx',
 *     redirect_url: 'http://localhost',
 *     token_path: './token.json'
 * }
 */
SpreadSheetApi.prototype.setup = function(options) {
    _.extend(this.opts, options);
    this.oAuth2 = new google.auth.OAuth2(this.opts.client_id, this.opts.client_secret, this.opts.redirect_url);
    if (fs.existsSync(this.opts.token_path)) {
        var token = JSON.parse(fs.readFileSync(this.opts.token_path, 'utf8'));
        this.oAuth2.setCredentials(token);
    }

    return this;
};

/**
 * generate auth url
 */
SpreadSheetApi.prototype.generateAuthUrl = function() {
    if (!this.oAuth2) {
        return new Error('not been setup yet.');
    }

    return this.oAuth2.generateAuthUrl.apply(this.oAuth2, arguments);
};

/**
 * get access token
 * @param {String} code
 * @param {Function} callback
 */
SpreadSheetApi.prototype.getAccessToken = function(code, callback) {
    var self = this;
    if (!this.oAuth2) {
        return callback(new Error('not been setup yet.'));
    }

    return this.oAuth2.getToken(code, function(err, result) {
        if (err) {
            return callback(err);
        }

        self.oAuth2.setCredentials(result);
        fs.writeFileSync(self.opts.token_path, JSON.stringify(result));
        callback(null, result);
    });
};

/**
 * refresh access token
 * @param {Function} callback
 */
SpreadSheetApi.prototype.refreshAccessToken = function(callback) {
    var self = this;
    if (!this.oAuth2) {
        return callback(new Error('not been setup yet.'));
    }

    return this.oAuth2.refreshAccessToken(function(err, result) {
        if (err) {
            return callback(err);
        }

        self.oAuth2.setCredentials(result);
        fs.writeFileSync(self.opts.token_path, JSON.stringify(result));
        callback(null, result);
    });
};

/**
 * Spread Sheet APIs
 */

/**
 * get worksheet
 * @param {String} key
 * @param {Function} callback
 */
SpreadSheetApi.prototype.getWorksheet = function(key, callback) {
    if (!this.oAuth2) {
        return callback(new Error('not been setup yet.'));
    }

    var opts = {
        url: SCOPE + '/worksheets/' + key + '/private/basic?alt=json'
    };
    return this.oAuth2.request(opts, callback);
};

/**
 * get list raw
 * @param {String} key
 * @param {String} worksheetId
 * @param {Function} callback
 */
SpreadSheetApi.prototype.getList = function(key, worksheetId, callback) {
    if (!this.oAuth2) {
        return callback(new Error('not been setup yet.'));
    }

    var opts = {
        url: SCOPE + '/list/' + key + '/' + worksheetId  + '/private/basic?alt=json&start-index=2'
    };
    return this.oAuth2.request(opts, callback);
};

/**
 * get cells raw
 * @param {String} key
 * @param {String} worksheetId
 * @param {Function} callback
 */
SpreadSheetApi.prototype.getCells = function(key, worksheetId, callback) {
    if (!this.oAuth2) {
        return callback(new Error('not been setup yet.'));
    }

    var opts = {
        url: SCOPE + '/cells/' + key + '/' + worksheetId  + '/private/basic?alt=json'
    };
    return this.oAuth2.request(opts, callback);
};
