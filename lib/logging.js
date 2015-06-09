/**
 * @fileOverview logging
 * @name logging.js
 * @author Yuhei Aihara <yu.e.yu.4119@gmail.com>
 * https://github.com/iyu/spread2json
 */
var path = require('path');

var originalLogger = {
  COLOR: {
    BLACK: '\u001b[30m',
    RED: '\u001b[31m',
    GREEN: '\u001b[32m',
    YELLOW: '\u001b[33m',
    BLUE: '\u001b[34m',
    MAGENTA: '\u001b[35m',
    CYAN: '\u001b[36m',
    WHITE: '\u001b[37m',
    RESET: '\u001b[0m'
  },
  _getDateLine: function() {
    var date = new Date();
    var year = date.getFullYear();
    var month = date.getMonth() + 1;
    month = month > 10 ? month : '0' + month;
    var day = date.getDate();
    day = day > 10 ? day : '0' + day;
    var time = date.toLocaleTimeString();
    return '[' + year + '-' + month + '-' + day + ' ' + time + ']';
  },
  _prepareStackTrace: function(err, stack) {
    var stackLine = stack[1];
    var filename = path.relative('./', stackLine.getFileName());
    return '(' + filename + ':' + stackLine.getLineNumber() + ')';
  },
  _getFileLineNumber: function() {
    var obj = {};
    var original = Error.prepareStackTrace;
    Error.prepareStackTrace = this._prepareStackTrace;
    Error.captureStackTrace(obj, this._getFileLineNumber);
    var stack = obj.stack;
    Error.prepareStackTrace = original;

    return stack;
  },
  info: function() {
    Array.prototype.unshift.call(arguments, this.COLOR.RESET);
    Array.prototype.unshift.call(arguments, '[INFO]');
    Array.prototype.unshift.call(arguments, this._getDateLine());
    Array.prototype.unshift.call(arguments, this.COLOR.GREEN);
    Array.prototype.push.call(arguments, this._getFileLineNumber());
    console.info.apply(null, arguments);
  },
  debug: function() {
    Array.prototype.unshift.call(arguments, this.COLOR.RESET);
    Array.prototype.unshift.call(arguments, '[DEBUG]');
    Array.prototype.unshift.call(arguments, this._getDateLine());
    Array.prototype.unshift.call(arguments, this.COLOR.BLUE);
    Array.prototype.push.call(arguments, this._getFileLineNumber());
    console.log.apply(null, arguments);
  },
  error: function() {
    Array.prototype.unshift.call(arguments, this.COLOR.RESET);
    Array.prototype.unshift.call(arguments, '[ERROR]');
    Array.prototype.unshift.call(arguments, this._getDateLine());
    Array.prototype.unshift.call(arguments, this.COLOR.RED);
    Array.prototype.push.call(arguments, this._getFileLineNumber());
    console.error.apply(null, arguments);
  }
};
Object.freeze(originalLogger);

module.exports = {
  logger: originalLogger,
  revert: function() {
    this.logger = originalLogger;
    return this.logger;
  }
};
