/* global describe, it, before */

'use strict';

const should = require('should');
const sinon = require('sinon');

const _ = require('lodash');
const spread2json = require('../');
const api = require('../lib/api');

const opts = require('./opts');
const mock = require('./mock');
const data = require('./data');

const SPREADSHEET_KEY = '1YXVzaaxqkPKsr-excIOXScnTQC7y_DKrUKs0ukzSIgo';
const WORKSHEET_NAMES = [
  'Test1',
  'Test1.list',
  'Test1.#list.list',
  'Test2',
  'Test2.map',
  'Test2.map.$.map',
];

describe('spread2json', () => {
  before(() => {
    spread2json.setup({ api: opts.installed });
    const sandbox = sinon.sandbox.create();
    sandbox.stub(api, 'getWorksheet').yields(null, mock.getWorksheet);
    const stubGetList = sandbox.stub(api, 'getList');
    _.forEach(mock.getList, (d, key) => {
      stubGetList.withArgs(undefined, SPREADSHEET_KEY, key).yields(null, mock.getList[key]);
    });
  });

  it('#getWorksheet', (done) => {
    spread2json.getWorksheet(SPREADSHEET_KEY, (err, result) => {
      should.not.exist(err);
      should.exist(result);
      result.should.have.length(6);
      result.forEach((d, i) => {
        d.should.have.property('title', WORKSHEET_NAMES[i]);
        d.should.have.property('sheetId');
        d.should.have.property('link');
      });

      done();
    });
  });

  it('#getWorksheetDatas', (done) => {
    spread2json.getWorksheetDatas(SPREADSHEET_KEY, WORKSHEET_NAMES, (err, result) => {
      should.not.exist(err);
      should.exist(result);
      result.should.have.length(6);
      result.forEach((d) => {
        d.should.have.property('name');
        d.should.have.property('opts');
        d.should.have.property('list');
      });

      done();
    });
  });

  it('#toJson', (done) => {
    spread2json.getWorksheetDatas(SPREADSHEET_KEY, WORKSHEET_NAMES, (err, result) => {
      should.not.exist(err);
      should.exist(result);

      spread2json.toJson(result, (_err, _result) => {
        should.not.exist(_err);
        should.exist(_result);
        _.forEach(_result, (ret, key) => {
          _.forEach(ret, (_ret, _key) => {
            _ret.should.eql(data[key][_key]);
          });
        });

        done();
      });
    });
  });
});
