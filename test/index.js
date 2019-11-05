const should = require('should');
const sinon = require('sinon');

const _ = require('lodash');
const spread2json = require('../');
const { default: api } = require('../build/src/api');

const opts = require('./opts');
const mock = require('./mock');
const data = require('./data');
const schema = require('./schema');

const COLUMN_KEYMAP = spread2json.constructor.COLUMN_KEYMAP;
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
    const sandbox = sinon.createSandbox();
    sandbox.stub(api, 'getWorksheet').resolves(mock.getWorksheet);
    const stubGetList = sandbox.stub(api, 'getList');
    _.forEach(mock.getList, (d, key) => {
      stubGetList.withArgs(undefined, SPREADSHEET_KEY, key).resolves(mock.getList[key]);
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

  it('#convertSchema2Format', () => {
    const result = spread2json.convertSchema2Format(schema.Test1, ['_id']);
    should.exist(result);
    result.should.eql([
      '_id:string',
      'str:string',
      'num:number',
      'date:date',
      'bool:boolean',
      'obj.type1:string',
      'obj.type2:string',
      '$arr:number',
      '#lists.code:string',
      '#lists.bool:boolean',
      '#list.code:string',
      '#list.$arr:number',
      '#list.#list.code:string',
      '#list.#list.$arr:number',
    ]);
  });

  it('#splitSheetData', () => {
    const refKeys = ['_id'];
    const format = spread2json.convertSchema2Format(schema.Test1, refKeys);
    const rows = _.values(data.Test1);

    const result = spread2json.splitSheetData('Test1', refKeys, format, rows);

    should.exist(result);
    result.should.have.length(2);
    result[0].should.have.property('name', WORKSHEET_NAMES[0]);
    result[0].should.have.property('opts');
    result[0].should.have.property('format');
    result[0].format.should.eql([
      '_id:string',
      'str:string',
      'num:number',
      'date:date',
      'bool:boolean',
      'obj.type1:string',
      'obj.type2:string',
      '$arr:number',
      '#lists.code:string',
      '#lists.bool:boolean',
    ]);
    result[0].should.have.property('datas');
    result[0].datas.should.have.length(2);

    result[1].should.have.property('name', WORKSHEET_NAMES[1]);
    result[1].should.have.property('opts');
    result[1].opts.should.have.property('type', 'array');
    result[1].opts.should.have.property('key', 'list');
    result[1].should.have.property('format');
    result[1].format.should.eql([
      '__ref_0',
      'code:string',
      '$arr:number',
      '#list.code:string',
      '#list.$arr:number',
    ]);
    result[1].should.have.property('datas');
    result[1].datas.should.have.length(2);
  });

  it('#toCells', () => {
    const refKeys = ['_id'];
    const format = spread2json.convertSchema2Format(schema.Test1, refKeys);
    const rows = _.values(data.Test1);
    const sheetDataList = spread2json.splitSheetData('Test1', refKeys, format, rows);

    const result = _.map(sheetDataList, (sheetData) => {
      return spread2json.toCells(sheetData);
    });

    should.exist(result);
    result.should.have.length(2);
    _.forEach(result, (ret) => {
      ret.should.have.property('maxCol');
      ret.should.have.property('maxRow');
      ret.should.have.property('cells');
      ret.should.have.property('name');
      ret.should.have.property('opts');
    });
  });

  it('all', (done) => {
    const refKeys = ['_id'];
    const format = spread2json.convertSchema2Format(schema.Test1, refKeys);
    const rows = _.values(data.Test1);
    const sheetDataList = spread2json.splitSheetData('Test1', refKeys, format, rows);

    const cellDataList = _.map(sheetDataList, (sheetData) => {
      const cellData = spread2json.toCells(sheetData);
      return _.assign(cellData, {
        cells: _.map(cellData.cells, (cell) => {
          const column = COLUMN_KEYMAP[cell.col - 1];
          return _.assign(cell, { column, cell: column + cell.row });
        }),
      });
    });

    const formatDataList = _.map(cellDataList, (cellData) => {
      return _.assign({ name: cellData.name }, spread2json._format(cellData.cells));
    });

    spread2json.toJson(formatDataList, (err, result) => {
      should.not.exist(err);
      should.exist(result);
      result.should.have.property('Test1');
      result.Test1.should.eql(data.Test1);
      done();
    });
  });
});
