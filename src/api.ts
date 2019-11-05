/**
 * @fileOverview spread sheet api
 * @name api
 * @author Yuhei Aihara <yu.e.yu.4119@gmail.com>
 * https://github.com/iyu/spread2json
 */

import fs from 'fs';
import path from 'path';

import _ from 'lodash';
import { google, drive_v3, sheets_v4 } from 'googleapis';
import { OAuth2Client } from 'googleapis-common';
import { Credentials } from 'google-auth-library';

/**
 * @see https://developers.google.com/sheets/guides/authorizing
 */
const SCOPE = [
  'https://www.googleapis.com/auth/spreadsheets',
  'https://www.googleapis.com/auth/drive',
];

class SpreadSheetApi {
  private opts = {
    client_id: undefined,
    client_secret: undefined,
    redirect_url: 'http://localhost',
    token_file: {
      use: true,
      path: './dist/token.json',
    },
  };
  private oAuth2?: OAuth2Client;

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
  setup(options: {
    client_id: string;
    client_secret: string;
    redirect_url?: string;
    token_file?: {
      use: boolean;
      path: string;
    }
  }) {
    const opts = _.assign(this.opts, options);
    this.oAuth2 = new google.auth.OAuth2(opts.client_id, opts.client_secret, opts.redirect_url);
    if (!this.opts.token_file || !this.opts.token_file.path) {
      this.opts.token_file = {
        use: false,
        path: './dist/token.json',
      };
    }
    if (!this.opts.token_file.use) {
      return this;
    }
    const tokenDir = path.dirname(this.opts.token_file.path);
    if (!fs.existsSync(tokenDir)) {
      fs.mkdirSync(tokenDir);
    }
    if (fs.existsSync(this.opts.token_file.path)) {
      const token = JSON.parse(fs.readFileSync(this.opts.token_file.path, 'utf8'));
      this.oAuth2.setCredentials(token);
    }

    return this;
  }

  /**
   * generate auth url
   * @see https://github.com/google/google-api-nodejs-client/#generating-an-authentication-url
   * @param {Object} [opts={}]
   * @param {string|Array} [opts.scope=SCOPE]
   * @return {string|Error} - URL to consent page.
   */
  generateAuthUrl(opts: { scope: string[] }) {
    if (!this.oAuth2) {
      return new Error('not been setup yet.');
    }
    opts = opts || {};
    opts.scope = opts.scope || SCOPE;

    return this.oAuth2.generateAuthUrl.call(this.oAuth2, opts);
  }

  /**
   * get access token
   * @param {string} code
   */
  async getAccessToken(code: string): Promise<Credentials | null | undefined> {
    if (!this.oAuth2) {
      throw new Error('not been setup yet.');
    }

    let result: Credentials;
    try {
      ({ tokens: result } = await this.oAuth2.getToken(code));
    } catch (e) {
      e.status = 401;
      throw e;
    }

    if (this.opts.token_file.use && this.oAuth2 && result) {
      this.oAuth2.setCredentials(result);
      fs.writeFileSync(this.opts.token_file.path, JSON.stringify(result));
    }
    return result;
  }

  /**
   * refresh access token
   * @param {Credentials} [credentials] - oAuth2.getToken result
   */
  async refreshAccessToken(credentials?: Credentials): Promise<Credentials> {
    if (!this.oAuth2) {
      throw new Error('not been setup yet.');
    }
    if (credentials) {
      this.oAuth2.setCredentials(credentials);
    }

    let result: Credentials;
    try {
      ({ credentials: result } = await this.oAuth2.refreshAccessToken());
    } catch (e) {
      e.status = 401;
      throw e;
    }

    if (this.opts.token_file.use) {
      this.oAuth2.setCredentials(result);
      fs.writeFileSync(this.opts.token_file.path, JSON.stringify(result));
    }
    return result;
  }

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
  async getSpreadsheet(
    credentials?: Credentials,
    opts?: drive_v3.Params$Resource$Files$List,
  ): Promise<drive_v3.Schema$File[] | undefined> {
    if (!this.oAuth2) {
      throw new Error('not been setup yet.');
    }
    if (credentials) {
      this.oAuth2.setCredentials(credentials);
    }

    const drive = google.drive('v3');

    const result = await drive.files.list(_.assign({
      auth: this.oAuth2,
      q: 'mimeType=\'application/vnd.google-apps.spreadsheet\'',
    }, opts));
    return result.data.files;
  }

  /**
   * get worksheet
   * @param {Credentials} [credentials] - oAuth2.getToken result
   * @param {string} spreadsheetId
   * @param {Function} callback
   */
  async getWorksheet(
    credentials: Credentials | undefined,
    spreadsheetId: string,
  ): Promise<sheets_v4.Schema$Spreadsheet> {
    if (!this.oAuth2) {
      throw new Error('not been setup yet.');
    }
    if (credentials) {
      this.oAuth2.setCredentials(credentials);
    }

    const sheets = google.sheets('v4');

    const result = await sheets.spreadsheets.get({
      auth: this.oAuth2,
      spreadsheetId,
    });

    return result.data;
  }

  /**
   * add worksheet
   * @param {Credentials} [credentials] - oAuth2.getToken result
   * @param {string} spreadsheetId
   * @param {string} title - worksheet title
   * @param {number} rowCount
   * @param {number} columnCount
   */
  async addWorksheet(
    credentials: Credentials | undefined,
    spreadsheetId: string,
    title: string,
    rowCount: number,
    columnCount: number,
  ): Promise<sheets_v4.Schema$BatchUpdateSpreadsheetResponse> {
    if (!this.oAuth2) {
      throw new Error('not been setup yet.');
    }
    if (credentials) {
      this.oAuth2.setCredentials(credentials);
    }

    const sheets = google.sheets('v4');

    const result = await sheets.spreadsheets.batchUpdate({
      auth: this.oAuth2,
      spreadsheetId,
      requestBody: {
        requests: [
          {
            addSheet: {
              properties: {
                title,
                gridProperties: {
                  rowCount,
                  columnCount,
                  frozenRowCount: 3,
                  frozenColumnCount: 1,
                },
                tabColor: {
                  red: 1,
                  green: 0,
                  blue: 0,
                  alpha: 1,
                },
              },
            },
          },
        ],
      },
    });

    return result.data;
  }

  /**
   * get list raw
   * @param {Credentials} [credentials] - oAuth2.getToken result
   * @param {string} spreadsheetId
   * @param {string} sheetName
   */
  async getList(
    credentials: Credentials | null,
    spreadsheetId: string,
    sheetName: string,
  ): Promise<sheets_v4.Schema$ValueRange> {
    if (!this.oAuth2) {
      throw new Error('not been setup yet.');
    }
    if (credentials) {
      this.oAuth2.setCredentials(credentials);
    }

    const sheets = google.sheets('v4');

    const result = await sheets.spreadsheets.values.get({
      auth: this.oAuth2,
      spreadsheetId,
      range: sheetName,
      majorDimension: 'ROWS',
      valueRenderOption: 'FORMATTED_VALUE',
      dateTimeRenderOption: 'FORMATTED_STRING',
    });

    return result.data;
  }

  /**
   * get cells raw
   * @param {Credentials} [credentials] - oAuth2.getToken result
   * @param {string} spreadsheetId
   * @param {string} sheetName
   * @param {Array[]} entry
   */
  async batchCells(
    credentials: Credentials | undefined,
    spreadsheetId: string,
    sheetName: string,
    entry: any[][],
  ): Promise<sheets_v4.Schema$BatchUpdateValuesResponse> {
    if (!this.oAuth2) {
      throw new Error('not been setup yet.');
    }
    if (credentials) {
      this.oAuth2.setCredentials(credentials);
    }

    const sheets = google.sheets('v4');

    const result = await sheets.spreadsheets.values.batchUpdate({
      auth: this.oAuth2,
      spreadsheetId,
      requestBody: {
        valueInputOption: 'RAW',
        data: [
          {
            range: sheetName,
            values: entry,
          },
        ],
      },
    });

    return result.data;
  }
}

export default new SpreadSheetApi();
