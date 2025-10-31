const {PassThrough} = require('stream');
const express = require('express');
const axios = require('axios');
const testutils = require('../utils/index');

const Excel = verquire('exceljs');

describe('Express', () => {
  let server;
  let port;

  before(async () => {
    const app = express();
    app.get('/workbook', (req, res) => {
      const wb = testutils.createTestBook(new Excel.Workbook(), 'xlsx');
      res.setHeader(
        'Content-Type',
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
      );
      res.setHeader('Content-Disposition', 'attachment; filename=Report.xlsx');
      wb.xlsx.write(res).then(() => {
        res.end();
      });
    });
    // Use dynamic port assignment (port 0 = OS assigns available port)
    await new Promise(resolve => {
      server = app.listen(0, () => {
        port = server.address().port;
        resolve();
      });
    });
  });

  after(() => {
    if (server) {
      server.close();
    }
  });

  it('downloads a workbook', async () => {
    const response = await axios({
      method: 'get',
      url: `http://127.0.0.1:${port}/workbook`,
      responseType: 'stream',
      decompress: false,
    });
    const wb2 = new Excel.Workbook();
    await wb2.xlsx.read(response.data.pipe(new PassThrough()));
    testutils.checkTestBook(wb2, 'xlsx');
  });
});
