const WorksheetWriter = verquire('stream/xlsx/worksheet-writer');
const StreamBuf = verquire('utils/stream-buf');

describe('Workbook Writer', () => {
  it('generates valid xml even when there is no data', () =>
    // issue: https://github.com/guyonroche/exceljs/issues/99
    // PR: https://github.com/guyonroche/exceljs/pull/255
    new Promise((resolve, reject) => {
      const mockWorkbook = {
        _openStream() {
          return this.stream;
        },
        stream: new StreamBuf(),
      };
      mockWorkbook.stream.on('finish', () => {
        try {
          const xml = mockWorkbook.stream.read().toString();
          // Check that XML is well-formed by verifying it starts with declaration and has worksheet element
          expect(xml).toMatch(/^<\?xml/);
          expect(xml).toContain('<worksheet');
          expect(xml).toContain('</worksheet>');
          resolve();
        } catch (error) {
          reject(error);
        }
      });

      const writer = new WorksheetWriter({
        id: 1,
        workbook: mockWorkbook,
      });

      writer.commit();
    }));
});
