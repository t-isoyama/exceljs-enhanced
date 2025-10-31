const TwoCellAnchorXform = verquire('xlsx/xform/drawing/two-cell-anchor-xform');

describe('TwoCellAnchorXform', () => {
  describe('reconcile', () => {
    it('should not throw on null picture', () => {
      const twoCell = new TwoCellAnchorXform();
      expect(() => twoCell.reconcile({picture: null}, {})).not.toThrow();
    });
    it('should not throw on null tl', () => {
      const twoCell = new TwoCellAnchorXform();
      expect(() =>
        twoCell.reconcile({br: {col: 1, row: 1}}, {})
      ).not.toThrow();
    });
    it('should not throw on null br', () => {
      const twoCell = new TwoCellAnchorXform();
      expect(() =>
        twoCell.reconcile({tl: {col: 1, row: 1}}, {})
      ).not.toThrow();
    });
  });
});
