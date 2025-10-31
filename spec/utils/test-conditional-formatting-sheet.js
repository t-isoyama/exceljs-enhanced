const tools = require('./tools');

const self = {
  conditionalFormattings: tools.fix(require('./data/conditional-formatting')),
  getConditionalFormatting(type) {
    return self.conditionalFormattings[type] || null;
  },
  addSheet(wb) {
    const ws = wb.addWorksheet('conditional-formatting');
    const {types} = self.conditionalFormattings;
    types.forEach(type => {
      const conditionalFormatting = self.getConditionalFormatting(type);
      if (conditionalFormatting) {
        ws.addConditionalFormatting(conditionalFormatting);
      }
    });
  },

  checkSheet(wb) {
    const ws = wb.getWorksheet('conditional-formatting');
    expect(ws).not.toBeUndefined();
    expect(ws.conditionalFormattings).not.toBeUndefined();
    (ws.conditionalFormattings && ws.conditionalFormattings).forEach(item => {
      const type = item.rules && item.rules[0].type;
      const conditionalFormatting = self.getConditionalFormatting(type);
      expect(item).toHaveProperty('ref');
      expect(item).toHaveProperty('rules');
      expect(self.conditionalFormattings[type]).toHaveProperty('ref');
      expect(self.conditionalFormattings[type]).toHaveProperty('rules');
      expect(item.ref).toEqual(conditionalFormatting.ref);
      expect(item.rules.length).toBe(conditionalFormatting.rules.length);
    });
  },
};

module.exports = self;
