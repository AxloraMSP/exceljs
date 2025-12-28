const CompositeXform = require('../composite-xform');
const BaseXform = require('../base-xform');
const Range = require('../../../doc/range');
const utils = require('../../../utils/utils');

function containsCrossSheetReference(formula) {
  if (typeof formula !== 'string') {
    return false;
  }
  const crossSheetPattern = /(?:^|[^'"])'?[^'"]+'?!|^[^!'"]+!/;
  return crossSheetPattern.test(formula);
}

function assign(obj, attributes, name, defaultValue) {
  const value = attributes[name];
  if (value !== undefined) {
    obj[name] = value;
  } else if (defaultValue !== undefined) {
    obj[name] = defaultValue;
  }
}

function assignBool(obj, attributes, name, defaultValue) {
  const value = attributes[name];
  if (value !== undefined) {
    obj[name] = utils.parseBoolean(value);
  } else if (defaultValue !== undefined) {
    obj[name] = defaultValue;
  }
}

class DataValidationExtXform extends BaseXform {
  get tag() {
    return 'x14:dataValidation';
  }

  parseOpen(node) {
    switch (node.name) {
      case 'x14:dataValidation': {
        const dataValidation = {type: node.attributes.type || 'any', formulae: []};

        if (node.attributes.type) {
          assignBool(dataValidation, node.attributes, 'allowBlank');
        }
        assignBool(dataValidation, node.attributes, 'showInputMessage');
        assignBool(dataValidation, node.attributes, 'showErrorMessage');

        switch (dataValidation.type) {
          case 'any':
          case 'list':
          case 'custom':
            break;
          default:
            assign(dataValidation, node.attributes, 'operator', 'between');
            break;
        }
        assign(dataValidation, node.attributes, 'promptTitle');
        assign(dataValidation, node.attributes, 'prompt');
        assign(dataValidation, node.attributes, 'errorStyle');
        assign(dataValidation, node.attributes, 'errorTitle');
        assign(dataValidation, node.attributes, 'error');

        this._dataValidation = dataValidation;
        this._sqref = null;
        this._formula = null;
        return true;
      }

      case 'x14:formula1':
      case 'x14:formula2':
        this._formula = [];
        this._inFormula = true;
        return true;

      case 'xm:f':
        this._inFormulaF = true;
        this._formulaF = [];
        return true;

      case 'xm:sqref':
        this._sqref = [];
        return true;

      default:
        return false;
    }
  }

  parseText(text) {
    if (this._inFormulaF && this._formulaF) {
      this._formulaF.push(text);
    } else if (this._formula) {
      this._formula.push(text);
    }
    if (this._sqref) {
      this._sqref.push(text);
    }
  }

  render(xmlStream, model) {
    const {address, dataValidation} = model;
    if (!dataValidation || !dataValidation.formulae || !dataValidation.formulae.length) {
      return;
    }

    xmlStream.openNode('x14:dataValidation');

    if (dataValidation.type && dataValidation.type !== 'any') {
      xmlStream.addAttribute('type', dataValidation.type);
      if (dataValidation.operator && dataValidation.type !== 'list' && dataValidation.operator !== 'between') {
        xmlStream.addAttribute('operator', dataValidation.operator);
      }
      if (dataValidation.allowBlank) {
        xmlStream.addAttribute('allowBlank', '1');
      }
    }
    if (dataValidation.showInputMessage) {
      xmlStream.addAttribute('showInputMessage', '1');
    }
    if (dataValidation.promptTitle) {
      xmlStream.addAttribute('promptTitle', dataValidation.promptTitle);
    }
    if (dataValidation.prompt) {
      xmlStream.addAttribute('prompt', dataValidation.prompt);
    }
    if (dataValidation.showErrorMessage) {
      xmlStream.addAttribute('showErrorMessage', '1');
    }
    if (dataValidation.errorStyle) {
      xmlStream.addAttribute('errorStyle', dataValidation.errorStyle);
    }
    if (dataValidation.errorTitle) {
      xmlStream.addAttribute('errorTitle', dataValidation.errorTitle);
    }
    if (dataValidation.error) {
      xmlStream.addAttribute('error', dataValidation.error);
    }

    if (dataValidation.formulae[0]) {
      xmlStream.openNode('x14:formula1');
      xmlStream.openNode('xm:f');
      xmlStream.writeText(String(dataValidation.formulae[0]));
      xmlStream.closeNode();
      xmlStream.closeNode();
    }

    xmlStream.openNode('xm:sqref');
    xmlStream.writeText(address);
    xmlStream.closeNode();

    xmlStream.closeNode();
  }

  parseClose(name) {
    switch (name) {
      case 'x14:dataValidation': {
        if (!this._dataValidation.formulae || !this._dataValidation.formulae.length) {
          delete this._dataValidation.formulae;
          delete this._dataValidation.operator;
        }
        const address = this._sqref ? this._sqref.join('') : '';
        const addresses = [];
        if (address) {
          const list = address.split(/\s+/g) || [];
          list.forEach(addr => {
            if (addr.includes(':')) {
              const range = new Range(addr);
              range.forEachAddress(a => {
                addresses.push(a);
              });
            } else {
              addresses.push(addr);
            }
          });
        }
        this.model = {
          dataValidation: this._dataValidation,
          addresses,
        };
        return false;
      }

      case 'x14:formula1':
      case 'x14:formula2': {
        let formula = '';
        if (this._formulaF && this._formulaF.length > 0) {
          formula = this._formulaF.join('');
        } else if (this._formula && this._formula.length > 0) {
          formula = this._formula.join('');
        }
        const hasCrossSheetRef = containsCrossSheetReference(formula);
        if (!hasCrossSheetRef) {
          switch (this._dataValidation.type) {
            case 'whole':
            case 'textLength':
              formula = parseInt(formula, 10);
              break;
            case 'decimal':
              formula = parseFloat(formula);
              break;
            case 'date':
              formula = utils.excelToDate(parseFloat(formula));
              break;
            default:
              break;
          }
        }
        this._dataValidation.formulae.push(formula);
        this._formula = null;
        this._formulaF = null;
        this._inFormula = false;
        this._inFormulaF = false;
        return true;
      }

      case 'xm:f':
        if (this._inFormulaF) {
          this._inFormulaF = false;
        }
        return true;

      case 'xm:sqref':
        return true;

      default:
        return true;
    }
  }
}

class DataValidationsExtXform extends CompositeXform {
  constructor() {
    super();

    this.map = {
      'x14:dataValidation': (this.dvXform = new DataValidationExtXform()),
    };
  }

  get tag() {
    return 'x14:dataValidations';
  }

  hasContent(model) {
    return model && Object.keys(model).length > 0;
  }

  render(xmlStream, model) {
    if (this.hasContent(model)) {
      xmlStream.openNode(this.tag, {
        count: Object.keys(model).length,
        'xmlns:xm': 'http://schemas.microsoft.com/office/excel/2006/main',
      });
      Object.entries(model).forEach(([address, dataValidation]) => {
        this.dvXform.render(xmlStream, {address, dataValidation});
      });
      xmlStream.closeNode();
    }
  }

  createNewModel() {
    return {};
  }

  onParserClose(name, parser) {
    const {dataValidation, addresses} = parser.model;
    addresses.forEach(address => {
      this.model[address] = dataValidation;
    });
  }
}

module.exports = DataValidationsExtXform;
