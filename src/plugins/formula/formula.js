import BasePlugin from './../_base';
import {registerPlugin} from './../../plugins';
import {Parser} from '../../formula';

var ruleJS = (function (root) {
  'use strict';
  var instance = this;
  var rootElement = document.getElementById(root) || null;
  var version = '0.1.0';
  var parser = {};
  var el = {};
  var Matrix = function () {
    var item = {
      id: '',
      formula: '',
      value: '',
      error: '',
      deps: [],
      formulaEdit: false
    };
    this.data = [];
    var formElements = ['input[type=text]', '[data-formula]'];
    var listen = function () {
      if (document.activeElement && document.activeElement !== document.body) {
        document.activeElement.blur();
      }
      else if (!document.activeElement) { //IE
        document.body.focus();
      }
    };
    this.getItem = function (id) {
      return instance.matrix.data.filter(function (item) {
        return item.id === id;
      })[0];
    };
    this.removeItem = function (id) {
      instance.matrix.data = instance.matrix.data.filter(function (item) {
        return item.id !== id;
      });
    };
    this.removeItemsInCol = function (col) {
      instance.matrix.data = instance.matrix.data.filter(function (item) {
        return item.col !== col;
      });
    };
    this.removeItemsInRow = function (row) {
      instance.matrix.data = instance.matrix.data.filter(function (item) {
        return item.row !== row;
      })
    };
    this.removeItemsBelowCol = function (col) {
      instance.matrix.data = instance.matrix.data.filter(function (item) {
        return item.col < col;
      });
    };
    this.removeItemsBelowRow = function (row) {
      instance.matrix.data = instance.matrix.data.filter(function (item) {
        return item.row < row;
      })
    };
    this.updateItem = function (item, props) {
      if (instance.utils.isString(item)) {
        item = instance.matrix.getItem(item);
      }

      if (item && props) {
        for (var p in props) {
          if (item[p] && instance.utils.isArray(item[p])) {
            if (instance.utils.isArray(props[p])) {
              props[p].forEach(function (i) {
                if (item[p].indexOf(i) === -1) {
                  item[p].push(i);
                }
              });
            } else {

              if (item[p].indexOf(props[p]) === -1) {
                item[p].push(props[p]);
              }
            }
          } else {
            item[p] = props[p];
          }
        }
      }
    };
    this.addItem = function (item) {
      var cellId = item.id,
          coords = instance.utils.cellCoords(cellId);

      item.row = coords.row;
      item.col = coords.col;

      var cellExist = instance.matrix.data.filter(function (cell) {
        return cell.id === cellId;
      })[0];

      if (!cellExist) {
        instance.matrix.data.push(item);
      } else {
        instance.matrix.updateItem(cellExist, item);
      }

      return instance.matrix.getItem(cellId);
    };
    this.getRefItemsToColumn = function (col) {
      var result = [];

      if (!instance.matrix.data.length) {
        return result;
      }

      instance.matrix.data.forEach(function (item) {
        if (item.deps) {
          var deps = item.deps.filter(function (cell) {

            var alpha = instance.utils.getCellAlphaNum(cell).alpha,
              num = instance.utils.toNum(alpha);

            return num >= col;
          });

          if (deps.length > 0 && result.indexOf(item.id) === -1) {
            result.push(item.id);
          }
        }
      });

      return result;
    };
    this.getRefItemsToRow = function (row) {
      var result = [];

      if (!instance.matrix.data.length) {
        return result;
      }

      instance.matrix.data.forEach(function (item) {
        if (item.deps) {
          var deps = item.deps.filter(function (cell) {
            var num = instance.utils.getCellAlphaNum(cell).num;
            return num > row;
          });

          if (deps.length > 0 && result.indexOf(item.id) === -1) {
            result.push(item.id);
          }
        }
      });

      return result;
    };
    this.updateElementItem = function (element, props) {
      var id = element.getAttribute('id'),
          item = instance.matrix.getItem(id);

      instance.matrix.updateItem(item, props);
    };
    this.getDependencies = function (id) {

      var getDependencies = function (id) {
        var filtered = instance.matrix.data.filter(function (cell) {
          if (cell.deps) {
            return cell.deps.indexOf(id) > -1;
          }
        });

        var deps = [];
        filtered.forEach(function (cell) {
          if (deps.indexOf(cell.id) === -1) {
            deps.push(cell.id);
          }
        });

        return deps;
      };

      var allDependencies = [];

      var getTotalDependencies = function (id) {
        var deps = getDependencies(id);

        if (deps.length) {
          deps.forEach(function (refId) {
            if (allDependencies.indexOf(refId) === -1) {
              allDependencies.push(refId);

              var item = instance.matrix.getItem(refId);
              if (item.deps.length) {
                getTotalDependencies(refId);
              }
            }
          });
        }
      };

      getTotalDependencies(id);

      return allDependencies;
    };
    this.getElementDependencies = function (element) {
      return instance.matrix.getDependencies(element.getAttribute('id'));
    };
    var recalculateElementDependencies = function (element) {
      var allDependencies = instance.matrix.getElementDependencies(element),
          id = element.getAttribute('id');

      allDependencies.forEach(function (refId) {
        var item = instance.matrix.getItem(refId);
        if (item && item.formula) {
          var refElement = document.getElementById(refId);
          calculateElementFormula(item.formula, refElement);
        }
      });
    };
    var calculateElementFormula = function (formula, element) {
      var parsed = parse(formula, element),
          value = parsed.result,
          error = parsed.error,
          nodeName = element.nodeName.toUpperCase();

      instance.matrix.updateElementItem(element, {value: value, error: error});

      if (['INPUT'].indexOf(nodeName) === -1) {
        element.innerText = value || error;
      }

      element.value = value || error;

      return parsed;
    };
    var registerElementInMatrix = function (element) {

      var id = element.getAttribute('id'),
          formula = element.getAttribute('data-formula');

      if (formula) {
        // add item with basic properties to data array
        instance.matrix.addItem({
          id: id,
          formula: formula
        });

        calculateElementFormula(formula, element);
      }

    };
    var registerElementEvents = function (element) {
      var id = element.getAttribute('id');

      // on db click show formula
      element.addEventListener('dblclick', function () {
        var item = instance.matrix.getItem(id);

        if (item && item.formula) {
          item.formulaEdit = true;
          element.value = '=' + item.formula;
        }
      });

      element.addEventListener('blur', function () {
        var item = instance.matrix.getItem(id);

        if (item) {
          if (item.formulaEdit) {
            element.value = item.value || item.error;
          }

          item.formulaEdit = false;
        }
      });

      // if pressed ESC restore original value
      element.addEventListener('keyup', function (event) {
        switch (event.keyCode) {
          case 13: // ENTER
          case 27: // ESC
            // leave cell
            listen();
            break;
        }
      });

      // re-calculate formula if ref cells value changed
      element.addEventListener('change', function () {
        // reset and remove item
        instance.matrix.removeItem(id);

        // check if inserted text could be the formula
        var value = element.value;

        if (value[0] === '=') {
          element.setAttribute('data-formula', value.substr(1));
          registerElementInMatrix(element);
        }

        recalculateElementDependencies(element);
      });
    };
    this.depsInFormula = function (item) {

      var formula = item.formula,
          deps = item.deps;

      if (deps) {
        deps = deps.filter(function (id) {
          return formula.indexOf(id) !== -1;
        });

        return deps.length > 0;
      }

      return false;
    };
    this.scan = function () {
      var $totalElements = rootElement.querySelectorAll(formElements);

      // iterate through elements contains specified attributes
      [].slice.call($totalElements).forEach(function ($item) {
        registerElementInMatrix($item);
        registerElementEvents($item);
      });
    };
  };
  var utils = {
    isArray: function (value) {
      return Object.prototype.toString.call(value) === '[object Array]';
    },
    isNumber: function (value) {
      return Object.prototype.toString.call(value) === '[object Number]';
    },
    isString: function (value) {
      return Object.prototype.toString.call(value) === '[object String]';
    },
    isFunction: function (value) {
      return Object.prototype.toString.call(value) === '[object Function]';
    },
    isUndefined: function (value) {
      return Object.prototype.toString.call(value) === '[object Undefined]';
    },
    isNull: function (value) {
      return Object.prototype.toString.call(value) === '[object Null]';
    },
    isSet: function (value) {
      return !instance.utils.isUndefined(value) && !instance.utils.isNull(value);
    },
    isCell: function (value) {
      return value.match(/^[A-Za-z]+[0-9]+/) ? true : false;
    },
    getCellAlphaNum: function (cell) {
      var num = cell.match(/\d+$/),
          alpha = cell.replace(num, '');

      return {
        alpha: alpha,
        num: parseInt(num[0], 10)
      }
    },
    changeRowIndex: function (cell, counter) {
      var alphaNum = instance.utils.getCellAlphaNum(cell),
          alpha = alphaNum.alpha,
          col = alpha,
          row = parseInt(alphaNum.num + counter, 10);

      if (row < 1) {
        row = 1;
      }

      return col + '' + row;
    },
    changeColIndex: function (cell, counter) {
      var alphaNum = instance.utils.getCellAlphaNum(cell),
          alpha = alphaNum.alpha,
          col = instance.utils.toChar(parseInt(instance.utils.toNum(alpha) + counter, 10)),
          row = alphaNum.num;

      if (!col || col.length === 0) {
        col = 'A';
      }

      var fixedCol = alpha[0] === '$' || false,
          fixedRow = alpha[alpha.length - 1] === '$' || false;

      col = (fixedCol ? '$' : '') + col;
      row = (fixedRow ? '$' : '') + row;

      return col + '' + row;
    },
    changeFormula: function (formula, delta, change) {
      if (!delta) {
        delta = 1;
      }

      return formula.replace(/(\$?[A-Za-z]+\$?[0-9]+)/g, function (match) {
        var alphaNum = instance.utils.getCellAlphaNum(match),
            alpha = alphaNum.alpha,
            num = alphaNum.num;

        if (instance.utils.isNumber(change.col)) {
          num = instance.utils.toNum(alpha);

          if (change.col <= num) {
            return instance.utils.changeColIndex(match, delta);
          }
        }

        if (instance.utils.isNumber(change.row)) {
          if (change.row < num) {
            return instance.utils.changeRowIndex(match, delta);
          }
        }

        return match;
      });
    },
    updateFormula: function (formula, direction, delta) {
      var type,
          counter;

      // left, right -> col
      if (['left', 'right'].indexOf(direction) !== -1) {
        type = 'col';
      } else if (['up', 'down'].indexOf(direction) !== -1) {
        type = 'row'
      }

      // down, up -> row
      if (['down', 'right'].indexOf(direction) !== -1) {
        counter = delta * 1;
      } else if(['up', 'left'].indexOf(direction) !== -1) {
        counter = delta * (-1);
      }

      if (type && counter) {
        return formula.replace(/(\$?[A-Za-z]+\$?[0-9]+)/g, function (match) {

          var alpha = instance.utils.getCellAlphaNum(match).alpha;

          var fixedCol = alpha[0] === '$' || false,
              fixedRow = alpha[alpha.length - 1] === '$' || false;

          if (type === 'row' && fixedRow) {
            return match;
          }

          if (type === 'col' && fixedCol) {
            return match;
          }

          return (type === 'row' ? instance.utils.changeRowIndex(match, counter) : instance.utils.changeColIndex(match, counter));
        });
      }

      return formula;
    },
    toNum: function (chr) {
      chr = instance.utils.clearFormula(chr);
      var base = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', i, j, result = 0;

      for (i = 0, j = chr.length - 1; i < chr.length; i += 1, j -= 1) {
        result += Math.pow(base.length, j) * (base.indexOf(chr[i]) + 1);
      }

      if (result) {
        --result;
      }

      return result;
    },
    toChar: function (num) {
      var s = '';

      while (num >= 0) {
        s = String.fromCharCode(num % 26 + 97) + s;
        num = Math.floor(num / 26) - 1;
      }

      return s.toUpperCase();
    },
    cellCoords: function (cell) {
      var num = cell.match(/\d+$/),
          alpha = cell.replace(num, '');

      return {
        row: parseInt(num[0], 10) - 1,
        col: instance.utils.toNum(alpha)
      };
    },
    clearFormula: function (formula) {
      return formula.replace(/\$/g, '');
    },
    translateCellCoords: function (coords) {
      return instance.utils.toChar(coords.col) + '' + parseInt(coords.row + 1, 10);
    },
    iterateCells: function (startCell, endCell, callback) {
      var result = {
        index: [], // list of cell index: A1, A2, A3
        value: []  // list of cell value
      };

      var cols = {
        start: 0,
        end: 0
      };

      if (endCell.col >= startCell.col) {
        cols = {
          start: startCell.col,
          end: endCell.col
        };
      } else {
        cols = {
          start: endCell.col,
          end: startCell.col
        };
      }

      var rows = {
        start: 0,
        end: 0
      };

      if (endCell.row >= startCell.row) {
        rows = {
          start: startCell.row,
          end: endCell.row
        };
      } else {
        rows = {
          start: endCell.row,
          end: startCell.row
        };
      }

      for (var column = cols.start; column <= cols.end; column++) {
        for (var row = rows.start; row <= rows.end; row++) {
          var cellIndex = instance.utils.toChar(column) + (row + 1),
              cellValue = instance.helper.cellValue.call(this, cellIndex);

          result.index.push(cellIndex);
          result.value.push(cellValue);
        }
      }

      if (instance.utils.isFunction(callback)) {
        return callback.apply(callback, [result]);
      } else {
        return result;
      }
    },
    sort: function (rev) {
      return function (a, b) {
        return ((a < b) ? -1 : ((a > b) ? 1 : 0)) * (rev ? -1 : 1);
      }
    }
  };
  var helper = {
    number: function (num) {
      switch (typeof num) {
        case 'number':
          return num;
        case 'string':
          if (!isNaN(num)) {
            return num.indexOf('.') > -1 ? parseFloat(num) : parseInt(num, 10);
          }
      }

      return num;
    },
    string: function (str) {
      return str.substring(1, str.length - 1);
    },
    numberInverted: function (num) {
      return this.number(num) * (-1);
    },
    specialMatch: function (type, exp1, exp2) {
      var result;

      switch (type) {
        case '&':
          result = exp1.toString() + exp2.toString();
          break;
      }
      return result;
    },
    logicMatch: function (type, exp1, exp2) {
      var result;

      switch (type) {
        case '=':
          result = (exp1 === exp2);
          break;

        case '>':
          result = (exp1 > exp2);
          break;

        case '<':
          result = (exp1 < exp2);
          break;

        case '>=':
          result = (exp1 >= exp2);
          break;

        case '<=':
          result = (exp1 === exp2);
          break;

        case '<>':
          result = (exp1 != exp2);
          break;

        case 'NOT':
          result = (exp1 != exp2);
          break;
      }

      return result;
    },
    mathMatch: function (type, number1, number2) {
      var result;

      number1 = helper.number(number1);
      number2 = helper.number(number2);

      if (isNaN(number1) || isNaN(number2)) {

        if (number1[0] === '=' || number2[0] === '=') {
          throw Error('NEED_UPDATE');
        }

        throw Error('VALUE');
      }

      switch (type) {
        case '+':
          result = number1 + number2;
          break;
        case '-':
          result = number1 - number2;
          break;
        case '/':
          result = number1 / number2;
          if (result == Infinity) {
            throw Error('DIV_ZERO');
          } else if (isNaN(result)) {
            throw Error('VALUE');
          }
          break;
        case '*':
          result = number1 * number2;
          break;
        case '^':
          result = Math.pow(number1, number2);
          break;
      }

      return result;
    },
    callFunction: function (fn, args) {
      fn = fn.toUpperCase();
      args = args || [];

      if (instance.formulas[fn]) {
        return instance.formulas[fn].apply(this, args);
      }

      throw Error('NAME');
    },
    callVariable: function (args) {
      args = args || [];
      var str = args[0];

      if (str) {
        str = str.toUpperCase();
        if (instance.formulas[str]) {
          return ((typeof instance.formulas[str] === 'function') ? instance.formulas[str].apply(this, args) : instance.formulas[str]);
        }
      }

      throw Error('NAME');
    },
    cellValue: function (cell) {
      var value,
          fnCellValue = instance.custom.cellValue,
          element = this,
          item = instance.matrix.getItem(cell);

      // check if custom cellValue fn exists
      if (instance.utils.isFunction(fnCellValue)) {

        var cellCoords = instance.utils.cellCoords(cell),
            cellId = instance.utils.translateCellCoords({row: element.row, col: element.col});

        // get value
        value = item ? item.value : fnCellValue(cellCoords.row, cellCoords.col);

        if (instance.utils.isNull(value)) {
          value = 0;
        }

        if (cellId) {
          //update dependencies
          instance.matrix.updateItem(cellId, {deps: [cell]});
        }

      } else {

        // get value
        value = item ? item.value : document.getElementById(cell).value;

        //update dependencies
        instance.matrix.updateElementItem(element, {deps: [cell]});
      }

      // check references error
      if (item && item.deps) {
        if (item.deps.indexOf(cellId) !== -1) {
          throw Error('REF');
        }
      }

      // check if any error occurs
      if (item && item.error) {
        throw Error(item.error);
      }

      // return value if is set
      if (instance.utils.isSet(value)) {
        var result = instance.helper.number(value);

        return !isNaN(result) ? result : value;
      }

      // cell is not available
      throw Error('NOT_AVAILABLE');
    },
    cellRangeValue: function (start, end) {
      var fnCellValue = instance.custom.cellValue,
          coordsStart = instance.utils.cellCoords(start),
          coordsEnd = instance.utils.cellCoords(end),
          element = this;

      // iterate cells to get values and indexes
      var cells = instance.utils.iterateCells.call(this, coordsStart, coordsEnd),
          result = [];

      // check if custom cellValue fn exists
      if (instance.utils.isFunction(fnCellValue)) {

        var cellId = instance.utils.translateCellCoords({row: element.row, col: element.col});

        //update dependencies
        instance.matrix.updateItem(cellId, {deps: cells.index});

      } else {

        //update dependencies
        instance.matrix.updateElementItem(element, {deps: cells.index});
      }

      result.push(cells.value);
      return result;
    },
    fixedCellValue: function (id) {
      id = id.replace(/\$/g, '');
      return instance.helper.cellValue.call(this, id);
    },
    fixedCellRangeValue: function (start, end) {
      start = start.replace(/\$/g, '');
      end = end.replace(/\$/g, '');

      return instance.helper.cellRangeValue.call(this, start, end);
    }
  };
  var parse = function (formula, element) {
      el = element;
      var result = parser.parse(formula);

      var id;
      if (element instanceof HTMLElement) {
        id = element.getAttribute('id');
      } else if (element && element.id) {
        id = element.id;
      }

      var deps = instance.matrix.getDependencies(id);

      if (deps.indexOf(id) !== -1) {
        result = null;

        deps.forEach(function (id) {
          instance.matrix.updateItem(id, {value: null, error: '#REF!'});
        });
        throw Error('REF');
      }

    return result;
  };

  var init = function () {
    instance = this;

    parser = new Parser();

    parser.on('callCellValue', function(cellCoord, done) {
      var val = instance.helper.cellValue.call(el, cellCoord.label);
      while(val[0]==='=')
        val = instance.helper.cellValue.call(el, val.substr(1));
      done(val);
    });

    parser.on('callRangeValue', function(startCellCoord, endCellCoord, done) {
      //var data = [];using the same variable passed to handsontable

      instance.helper.cellRangeValue.call(el, startCellCoord.label, endCellCoord.label);

      var fragment = [];

      for (var row = startCellCoord.row.index; row <= endCellCoord.row.index; row++) {
        var rowData = data[row];
        var colFragment = [];

        for (var col = startCellCoord.column.index; col <= endCellCoord.column.index; col++) {
          colFragment.push(rowData[col]);
        }
        fragment.push(colFragment);
      }

      if (fragment) {
        done(fragment);
      }
    });

    //instance.formulas = formulaParser.SUPPORTED_FORMULAS;
    instance.matrix = new Matrix();

    instance.custom = {};

    if (rootElement) {
      instance.matrix.scan();
    }
  };

  return {
    init: init,
    version: version,
    utils: utils,
    helper: helper,
    parse: parse
  };

});


class Formula extends BasePlugin {
   constructor(hotInstance) {
     super(hotInstance);

     this.formulaCell = {
       renderer: this.formulaRenderer,
       //editor: Handsontable.editors.TextEditor,
       dataType: 'formula'
     };

     this.plugin = new ruleJS();
     this.plugin.init();
     this.plugin.custom = {
       cellValue: hotInstance.getDataAtCell
     };

   }

   isFormula(value) {
     if (value) {
       if (value[0] === '=') {
         return true;
       }
     }
     return false;
   };

   formulaRenderer(instance, TD, row, col, prop, value, cellProperties) {
     if (instance.formulasEnabled && isFormula(value)) {
       // translate coordinates into cellId
       var cellId = instance.plugin.utils.translateCellCoords({
             row: row,
             col: col
           }),
           prevFormula = null,
           formula = null,
           needUpdate = false,
           error, result;

       if (!cellId) {
         return;
       }

       // set formula cell id attribute
       //TD.id = cellId;

       // get cell data
       var item = instance.plugin.matrix.getItem(cellId);

       if (item) {

         needUpdate = !! item.needUpdate;

         if (item.error) {
           prevFormula = item.formula;
           error = item.error;

           if (needUpdate) {
             error = null;
           }
         }
       }

       // check if typed formula or cell value should be recalculated
       if ((value && value[0] === '=') || needUpdate) {

         formula = value.substr(1).toUpperCase();

         if (!error || formula !== prevFormula) {

           var currentItem = item;

           if (!currentItem) {

             // define item to rulesJS matrix if not exists
             item = {
               id: cellId,
               formula: formula
             };

             // add item to matrix
             currentItem = instance.plugin.matrix.addItem(item);
           }

           // parse formula
           var newValue = instance.plugin.parse(formula, {
             row: row,
             col: col,
             id: cellId
           });

           // check if update needed
           needUpdate = (newValue.error === '#NEED_UPDATE');

           // update item value and error
           instance.plugin.matrix.updateItem(currentItem, {
             formula: formula,
             value: newValue.result,
             error: newValue.error,
             needUpdate: needUpdate
           });

           error = newValue.error;
           result = newValue.result;

           // update cell value in hot
           value = error || result;
         }
       }

       if (error) {
         // clear cell value
         if (!value) {
           // reset error
           error = null;
         } else {
           // show error
           value = error;
         }
       }

       // change background color
       if (instance.plugin.utils.isSet(error)) {
         Handsontable.Dom.addClass(TD, 'formula-error');
       } else if (instance.plugin.utils.isSet(result)) {
         Handsontable.Dom.removeClass(TD, 'formula-error');
         Handsontable.Dom.addClass(TD, 'formula');
       }
     }

     // apply changes
     if (cellProperties.type === 'numeric') {
       numericCell.renderer.apply(this, [instance, TD, row, col, prop, value, cellProperties]);
     } else {
       textCell.renderer.apply(this, [instance, TD, row, col, prop, value, cellProperties]);
     }
   };

   afterChange(changes, source) {
     var instance = this;

     if (!instance.formulasEnabled) {
       return;
     }

     if (source === 'edit' || source === 'undo' || source === 'autofill') {

       var rerender = false;

       changes.forEach(function(item) {

         var row = item[0],
             col = item[1],
             prevValue = item[2],
             value = item[3];

         var cellId = instance.plugin.utils.translateCellCoords({
           row: row,
           col: col
         });

         // if changed value, all references cells should be recalculated
         if (value[0] !== '=' || prevValue !== value) {
           instance.plugin.matrix.removeItem(cellId);

           // get referenced cells
           var deps = instance.plugin.matrix.getDependencies(cellId);

           // update cells
           deps.forEach(function(itemId) {
             instance.plugin.matrix.updateItem(itemId, {
               needUpdate: true
             });
           });

           rerender = true;
         }
       });

       if (rerender) {
         instance.render();
       }
     }
   };

   beforeAutofillInsidePopulate(index, direction, data, deltas, iterators, selected) {
     var instance = this;

     var r = index.row,
         c = index.col,
         value = data[r][c],
         delta = 0,
         rlength = data.length, // rows
         clength = data ? data[0].length : 0; //cols

     if (value[0] === '=') { // formula

       if (['down', 'up'].indexOf(direction) !== -1) {
         delta = rlength * iterators.row;
       } else if (['right', 'left'].indexOf(direction) !== -1) {
         delta = clength * iterators.col;
       }

       return {
         value: instance.plugin.utils.updateFormula(value, direction, delta),
         iterators: iterators
       }

     } else { // other value

       // increment or decrement  values for more than 2 selected cells
       if (rlength >= 2 || clength >= 2) {

         var newValue = instance.plugin.helper.number(value),
             ii,
             start;

         if (instance.plugin.utils.isNumber(newValue)) {

           if (['down', 'up'].indexOf(direction) !== -1) {

             delta = deltas[0][c];

             if (direction === 'down') {
               newValue += (delta * rlength * iterators.row);
             } else {

               ii = (selected.row - r) % rlength;
               start = ii > 0 ? rlength - ii : 0;

               newValue = instance.plugin.helper.number(data[start][c]);

               newValue += (delta * rlength * iterators.row);

               // last element in array -> decrement iterator
               // iterator cannot be less than 1
               if (iterators.row > 1 && (start + 1) === rlength) {
                 iterators.row--;
               }
             }

           } else if (['right', 'left'].indexOf(direction) !== -1) {
             delta = deltas[r][0];

             if (direction === 'right') {
               newValue += (delta * clength * iterators.col);
             } else {

               ii = (selected.col - c) % clength;
               start = ii > 0 ? clength - ii : 0;

               newValue = instance.plugin.helper.number(data[r][start]);

               newValue += (delta * clength * (iterators.col || 1));

               // last element in array -> decrement iterator
               // iterator cannot be less than 1
               if (iterators.col > 1 && (start + 1) === clength) {
                 iterators.col--;
               }
             }
           }

           return {
             value: newValue,
             iterators: iterators
           }
         }
       }

     }

     return {
       value: value,
       iterators: iterators
     };
   };

   afterCreateRow(row, amount, auto) {
     //if (auto) {
     //  return;
     //}

     var instance = this;

     var selectedRow = instance.plugin.utils.isArray(instance.getSelected()) ? instance.getSelected()[0] : undefined;

     if (instance.plugin.utils.isUndefined(selectedRow)) {
       return;
     }

     var direction = (selectedRow >= row) ? 'before' : 'after',
         items = instance.plugin.matrix.getRefItemsToRow(row),
         counter = 1,
         changes = [];

     items.forEach(function(id) {
       var item = instance.plugin.matrix.getItem(id),
           formula = instance.plugin.utils.changeFormula(item.formula, 1, {
             row: row
           }), // update formula if needed
           newId = id;

       if (formula !== item.formula) { // formula updated

         // change row index and get new coordinates
         if ((direction === 'before' && selectedRow <= item.row) || (direction === 'after' && selectedRow < item.row)) {
           newId = instance.plugin.utils.changeRowIndex(id, counter);
         }

         var cellCoords = instance.plugin.utils.cellCoords(newId);

         if (newId !== id) {
           // remove current item from matrix
           instance.plugin.matrix.removeItem(id);
         }

         // set updated formula in new cell
         changes.push([cellCoords.row, cellCoords.col, '=' + formula]);
       }
     });

     if (items) {
       instance.plugin.matrix.removeItemsBelowRow(row);
     }

     if (changes) {
       instance.setDataAtCell(changes);
     }
   };

   afterCreateCol(col) {
     var instance = this;

     var selectedCol = instance.plugin.utils.isArray(instance.getSelected()) ? instance.getSelected()[1] : undefined;

     if (instance.plugin.utils.isUndefined(selectedCol)) {
       return;
     }

     var items = instance.plugin.matrix.getRefItemsToColumn(col),
         counter = 1,
         direction = (selectedCol >= col) ? 'before' : 'after',
         changes = [];

     items.forEach(function(id) {

       var item = instance.plugin.matrix.getItem(id),
           formula = instance.plugin.utils.changeFormula(item.formula, 1, {
             col: col
           }), // update formula if needed
           newId = id;

       if (formula !== item.formula) { // formula updated

         // change col index and get new coordinates
         if ((direction === 'before' && selectedCol <= item.col) || (direction === 'after' && selectedCol < item.col)) {
           newId = instance.plugin.utils.changeColIndex(id, counter);
         }

         var cellCoords = instance.plugin.utils.cellCoords(newId);

         if (newId !== id) {
           // remove current item from matrix if id changed
           instance.plugin.matrix.removeItem(id);
         }

         // set updated formula in new cell
         changes.push([cellCoords.row, cellCoords.col, '=' + formula]);
       }
     });

     if (items) {
       instance.plugin.matrix.removeItemsBelowCol(col);
     }

     if (changes) {
       instance.setDataAtCell(changes);
     }
   };

   enablePlugin() {
     this.addHook('beforeInit', () => this.init());
     //this.addHook('afterUpdateSettings', () => this.init.call(this, 'afterUpdateSettings'));
     this.addHook('afterChange', (changes, source) => this.afterChange(changes, source));
     this.addHook('beforeAutofillInsidePopulate', (index, direction, data, deltas, iterators, selected) => this.beforeAutofillInsidePopulate(index, direction, data, deltas, iterators, selected));
     this.addHook('afterCreateRow', (row, amount, auto) => this.afterCreateRow(row, amount, auto));
     this.addHook('afterCreateCol', (col) => this.afterCreateCol(col));

     super.enablePlugin();
   }

   /**
    * Destroy plugin instance.
    */
   destroy() {
     super.destroy();
   }
};

registerPlugin('formula', Formula);

export default Formula;
