import {fastInnerText} from './../helpers/dom/element';
import {stringify} from './../helpers/mixed';
import {getRenderer} from './index';

function isFormula(value) {
  if (value) {
    if (value[0] === '=') {
      return true;
    }
  }
  return false;
}

/**
 * @private
 * @renderer formulaRenderer
 * @param instance
 * @param TD
 * @param row
 * @param col
 * @param prop
 * @param value
 * @param cellProperties
 */
function formulaRenderer(instance, TD, row, col, prop, value, cellProperties) {
  if (isFormula(value)) {
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
    const className = cellProperties.className || '';

    let classArr = className.length ? className.split(' ') : [];

    if (instance.plugin.utils.isSet(error)) {
      if (classArr.indexOf('formula-error') < 0) {
        classArr.push('formula-error');
      }
    } else if (instance.plugin.utils.isSet(result)) {
      if (classArr.indexOf('formula-error') >= 0) {
        classArr.splice(classArr.indexOf('formula-error'), 1);
      }
      if (classArr.indexOf('formula') < 0) {
        classArr.push('formula');
      }
    }

    cellProperties.className = classArr.join(' ');
  }

  var escaped = stringify(value);
  fastInnerText(TD, escaped);

  getRenderer('base').apply(this, arguments);
}

export default formulaRenderer;
