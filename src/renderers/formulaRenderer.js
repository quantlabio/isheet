import {fastInnerText, addClass, removeClass} from './../helpers/dom/element';
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
  //set cellProperties
  if(cellProperties.fontWeight != null)
    TD.style.fontWeight = cellProperties.fontWeight;
  if(cellProperties.color != null)
    TD.style.color = cellProperties.color;
  if(cellProperties.background != null)
    TD.style.background = cellProperties.background;

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
    if (instance.plugin.utils.isSet(error)) {
      addClass(TD, 'formula-error');
    } else if (instance.plugin.utils.isSet(result)) {
      removeClass(TD, 'formula-error');
      addClass(TD, 'formula');
    }
  }

  var escaped = stringify(value);
  fastInnerText(TD, escaped);

  getRenderer('base').apply(this, arguments);
}

export default formulaRenderer;
