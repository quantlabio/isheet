import {getEditor} from './../editors';
import {getRenderer} from './../renderers';
import {getValidator} from './../validators';

const CELL_TYPE = 'numeric';

export default {
  editor: getEditor(CELL_TYPE),
  renderer: getRenderer('formula'),
  validator: getValidator(CELL_TYPE),
  dataType: 'number',
};
