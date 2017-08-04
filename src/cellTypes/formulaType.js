import {getEditor} from './../editors';
import {getRenderer} from './../renderers';

export default {
  editor: getEditor('text'),
  renderer: getRenderer('formula')
};
