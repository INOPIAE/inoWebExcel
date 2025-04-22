import { createApp } from 'vue';
import Taskpane from './taskpane.vue';
import { i18n } from '../i18n/i18n'

import 'vuetify/styles'
import { createVuetify } from 'vuetify'
import * as components from 'vuetify/components'
import * as directives from 'vuetify/directives'

import '@mdi/font/css/materialdesignicons.css'

const vuetify = createVuetify({
    components,
    directives,
  })

createApp(Taskpane)
    .use(i18n)
    .use(vuetify)
    .mount('#app');
