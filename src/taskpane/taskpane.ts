import { createApp } from 'vue';
import Taskpane from './taskpane.vue';
import { i18n } from '../i18n/i18n'

createApp(Taskpane)
    .use(i18n)
    .mount('#app');
