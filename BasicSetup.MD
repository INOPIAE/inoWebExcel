# Basic setup to set up an Office Web Add in with Vue3

## Baisc installation with Office Add-In Deployment kit
1. Install basis with Office Add-In Deployment kit
1. Run first debug to ensure proper installation

## Update to Vue3

1. Install Vue3
    ```shell
    npm install vue
    ```

1. Add `src/taskpane.vue`

    ```vue
    <template>
        <div>
        <h1>inoWebExcel</h1>
        <button :disabled="!officeReady" @click="run">Run</button>
        <p v-if="!officeReady">Office wird initialisiert...</p>
        </div>
    </template>
    
    <script setup>
    import { ref, onMounted } from 'vue'
    
    const officeReady = ref(false)
    
    onMounted(() => {
        if (window.Office) {
            Office.onReady().then(() => {
            officeReady.value = true
            })
        } else {
            console.error("❌ Office.js wurde nicht geladen.")
        }
    })
    
    function run() {
        if (!officeReady.value) {
        console.error("Office is not ready yet.")
        return
        }
    
        Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet()
        sheet.getRange("A1").values = [["Hello from Vue!"]]
        await context.sync()
        }).catch((error) => {
        console.error("Fehler beim Ausführen von Excel.run:", error)
        })
    }
    </script>
    ```

1. Replace `src/taskpane.ts`

    ```ts
    import { createApp } from 'vue';
    import Taskpane from './taskpane.vue';

    createApp(Taskpane).mount('#app');
    ```


1. Raplace `src/taspane.html`
    ```html
    <html>
    <head>
        <meta charset="utf-8" />
        <title>inoWebExcel</title>
        <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>
    </head>
    <body>
        <div id="app"></div>
    </body>
    </html>
    ```

1. Configurate Webpack
    1. Install vue-loader und vue-style-loader
        ```shell
        npm install -D vue-loader@next vue-style-loader @vue/compiler-sfc
        ```
    1. Addin `webpack.config.js`
        ```js
        const { VueLoaderPlugin } = require('vue-loader');

        module.exports = {
        // ...
        module: {
            rules: [
            {
                test: /\.vue$/,
                use: 'vue-loader'
            },
            // ggf. weitere Loader
            ]
        },
        plugins: [
            new VueLoaderPlugin(),
            // ...
        ]
        };
        ```

# Add translation

1. Install vue-i18n@9
    ```shell
    npm install vue-i18n@9
    ```

1. Create language file in folder `src/i18n` e.g. `src/i18n/en.json`
    ```json
    {
        "title": "My Excel Add-in",
        "button": "Run"
    }
    ```

1. Add setup file `src/i18n/i18n.ts`
    ```ts
    import { createI18n } from 'vue-i18n'
    import de from './de.json'
    import en from './en.json'

    export const i18n = createI18n({
    legacy: false,
    locale: 'de',
    fallbackLocale: 'en',
    messages: {
        de,
        en
    }
    })

    ```

1. Register in `taskpane.ts`
    ```ts
    import { createApp } from 'vue';
    import Taskpane from './taskpane.vue';
    import { i18n } from '../i18n/i18n'

    createApp(Taskpane)
        .use(i18n)
        .mount('#app');
    ```

1. Use in template and onMounted() in `taskpane.vue`
    ```vue
    <template>
    <div>
        <h1>{{ $t("title") }}</h1>
        <button @click="run">{{ $t("button") }}</button>
    </div>
    </template>
    <script setup>
        import { ref, onMounted } from 'vue'
        
        const officeReady = ref(false)
        
        const officeLanguage = ref('')

        onMounted(() => {
            if (window.Office) {
            Office.onReady().then(() => {
                officeReady.value = true

                officeLanguage.value = Office.context.displayLanguage || 'unknown'
                console.log("📘 Office-Sprache:", officeLanguage.value)
            })
            } else {
            console.error("❌ Office.js wurde nicht geladen.")
            }
        ...
    </script>
})
    ```

1. Adjust  `webpack.config.js`
    ```js
    module.exports = {
        // ...
        module: {
            rules: [
            {
                test: /\.json$/,
                type: 'json'
            }
            ]
        },
        resolve: {
            extensions: ['.ts', '.js', '.json']
        }
    }
    ```

1. Adjust  `tsconfig.json`
    ```json
    {
    "compilerOptions": {
        "resolveJsonModule": true,
        "esModuleInterop": true
    }
    }
    ```
# Change `manifest.xml`

TBD
