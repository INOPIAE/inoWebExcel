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

# Change `manifest.xml`

TBD
