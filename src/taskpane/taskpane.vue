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
  