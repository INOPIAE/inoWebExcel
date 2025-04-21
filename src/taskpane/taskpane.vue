<template>
    <div>
      <p v-if="!officeReady">Office wird initialisiert...</p>
      <button :disabled="!officeReady" @click="setFocusA1">Set focus to A1</button>
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
  
  function setFocusA1() {

    if (!officeReady.value) {
      console.error("Office is not ready yet.")
      return
    }

    Excel.run(async (context) => {
      const workbook = context.workbook;

      const activeSheet = workbook.worksheets.getActiveWorksheet();
      activeSheet.load("name");

      const allSheets = workbook.worksheets;
      allSheets.load("items/name");

      await context.sync();

      for (const sheet of allSheets.items) {
        const ws = workbook.worksheets.getItem(sheet.name);
        ws.activate();
        ws.getRange("A1").select();
      }

      await context.sync();

      workbook.worksheets.getItem(activeSheet.name).activate();

      await context.sync();
    }).catch((error) => {
      console.error("Excel.run Fehler:", error);
    });
  }
  </script>
  