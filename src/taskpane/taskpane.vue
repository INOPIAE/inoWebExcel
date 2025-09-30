<template>
    <div>
      <p v-if="!officeReady">{{ $t("msgOfficeInit") }}</p>
      <v-btn :disabled="!officeReady" @click="setFocusA1">{{ $t("btnSetFocus") }}</v-btn>
      <v-btn :disabled="!officeReady" @click="selectLine">{{ $t("btnSelectLine") }}</v-btn>
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

  function selectLine() {

    if (!officeReady.value) {
      console.error("Office is not ready yet.")
      return
    }

    Excel.run(async (context) => {
      const workbook = context.workbook;
      const selectedRange = workbook.getSelectedRange();
      selectedRange.load("rowIndex");
      await context.sync();

      const rowIndex = selectedRange.rowIndex; 

      const activeSheet = workbook.worksheets.getActiveWorksheet();
      activeSheet.load("name");
      await context.sync();

      const sheets = workbook.worksheets;
      sheets.load("items/name");
      await context.sync();

      for (let sheet of sheets.items) {
          sheet.activate();

          const sheetRange = sheet.getRangeByIndexes(rowIndex, 0, 1, 100);
          sheetRange.getCell(rowIndex, 5).select();
          await context.sync();

          const sheetRange1 = sheet.getRangeByIndexes(rowIndex, 0, 1, 100);
          sheetRange1.select();
          await context.sync();
      }

      const originalSheet = workbook.worksheets.getItem(activeSheet.name);
      originalSheet.activate();
      await context.sync();
    }).catch((error) => {
      console.error("Excel.run Fehler:", error);
    });
  }
  </script>
  