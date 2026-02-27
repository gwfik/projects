function main(workbook: ExcelScript.Workbook) {
    let arkusz = workbook.getActiveWorksheet();
    let zakres = arkusz.getUsedRange();
    let ileWierszy = zakres.getRowCount();

    let zlacz1 = arkusz.getUsedRange().getLastColumn().getOffsetRange(0, 1);
    zlacz1.getCell(0, 0).setValue("złącz");
    let zakresFormuly = zlacz1.getOffsetRange(1, 0).getResizedRange(ileWierszy - 2, 0);
    zakresFormuly.setFormula("=A2&\";\"&D2");

    let pak = arkusz.getUsedRange().getLastColumn().getOffsetRange(0, 1);
    pak.getCell(0, 0).setValue("pak");
    let zakresFormuly2 = pak.getOffsetRange(1, 0).getResizedRange(ileWierszy - 2, 0);
    zakresFormuly2.setFormula("=E2*F2");

    let zakresDoTabeli = arkusz.getUsedRange();



    let staryArkusz = workbook.getWorksheet("Podsumowanie");
    if (staryArkusz) { staryArkusz.delete(); }

    let nowyArkusz = workbook.addWorksheet("Podsumowanie");
    let tabela = workbook.addPivotTable("MojaTabela", zakresDoTabeli, nowyArkusz.getRange("A1"));


    let uklad = tabela.getLayout();
    uklad.setLayoutType(ExcelScript.PivotLayoutType.tabular); 
    uklad.setSubtotalLocation(ExcelScript.SubtotalLocationType.off); 
    uklad.setShowColumnGrandTotals(false); 
    uklad.setShowRowGrandTotals(false); 
    uklad.repeatAllItemLabels(true); 


    tabela.addRowHierarchy(tabela.getHierarchy("Model"));
    tabela.addRowHierarchy(tabela.getHierarchy("Year"));
    tabela.addRowHierarchy(tabela.getHierarchy("Month"));
    tabela.addDataHierarchy(tabela.getHierarchy("pak"));


    let zakresTabeli = nowyArkusz.getUsedRange();
    let wartosci = zakresTabeli.getValues();
    tabela.delete();
    nowyArkusz.getRange("A1").getResizedRange(wartosci.length - 1, wartosci[0].length - 1).setValues(wartosci);



    let zakresPodsumowanie = nowyArkusz.getUsedRange();
    let ileWierszyPodsumowanie = zakresPodsumowanie.getRowCount();


    nowyArkusz.getRange("A:A").insert(ExcelScript.InsertShiftDirection.right);
    nowyArkusz.getRange("A1").setValue("złącz");
    let zlaczFormula = nowyArkusz.getRange("A2").getResizedRange(ileWierszyPodsumowanie - 2, 0);
    zlaczFormula.setFormulaLocal('=ZŁĄCZ.TEKST(C2;";";B2)');

    
    nowyArkusz.getRange("C:C").insert(ExcelScript.InsertShiftDirection.right);
    nowyArkusz.getRange("C1").setValue("region"); // Poprawiłem z "złącz" na "region"
    let regionFormula = nowyArkusz.getRange("C2").getResizedRange(ileWierszyPodsumowanie - 2, 0);
    regionFormula.setFormulaLocal('=X.WYSZUKAJ(A2;import!L:L;import!C:C; 0)');

    nowyArkusz.getUsedRange().getFormat().autofitColumns();
}