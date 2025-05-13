package io.github.vkostin.xlss.parser;

import org.apache.poi.ss.usermodel.*;

import java.util.*;

public class ParserHelpers {
    public static List<Map<String, String>> parseToStrings(Sheet sheet, int headerRowY, List<String> headersTitles) {
        List<Map<String, String>> rows = new ArrayList<>();
        Row headerRow = sheet.getRow(headerRowY);

        if (headerRow == null) {
            throw new IllegalArgumentException("Sheet is empty or does not contain a header row.");
        }

        //      List<String> headers = new ArrayList<>();
        Map<String, Integer> headersIdxMap = new HashMap<>();
        Set<String> headerToAdd = new HashSet<>(headersTitles);
        for (Cell cell : headerRow) {
            String cellValue = getCellValue(cell);
            if (headerToAdd.contains(cellValue)) {
                headersIdxMap.put(cellValue, cell.getColumnIndex());
                headerToAdd.remove(cellValue);
            }
        }
        if (!headerToAdd.isEmpty()) {
            throw new IllegalArgumentException("Cannot find required headers: " + headerToAdd);
        }

        for (int i = headerRowY + 1; i <= sheet.getLastRowNum(); i++) { // Start from the second row
            Row row = sheet.getRow(i);
            if (row == null) continue;

            Map<String, String> rowData = new HashMap<>();
//            for (int j = 0; j < headers.size(); j++) {
//                Cell cell = row.getCell(j);
//                String cellValue = getCellValue(cell);
//                rowData.put(headers.get(j), cellValue);
//            }

            try {
                for (String h : headersIdxMap.keySet()) {
                    Integer idx = headersIdxMap.get(h);
                    if (idx == null)
                        throw new RuntimeException("Cannot parse file, required column doesnt exist");
                    Cell cell = row.getCell(idx);
                    rowData.put(h, getCellValue(cell));
                }
            } catch (Exception e) {
                System.out.println(e.getMessage());
            }
            if (rowData.keySet().size() == headersIdxMap.size()) {
                rows.add(rowData);
            }
        }

        return rows;
    }

    private static String getCellValue(Cell cell) {
        if (cell == null) {
            return "";
        }
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return cell.getDateCellValue().toString();
                } else {
                    return String.valueOf(cell.getNumericCellValue());
                }
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                // up to apache poi 5.1.0 a FormulaEvaluator is needed to evaluate the formulas while using DataFormatter
//                FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
                DataFormatter dataFormatter = new DataFormatter(new java.util.Locale("en", "US"));
                // from 5.2.0 on the DataFormatter can set to use cached values for formula cells
                dataFormatter.setUseCachedValuesForFormulaCells(true);
                //String value = dataFormatter.formatCellValue(cell, evaluator); // up to apache poi 5.1.0
                //                return cell.getCellFormula();
                return dataFormatter.formatCellValue(cell);
            case BLANK:
                return "";
            default:
                return "";
        }
    }
}
