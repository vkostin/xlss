package io.github.vkostin.xlss.extractors;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.DateUtil;

import java.text.DecimalFormat;

public class XlsxParser {
    private static final DecimalFormat df = new DecimalFormat("#.####################");
    public static String getCellValue(Cell cell) {
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
                    double value = cell.getNumericCellValue();
                    return df.format(value);
                }
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                // up to apache poi 5.1.0 a FormulaEvaluator is needed to evaluate the formulas while using DataFormatter
//                FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
                DataFormatter dataFormatter = new DataFormatter(new java.util.Locale("en", "US"));
                // from 5.2.0 on the DataFormatter can set to use cached values for formula cells
                dataFormatter.setUseCachedValuesForFormulaCells(true);
                return dataFormatter.formatCellValue(cell);
            case BLANK:
                return "";
            default:
                return "";
        }
    }
}
