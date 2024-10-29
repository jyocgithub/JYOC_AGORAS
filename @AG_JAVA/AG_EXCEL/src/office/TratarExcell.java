package office;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;

public class TratarExcell {

    public static void main(String[] args) throws Exception {
        new TratarExcell();
    }

    public TratarExcell() throws Exception {


        // Crear canal para leer el fichero excell, con su directorio
        FileInputStream fis = new FileInputStream(new File("ficheros/ejemplo.xlsx"));
        // Crear un workbook, que sera quien procese el interior del fichero xls
        XSSFWorkbook wb = new XSSFWorkbook(fis);
        // Sacamos la primera hoja del excell (le primera es la numero 0)
        XSSFSheet sheet = wb.getSheetAt(0);
        // Creamos un objeto del "evaluador de formulas", una máquina que descifra formulas
        FormulaEvaluator formulaEvaluator = wb.getCreationHelper().createFormulaEvaluator();
        for (Row row : sheet) {     // Recorremos un a una cada fila de la hoja
            for (Cell cell : row) {    // Recorremos una a una cada celda de la fila
                // vamos a actuar de modo distinto segun el tipo de celda que sea
                switch (formulaEvaluator.evaluateInCell(cell).getCellType()) {
                    case Cell.CELL_TYPE_NUMERIC:   //  Se trata de una celda numerica
                        // sacamos el valor de la celda como un numero
                        System.out.print(cell.getNumericCellValue() + "\t\t");
                        break;
                    case Cell.CELL_TYPE_STRING:    //  Se trata de una celda con texto
                        // sacamos el valor de la celda como un string
                        System.out.print(cell.getStringCellValue() + "\t\t");
                        break;
                }
            }
            System.out.println();
        }
    }
}



