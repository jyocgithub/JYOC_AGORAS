package xls_con_jxl_jar;

import jxl.Workbook;
import jxl.write.Label;
import jxl.write.Number;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;

import java.io.File;

public class EscribirExcel {
	
	public static void main(String[] args) {
		File fichero = new File("./ficheros/nueva.xls");
		
        try {
        	WritableWorkbook w = Workbook.createWorkbook(fichero);
        	
        	/*
        	Workbook wb = Workbook.getWorkbook(fichero);
        	WritableWorkbook w = Workbook.createWorkbook(fichero, wb);
        	*/
        	
        	//Creamos una hoja: parametros:    Nombre de la hoja y numero de la misma
        	WritableSheet sheet = w.createSheet("Datos", 0);
        	
        	//parametros son:                        columna fila contenido //CUIDADO PRIMERO LA COLUMNA
			// necesita import jxl.write.Number;
        	Number number = new Number(0, 0, 1);
        	sheet.addCell(number);

        	//parametros son:                        columna fila contenido //CUIDADO PRIMERO LA COLUMNA
        	Label cadena = new Label(1, 0, "valor");
            sheet.addCell(cadena);

			// metemos otra fila
        	number = new Number(0, 1, 2);
        	sheet.addCell(number);
        	cadena = new Label(1, 1, "otro valor");
            sheet.addCell(cadena);

            w.write();
            w.close();
        	
        } catch (Exception e) {
        	e.printStackTrace();
        }
	}
  
}
