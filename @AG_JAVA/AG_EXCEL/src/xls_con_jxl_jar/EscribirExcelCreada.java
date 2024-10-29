package xls_con_jxl_jar;

import jxl.Workbook;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;

import java.io.File;

public class EscribirExcelCreada {
	
	public static void main(String[] args) {
		File fichero = new File("./ficheros/nueva.xls");
		
        try {

			WritableWorkbook w = Workbook.createWorkbook(fichero);

//			Workbook wb = Workbook.getWorkbook(fichero);
//        	WritableWorkbook w = Workbook.createWorkbook(fichero, wb);

			//Obtenemos referencia de una hoja en modo edicion (WritableSheet)
			// parametros:            Numero de la hoja
			WritableSheet sheet = w.getSheet(0);

        	jxl.write.Number number = new jxl.write.Number(0, 6, 43);
        	sheet.addCell(number);
        	
        	jxl.write.Label cadena = new jxl.write.Label(1, 6, "otro");
            sheet.addCell(cadena);

            w.write();
            w.close();
        	
        }
        /*
        catch (FileNotFoundException e) {
        	System.out.println("ERROR: el archivo est� siendo utilizado por otro proceso");
        }
        */   
        catch (Exception e) {
        	e.printStackTrace();
        }
	}
  
}
