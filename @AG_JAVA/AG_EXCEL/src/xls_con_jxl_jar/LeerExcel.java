package xls_con_jxl_jar;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;

import java.io.File;

public class LeerExcel {

    public static final int XLS_COL_NOMBRE = 0;
    public static final int XLS_COL_DNI = 1;
    public static final String CADENA = "prueba";

    public static void main(String[] args) {
        File fichero = new File("./ficheros/nueva.xls");

        try {
            Workbook w = Workbook.getWorkbook(fichero);

            //Obtenemos referencia de una hoja en modo lectura (Sheet)
            Sheet sheet = w.getSheet(0);

            for (int f = 0; f < sheet.getRows(); f++) {
                String resultado = " ";
                for (int c = 0; c < sheet.getColumns(); c++) {
					// getCell obtiene una celda segun su columna y fila (un objeto Cell) //CUIDADO PRIMERO LA COLUMNA
                    Cell unacelda = sheet.getCell(c,f);
					// getContents obtiene el contenido de una celda (un String)
                    String contenido = unacelda.getContents() ;

					//todas las cosas juntas
					//contenido += sheet.getCell(c, f).getContents();

					resultado = resultado + contenido + "\t";
                }
                System.out.println(resultado);

            }

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

	static public String getContenido(Sheet hoja, int columna, int fila){
		return hoja.getCell(columna, fila).getContents() ;
	}

}
