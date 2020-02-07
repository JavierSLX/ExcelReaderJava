/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package excelreader;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.text.DecimalFormat;
import java.text.NumberFormat;
import java.text.SimpleDateFormat;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author JavierSL
 */
public class ExcelReader 
{

    /**
     * @param args the command line arguments
     */
    public static void main(String[] args)
    {
        NumberFormat nf = new DecimalFormat("##########");
        
        try
        {
            FileInputStream fis = new FileInputStream(new File("C:\\Portabilidad\\Example.xlsx"));
            
            //Crea una instancia del Libro del excel
            XSSFWorkbook wb = new XSSFWorkbook(fis);
            
            //Crea una instancia de la hoja del libro
            XSSFSheet sheet = wb.getSheetAt(0);
            
            //Para evaluar el tipo de celda
            FormulaEvaluator formulaEvaluator = wb.getCreationHelper().createFormulaEvaluator();
            
            //Toma cada de la celda de la hoja
            for(Row row: sheet)
            {
                for(Cell cell: row)
                {
                    switch(formulaEvaluator.evaluateInCell(cell).getCellType())
                    {
                        //Si el valor de la celda es numerico
                        case Cell.CELL_TYPE_NUMERIC:
                            //Verifica si es una fecha
                            if(HSSFDateUtil.isCellDateFormatted(cell))
                            {
                                //Le da formato a la fecha
                                SimpleDateFormat dateFormat = new SimpleDateFormat("dd/MM/yyyy");
                                System.out.print(dateFormat.format(cell.getDateCellValue()) + "\t\t");
                            }
                            //Si no lo es, es un numero
                            else
                                System.out.print(nf.format(cell.getNumericCellValue()) + "\t\t");
                            break;
                        
                        // Si el valor de la celda es texto
                        case Cell.CELL_TYPE_STRING:
                            System.out.print(cell.getStringCellValue()+ "\t\t");
                            break;
                    }
                }
                
                System.out.println();
            }
            
        } catch (FileNotFoundException ex)
        {
            Logger.getLogger(ExcelReader.class.getName()).log(Level.SEVERE, null, ex);
        } catch (IOException ex)
        {
            Logger.getLogger(ExcelReader.class.getName()).log(Level.SEVERE, null, ex);
        }
        
    }
    
}
