package br.edu.uniso;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.SQLException;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Iterator;

public class IntegracaoExcel {


    public static void main(String args[]) {
        try {

            String nomeArquivo = "ListaTrabalhadores.xlsx";
            File file = new File(nomeArquivo);
            FileInputStream fis = new FileInputStream(file);
            Workbook workbook = null;

            //marotagem
            if (nomeArquivo.endsWith("xls")) {
                workbook = new HSSFWorkbook(fis);
            } else if (nomeArquivo.endsWith("xlsx")) {
                workbook = new XSSFWorkbook(fis);
            }

            Sheet planilha1 = workbook.getSheetAt(0);

            // linha a linha
            Iterator<Row> rows = planilha1.iterator();
            while (rows.hasNext()) {
                Row row = rows.next();

                // pegando as coluna (Cell)
                Iterator<Cell> celulas = row.cellIterator();


                    try {

                        Connection c = DriverManager.getConnection("jdbc:mysql://127.0.0.1:3306/sys",
                                "root", "ninja1234@");
                        PreparedStatement ps = c.prepareStatement("insert into vendedor(CPF,Nome_do_vendedor,idade) values " +
                                "(?,?,?)");

                        while (celulas.hasNext()) {

                            Cell coluna = celulas.next();
                            Cell coluna2 = celulas.next();
                            Cell coluna3 = celulas.next();

                            if (coluna.getCellType() == Cell.CELL_TYPE_NUMERIC)
                                System.out.print("\t" + coluna.getNumericCellValue());
                            ps.setDouble(1, coluna.getNumericCellValue());


                            if (coluna2.getCellType() == Cell.CELL_TYPE_NUMERIC)
                                System.out.print("\t" + coluna2.getStringCellValue());
                            ps.setString(2, coluna2.getStringCellValue());

                            if (coluna3.getCellType() == Cell.CELL_TYPE_NUMERIC)
                                System.out.print("\t" + coluna3.getNumericCellValue());
                            ps.setDouble(3, coluna3.getNumericCellValue());






                        //c.close();


                            ps.executeUpdate();

                            ps.close();
                            c.close();
                    }
                } catch (SQLException throwables) {
                        throwables.printStackTrace();
                    }


            }

        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }




    }



}




