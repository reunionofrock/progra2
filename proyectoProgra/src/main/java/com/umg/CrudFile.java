/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.umg;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import javax.swing.JOptionPane;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.JSONObject;

/**
 *
 * @author Dario
 */
public class CrudFile {

    public CrudFile() {

    }

    public JSONObject searchInColumnByHeader(String sheetName, String fileName, String headerValue,
            String searchValue) {
        JSONObject json = new JSONObject();
        try {
            File file = new File(fileName);
            if (file.exists() && file.length() > 0) {
                try (FileInputStream fis = new FileInputStream(file);
                        Workbook workbook = new XSSFWorkbook(fis)) {
                    Sheet sheet = workbook.getSheet(sheetName);
                    if (sheet != null) {
                        Row headerRow = sheet.getRow(0); // Obtener la fila de encabezado
                        int columnIndex = -1;
                        if (headerRow != null) {
                            for (Cell cell : headerRow) {
                                if (cell.getStringCellValue().equals(headerValue)) {
                                    columnIndex = cell.getColumnIndex(); // Obtener el índice de la columna
                                    break;
                                }
                            }
                        }
                        if (columnIndex != -1) {
                            for (Row row : sheet) {
                                Cell cell = row.getCell(columnIndex);
                                if (cell != null && cell.getCellType() == CellType.STRING
                                        && cell.getStringCellValue().equals(searchValue)) {
                                    // Valor encontrado, devolver toda la fila como JSONObject
                                    Row headerRowJson = sheet.getRow(0); // Obtener la fila de encabezado
                                    for (Cell cellInRow : row) {
                                        String header = headerRowJson.getCell(cellInRow.getColumnIndex())
                                                .getStringCellValue();
                                        String value = cellInRow.toString();
                                        json.put(header, value);
                                    }
                                    return json;
                                }
                            }
                        }
                    }
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        } catch (Exception ex) {
            JOptionPane.showMessageDialog(null, "Error: " + ex.getMessage());
        }
        return json; // Valor no encontrado
    }

    @SuppressWarnings("resource")
    public boolean CreateFile(JSONObject obj, String sheetName, String fileName) {
        boolean result = false;
        try {
            Workbook workbook;
            File file = new File(fileName);
            if (file.exists() && file.length() > 0) {
                workbook = new XSSFWorkbook(new FileInputStream(file));
            } else {
                workbook = new XSSFWorkbook();
            }
            Sheet sheet = workbook.getSheet(sheetName);
            int id = 0; // Inicializar el ID
            if (sheet == null) {
                sheet = workbook.createSheet(sheetName);
                // Crear fila de encabezado
                Row headerRow = sheet.createRow(0);
                int cellIndex = 0;

                // Crear la celda para "ID"
                Cell idCell = headerRow.createCell(cellIndex++);
                idCell.setCellValue("ID");

                // Iterar sobre las claves del objeto JSON
                for (String key : obj.keySet()) {
                    Cell cell = headerRow.createCell(cellIndex++);
                    cell.setCellValue(key);
                }
            } else {
                // Obtener el último ID
                int lastRowNum = sheet.getLastRowNum();
                Row lastRow = sheet.getRow(lastRowNum);
                if (lastRow != null) {
                    Cell idCell = lastRow.getCell(0);
                    if (idCell != null) {
                        id = (int) idCell.getNumericCellValue();
                    }
                }
                id++;
            }

            int lastRowNum = sheet.getLastRowNum();
            Row dataRow = sheet.createRow(lastRowNum + 1);
            int cellIndex = 0;

            // Crear la celda para el ID y establecer su valor
            Cell idCell = dataRow.createCell(cellIndex++);
            idCell.setCellValue(id);

            // Iterar sobre las claves del objeto JSON
            for (String key : obj.keySet()) {
                Cell cell = dataRow.createCell(cellIndex++);
                cell.setCellValue(obj.get(key).toString());
            }
            // Iterar sobre las columnas para redimensionarlas
            for (int i = 0; i < cellIndex; i++) {
                sheet.autoSizeColumn(i);
            }
            // Escribir el archivo Excel
            try (FileOutputStream fileOut = new FileOutputStream(fileName)) {
                workbook.write(fileOut);
                result = true;
            }
        } catch (IOException ex) {
            JOptionPane.showMessageDialog(null, "Error: " + ex.getMessage());
        }
        return result;
    }

    public Row findRow(Sheet sheet, String cellContent) {
        for (Row row : sheet) {
            Cell cell = row.getCell(7); // Asume que quieres buscar en la primera columna
            if (cell.getCellType() == CellType.STRING && cell.getRichStringCellValue().getString().trim().equals(cellContent)) {
                return row;
            }
        }
        return null;
    }

    public boolean updateRow(String fileName, String sheetName, String searchValue, JSONObject newData) {
        boolean result = false;
        try {
            FileInputStream fis = new FileInputStream(new File(fileName));
            Workbook workbook = new XSSFWorkbook(fis);
            Sheet sheet = workbook.getSheet(sheetName);
            Row row = findRow(sheet, searchValue);
            if (row != null) {
                for (String key : newData.keySet()) {
                    Cell cell = row.createCell(getColumnIndex(sheet, key)); // Asume que tienes un método getColumnIndex que obtiene el índice de una columna por su encabezado
                    cell.setCellValue(newData.getString(key));
                    result = true;
                }
                fis.close();
                FileOutputStream fos = new FileOutputStream(new File(fileName));
                workbook.write(fos);
                fos.close();
            }
            workbook.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
        return result;
    }

    public int getColumnIndex(Sheet sheet, String header) {
        Row headerRow = sheet.getRow(0); // Asume que la primera fila contiene los encabezados
        for (Cell cell : headerRow) {
            if (cell.getCellType() == CellType.STRING && cell.getStringCellValue().trim().equals(header)) {
                return cell.getColumnIndex();
            }
        }
        return -1; // Devuelve -1 si no se encuentra el encabezado
    }

}
