/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package com.umg;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.FileOutputStream;
import java.io.IOException;

public class Registro_cliente extends JFrame {
    private JTextField NombreCliente, IDnum, Apellido, Adress, Phone, Mail, Occupation, Income;
    private JButton saveButton;

    public Registro_cliente() {
//         Configuración de la ventana principal
        setTitle("Ingresar Datos del Usuario");
        setSize(400, 400);
        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        setLayout(new GridLayout(9, 2));

        // Crear y añadir los componentes
        add(new JLabel("ID:"));
        NombreCliente = new JTextField();
        add(NombreCliente);

        add(new JLabel("Nombre:"));
        IDnum = new JTextField();
        add(IDnum);

        add(new JLabel("Apellido:"));
        Apellido = new JTextField();
        add(Apellido);

        add(new JLabel("Dirección:"));
        Adress = new JTextField();
        add(Adress);

        add(new JLabel("Teléfono:"));
        Phone = new JTextField();
        add(Phone);

        add(new JLabel("Correo Electrónico:"));
        Mail = new JTextField();
        add(Mail);

        add(new JLabel("Ocupación:"));
        Occupation = new JTextField();
        add(Occupation);

        add(new JLabel("Ingresos Mensuales:"));
        Income = new JTextField();
        add(Income);

        saveButton = new JButton("Guardar en Excel");
        saveButton.addActionListener(new SaveButtonListener());
        add(saveButton);

        setVisible(true);
    }

    Registro_cliente(String nombreCliente) {
        
    }

    void Registro_Cliente(JTextField NombreCliente) {
        throw new UnsupportedOperationException("Not supported yet."); // Generated from nbfs://nbhost/SystemFileSystem/Templates/Classes/Code/GeneratedMethodBody
    }

    private class SaveButtonListener implements ActionListener {
//        @Override
        public void actionPerformed(ActionEvent e) {
            String nombre = NombreCliente.getText();
            String id = IDnum.getText();
            String apellido = Apellido.getText();
            String direccion = Adress.getText();
            String telefono = Phone.getText();
            String correo = Mail.getText();
            String ocupacion = Occupation.getText();
            String ingresos = Income.getText();

            try (Workbook workbook = new XSSFWorkbook()) {
                Sheet sheet = workbook.createSheet("Datos de Usuario");

                // Crear fila de encabezado
                String[] columnHeaders = {"ID", "Nombre", "Apellido", "Dirección", "Teléfono", "Correo Electrónico", "Ocupación", "Ingresos Mensuales"};
                Row headerRow = sheet.createRow(0);
                for (int i = 0; i < columnHeaders.length; i++) {
                    Cell cell = headerRow.createCell(i);
                    cell.setCellValue(columnHeaders[i]);
                }

                // Crear fila de datos
                Row dataRow = sheet.createRow(1);
                dataRow.createCell(0).setCellValue(id);
                dataRow.createCell(1).setCellValue(nombre);
                dataRow.createCell(2).setCellValue(apellido);
                dataRow.createCell(3).setCellValue(direccion);
                dataRow.createCell(4).setCellValue(telefono);
                dataRow.createCell(5).setCellValue(correo);
                dataRow.createCell(6).setCellValue(ocupacion);
                dataRow.createCell(7).setCellValue(ingresos);

                // Redimensionar columnas
                for (int i = 0; i < columnHeaders.length; i++) {
                    sheet.autoSizeColumn(i);
                }

                // Escribir el archivo Excel
                try (FileOutputStream fileOut = new FileOutputStream("DatosUsuario.xlsx")) {
                    workbook.write(fileOut);
                    JOptionPane.showMessageDialog(null, "Archivo Excel creado exitosamente.");
                } catch (IOException ex) {
                    JOptionPane.showMessageDialog(null, "Error al escribir el archivo: " + ex.getMessage());
                }
            } catch (IOException ex) {
                JOptionPane.showMessageDialog(null, "Error al crear el workbook: " + ex.getMessage());
            }
        }
    }

    
    }

