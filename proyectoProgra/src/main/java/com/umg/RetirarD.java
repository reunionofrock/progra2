/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/GUIForms/JFrame.java to edit this template
 */
package com.umg;

import java.awt.event.ActionEvent;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.text.SimpleDateFormat;
import java.util.Date;

import javax.swing.JOptionPane;

import org.apache.commons.math3.exception.ZeroException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author MAR ACEITUNO
 */
public class RetirarD extends javax.swing.JFrame {

    /**
     * Creates new form RetirarD
     */
    public RetirarD() {
        initComponents();
    }

    /**
     * This method is called from within the constructor to initialize the form.
     * WARNING: Do NOT modify this code. The content of this method is always
     * regenerated by the Form Editor.
     */
    @SuppressWarnings("unchecked")
    // <editor-fold defaultstate="collapsed" desc="Generated Code">//GEN-BEGIN:initComponents
    private void initComponents() {

        jPanel1 = new javax.swing.JPanel();
        jButton1 = new javax.swing.JButton();
        jButton2 = new javax.swing.JButton();
        jLabel1 = new javax.swing.JLabel();
        jLabel2 = new javax.swing.JLabel();
        NumCuentaField = new javax.swing.JTextField();
        jLabel3 = new javax.swing.JLabel();
        NombreeField = new javax.swing.JTextField();
        jLabel4 = new javax.swing.JLabel();
        MontoField = new javax.swing.JTextField();

        setDefaultCloseOperation(javax.swing.WindowConstants.EXIT_ON_CLOSE);

        jPanel1.setBackground(new java.awt.Color(0, 204, 204));
        jPanel1.setBorder(javax.swing.BorderFactory.createTitledBorder("RETIRAR DINERO"));

        jButton1.setBackground(new java.awt.Color(204, 204, 204));
        jButton1.setText("REGRESAR");
        jButton1.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jButton1ActionPerformed(evt);
            }
        });

        jButton2.setBackground(new java.awt.Color(204, 204, 204));
        jButton2.setText("ACEPTAR");

        jLabel1.setFont(new java.awt.Font("Segoe UI", 0, 18)); // NOI18N
        jLabel1.setText("RETIROS");

        jLabel2.setText("NO. CUENTA");

        jLabel3.setText("NOMBRE");

        jLabel4.setText("MONTO");

        javax.swing.GroupLayout jPanel1Layout = new javax.swing.GroupLayout(jPanel1);
        jPanel1.setLayout(jPanel1Layout);
        jPanel1Layout.setHorizontalGroup(
            jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel1Layout.createSequentialGroup()
                .addGap(377, 377, 377)
                .addComponent(jLabel1)
                .addContainerGap(400, Short.MAX_VALUE))
            .addGroup(jPanel1Layout.createSequentialGroup()
                .addGap(84, 84, 84)
                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addGroup(jPanel1Layout.createSequentialGroup()
                        .addComponent(jButton1)
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                        .addComponent(jButton2)
                        .addGap(151, 151, 151))
                    .addGroup(jPanel1Layout.createSequentialGroup()
                        .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.TRAILING)
                            .addComponent(jLabel3)
                            .addComponent(jLabel2)
                            .addComponent(jLabel4))
                        .addGap(18, 18, 18)
                        .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING, false)
                            .addComponent(NumCuentaField)
                            .addComponent(NombreeField)
                            .addComponent(MontoField, javax.swing.GroupLayout.DEFAULT_SIZE, 351, Short.MAX_VALUE))
                        .addContainerGap(javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))))
        );
        jPanel1Layout.setVerticalGroup(
            jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, jPanel1Layout.createSequentialGroup()
                .addComponent(jLabel1)
                .addGap(43, 43, 43)
                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(jLabel2)
                    .addComponent(NumCuentaField, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addGap(25, 25, 25)
                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(jLabel3)
                    .addComponent(NombreeField, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addGap(62, 62, 62)
                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(jLabel4)
                    .addComponent(MontoField, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED, 219, Short.MAX_VALUE)
                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(jButton1)
                    .addComponent(jButton2))
                .addGap(117, 117, 117))
        );

        javax.swing.GroupLayout layout = new javax.swing.GroupLayout(getContentPane());
        getContentPane().setLayout(layout);
        layout.setHorizontalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addComponent(jPanel1, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
        );
        layout.setVerticalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addComponent(jPanel1, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
        );
// Accion de retirar dinero se agrega como listener de evento
        jButton2.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                WithdrawFromAccount(evt);
            }

           
            }
        );
        pack();
    }// </editor-fold>//GEN-END:initComponents

    private void jButton1ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jButton1ActionPerformed
      Panel_control  pc = new Panel_control();
          this.setVisible(false);
          pc.setVisible(true);
          
          String NumCuenta = NumCuentaField.getText();
          String Nombree = NombreeField.getText();
          String Monto = MontoField.getText();
    }//GEN-LAST:event_jButton1ActionPerformed

    private void WithdrawFromAccount(ActionEvent evt) {
        String NodeCuenta = NumCuentaField.getText();
        //String Nombre = NombreeField.getText();
        String MontoRetirar = MontoField.getText();
        System.out.println(NodeCuenta);
        //System.out.println(Nombre);
        System.out.println(MontoRetirar);

        // metodo provisional para insertar en el archivo
        try {

            Workbook workbook;
            File file = new File("DatosCuentas.xlsx");
            if (file.exists() && file.length() > 0) {
                workbook = new XSSFWorkbook(new FileInputStream(file));
            } else {
                workbook = new XSSFWorkbook();
            }
            Sheet sheet = workbook.getSheet("Datos de cuentas");
            // Por cada fila, se obtiene la cuarta columna(indice 3), correspondiente al ID de la cuenta
            sheet.forEach((row) -> {
                Cell currentCell = row.getCell(4);
                CellType cellType = currentCell.getCellType();
                double NoCuenta = 0;
                // el switch ayuda a manejar errores por la primera fila de encabezados
                switch (cellType) {
                    case CellType.STRING -> {
                        try {
                            NoCuenta = Double.valueOf(currentCell.getStringCellValue());
                            System.out.println("Se encontraron los numeros de cuenta(string): " + currentCell.getStringCellValue());
                            if(Double.valueOf(NodeCuenta) == NoCuenta){
                                System.out.println("Se encontro coincidencia con una cuenta");
                                System.out.println("Indice de fila");
                                System.out.println(currentCell.getRowIndex());
                                Cell currentCellBalance = sheet.getRow(currentCell.getRowIndex()).getCell(3);
                                System.out.println(currentCellBalance.getAddress().formatAsString());
                                System.out.println(currentCellBalance.getStringCellValue());
                                if(Double.valueOf(MontoRetirar) > Double.valueOf(currentCellBalance.getStringCellValue())){
                                    throw new ZeroException();
                                }
                                double newBalance = Double.valueOf(currentCellBalance.getStringCellValue()) - Double.valueOf(MontoRetirar);
                                currentCellBalance.setAsActiveCell();
                                currentCellBalance.setBlank();
                                currentCellBalance.setCellValue(String.valueOf(newBalance));
                                System.out.println("Cuenta actualizada");
                                System.out.println(currentCellBalance.getStringCellValue());
                                // Se agrega un comentario para la bitacora de datos
                                SimpleDateFormat formatter = new SimpleDateFormat("dd/MM/yyyy HH:mm");
                                String strDate = formatter.format(new Date());
                               // String previousComments = row.getCell(7).getRichStringCellValue().toString();
                               // row.createCell(5).setCellValue(previousComments + ".Se hace retiro de cuenta por Q." + MontoRetirar + " en fecha: " + strDate);
                                FileOutputStream fileOut = new FileOutputStream("DatosCuentas.xlsx");
                                workbook.write(fileOut);
                                workbook.close();

                                File accountDetails = new File(String.format("%s.txt",NodeCuenta));
                                if (accountDetails.exists() && accountDetails.length() > 0) {
                                    file.createNewFile();
                                } else {
                                    accountDetails.createNewFile();
                                }
                                fileOut = new FileOutputStream(String.format("%s.txt",NodeCuenta), true);
                                fileOut.write((".Se hace retiro de cuenta por Q." + MontoRetirar + " en fecha: " + strDate).getBytes());
                                fileOut.close();
                                JOptionPane.showMessageDialog(null, "Archivo Excel creado exitosamente.");
                            }
                        } catch (ZeroException e) {
                            JOptionPane.showMessageDialog(null, "El monto a retirar es mayor al monto disponible");
                            System.out.println(e.getMessage());
                        } catch (Exception e) {
                            System.out.println("Continuando se encontro una fila que no puede ser convertida a numero");
                            System.out.println(e.getMessage());
                        }
                    }
                    case CellType.NUMERIC -> {
                        try {
                            NoCuenta = Double.valueOf(currentCell.getStringCellValue());
                            System.out.println("Se encontraron los numeros de cuenta(string): " + currentCell.getStringCellValue());
                            if(Double.valueOf(NodeCuenta) == NoCuenta){
                                System.out.println("Se encontro coincidencia con una cuenta");
                                System.out.println("Indice de fila");
                                System.out.println(currentCell.getRowIndex());
                                Cell currentCellBalance = sheet.getRow(currentCell.getRowIndex()).getCell(2);
                                System.out.println(currentCellBalance.getAddress().formatAsString());
                                System.out.println(currentCellBalance.getStringCellValue());
                                if(Double.valueOf(MontoRetirar) > Double.valueOf(currentCellBalance.getStringCellValue())){
                                    throw new ZeroException();
                                }
                                double newBalance = Double.valueOf(currentCellBalance.getStringCellValue()) - Double.valueOf(MontoRetirar);
                                currentCellBalance.setAsActiveCell();
                                currentCellBalance.setBlank();
                                currentCellBalance.setCellValue(newBalance);
                                System.out.println("Cuenta actualizada");
                                System.out.println(currentCellBalance.getStringCellValue());
                                SimpleDateFormat formatter = new SimpleDateFormat("dd/MM/yyyy HH:mm");
                                String strDate = formatter.format(new Date());
                                String previousComments = row.getCell(5).getRichStringCellValue().toString();
                                row.createCell(5).setCellValue(previousComments + ".Se hace retiro de cuenta por Q." + MontoRetirar + " en fecha: " + strDate);
                                FileOutputStream fileOut = new FileOutputStream("DatosCuentas.xlsx");
                                workbook.write(fileOut);
                                workbook.close();
                                JOptionPane.showMessageDialog(null, "Archivo Excel creado exitosamente.");
                            }
                        } 
                        catch (ZeroException e) {
                            JOptionPane.showMessageDialog(null, "El monto a retirar es mayor al monto disponible");
                            System.out.println(e.getMessage());
                        }catch (Exception e) {
                            System.out.println("Continuando se encontro una fila que no puede ser convertida a string");
                            System.out.println(e.getMessage());
                            e.printStackTrace();
                        }
                    }
                    default -> throw new NumberFormatException();
                }
               
            }
            );
        } catch (Exception e) {
            System.out.println("Error en la creacion o lectura del archivo:");
            System.out.println(e.getMessage());
        }
    }
    /**
     * @param args the command line arguments
     */
    public static void main(String args[]) {
        /* Set the Nimbus look and feel */
        //<editor-fold defaultstate="collapsed" desc=" Look and feel setting code (optional) ">
        /* If Nimbus (introduced in Java SE 6) is not available, stay with the default look and feel.
         * For details see http://download.oracle.com/javase/tutorial/uiswing/lookandfeel/plaf.html 
         */
        try {
            for (javax.swing.UIManager.LookAndFeelInfo info : javax.swing.UIManager.getInstalledLookAndFeels()) {
                if ("Nimbus".equals(info.getName())) {
                    javax.swing.UIManager.setLookAndFeel(info.getClassName());
                    break;
                }
            }
        } catch (ClassNotFoundException ex) {
            java.util.logging.Logger.getLogger(RetirarD.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (InstantiationException ex) {
            java.util.logging.Logger.getLogger(RetirarD.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (IllegalAccessException ex) {
            java.util.logging.Logger.getLogger(RetirarD.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (javax.swing.UnsupportedLookAndFeelException ex) {
            java.util.logging.Logger.getLogger(RetirarD.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        }
        //</editor-fold>

        /* Create and display the form */
        java.awt.EventQueue.invokeLater(new Runnable() {
            public void run() {
                new RetirarD().setVisible(true);
            }
        });
    }

    // Variables declaration - do not modify//GEN-BEGIN:variables
    private javax.swing.JTextField MontoField;
    private javax.swing.JTextField NombreeField;
    private javax.swing.JTextField NumCuentaField;
    private javax.swing.JButton jButton1;
    private javax.swing.JButton jButton2;
    private javax.swing.JLabel jLabel1;
    private javax.swing.JLabel jLabel2;
    private javax.swing.JLabel jLabel3;
    private javax.swing.JLabel jLabel4;
    private javax.swing.JPanel jPanel1;
    // End of variables declaration//GEN-END:variables
}
