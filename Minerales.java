package Formulario;

import Modelo.ConexionMySQL;
import com.mysql.jdbc.Connection;
import java.awt.Desktop;
import java.awt.Dimension;
import java.awt.Image;
import java.awt.Toolkit;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.swing.JFileChooser;
import javax.swing.JOptionPane;
import javax.swing.JTable;
import javax.swing.filechooser.FileNameExtensionFilter;
import javax.swing.table.DefaultTableModel;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Picture;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author Alana Mina
 */
public class Minerales extends javax.swing.JFrame {
    DefaultTableModel modelo;
    clsExportarExcel obj;
    String accion;
    
    public static void main(String args[]) throws IOException, SQLException {
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
            java.util.logging.Logger.getLogger(Minerales.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (InstantiationException ex) {
            java.util.logging.Logger.getLogger(Minerales.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (IllegalAccessException ex) {
            java.util.logging.Logger.getLogger(Minerales.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (javax.swing.UnsupportedLookAndFeelException ex) {
            java.util.logging.Logger.getLogger(Minerales.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        }
        //</editor-fold>

        /* Create and display the form */
        java.awt.EventQueue.invokeLater(new Runnable() {
            public void run() {
                new Minerales().setVisible(true);
            }
        });
        
    }
    
    public Minerales() {
        initComponents();
        CargarTablaMinerales("");
        inhabilitar();
        Toolkit mipantalla=Toolkit.getDefaultToolkit();
        Dimension tamanoPantalla=mipantalla.getScreenSize();
        int alturaPantalla=tamanoPantalla.height;
        int anchoPantalla=tamanoPantalla.width;
        setLocation(anchoPantalla/15, 0);
        setTitle("Museo de Mineralogía Alfred W. Stelzner");
        Image miIcono=mipantalla.getImage("principal.jpg");
        setIconImage(miIcono);
        txtObs.setLineWrap(true);
        txtObs.setWrapStyleWord(true);
        txtDes.setLineWrap(true);
        txtDes.setWrapStyleWord(true);
        setDefaultCloseOperation(DO_NOTHING_ON_CLOSE);
    }

    void CargarTablaMinerales(String valor) {
        String[] titulos = {"ID", "Código", "Especie", "Procedencia", "Descripción", "Colección", "Fecha de Ingreso", "Ubicación", "Observaciones"};
        String[] registro = new String[9]; //[] indican un vector
        String sSQL="";
        modelo = new DefaultTableModel(null, titulos);
        
        ConexionMySQL mysql = new ConexionMySQL();
        java.sql.Connection cn = mysql.getConexion();
        
        sSQL = "SELECT id, codigo, especie, procedencia, descripcion, coleccion, ano, ubicacion, observaciones FROM minerales " +
                "WHERE CONCAT(codigo, ' ', especie, ' ', procedencia, ' ', coleccion, ' ', ubicacion, ' ', observaciones) LIKE '%"+valor+"%'";
        
        try {
            Statement st = cn.createStatement();
            ResultSet rs = st.executeQuery(sSQL);
            
            while (rs.next()) {
                registro[0] = rs.getString("id");
                registro[1] = rs.getString("codigo");
                registro[2] = rs.getString("especie");
                registro[3] = rs.getString("procedencia");
                registro[4] = rs.getString("descripcion");
                registro[5] = rs.getString("coleccion");
                registro[6] = rs.getString("ano");
                registro[7] = rs.getString("ubicacion");
                registro[8] = rs.getString("observaciones");
                modelo.addRow(registro);
            }
            tblMineral.setModel(modelo);
        } catch (SQLException ex) {
            JOptionPane.showMessageDialog(null, ex);
        }
    }
    
    void habilitar() {
        txtCod.setEnabled(true);
        txtCol.setEnabled(true);
        txtAno.setEnabled(true);
        txtDes.setEnabled(true);
        txtObs.setEnabled(true);
        txtPro.setEnabled(true);
        txtSpc.setEnabled(true);
        txtUbi.setEnabled(true);
        txtCod.setText("");
        txtCol.setText("");
        txtAno.setText("");
        txtDes.setText("");
        txtObs.setText("");
        txtPro.setText("");
        txtSpc.setText("");
        txtUbi.setText("");
        btnGuardar.setEnabled(true);
        btnCancelar.setEnabled(true);
        txtSpc.requestFocus();
    }
    
    void inhabilitar() {
        txtCod.setEnabled(false);
        txtCol.setEnabled(false);
        txtAno.setEnabled(false);
        txtDes.setEnabled(false);
        txtObs.setEnabled(false);
        txtPro.setEnabled(false);
        txtSpc.setEnabled(false);
        txtUbi.setEnabled(false);
        txtCod.setText("");
        txtCol.setText("");
        txtAno.setText("");
        txtDes.setText("");
        txtObs.setText("");
        txtPro.setText("");
        txtSpc.setText("");
        txtUbi.setText("");
        btnGuardar.setEnabled(false);
        btnCancelar.setEnabled(false);
    }
    
    public static void generar() {
        Workbook book = new XSSFWorkbook();
        Sheet sheet = book.createSheet("Minerales");
        
        InputStream is;
        try {
            is = new FileInputStream("logo.jpg");
            byte[] bytes = IOUtils.toByteArray(is);
            int imgIndex = book.addPicture(bytes, Workbook.PICTURE_TYPE_JPEG);
            is.close();
            
            CreationHelper help = book.getCreationHelper();
            Drawing draw = sheet.createDrawingPatriarch();
            
            ClientAnchor anchor = help.createClientAnchor();
            anchor.setCol1(0);
            anchor.setRow1(0);
            Picture pict = draw.createPicture(anchor, imgIndex);
            pict.resize(1,3);
            
            CellStyle tituloEstilo = book.createCellStyle();
            tituloEstilo.setAlignment(HorizontalAlignment.CENTER);
            tituloEstilo.setVerticalAlignment(VerticalAlignment.CENTER);
            Font fuenteTitulo=book.createFont();
            fuenteTitulo.setFontName("Calibri");
            fuenteTitulo.setBold(true);
            fuenteTitulo.setFontHeightInPoints((short)16);
            tituloEstilo.setFont(fuenteTitulo);
            
            Row filaTitulo = sheet.createRow(0);
            Cell celdaTitulo = filaTitulo.createCell(1);
            celdaTitulo.setCellStyle(tituloEstilo);
            celdaTitulo.setCellValue("Museo de Mineralogía Alfred W. Stelzner");
            
            sheet.addMergedRegion(new CellRangeAddress(0, 2, 1, 7));
            
            String[] cabecera = new String[]{"Código", "Especie", "Procedencia", "Descripción", "Colección", "Año de Ingreso", "Ubicación", "Observaciones"};
            
            CellStyle headerStyle = book.createCellStyle();
            headerStyle.setFillForegroundColor(IndexedColors.LIGHT_BLUE.getIndex());
            headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            headerStyle.setBorderBottom(BorderStyle.THIN);
            headerStyle.setBorderLeft(BorderStyle.THIN);
            headerStyle.setBorderRight(BorderStyle.THIN);
            headerStyle.setBorderBottom(BorderStyle.THIN);
            
            Font font = book.createFont();
            font.setFontName("Calibri");
            font.setBold(true);
            font.setColor(IndexedColors.WHITE.getIndex());
            font.setFontHeightInPoints((short)12);
            headerStyle.setFont(font);
            
            Row filaEncabezados = sheet.createRow(3);
            for (int i=0; i< cabecera.length; i++) {
                Cell celdaEncabezado = filaEncabezados.createCell(i);
                celdaEncabezado.setCellStyle(headerStyle);
                celdaEncabezado.setCellValue(cabecera[i]);
            }
            
            ConexionMySQL con = new ConexionMySQL();
            PreparedStatement ps;
            ResultSet rs;
            java.sql.Connection conn = con.getConexion();
            
            int numfilaDatos = 4;
            CellStyle datosEstilo = book.createCellStyle();
            datosEstilo.setBorderBottom(BorderStyle.THIN);
            datosEstilo.setBorderLeft(BorderStyle.THIN);
            datosEstilo.setBorderRight(BorderStyle.THIN);
            datosEstilo.setBorderBottom(BorderStyle.THIN);
            
            ps=conn.prepareStatement("SELECT codigo, especie, procedencia, descripcion, coleccion, ano, ubicacion, observaciones FROM minerales");
            rs = ps.executeQuery();
            
            int numCol = rs.getMetaData().getColumnCount();
            
            while (rs.next()) {
                Row filaDatos = sheet.createRow(numfilaDatos);
                
                for (int a=0; a<numCol; a++) {
                    Cell CeldaDatos = filaDatos.createCell(a);
                    CeldaDatos.setCellStyle(datosEstilo);
                    
                    if(a==5) {
                        CeldaDatos.setCellValue(rs.getInt(a+1));
                    }
                    else {
                        CeldaDatos.setCellValue(rs.getString(a+1));
                    }
                }
                numfilaDatos++;
            }
            
            sheet.autoSizeColumn(0);
            sheet.autoSizeColumn(1);
            sheet.autoSizeColumn(2);
            sheet.autoSizeColumn(3);
            sheet.autoSizeColumn(4);
            sheet.autoSizeColumn(5);
            sheet.autoSizeColumn(6);
            sheet.autoSizeColumn(7);
            
            FileOutputStream fileOut = new FileOutputStream("Respaldo.xlsx");
            book.write(fileOut);
            fileOut.close();
        } catch (FileNotFoundException ex) {
            Logger.getLogger(Minerales.class.getName()).log(Level.SEVERE, null, ex);
        } catch (IOException ex) {
            Logger.getLogger(Minerales.class.getName()).log(Level.SEVERE, null, ex);
        } catch (SQLException ex) {
            Logger.getLogger(Minerales.class.getName()).log(Level.SEVERE, null, ex);
        }
    }
    
    void nuevo() {
        txtCod.setEnabled(true);
        txtCol.setEnabled(true);
        txtAno.setEnabled(true);
        txtDes.setEnabled(true);
        txtObs.setEnabled(true);
        txtPro.setEnabled(true);
        txtSpc.setEnabled(true);
        txtUbi.setEnabled(true);
        codigo();
        txtCol.setText("");
        txtAno.setText("");
        txtDes.setText("");
        txtObs.setText("");
        txtPro.setText("");
        txtSpc.setText("");
        txtUbi.setText("");
        btnGuardar.setEnabled(true);
        btnCancelar.setEnabled(true);
        txtSpc.requestFocus();
    }
    
    void codigo() {
        String cod;
        modelo = (DefaultTableModel) tblMineral.getModel();
        int fila = tblMineral.getRowCount()-1;
        cod = (String) modelo.getValueAt(fila, 1);
        int coda = Integer.parseInt(cod) + 1;
        cod = Integer.toString(coda);
        txtCod.setText(cod);
    }
    
    void relleno() {
        if(txtCol.getText().isEmpty()){
            txtCol.setText("-");
        }
        if(txtAno.getText().isEmpty()){
            txtAno.setText("0");
        }
        if(txtUbi.getText().isEmpty()){
            txtUbi.setText("-");
        }
        if(txtDes.getText().isEmpty()){
            txtDes.setText("-");
        }
        if(txtObs.getText().isEmpty()){
            txtObs.setText("-");
        }
        
    }
    
    String id_actualizar = "";
    void BuscarMineralEditar(String id) {
        String sSQL="";
        String cod = "", spc = "", pro = "", des = "", col = "", ano = "", ubi = "", obs = "";
        ConexionMySQL mysql = new ConexionMySQL();
        java.sql.Connection cn = mysql.getConexion();
        
        sSQL = "SELECT id, codigo, especie, procedencia, descripcion, coleccion, ano, ubicacion, observaciones FROM minerales " +
                "WHERE id = "+id;
        
        try {
            Statement st = cn.createStatement();
            ResultSet rs = st.executeQuery(sSQL);
            
            while (rs.next()) {
                cod = rs.getString("codigo");
                spc = rs.getString("especie");
                pro = rs.getString("procedencia");
                des = rs.getString("descripcion");
                col = rs.getString("coleccion");
                ano = rs.getString("ano");
                ubi = rs.getString("ubicacion");
                obs = rs.getString("observaciones");
            }
            txtCod.setText(cod);
            txtSpc.setText(spc);
            txtPro.setText(pro);
            txtDes.setText(des);
            txtCol.setText(col);
            txtAno.setText(ano);
            txtUbi.setText(ubi);
            txtObs.setText(obs);
            id_actualizar = id;
        } 
        catch (SQLException ex) {
            JOptionPane.showMessageDialog(null, ex);
        }
    }
    
    void EliminarMineral(String id) {
        String sSQL = "";
        String mensaje="";
        ConexionMySQL mysql = new ConexionMySQL();
        java.sql.Connection cn = mysql.getConexion();
        sSQL ="DELETE FROM minerales WHERE id =" + id;
        mensaje = "El mineral se ha eliminado de la base de datos";  
        
        try {                
            PreparedStatement pst = cn.prepareStatement(sSQL); 
               
            int m = pst.executeUpdate();
            if (m > 0) {
                JOptionPane.showMessageDialog(null, mensaje);
                String valor = txtBuscar.getText();
                CargarTablaMinerales(valor);
                inhabilitar();
            }   
        } 
        catch (Exception e) {
            JOptionPane.showMessageDialog(null, e);
        }
    }
    
    /**
     * This method is called from within the constructor to initialize the form.
     * WARNING: Do NOT modify this code. The content of this method is always
     * regenerated by the Form Editor.
     */
    @SuppressWarnings("unchecked")
    // <editor-fold defaultstate="collapsed" desc="Generated Code">//GEN-BEGIN:initComponents
    private void initComponents() {

        jPopupMenu2 = new javax.swing.JPopupMenu();
        mnEditar = new javax.swing.JMenuItem();
        mnBorrar = new javax.swing.JMenuItem();
        jPanel1 = new javax.swing.JPanel();
        jLabel1 = new javax.swing.JLabel();
        txtCod = new javax.swing.JTextField();
        jLabel2 = new javax.swing.JLabel();
        txtSpc = new javax.swing.JTextField();
        jLabel3 = new javax.swing.JLabel();
        jLabel4 = new javax.swing.JLabel();
        jLabel5 = new javax.swing.JLabel();
        jLabel6 = new javax.swing.JLabel();
        txtCol = new javax.swing.JTextField();
        txtPro = new javax.swing.JTextField();
        txtAno = new javax.swing.JTextField();
        txtUbi = new javax.swing.JTextField();
        jLabel7 = new javax.swing.JLabel();
        jLabel8 = new javax.swing.JLabel();
        btnNuevo = new javax.swing.JButton();
        btnGuardar = new javax.swing.JButton();
        btnCancelar = new javax.swing.JButton();
        btnSalir = new javax.swing.JButton();
        jPanel2 = new javax.swing.JPanel();
        jLabel9 = new javax.swing.JLabel();
        txtBuscar = new javax.swing.JTextField();
        jScrollPane1 = new javax.swing.JScrollPane();
        tblMineral = new javax.swing.JTable();
        btnBuscar = new javax.swing.JButton();
        btnGenerar = new javax.swing.JButton();
        btnAdmin = new javax.swing.JButton();
        jScrollPane2 = new javax.swing.JScrollPane();
        jTextPane1 = new javax.swing.JTextPane();
        jScrollPane3 = new javax.swing.JScrollPane();
        txtDes = new javax.swing.JTextArea();
        jScrollPane4 = new javax.swing.JScrollPane();
        txtObs = new javax.swing.JTextArea();

        mnEditar.setText("Modificar");
        mnEditar.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                mnEditarActionPerformed(evt);
            }
        });
        jPopupMenu2.add(mnEditar);

        mnBorrar.setText("Eliminar");
        mnBorrar.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                mnBorrarActionPerformed(evt);
            }
        });
        jPopupMenu2.add(mnBorrar);

        setDefaultCloseOperation(javax.swing.WindowConstants.EXIT_ON_CLOSE);

        jPanel1.setBackground(new java.awt.Color(255, 255, 255));
        jPanel1.setBorder(javax.swing.BorderFactory.createTitledBorder(null, "Ingreso", javax.swing.border.TitledBorder.DEFAULT_JUSTIFICATION, javax.swing.border.TitledBorder.DEFAULT_POSITION, new java.awt.Font("Agency FB", 1, 18))); // NOI18N

        jLabel1.setText("Código:");

        txtCod.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                txtCodActionPerformed(evt);
            }
        });

        jLabel2.setText("Especie:");

        txtSpc.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                txtSpcActionPerformed(evt);
            }
        });

        jLabel3.setText("Procedencia:");

        jLabel4.setText("Colección:");

        jLabel5.setText("Año de Ingreso:");

        jLabel6.setText("Ubicación:");

        txtCol.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                txtColActionPerformed(evt);
            }
        });
        txtCol.addKeyListener(new java.awt.event.KeyAdapter() {
            public void keyPressed(java.awt.event.KeyEvent evt) {
                txtColKeyPressed(evt);
            }
        });

        txtPro.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                txtProActionPerformed(evt);
            }
        });

        txtAno.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                txtAnoActionPerformed(evt);
            }
        });

        txtUbi.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                txtUbiActionPerformed(evt);
            }
        });

        jLabel7.setText("Descripción de la muestra:");

        jLabel8.setText("Observaciones:");

        btnNuevo.setText("Nuevo");
        btnNuevo.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                btnNuevoActionPerformed(evt);
            }
        });

        btnGuardar.setText("Guardar");
        btnGuardar.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                btnGuardarActionPerformed(evt);
            }
        });

        btnCancelar.setText("Cancelar");
        btnCancelar.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                btnCancelarActionPerformed(evt);
            }
        });

        btnSalir.setText("Salir");
        btnSalir.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                btnSalirActionPerformed(evt);
            }
        });

        jPanel2.setBorder(javax.swing.BorderFactory.createTitledBorder(null, "Base de Datos", javax.swing.border.TitledBorder.DEFAULT_JUSTIFICATION, javax.swing.border.TitledBorder.DEFAULT_POSITION, new java.awt.Font("Agency FB", 1, 18))); // NOI18N

        jLabel9.setText("Buscar:");

        tblMineral.setModel(new javax.swing.table.DefaultTableModel(
            new Object [][] {

            },
            new String [] {

            }
        ));
        tblMineral.setComponentPopupMenu(jPopupMenu2);
        jScrollPane1.setViewportView(tblMineral);

        btnBuscar.setText("Buscar");
        btnBuscar.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                btnBuscarActionPerformed(evt);
            }
        });

        btnGenerar.setText("Generar");
        btnGenerar.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                btnGenerarActionPerformed(evt);
            }
        });

        javax.swing.GroupLayout jPanel2Layout = new javax.swing.GroupLayout(jPanel2);
        jPanel2.setLayout(jPanel2Layout);
        jPanel2Layout.setHorizontalGroup(
            jPanel2Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel2Layout.createSequentialGroup()
                .addContainerGap()
                .addGroup(jPanel2Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addComponent(jScrollPane1)
                    .addGroup(jPanel2Layout.createSequentialGroup()
                        .addComponent(jLabel9)
                        .addGap(18, 18, 18)
                        .addComponent(txtBuscar, javax.swing.GroupLayout.PREFERRED_SIZE, 248, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addGap(18, 18, 18)
                        .addComponent(btnBuscar)
                        .addGap(18, 18, 18)
                        .addComponent(btnGenerar)
                        .addGap(0, 0, Short.MAX_VALUE)))
                .addContainerGap())
        );
        jPanel2Layout.setVerticalGroup(
            jPanel2Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel2Layout.createSequentialGroup()
                .addGroup(jPanel2Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(jLabel9)
                    .addComponent(txtBuscar, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(btnBuscar)
                    .addComponent(btnGenerar))
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addComponent(jScrollPane1, javax.swing.GroupLayout.DEFAULT_SIZE, 241, Short.MAX_VALUE)
                .addContainerGap())
        );

        btnAdmin.setText("Admin");
        btnAdmin.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                btnAdminActionPerformed(evt);
            }
        });

        jTextPane1.setBackground(new java.awt.Color(240, 240, 240));
        jTextPane1.setFont(new java.awt.Font("Tahoma", 1, 11)); // NOI18N
        jTextPane1.setText("-Para guardar los cambios utilice \n el botón \"salir\", de lo contrario\n no se generará un respaldo de\n la base de datos.\n-*SIC: significa que la palabra que \n lo antecede es literal o textual, \n aunque pueda no parecerlo.");
        jTextPane1.setFocusable(false);
        jTextPane1.setRequestFocusEnabled(false);
        jScrollPane2.setViewportView(jTextPane1);

        txtDes.setColumns(20);
        txtDes.setRows(5);
        jScrollPane3.setViewportView(txtDes);

        txtObs.setColumns(20);
        txtObs.setRows(5);
        jScrollPane4.setViewportView(txtObs);

        javax.swing.GroupLayout jPanel1Layout = new javax.swing.GroupLayout(jPanel1);
        jPanel1.setLayout(jPanel1Layout);
        jPanel1Layout.setHorizontalGroup(
            jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel1Layout.createSequentialGroup()
                .addContainerGap()
                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addComponent(jPanel2, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                    .addGroup(jPanel1Layout.createSequentialGroup()
                        .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.TRAILING, false)
                            .addGroup(jPanel1Layout.createSequentialGroup()
                                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                                    .addComponent(jLabel3)
                                    .addComponent(jLabel1))
                                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.TRAILING)
                                    .addComponent(txtPro, javax.swing.GroupLayout.PREFERRED_SIZE, 170, javax.swing.GroupLayout.PREFERRED_SIZE)
                                    .addComponent(txtCod, javax.swing.GroupLayout.PREFERRED_SIZE, 170, javax.swing.GroupLayout.PREFERRED_SIZE)))
                            .addGroup(jPanel1Layout.createSequentialGroup()
                                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                                    .addGroup(jPanel1Layout.createSequentialGroup()
                                        .addComponent(jLabel5)
                                        .addGap(84, 84, 84))
                                    .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, jPanel1Layout.createSequentialGroup()
                                        .addComponent(jLabel7)
                                        .addGap(36, 36, 36))
                                    .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, jPanel1Layout.createSequentialGroup()
                                        .addComponent(btnNuevo, javax.swing.GroupLayout.PREFERRED_SIZE, 75, javax.swing.GroupLayout.PREFERRED_SIZE)
                                        .addGap(2, 2, 2)))
                                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING, false)
                                    .addComponent(txtAno, javax.swing.GroupLayout.DEFAULT_SIZE, 170, Short.MAX_VALUE)
                                    .addComponent(btnGuardar, javax.swing.GroupLayout.Alignment.TRAILING, javax.swing.GroupLayout.PREFERRED_SIZE, 75, javax.swing.GroupLayout.PREFERRED_SIZE)
                                    .addComponent(jScrollPane3))))
                        .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.TRAILING)
                            .addGroup(jPanel1Layout.createSequentialGroup()
                                .addGap(153, 153, 153)
                                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING, false)
                                    .addGroup(jPanel1Layout.createSequentialGroup()
                                        .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                                            .addComponent(jLabel2)
                                            .addComponent(jLabel4)
                                            .addComponent(jLabel6))
                                        .addGap(44, 44, 44)
                                        .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                                            .addComponent(txtSpc, javax.swing.GroupLayout.PREFERRED_SIZE, 170, javax.swing.GroupLayout.PREFERRED_SIZE)
                                            .addComponent(txtCol, javax.swing.GroupLayout.PREFERRED_SIZE, 170, javax.swing.GroupLayout.PREFERRED_SIZE)
                                            .addComponent(txtUbi, javax.swing.GroupLayout.PREFERRED_SIZE, 170, javax.swing.GroupLayout.PREFERRED_SIZE)))
                                    .addGroup(jPanel1Layout.createSequentialGroup()
                                        .addComponent(jLabel8)
                                        .addGap(18, 18, 18)
                                        .addComponent(jScrollPane4))))
                            .addGroup(javax.swing.GroupLayout.Alignment.LEADING, jPanel1Layout.createSequentialGroup()
                                .addGap(105, 105, 105)
                                .addComponent(btnCancelar)
                                .addGap(114, 114, 114)
                                .addComponent(btnAdmin)))
                        .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                            .addGroup(jPanel1Layout.createSequentialGroup()
                                .addGap(41, 41, 41)
                                .addComponent(jScrollPane2, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))
                            .addGroup(jPanel1Layout.createSequentialGroup()
                                .addGap(53, 53, 53)
                                .addComponent(btnSalir, javax.swing.GroupLayout.PREFERRED_SIZE, 75, javax.swing.GroupLayout.PREFERRED_SIZE)))
                        .addGap(0, 0, Short.MAX_VALUE)))
                .addContainerGap())
        );
        jPanel1Layout.setVerticalGroup(
            jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel1Layout.createSequentialGroup()
                .addContainerGap()
                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addComponent(jScrollPane2, javax.swing.GroupLayout.PREFERRED_SIZE, 120, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addGroup(jPanel1Layout.createSequentialGroup()
                        .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                            .addComponent(jLabel1)
                            .addComponent(txtCod, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addComponent(jLabel2)
                            .addComponent(txtSpc, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))
                        .addGap(20, 20, 20)
                        .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                            .addComponent(jLabel3)
                            .addComponent(txtPro, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addComponent(jLabel4)
                            .addComponent(txtCol, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))
                        .addGap(18, 18, 18)
                        .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                            .addComponent(jLabel5)
                            .addComponent(txtAno, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addComponent(jLabel6)
                            .addComponent(txtUbi, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))
                        .addGap(17, 17, 17)
                        .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                            .addComponent(jLabel7)
                            .addComponent(jLabel8)
                            .addComponent(jScrollPane3, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addComponent(jScrollPane4, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))))
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(btnNuevo)
                    .addComponent(btnGuardar)
                    .addComponent(btnCancelar)
                    .addComponent(btnAdmin)
                    .addComponent(btnSalir))
                .addGap(18, 18, 18)
                .addComponent(jPanel2, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                .addContainerGap())
        );

        javax.swing.GroupLayout layout = new javax.swing.GroupLayout(getContentPane());
        getContentPane().setLayout(layout);
        layout.setHorizontalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(layout.createSequentialGroup()
                .addContainerGap()
                .addComponent(jPanel1, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                .addContainerGap())
        );
        layout.setVerticalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(layout.createSequentialGroup()
                .addContainerGap()
                .addComponent(jPanel1, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                .addContainerGap())
        );

        pack();
    }// </editor-fold>//GEN-END:initComponents

    private void btnSalirActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_btnSalirActionPerformed
        generar();
        this.dispose();
    }//GEN-LAST:event_btnSalirActionPerformed

    private void btnNuevoActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_btnNuevoActionPerformed
        nuevo();
        accion="Insertar";
    }//GEN-LAST:event_btnNuevoActionPerformed

    private void btnCancelarActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_btnCancelarActionPerformed
        inhabilitar();
    }//GEN-LAST:event_btnCancelarActionPerformed

    private void txtCodActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_txtCodActionPerformed
        txtCod.transferFocus();
    }//GEN-LAST:event_txtCodActionPerformed

    private void txtSpcActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_txtSpcActionPerformed
        txtSpc.transferFocus();
    }//GEN-LAST:event_txtSpcActionPerformed

    private void txtProActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_txtProActionPerformed
        txtPro.transferFocus();
    }//GEN-LAST:event_txtProActionPerformed

    private void txtColActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_txtColActionPerformed
        txtCol.transferFocus();
    }//GEN-LAST:event_txtColActionPerformed

    private void txtAnoActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_txtAnoActionPerformed
        txtAno.transferFocus();
    }//GEN-LAST:event_txtAnoActionPerformed

    private void txtUbiActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_txtUbiActionPerformed
        txtUbi.transferFocus();
    }//GEN-LAST:event_txtUbiActionPerformed

    private void btnGuardarActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_btnGuardarActionPerformed
        relleno();
        ConexionMySQL mysql=new ConexionMySQL();
        Connection cn = mysql.getConexion();
        String cod, spc, pro, des, col, ano, ubi, obs;
        String sSQL= "";
        String mensaje="";
        
        cod= txtCod.getText();
        spc=txtSpc.getText();
        pro= txtPro.getText();
        des= txtDes.getText();
        col= txtCol.getText();
        ano= txtAno.getText();
        ubi= txtUbi.getText();
        obs= txtObs.getText();
                
        if (accion.equals("Insertar")) {
            sSQL = "INSERT INTO minerales (codigo, especie, procedencia, descripcion, coleccion, ano, ubicacion, observaciones)" + "VALUES(?, ?, ?, ?, ?, ?, ?, ?)";
            mensaje= "Los datos se han insertado de manera satisfactoria";
        }
        else if (accion.equals("Modificar")) {
            sSQL = "UPDATE minerales " +
                    "SET codigo = ?," + "especie = ?," +
                    "procedencia = ?," + "descripcion = ?," +
                    "coleccion = ?," + "ano = ?, " +
                    "ubicacion = ?," + "observaciones = ? " +
                    "WHERE id = "+id_actualizar;
            mensaje= "Los datos se han modificado de manera satisfactoria";
        }
        try {
            PreparedStatement pst = cn.prepareStatement(sSQL);
            pst.setString(1, cod);
            pst.setString(2, spc);
            pst.setString(3, pro);
            pst.setString(4, des);
            pst.setString(5, col);
            pst.setString(6, ano);
            pst.setString(7, ubi);
            pst.setString(8, obs);
            
            int n = pst.executeUpdate();
            
            if (n > 0) {
                JOptionPane.showMessageDialog(null, mensaje);
                String valor = txtBuscar.getText();
                CargarTablaMinerales(valor);
                inhabilitar();
            }
        } catch (SQLException ex) {
            JOptionPane.showMessageDialog(null, ex);
        }
    }//GEN-LAST:event_btnGuardarActionPerformed

    private void btnBuscarActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_btnBuscarActionPerformed
        String valor = txtBuscar.getText();
        CargarTablaMinerales(valor);
    }//GEN-LAST:event_btnBuscarActionPerformed

    private void btnGenerarActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_btnGenerarActionPerformed
        try {
            obj = new clsExportarExcel();
            obj.clsExportarExcel(tblMineral);
        } catch (IOException ex) {
            Logger.getLogger(Minerales.class.getName()).log(Level.SEVERE, null, ex);
        }
    }//GEN-LAST:event_btnGenerarActionPerformed

    private void mnEditarActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_mnEditarActionPerformed
        int filasel;
        String id;
        
        try {
            filasel = tblMineral.getSelectedRow();
            if (filasel == -1) {
                JOptionPane.showMessageDialog(null, "No se ha seleccionado ninguna fila");
            }
            else {
                accion = "Modificar";
                modelo = (DefaultTableModel) tblMineral.getModel();
                id = (String) modelo.getValueAt(filasel, 0);
                habilitar();
                BuscarMineralEditar(id);
            }
        }
        catch (Exception e) {
            
        }
    }//GEN-LAST:event_mnEditarActionPerformed

    private void mnBorrarActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_mnBorrarActionPerformed
        int filasel;
        String id;
        try {
            filasel = tblMineral.getSelectedRow();
            if(filasel == -1) {
                JOptionPane.showMessageDialog(null, "No se ha seleccionado ninguna fila");
            }
            else {
                modelo = (DefaultTableModel) tblMineral.getModel();    
                id = (String) modelo.getValueAt(filasel, 0);
                
                String botones[]={"Cerrar","Cancelar"};
                int eleccion= JOptionPane.showConfirmDialog(this, "¿Desea eliminar este mineral?");
                if (eleccion==JOptionPane.YES_OPTION){
                    EliminarMineral(id);
                }
                else if (eleccion==JOptionPane.NO_OPTION) {
                }
            }
        } 
        catch (Exception e) {
            JOptionPane.showMessageDialog(null, e);
        }
    }//GEN-LAST:event_mnBorrarActionPerformed

    private void btnAdminActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_btnAdminActionPerformed
        new Admin().setVisible(true);
    }//GEN-LAST:event_btnAdminActionPerformed

    private void txtColKeyPressed(java.awt.event.KeyEvent evt) {//GEN-FIRST:event_txtColKeyPressed
        
    }//GEN-LAST:event_txtColKeyPressed

    /**
     * @param args the command line arguments
     */
    

    // Variables declaration - do not modify//GEN-BEGIN:variables
    private javax.swing.JButton btnAdmin;
    private javax.swing.JButton btnBuscar;
    private javax.swing.JButton btnCancelar;
    private javax.swing.JButton btnGenerar;
    private javax.swing.JButton btnGuardar;
    private javax.swing.JButton btnNuevo;
    private javax.swing.JButton btnSalir;
    private javax.swing.JLabel jLabel1;
    private javax.swing.JLabel jLabel2;
    private javax.swing.JLabel jLabel3;
    private javax.swing.JLabel jLabel4;
    private javax.swing.JLabel jLabel5;
    private javax.swing.JLabel jLabel6;
    private javax.swing.JLabel jLabel7;
    private javax.swing.JLabel jLabel8;
    private javax.swing.JLabel jLabel9;
    private javax.swing.JPanel jPanel1;
    private javax.swing.JPanel jPanel2;
    private javax.swing.JPopupMenu jPopupMenu2;
    private javax.swing.JScrollPane jScrollPane1;
    private javax.swing.JScrollPane jScrollPane2;
    private javax.swing.JScrollPane jScrollPane3;
    private javax.swing.JScrollPane jScrollPane4;
    private javax.swing.JTextPane jTextPane1;
    private javax.swing.JMenuItem mnBorrar;
    private javax.swing.JMenuItem mnEditar;
    private javax.swing.JTable tblMineral;
    private javax.swing.JTextField txtAno;
    private javax.swing.JTextField txtBuscar;
    private javax.swing.JTextField txtCod;
    private javax.swing.JTextField txtCol;
    private javax.swing.JTextArea txtDes;
    private javax.swing.JTextArea txtObs;
    private javax.swing.JTextField txtPro;
    private javax.swing.JTextField txtSpc;
    private javax.swing.JTextField txtUbi;
    // End of variables declaration//GEN-END:variables

    private static class clsExportarExcel {
        public void clsExportarExcel(JTable t) throws IOException {
            JFileChooser chooser = new JFileChooser();
            FileNameExtensionFilter filter = new FileNameExtensionFilter("Archivos de excel", "xls");
            chooser.setFileFilter(filter);
            chooser.setDialogTitle("Guardar archivo");
            chooser.setAcceptAllFileFilterUsed(false);
            if (chooser.showSaveDialog(null) == JFileChooser.APPROVE_OPTION) {
                String ruta = chooser.getSelectedFile().toString().concat(".xlsx");
                try {
                    File archivoXLSX = new File(ruta);
                    if (archivoXLSX.exists()) {
                        archivoXLSX.delete();
                    }
                    archivoXLSX.createNewFile();
                    Workbook libro = new XSSFWorkbook();
                    FileOutputStream archivo = new FileOutputStream(archivoXLSX);
                    Sheet hoja = libro.createSheet("MineralesFiltrados");
                    hoja.setDisplayGridlines(false);
                    
                    InputStream is;
                    is = new FileInputStream("logo.jpg");
                    byte[] bytes = IOUtils.toByteArray(is);
                    int imgIndex = libro.addPicture(bytes, Workbook.PICTURE_TYPE_JPEG);
                    is.close();
            
                    CreationHelper help = libro.getCreationHelper();
                    Drawing draw = hoja.createDrawingPatriarch();
            
                    ClientAnchor anchor = help.createClientAnchor();
                    anchor.setCol1(0);
                    anchor.setRow1(0);
                    Picture pict = draw.createPicture(anchor, imgIndex);
                    pict.resize(2,3);
                    
                    CellStyle tituloEstilo = libro.createCellStyle();
                    tituloEstilo.setAlignment(HorizontalAlignment.CENTER);
                    tituloEstilo.setVerticalAlignment(VerticalAlignment.CENTER);
                    Font fuenteTitulo=libro.createFont();
                    fuenteTitulo.setFontName("Calibri");
                    fuenteTitulo.setBold(true);
                    fuenteTitulo.setFontHeightInPoints((short)16);
                    tituloEstilo.setFont(fuenteTitulo);
            
                    Row filaTitulo = hoja.createRow(0);
                    Cell celdaTitulo = filaTitulo.createCell(2);
                    celdaTitulo.setCellStyle(tituloEstilo);
                    celdaTitulo.setCellValue("Museo de Mineralogía Alfred W. Stelzner");
            
                    hoja.addMergedRegion(new CellRangeAddress(0, 2, 2, 8));
                    
                    String[] cabecera = new String[]{"ID", "Código", "Especie", "Procedencia", "Descripción", "Colección", "Año de Ingreso", "Ubicación", "Observaciones"};
            
                    CellStyle headerStyle = libro.createCellStyle();
                    headerStyle.setFillForegroundColor(IndexedColors.LIGHT_BLUE.getIndex());
                    headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                    headerStyle.setBorderBottom(BorderStyle.THIN);
                    headerStyle.setBorderLeft(BorderStyle.THIN);
                    headerStyle.setBorderRight(BorderStyle.THIN);
                    headerStyle.setBorderBottom(BorderStyle.THIN);
            
                    Font font = libro.createFont();
                    font.setFontName("Calibri");
                    font.setBold(true);
                    font.setColor(IndexedColors.WHITE.getIndex());
                    font.setFontHeightInPoints((short)12);
                    headerStyle.setFont(font);
            
                    Row filaEncabezados = hoja.createRow(3);
                    for (int i=0; i< cabecera.length; i++) {
                        Cell celdaEncabezado = filaEncabezados.createCell(i);
                        celdaEncabezado.setCellStyle(headerStyle);
                        celdaEncabezado.setCellValue(cabecera[i]);
                    }
                    
                    int filaInicio = 4;
                    CellStyle datosEstilo = libro.createCellStyle();
                    datosEstilo.setBorderBottom(BorderStyle.THIN);
                    datosEstilo.setBorderLeft(BorderStyle.THIN);
                    datosEstilo.setBorderRight(BorderStyle.THIN);
                    datosEstilo.setBorderBottom(BorderStyle.THIN);
                    for (int f = 0; f < t.getRowCount(); f++) {
                        Row fila = hoja.createRow(filaInicio);
                        filaInicio++;
                        for (int c = 0; c < t.getColumnCount(); c++) {
                            Cell celda = fila.createCell(c);
                            celda.setCellStyle(datosEstilo);
                            if (t.getValueAt(f, c) instanceof Double) {
                                celda.setCellValue(Double.parseDouble(t.getValueAt(f, c).toString()));
                            } else if (t.getValueAt(f, c) instanceof Float) {
                                celda.setCellValue(Float.parseFloat((String) t.getValueAt(f, c)));
                            } else {
                                celda.setCellValue(String.valueOf(t.getValueAt(f, c)));
                            }
                        }
                    }
                    
                    hoja.autoSizeColumn(0);
                    hoja.autoSizeColumn(1);
                    hoja.autoSizeColumn(2);
                    hoja.autoSizeColumn(3);
                    hoja.autoSizeColumn(4);
                    hoja.autoSizeColumn(5);
                    hoja.autoSizeColumn(6);
                    hoja.autoSizeColumn(7);
                    hoja.autoSizeColumn(8);
            
                    libro.write(archivo);
                    archivo.close();
                    Desktop.getDesktop().open(archivoXLSX);
                } catch (IOException | NumberFormatException e) {
                    throw e;
                }
            }
        }
    }
}