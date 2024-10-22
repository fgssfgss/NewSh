package com.newsh;

import java.io.File;
import javax.swing.*;
import javax.swing.filechooser.FileFilter;
import javax.swing.filechooser.FileNameExtensionFilter;

/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */

/**
 * @author Eugeny
 */
public class MainFrame extends javax.swing.JFrame {

    private JFileChooser fc;

    public final static FileFilter WORD_FILES = new FileNameExtensionFilter("Файлы Word", "docx");
    public final static FileFilter XML_FILES = new FileNameExtensionFilter("Файлы XML", "xml");
    private final static FileFilter DIRECTORIES = new FileFilter() {

        @Override
        public boolean accept(File f) {
            return f.isDirectory();
        }

        @Override
        public String getDescription() {
            return "Каталог";
        }
    };

    /**
     * Creates new form MainFrame
     */
    public MainFrame() {
        fc = new JFileChooser();
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

        buttonGroup1 = new javax.swing.ButtonGroup();
        buttonGroup2 = new javax.swing.ButtonGroup();
        jPanel1 = new javax.swing.JPanel();
        jProgressBar1 = new javax.swing.JProgressBar();
        jPanel6 = new javax.swing.JPanel();
        plainGroupRadioButton = new javax.swing.JRadioButton();
        stGroupRadioButton = new javax.swing.JRadioButton();
        jPanel7 = new javax.swing.JPanel();
        le25RadioButton = new javax.swing.JRadioButton();
        gt25RadioButton = new javax.swing.JRadioButton();
        makeItButton = new javax.swing.JButton();
        jPanel2 = new javax.swing.JPanel();
        jPanel3 = new javax.swing.JPanel();
        wordFileNameTextField = new javax.swing.JTextField();
        chooseWordButton = new javax.swing.JButton();
        jPanel4 = new javax.swing.JPanel();
        xmlFileNameTextField = new javax.swing.JTextField();
        chooseXMLButton = new javax.swing.JButton();
        jPanel5 = new javax.swing.JPanel();
        outputDirectoryNameTextField = new javax.swing.JTextField();
        chooseOutputDirectoryButton = new javax.swing.JButton();
        jComboBox1 = new javax.swing.JComboBox();
        jLabel1 = new javax.swing.JLabel();

        setDefaultCloseOperation(javax.swing.WindowConstants.EXIT_ON_CLOSE);

        jPanel6.setBorder(javax.swing.BorderFactory.createTitledBorder("Группа"));

        buttonGroup1.add(plainGroupRadioButton);
        plainGroupRadioButton.setSelected(true);
        plainGroupRadioButton.setText("Обычная");

        buttonGroup1.add(stGroupRadioButton);
        stGroupRadioButton.setText("СТ");

        javax.swing.GroupLayout jPanel6Layout = new javax.swing.GroupLayout(jPanel6);
        jPanel6.setLayout(jPanel6Layout);
        jPanel6Layout.setHorizontalGroup(
                jPanel6Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                        .addGroup(jPanel6Layout.createSequentialGroup()
                                .addContainerGap()
                                .addGroup(jPanel6Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                                        .addComponent(plainGroupRadioButton)
                                        .addComponent(stGroupRadioButton))
                                .addContainerGap(javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
        );
        jPanel6Layout.setVerticalGroup(
                jPanel6Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                        .addGroup(jPanel6Layout.createSequentialGroup()
                                .addContainerGap(javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                                .addComponent(plainGroupRadioButton)
                                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.UNRELATED)
                                .addComponent(stGroupRadioButton))
        );

        jPanel7.setBorder(javax.swing.BorderFactory.createTitledBorder("Количество предметов"));

        buttonGroup2.add(le25RadioButton);
        le25RadioButton.setText("до 25");

        buttonGroup2.add(gt25RadioButton);
        gt25RadioButton.setSelected(true);
        gt25RadioButton.setText("Больше 25");

        javax.swing.GroupLayout jPanel7Layout = new javax.swing.GroupLayout(jPanel7);
        jPanel7.setLayout(jPanel7Layout);
        jPanel7Layout.setHorizontalGroup(
                jPanel7Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                        .addGroup(jPanel7Layout.createSequentialGroup()
                                .addContainerGap()
                                .addGroup(jPanel7Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                                        .addComponent(le25RadioButton)
                                        .addComponent(gt25RadioButton))
                                .addContainerGap(javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
        );
        jPanel7Layout.setVerticalGroup(
                jPanel7Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                        .addGroup(jPanel7Layout.createSequentialGroup()
                                .addContainerGap()
                                .addComponent(le25RadioButton)
                                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.UNRELATED)
                                .addComponent(gt25RadioButton)
                                .addContainerGap(8, Short.MAX_VALUE))
        );

        makeItButton.setText("Сделать!");
        makeItButton.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                makeItButtonActionPerformed(evt);
            }
        });

        javax.swing.GroupLayout jPanel1Layout = new javax.swing.GroupLayout(jPanel1);
        jPanel1.setLayout(jPanel1Layout);
        jPanel1Layout.setHorizontalGroup(
                jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                        .addGroup(jPanel1Layout.createSequentialGroup()
                                .addContainerGap()
                                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                                        .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, jPanel1Layout.createSequentialGroup()
                                                .addGap(0, 0, Short.MAX_VALUE)
                                                .addComponent(jProgressBar1, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))
                                        .addComponent(jPanel6, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                                        .addComponent(jPanel7, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                                        .addComponent(makeItButton, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
                                .addContainerGap())
        );
        jPanel1Layout.setVerticalGroup(
                jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                        .addGroup(jPanel1Layout.createSequentialGroup()
                                .addContainerGap()
                                .addComponent(jProgressBar1, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.UNRELATED)
                                .addComponent(jPanel6, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                                .addComponent(jPanel7, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.UNRELATED)
                                .addComponent(makeItButton, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                                .addContainerGap())
        );

        jPanel3.setBorder(javax.swing.BorderFactory.createTitledBorder("Выбрать Word документ"));

        chooseWordButton.setText("Обзор...");
        chooseWordButton.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                chooseWordButtonActionPerformed(evt);
            }
        });

        javax.swing.GroupLayout jPanel3Layout = new javax.swing.GroupLayout(jPanel3);
        jPanel3.setLayout(jPanel3Layout);
        jPanel3Layout.setHorizontalGroup(
                jPanel3Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                        .addGroup(jPanel3Layout.createSequentialGroup()
                                .addContainerGap()
                                .addComponent(wordFileNameTextField, javax.swing.GroupLayout.DEFAULT_SIZE, 230, Short.MAX_VALUE)
                                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                                .addComponent(chooseWordButton)
                                .addContainerGap())
        );
        jPanel3Layout.setVerticalGroup(
                jPanel3Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                        .addGroup(jPanel3Layout.createSequentialGroup()
                                .addContainerGap()
                                .addGroup(jPanel3Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                                        .addComponent(wordFileNameTextField, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                                        .addComponent(chooseWordButton))
                                .addContainerGap(javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
        );

        jPanel4.setBorder(javax.swing.BorderFactory.createTitledBorder("Выбрать XML:"));

        chooseXMLButton.setText("Обзор...");
        chooseXMLButton.setEnabled(false);
        chooseXMLButton.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                chooseXMLButtonActionPerformed(evt);
            }
        });

        javax.swing.GroupLayout jPanel4Layout = new javax.swing.GroupLayout(jPanel4);
        jPanel4.setLayout(jPanel4Layout);
        jPanel4Layout.setHorizontalGroup(
                jPanel4Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                        .addGroup(jPanel4Layout.createSequentialGroup()
                                .addContainerGap()
                                .addComponent(xmlFileNameTextField)
                                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                                .addComponent(chooseXMLButton)
                                .addContainerGap())
        );
        jPanel4Layout.setVerticalGroup(
                jPanel4Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                        .addGroup(jPanel4Layout.createSequentialGroup()
                                .addContainerGap()
                                .addGroup(jPanel4Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                                        .addComponent(xmlFileNameTextField, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                                        .addComponent(chooseXMLButton))
                                .addContainerGap(16, Short.MAX_VALUE))
        );

        jPanel5.setBorder(javax.swing.BorderFactory.createTitledBorder("Папка выхода"));

        chooseOutputDirectoryButton.setText("Обзор...");
        chooseOutputDirectoryButton.setEnabled(false);
        chooseOutputDirectoryButton.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                chooseOutputDirectoryButtonActionPerformed(evt);
            }
        });

        javax.swing.GroupLayout jPanel5Layout = new javax.swing.GroupLayout(jPanel5);
        jPanel5.setLayout(jPanel5Layout);
        jPanel5Layout.setHorizontalGroup(
                jPanel5Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                        .addGroup(jPanel5Layout.createSequentialGroup()
                                .addContainerGap()
                                .addComponent(outputDirectoryNameTextField)
                                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                                .addComponent(chooseOutputDirectoryButton)
                                .addContainerGap())
        );
        jPanel5Layout.setVerticalGroup(
                jPanel5Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                        .addGroup(jPanel5Layout.createSequentialGroup()
                                .addContainerGap()
                                .addGroup(jPanel5Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                                        .addComponent(outputDirectoryNameTextField, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                                        .addComponent(chooseOutputDirectoryButton))
                                .addContainerGap(26, Short.MAX_VALUE))
        );

        jComboBox1.setModel(new javax.swing.DefaultComboBoxModel(new String[]{"Спецiалiста", "Бакалавра", "Магiстра"}));

        jLabel1.setText("Диплом:");

        javax.swing.GroupLayout jPanel2Layout = new javax.swing.GroupLayout(jPanel2);
        jPanel2.setLayout(jPanel2Layout);
        jPanel2Layout.setHorizontalGroup(
                jPanel2Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                        .addGroup(jPanel2Layout.createSequentialGroup()
                                .addGroup(jPanel2Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                                        .addGroup(jPanel2Layout.createSequentialGroup()
                                                .addGap(10, 10, 10)
                                                .addComponent(jPanel3, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
                                        .addGroup(jPanel2Layout.createSequentialGroup()
                                                .addContainerGap()
                                                .addComponent(jPanel4, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
                                        .addGroup(jPanel2Layout.createSequentialGroup()
                                                .addContainerGap()
                                                .addComponent(jPanel5, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
                                        .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, jPanel2Layout.createSequentialGroup()
                                                .addContainerGap()
                                                .addComponent(jLabel1)
                                                .addGap(18, 18, 18)
                                                .addComponent(jComboBox1, 0, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)))
                                .addContainerGap())
        );
        jPanel2Layout.setVerticalGroup(
                jPanel2Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                        .addGroup(jPanel2Layout.createSequentialGroup()
                                .addGroup(jPanel2Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                                        .addComponent(jComboBox1, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                                        .addComponent(jLabel1))
                                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                                .addComponent(jPanel3, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                                .addComponent(jPanel4, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                                .addComponent(jPanel5, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                                .addContainerGap(javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
        );

        javax.swing.GroupLayout layout = new javax.swing.GroupLayout(getContentPane());
        getContentPane().setLayout(layout);
        layout.setHorizontalGroup(
                layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                        .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, layout.createSequentialGroup()
                                .addContainerGap()
                                .addComponent(jPanel2, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.UNRELATED)
                                .addComponent(jPanel1, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                                .addContainerGap())
        );
        layout.setVerticalGroup(
                layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                        .addGroup(layout.createSequentialGroup()
                                .addContainerGap()
                                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                                        .addComponent(jPanel1, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                                        .addComponent(jPanel2, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
                                .addContainerGap())
        );

        pack();
    }// </editor-fold>//GEN-END:initComponents

    public JComboBox getComboBox() {
        return jComboBox1;
    }

    private void chooseWordButtonActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_chooseWordButtonActionPerformed
        File wordFile = selectDocument(WORD_FILES);
        if (wordFile != null) {
            wordFileNameTextField.setText(wordFile.getAbsolutePath());
        } else {
            wordFileNameTextField.setText("");
        }
        chooseXMLButton.setEnabled(!wordFileNameTextField.getText().isEmpty());
    }//GEN-LAST:event_chooseWordButtonActionPerformed

    private void chooseXMLButtonActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_chooseXMLButtonActionPerformed
        File xmlFile = selectDocument(XML_FILES);
        if (xmlFile != null) {
            xmlFileNameTextField.setText(xmlFile.getAbsolutePath());
        } else {
            xmlFileNameTextField.setText("");
        }
        chooseOutputDirectoryButton.setEnabled(!xmlFileNameTextField.getText().isEmpty());
    }//GEN-LAST:event_chooseXMLButtonActionPerformed

    private void chooseOutputDirectoryButtonActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_chooseOutputDirectoryButtonActionPerformed
        File directory = selectDocument(DIRECTORIES);
        if (directory != null) {
            outputDirectoryNameTextField.setText(directory.getAbsolutePath());
        } else {
            outputDirectoryNameTextField.setText("");
        }
        makeItButton.setEnabled(!outputDirectoryNameTextField.getText().isEmpty());
    }//GEN-LAST:event_chooseOutputDirectoryButtonActionPerformed

    private void makeItButtonActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_makeItButtonActionPerformed
        String wordPath = wordFileNameTextField.getText();
        String xmlPath = xmlFileNameTextField.getText();
        String outPath = outputDirectoryNameTextField.getText();
        makeItButton.setEnabled(false);
        try {
            run(wordPath, xmlPath, outPath);
        } catch (Exception e) {
            JOptionPane.showMessageDialog(getFocusOwner(), "ERROR: " + e.toString(), "ERROR", JOptionPane.ERROR_MESSAGE);
            makeItButton.setEnabled(true);
        }
    }//GEN-LAST:event_makeItButtonActionPerformed

    /**
     * @param args the command line arguments
     */
    public static void main(String args[]) {
        /* Set the Nimbus look and feel */
        //<editor-fold defaultstate="collapsed" desc=" Look and feel setting code (optional) ">
        /* If Nimbus (introduced in Java SE 6) is not available, stay with the default look and feel.
         * For details see http://download.oracle.com/javase/tutorial/uiswing/lookandfeel/plaf.html 
         */
        MyLogger.setDebugMode();
        //MyLogger.setNormalMode();
        MyLogger.log("App started");
        try {
            javax.swing.UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());
        } catch (ClassNotFoundException | InstantiationException | IllegalAccessException | javax.swing.UnsupportedLookAndFeelException ex) {
            java.util.logging.Logger.getLogger(MainFrame.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        }
        //</editor-fold>
        /* Create and display the form */
        java.awt.EventQueue.invokeLater(new Runnable() {
            @Override
            public void run() {
                new MainFrame().setVisible(true);

            }
        });
    }

    // Variables declaration - do not modify//GEN-BEGIN:variables
    private javax.swing.ButtonGroup buttonGroup1;
    private javax.swing.ButtonGroup buttonGroup2;
    private javax.swing.JButton chooseOutputDirectoryButton;
    private javax.swing.JButton chooseWordButton;
    private javax.swing.JButton chooseXMLButton;
    private javax.swing.JRadioButton gt25RadioButton;
    private javax.swing.JComboBox jComboBox1;
    private javax.swing.JLabel jLabel1;
    private javax.swing.JPanel jPanel1;
    private javax.swing.JPanel jPanel2;
    private javax.swing.JPanel jPanel3;
    private javax.swing.JPanel jPanel4;
    private javax.swing.JPanel jPanel5;
    private javax.swing.JPanel jPanel6;
    private javax.swing.JPanel jPanel7;
    private javax.swing.JProgressBar jProgressBar1;
    private javax.swing.JRadioButton le25RadioButton;
    private javax.swing.JButton makeItButton;
    private javax.swing.JTextField outputDirectoryNameTextField;
    private javax.swing.JRadioButton plainGroupRadioButton;
    private javax.swing.JRadioButton stGroupRadioButton;
    private javax.swing.JTextField wordFileNameTextField;
    private javax.swing.JTextField xmlFileNameTextField;
    // End of variables declaration//GEN-END:variables

    private File selectDocument(FileFilter filter) {
        fc.setFileFilter(filter);
        if (filter == DIRECTORIES) {
            fc.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
        }
        int showOpenDialog = fc.showOpenDialog(this);
        if (showOpenDialog == JFileChooser.APPROVE_OPTION) {
            return fc.getSelectedFile();
        } else {
            return null;
        }
    }

    /**
     * Настройка ProgressBar первый раз вызвать его. Потом можно вызывать
     * updateProgess(completed)
     *
     * @param completed завершено
     * @param overall   всего
     */
    public void updateProgress(int completed, int overall) {
        jProgressBar1.setMaximum(overall);
        jProgressBar1.setMinimum(completed);
        jProgressBar1.setValue(completed);
        if (completed == overall) {
            JOptionPane.showMessageDialog(getFocusOwner(), "Создание файлов успешно завершено", "COMPLETED", JOptionPane.INFORMATION_MESSAGE);
            jProgressBar1.setValue(0);
        }
    }

    /**
     * Упрощенное обновление ProgressBar
     *
     * @param completed завершено
     */
    public void updateProgress(int completed) {
        jProgressBar1.setValue(completed);
        if (completed == jProgressBar1.getMaximum()) {
            JOptionPane.showMessageDialog(getFocusOwner(), "Создание файлов успешно завершено", "COMPLETED", JOptionPane.INFORMATION_MESSAGE);
            jProgressBar1.setValue(0);
        }

    }

    /**
     * Генерация документов
     *
     * @param wordPath путь к word-файлу
     * @param xmlPath  путь к xml-файлу
     * @param outPath  путь к каталогу ввыода
     */
    public void run(String wordPath, String xmlPath, String outPath) {
        final String PATH = "doc\\";
        String templateFileName = PATH + "sh13.docx";
        String templateFileNameST = PATH + "sh13_st.docx";
        String templateFileNameSP = PATH + "sh13_sp.docx";
        String templateFileNameM = PATH + "sh13_m.docx";

        Hack h = new Hack();
        h.Hack(wordPath);

        TemplateEngine tempEngine = new TemplateEngine(this);
        tempEngine.setDestination(outPath.concat("\\"));
        tempEngine.setDocFileName(wordPath);

        // do not forget!
        // 0 - specialist
        // 1 - bachelor
        // 2 - magister
        // Do you want to see how programming cancer look like?
        Integer selectedItem = getComboBox().getSelectedIndex();
        if (selectedItem == 1) {
            MyLogger.log("bachelor");
            if (getStGroupRadioButton().isSelected())
                tempEngine.setTemplateFileName(templateFileNameST);
            else
                tempEngine.setTemplateFileName(templateFileName);
        } else if (selectedItem == 0) {
            MyLogger.log("specialist");
            tempEngine.setTemplateFileName(templateFileNameSP);
        } else if (selectedItem == 2) {
            MyLogger.log("magister");
            tempEngine.setTemplateFileName(templateFileNameM);
        }
        tempEngine.setXmlFileName(xmlPath);
        tempEngine.start();
    }

    public JButton getMakeItButton() {
        return makeItButton;
    }

    public void setMakeItButton(JButton makeItButton) {
        this.makeItButton = makeItButton;
    }

    public JRadioButton getStGroupRadioButton() {
        return stGroupRadioButton;
    }
}
