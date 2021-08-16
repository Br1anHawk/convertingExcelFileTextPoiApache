import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import javax.swing.table.DefaultTableModel;
import java.awt.event.*;
import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;

public class Dialog extends JDialog {
    private JPanel contentPane;
    private JButton buttonOK;
    private JButton buttonCancel;
    private JButton buttonAddFile;
    private JTable tableOfSelectedFiles;
    private JButton buttonConvertFiles;
    private JTextField textFieldRouterID;
    private JButton buttonSetDefaultDirectory;
    private JLabel labelDefaultDirectory;
    private JButton buttonClearTable;
    private JPanel panelOKCancelButtons;
    private JButton buttonSelectedItemUp;
    private JButton buttonSelectedItemDown;
    private JButton buttonRemoveSelectedFiles;
    private DefaultTableModel tableModel;
    private JFileChooser fileChooser;
    private ArrayList<File> files = new ArrayList<>();

    public Dialog() {
        setContentPane(contentPane);
        setModal(true);
        getRootPane().setDefaultButton(buttonOK);

        panelOKCancelButtons.setVisible(false);
        labelDefaultDirectory.setText("S:\\programming"); //FOR_DEBUGGING_ONLY!

        initModelForJTable();

        buttonSetDefaultDirectory.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent actionEvent) {
                fileChooser = new JFileChooser();
                fileChooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
                int isDirectorySelectedInt = fileChooser.showOpenDialog(contentPane);
                if (isDirectorySelectedInt == JFileChooser.APPROVE_OPTION) {
                    labelDefaultDirectory.setText(fileChooser.getSelectedFile().getPath());
                }
            }
        });

        buttonAddFile.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent actionEvent) {
                fileChooser = new JFileChooser();
                fileChooser.setMultiSelectionEnabled(true);
                fileChooser.setFileFilter(new FileNameExtensionFilter("Excel files", "xlsx"));
                fileChooser.setCurrentDirectory(new File(labelDefaultDirectory.getText()));
                int isFileSelectedInt = fileChooser.showOpenDialog(contentPane);
                if (isFileSelectedInt == JFileChooser.APPROVE_OPTION) {
                    File[] selectedFiles = fileChooser.getSelectedFiles();
                    files.addAll(Arrays.asList(selectedFiles));
                    for (File file: files) {
                        tableModel.addRow(new Object[]{file.getName()});
                    }
                }
            }
        });

        buttonRemoveSelectedFiles.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent actionEvent) {
                int[] selectedRows = tableOfSelectedFiles.getSelectedRows();
                tableOfSelectedFiles.clearSelection();
                for (int i = selectedRows.length - 1; i >= 0; i--) {
                    tableModel.removeRow(selectedRows[i]);
                    files.remove(selectedRows[i]);
                }
            }
        });

        buttonClearTable.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent actionEvent) {
                while(tableModel.getRowCount() > 0) {
                    tableModel.removeRow(0);
                }
                files.clear();
            }
        });

        buttonSelectedItemUp.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent actionEvent) {
                moveRowsBy(-1);
            }
        });

        buttonSelectedItemDown.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent actionEvent) {
                moveRowsBy(1);
            }
        });

        buttonConvertFiles.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent actionEvent) {
                if (files.isEmpty()) {
                    JOptionPane.showMessageDialog(contentPane, "Add files!");
                    return;
                }
                if (textFieldRouterID.getText().isEmpty()) {
                    JOptionPane.showMessageDialog(contentPane, "Set router ID!");
                    return;
                }
                ExcelFileConverter excelFileConverter = new ExcelFileConverter();
                if (excelFileConverter.isConvertedFileExists(files, textFieldRouterID.getText())) {
                    int isReplace = JOptionPane.showConfirmDialog(
                            contentPane,
                            "File is exists. Replace it?",
                            "Info",
                            JOptionPane.YES_NO_OPTION
                    );
                    if (isReplace == JOptionPane.NO_OPTION) return;
                }
                try {
                    excelFileConverter.convert(files, textFieldRouterID.getText());
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        });

        buttonOK.addActionListener(new ActionListener() {
            public void actionPerformed(ActionEvent e) {
                onOK();
            }
        });

        buttonCancel.addActionListener(new ActionListener() {
            public void actionPerformed(ActionEvent e) {
                onCancel();
            }
        });

        // call onCancel() when cross is clicked
        setDefaultCloseOperation(DO_NOTHING_ON_CLOSE);
        addWindowListener(new WindowAdapter() {
            public void windowClosing(WindowEvent e) {
                onCancel();
            }
        });

        // call onCancel() on ESCAPE
        contentPane.registerKeyboardAction(new ActionListener() {
            public void actionPerformed(ActionEvent e) {
                onCancel();
            }
        }, KeyStroke.getKeyStroke(KeyEvent.VK_ESCAPE, 0), JComponent.WHEN_ANCESTOR_OF_FOCUSED_COMPONENT);


    }

    private void moveRowsBy(int moveBy) {
        int[] selectedRows = tableOfSelectedFiles.getSelectedRows();
        if (selectedRows[0] + moveBy < 0 || selectedRows[selectedRows.length - 1] + moveBy == tableModel.getRowCount()) {
            return;
        }
        tableOfSelectedFiles.clearSelection();
        if (moveBy < 0) {
            for (int rowId : selectedRows) {
                tableModel.moveRow(rowId, rowId, rowId + moveBy);
                tableOfSelectedFiles.addRowSelectionInterval(rowId + moveBy, rowId + moveBy);
                files.add(rowId + moveBy, files.remove(rowId));
            }
        } else {
            for (int i = selectedRows.length - 1; i >= 0; i--) {
                int rowId = selectedRows[i];
                tableModel.moveRow(rowId, rowId, rowId + moveBy);
                tableOfSelectedFiles.addRowSelectionInterval(rowId + moveBy, rowId + moveBy);
                files.add(rowId + moveBy, files.remove(rowId));
            }
        }
        debuggingLogPrintListOfFiles(); //FOR_DEBUGGING_ONLY!
    }

    private void onOK() {
        // add your code here
        dispose();
    }

    private void onCancel() {
        // add your code here if necessary
        dispose();
    }

    private void initModelForJTable() {
        tableModel = new DefaultTableModel();
        tableModel.addColumn("File name");
        tableOfSelectedFiles.setModel(tableModel);
    }

    public static void main(String[] args) {
        Dialog dialog = new Dialog();
        dialog.setTitle("Converter");
        dialog.pack();
        dialog.setVisible(true);
        System.exit(0);
    }

    private void debuggingLogPrintListOfFiles() {
        for (File file : files) {
            System.out.println(file.getName());
        }
        System.out.println("--------------------------");
        System.out.println();
    }
}