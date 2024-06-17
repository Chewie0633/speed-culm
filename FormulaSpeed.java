// File: src/main/java/com/formulaspeed/FormulaSpeed.java

package com.formulaspeed;

import javax.swing.*; // Importing Swing components for GUI
import javax.swing.table.DefaultTableModel; // Importing table model for JTable
import java.awt.*; // Importing AWT for layout management
import java.awt.event.ActionEvent; // Importing event classes
import java.awt.event.ActionListener; // Importing event listener interface
import java.io.FileOutputStream; // Importing FileOutputStream to write to files
import java.io.IOException; // Importing IOException for handling file errors
import org.apache.poi.ss.usermodel.*; // Importing POI classes for spreadsheet handling
import org.apache.poi.xssf.usermodel.XSSFWorkbook; // Importing POI class for creating Excel workbook

public class FormulaSpeed {
    private JFrame frame; // Main frame of the application
    private JTextField distanceField; // Text field for inputting distance
    private JTextField timeField; // Text field for inputting time
    private JTable resultsTable; // Table to display the results
    private DefaultTableModel tableModel; // Model for the JTable

    public FormulaSpeed() {
        // Initializing the main frame
        frame = new JFrame("FormulaSpeed");
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE); // Exit the application when the frame is closed
        frame.setSize(600, 400); // Setting the size of the frame
        
        // Creating a panel with GridLayout to arrange components
        JPanel panel = new JPanel(new GridLayout(5, 2));

        // Adding labels and text fields for distance and time inputs
        panel.add(new JLabel("Distance (km):"));
        distanceField = new JTextField();
        panel.add(distanceField);

        panel.add(new JLabel("Time (hours):"));
        timeField = new JTextField();
        panel.add(timeField);

        // Adding a button to calculate the average speed
        JButton calculateButton = new JButton("Calculate Average Speed");
        panel.add(calculateButton);

        // Setting up the table model and JTable to display the results
        String[] columnNames = {"Distance (km)", "Time (hours)", "Average Speed (km/h)"};
        tableModel = new DefaultTableModel(columnNames, 0);
        resultsTable = new JTable(tableModel);
        JScrollPane scrollPane = new JScrollPane(resultsTable); // Adding table to scroll pane
        frame.add(scrollPane, BorderLayout.CENTER); // Adding the scroll pane to the center of the frame

        // Adding a button to save the results to a spreadsheet
        JButton saveButton = new JButton("Save to Spreadsheet");
        panel.add(saveButton);

        // Adding the panel to the top of the frame
        frame.add(panel, BorderLayout.NORTH);

        // Adding action listener to the calculate button
        calculateButton.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                calculateAndDisplaySpeed(); // Call method to calculate and display speed
            }
        });

        // Adding action listener to the save button
        saveButton.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                saveToSpreadsheet(); // Call method to save results to spreadsheet
            }
        });

        frame.setVisible(true); // Make the frame visible
    }

    // Method to calculate and display the average speed
    private void calculateAndDisplaySpeed() {
        try {
            // Parsing input values for distance and time
            double distance = Double.parseDouble(distanceField.getText());
            double time = Double.parseDouble(timeField.getText());
            if (time <= 0) {
                // Show error message if time is less than or equal to zero
                JOptionPane.showMessageDialog(frame, "Time must be greater than zero.", "Error", JOptionPane.ERROR_MESSAGE);
                return;
            }
            // Calculating average speed
            double averageSpeed = distance / time;
            // Adding the results to the table
            tableModel.addRow(new Object[]{distance, time, averageSpeed});
        } catch (NumberFormatException e) {
            // Show error message if inputs are not valid numbers
            JOptionPane.showMessageDialog(frame, "Please enter valid numbers.", "Error", JOptionPane.ERROR_MESSAGE);
        }
    }

    // Method to save the results to a spreadsheet
    private void saveToSpreadsheet() {
        try (Workbook workbook = new XSSFWorkbook()) {
            // Creating a new sheet in the workbook
            Sheet sheet = workbook.createSheet("Results");

            // Creating the header row
            Row headerRow = sheet.createRow(0);
            for (int i = 0; i < tableModel.getColumnCount(); i++) {
                Cell cell = headerRow.createCell(i);
                cell.setCellValue(tableModel.getColumnName(i)); // Setting column names as header
            }

            // Adding data rows to the sheet
            for (int i = 0; i < tableModel.getRowCount(); i++) {
                Row row = sheet.createRow(i + 1);
                for (int j = 0; j < tableModel.getColumnCount(); j++) {
                    Cell cell = row.createCell(j);
                    cell.setCellValue(tableModel.getValueAt(i, j).toString()); // Setting cell values from table data
                }
            }

            // Writing the workbook to a file
            try (FileOutputStream fileOut = new FileOutputStream("results.xlsx")) {
                workbook.write(fileOut); // Saving the Excel file
            }

            // Show success message
            JOptionPane.showMessageDialog(frame, "Results saved to results.xlsx", "Success", JOptionPane.INFORMATION_MESSAGE);
        } catch (IOException e) {
            // Show error message if there is an issue saving the file
            JOptionPane.showMessageDialog(frame, "Error saving to spreadsheet.", "Error", JOptionPane.ERROR_MESSAGE);
        }
    }

    public static void main(String[] args) {
        // Running the application on the Event Dispatch Thread
        SwingUtilities.invokeLater(new Runnable() {
            @Override
            public void run() {
                new FormulaSpeed(); // Creating an instance of the main class
            }
        });
    }
}
