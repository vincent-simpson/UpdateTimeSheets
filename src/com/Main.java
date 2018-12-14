package com;


import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.Arrays;
import java.util.Iterator;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.*;


import javafx.application.Application;
import javafx.scene.Scene;
import javafx.scene.control.Button;
import javafx.scene.control.TextField;
import javafx.scene.layout.GridPane;
import javafx.scene.layout.Pane;
import javafx.scene.text.Text;
import javafx.stage.FileChooser;
import javafx.stage.Stage;

public class Main extends Application {
    private int year=-1;

    @Override
    public void start(Stage primaryStage) throws Exception{
        Pane root = new Pane();
        GridPane gp = new GridPane();

        TextField promptYear = new TextField();
        promptYear.setPromptText("Update for which year?");
        Button submitYear = new Button("Submit year");
        Text successfullyUpdated = new Text("Spreadsheet successfully updated");


        Button browseButton = new Button("Browse");
        gp.add(promptYear, 0, 0);
        gp.add(submitYear, 1, 0);
        gp.add(browseButton, 0, 1);

        browseButton.setOnAction(event -> {
            if(year != -1) {
                FileChooser fc = new FileChooser();
                fc.setTitle("Browse to TimeSheet File");
                File selectedFile = fc.showOpenDialog(primaryStage);
                if(selectedFile != null){
                    final String filePath = selectedFile.getAbsolutePath();
                    int successfullInt = updateSpeadsheet(filePath);
                    if(successfullInt == 1) gp.add(successfullyUpdated, 0, 2);
                }
            }
        });

        submitYear.setOnAction(event -> {
            year = Integer.parseInt(promptYear.getText());
        });

        root.getChildren().add(gp);
        Scene scene = new Scene(root, 250, 200);
        primaryStage.setScene(scene);
        primaryStage.show();

    }

    private int updateSpeadsheet(String filePath) {
        try {
            clearSpreadSheet(filePath);

            Workbook workbook;            

            LocalDate[] endingPayPeriods = {
                    LocalDate.of(year, 1, 8),
                    LocalDate.of(year, 1, 24),
                    LocalDate.of(year, 2, 8),
                    LocalDate.of(year, 2, 21),
                    LocalDate.of(year, 3, 8),
                    LocalDate.of(year, 3, 24),
                    LocalDate.of(year, 4, 8),
                    LocalDate.of(year, 4, 23),
                    LocalDate.of(year, 5, 8),
                    LocalDate.of(year, 5, 24),
                    LocalDate.of(year, 6, 8),
                    LocalDate.of(year, 6, 23),
                    LocalDate.of(year, 7, 8),
                    LocalDate.of(year, 7, 24),
                    LocalDate.of(year, 8, 8),
                    LocalDate.of(year, 8, 24),
                    LocalDate.of(year, 9, 8),
                    LocalDate.of(year, 9, 23),
                    LocalDate.of(year, 10, 8),
                    LocalDate.of(year, 10, 24),
                    LocalDate.of(year, 11, 8),
                    LocalDate.of(year, 11, 23),
                    LocalDate.of(year, 12, 8),
                    LocalDate.of(year, 12, 24)
            };

            LocalDate[] endingPayPeriodsForPayDate = Arrays.copyOf(endingPayPeriods, endingPayPeriods.length);

            File timeSheets = new File(filePath);
            FileInputStream inputStream = null;

            if (!timeSheets.exists()) {
                throw new FileNotFoundException();
            } else {
                inputStream = new FileInputStream(timeSheets);
                workbook = WorkbookFactory.create(inputStream);
            }

            DateTimeFormatter formatter = DateTimeFormatter.ofPattern("MM/dd/yyyy");
            DateTimeFormatter formatterHeaderPart1 = DateTimeFormatter.ofPattern("MM/dd-");
            DateTimeFormatter formatterHeaderPart2 = DateTimeFormatter.ofPattern("MM/dd/yy");

            LocalDate baseDate = LocalDate.of(year - 1, 12, 24);

            CellStyle cellStyle = workbook.createCellStyle();
            cellStyle.setBorderTop(BorderStyle.MEDIUM);
            cellStyle.setBorderBottom(BorderStyle.MEDIUM);
            cellStyle.setBorderLeft(BorderStyle.THIN);
            cellStyle.setBorderRight(BorderStyle.THIN);
            cellStyle.setAlignment(HorizontalAlignment.RIGHT);

            Iterator<Sheet> sheetIterator = workbook.sheetIterator();

            int index = 0;

            while (sheetIterator.hasNext()) {
                Sheet sheet = sheetIterator.next();

                Row payPeriodInfoRow = sheet.getRow(1);
                Cell payPeriod = payPeriodInfoRow.createCell(1);
                Cell payDate = payPeriodInfoRow.createCell(3);
                Cell paidDateCell = payPeriodInfoRow.createCell(4);

                paidDateCell.setCellType(CellType.BLANK);
                paidDateCell.setCellStyle(cellStyle);
                payDate.setCellStyle(cellStyle);
                payPeriod.setCellStyle(cellStyle);

                if (index == 0) {
                    payPeriod.setCellValue(baseDate.plusDays(1).format(formatterHeaderPart1) + endingPayPeriods[0].format(formatterHeaderPart2));
                    endingPayPeriodsForPayDate[0] = endingPayPeriods[0].plusDays(1);
                    payDate.setCellValue(endingPayPeriods[0].plusDays(7).format(formatter));
                } else if (index <= 23) {
                    payPeriod.setCellValue(endingPayPeriodsForPayDate[index - 1].format(formatterHeaderPart1) + endingPayPeriodsForPayDate[index].format(formatterHeaderPart2));
                    endingPayPeriodsForPayDate[index] = endingPayPeriods[index].plusDays(1);
                    payDate.setCellValue(endingPayPeriods[index].plusDays(7).format(formatter));
                }

                Row dayOfWeekRow = null;
                boolean weekday;
                boolean firstMonday = true; //Its assumed that the first day of the week will be a Monday. If not,
                // then the case statement for that day of the week will set firstMonday=false
                //meaning that the next time we enter the Monday case statement, we need to go to the second Monday cell.
                //This method has been applied to all other days of the week below.
                boolean firstTuesday = true;
                boolean firstWednesday = true;
                boolean firstThursday = true;
                boolean firstFriday=true;
                
                boolean secondMonday=true;
                boolean secondTuesday=true;
                boolean secondWednesday=true;
                boolean secondThursday=true;
                boolean secondFriday = true;
                
//                boolean thirdMonday = true;
//                boolean thirdTuesday = true;
//                boolean thirdWednesday = true;
//                boolean thirdThursday = true;
//                boolean thirdFriday = true;


                while (baseDate.isBefore(endingPayPeriods[index])) {
                    weekday = true;

                    baseDate = baseDate.plusDays(1);
                    switch (baseDate.getDayOfWeek()) {
                        case MONDAY:
                            if (firstMonday) {
                                dayOfWeekRow = sheet.getRow(4);
                                firstMonday=false;
                            } else if (secondMonday) {
                                dayOfWeekRow = sheet.getRow(14);
                                secondMonday=false;
                            } else  {
                                dayOfWeekRow = sheet.getRow(24);
                            }
                            break;
                        case TUESDAY:
                            firstMonday = false;
                            if (firstTuesday) {
                                dayOfWeekRow = sheet.getRow(6);
                                firstTuesday = false;
                            } else if (secondTuesday) {
                                dayOfWeekRow = sheet.getRow(16);
                                secondTuesday = false;
                            } else {
                                dayOfWeekRow = sheet.getRow(26);
                            }
                            break;
                        case WEDNESDAY:
                            firstMonday = false;
                            firstTuesday = false;
                            if (firstWednesday) {
                                dayOfWeekRow = sheet.getRow(8);
                                firstWednesday = false;
                            } else if (secondWednesday) {
                                dayOfWeekRow = sheet.getRow(18);
                                secondWednesday=false;
                            } else {
                                dayOfWeekRow = sheet.getRow(28);
                            }
                            break;
                        case THURSDAY:
                            firstMonday = false;
                            firstTuesday = false;
                            firstWednesday = false;
                            if (firstThursday) {
                                dayOfWeekRow = sheet.getRow(10);
                                firstThursday = false;
                            } else if (secondThursday) {
                                dayOfWeekRow = sheet.getRow(20);
                                secondThursday = false;
                            } else {
                                dayOfWeekRow = sheet.getRow(30);
                            }
                            break;
                        case FRIDAY:
                            firstMonday = false;
                            firstTuesday = false;
                            firstWednesday = false;
                            firstThursday = false;
                            if (firstFriday) {
                                dayOfWeekRow = sheet.getRow(12);
                                firstFriday = false;
                            } else if (secondFriday) {
                                dayOfWeekRow = sheet.getRow(22);
                                secondFriday = false;
                            } else {
                                dayOfWeekRow = sheet.getRow(32);
                            }
                            break;
                        case SATURDAY:
                            weekday = false;
                            break;
                        case SUNDAY:
                            weekday = false;
                            break;
                    }

                    if (weekday) {
                        Cell dateCell = dayOfWeekRow.createCell(1);
                        dateCell.setCellValue(baseDate.format(formatter));

                        cellStyle.setBorderBottom(BorderStyle.MEDIUM);
                        cellStyle.setBorderTop(BorderStyle.MEDIUM);
                        cellStyle.setBorderLeft(BorderStyle.THIN);
                        cellStyle.setBorderRight(BorderStyle.THIN);
                        dateCell.setCellStyle(cellStyle);
                    }
                }
                index++;
            }

            for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                workbook.getSheetAt(i).autoSizeColumn(1);
                workbook.getSheetAt(i).autoSizeColumn(3);
            }

            if (inputStream != null) {
                inputStream.close();
            }

            FileOutputStream outputStream = new FileOutputStream(timeSheets);
            workbook.write(outputStream);
            outputStream.close();
            workbook.close();


        } catch (IOException e) {
            e.printStackTrace();
        }
        return 1;
    }
    
    private void clearSpreadSheet(String path) throws EncryptedDocumentException, IOException {
    	Workbook workbook;
    	
    	File timeSheets = new File(path);
        FileInputStream inputStream = null;

        if (!timeSheets.exists()) {
            throw new FileNotFoundException();
        } else {
            inputStream = new FileInputStream(timeSheets);
            workbook = WorkbookFactory.create(inputStream);
        }
        
        Iterator<Sheet> sheetIterator = workbook.sheetIterator();
        while(sheetIterator.hasNext()) {
        	Sheet s = sheetIterator.next();
        	
        	for(int i=4; i < 32; i+=2 ) {
        		Cell c = s.getRow(i).getCell(1);
        		if(c != null) {
            		c.setCellType(CellType.BLANK);
        		}
        	}      	
        }
      
        if (inputStream != null) {
            inputStream.close();
        }

        FileOutputStream outputStream = new FileOutputStream(timeSheets);
        workbook.write(outputStream);
        outputStream.close();
        workbook.close();
    	
    }


    public static void main(String[] args) {
        launch(args);
    }
}
