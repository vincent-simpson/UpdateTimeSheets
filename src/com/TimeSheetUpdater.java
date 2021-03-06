package com;

import javafx.application.Application;
import javafx.geometry.HPos;
import javafx.geometry.Pos;
import javafx.scene.Scene;
import javafx.scene.control.Button;
import javafx.scene.control.TextField;
import javafx.scene.layout.ColumnConstraints;
import javafx.scene.layout.GridPane;
import javafx.scene.layout.VBox;
import javafx.scene.text.Text;
import javafx.stage.FileChooser;
import javafx.stage.Stage;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;

import java.io.*;
import java.time.DayOfWeek;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.Arrays;
import java.util.Iterator;

public class TimeSheetUpdater extends Application {
	public static void main(String[] args) {
		launch(args);
	}

	private int year = -1;
	private boolean weekday;
	private int start_row = -1;

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
		while (sheetIterator.hasNext()) {
			Sheet s = sheetIterator.next();

			for (int row = 4; row < 33; row += 2) {
				for (int col = 1; col < 6; col++) {
					Cell c = s.getRow(row).getCell(col);
					if (c != null) {
						c.setCellValue(0);
						c.setCellType(CellType.BLANK);
					}
				}
			}
		}

		XSSFFormulaEvaluator.evaluateAllFormulaCells(workbook);

		if (inputStream != null) {
			inputStream.close();
		}

		FileOutputStream outputStream = new FileOutputStream(timeSheets);
		workbook.write(outputStream);
		outputStream.close();
		workbook.close();
	}

	@Override
	public void start(Stage primaryStage) {

		GridPane gp = new GridPane();
		gp.setGridLinesVisible(false);

		TextField promptYear = new TextField();
		promptYear.setPromptText("Update for which year?");

		Button submitYear = new Button("Submit year");

		Text successfullyUpdated = new Text("Spreadsheet successfully updated");

		Button browseButton = new Button("Browse");

		ColumnConstraints column1 = new ColumnConstraints();
		column1.setHalignment(HPos.CENTER);

		VBox vbox = new VBox(8);
		vbox.getChildren().addAll(promptYear, submitYear, browseButton);

		gp.getChildren().add(vbox);

		gp.setAlignment(Pos.CENTER);

		browseButton.setOnAction(event -> {
			if (year != -1) {
				FileChooser fc = new FileChooser();
				fc.setTitle("Browse to TimeSheet File");
				File selectedFile = fc.showOpenDialog(primaryStage);
				if (selectedFile != null) {
					final String filePath = selectedFile.getAbsolutePath();
					int successfulInt = updateSpreadsheet(filePath);
					if (successfulInt == 1)
						gp.add(successfullyUpdated, 0, 2);
				}
			}
		});

		submitYear.setOnAction(event -> year = Integer.parseInt(promptYear.getText()));

		Scene scene = new Scene(gp, 200, 200);
		primaryStage.setScene(scene);
		primaryStage.show();
	}

	private int updateSpreadsheet(String filePath) {
		try {
			clearSpreadSheet(filePath);

			Workbook workbook;

			LocalDate[] endingPayPeriods = getSemiMonthlyDates();

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

				/*
				 * This code could probably use some optimization. Basically what we're looking
				 * at is which sheet we're currently viewing. If its the first sheet, the date
				 * range for the pay period needs to be the original base date up to the first
				 * predefined end of the pay date. Every subsequent start date of the pay period
				 * is the end date of the previous period + 1.
				 */
				if (index == 0) {
					payPeriod.setCellValue(baseDate.plusDays(1).format(formatterHeaderPart1)
							+ endingPayPeriods[0].format(formatterHeaderPart2));

					endingPayPeriodsForPayDate[0] = endingPayPeriods[0].plusDays(1);

					payDate.setCellValue(endingPayPeriods[0].plusDays(7).format(formatter));
				} else if (index <= 23) {
					payPeriod.setCellValue(endingPayPeriodsForPayDate[index - 1].format(formatterHeaderPart1)
							+ endingPayPeriodsForPayDate[index].format(formatterHeaderPart2));
					endingPayPeriodsForPayDate[index] = endingPayPeriods[index].plusDays(1);
					payDate.setCellValue(endingPayPeriods[index].plusDays(7).format(formatter));
				}

				/*
				 * Set sheet name to payDate
				 */
				workbook.setSheetName(workbook.getSheetIndex(sheet), payDate.getStringCellValue().replace('/', '-'));

				Row start_row = getStartingRow(baseDate, endingPayPeriods, sheet);
				start_row.setRowNum(start_row.getRowNum() - 1);
				while (!baseDate.isAfter(endingPayPeriods[index]) && start_row.getRowNum() < 34) {
					
					if (start_row != null && (baseDate.getDayOfWeek() != DayOfWeek.SATURDAY && baseDate.getDayOfWeek() != DayOfWeek.SUNDAY)) {
						Cell dateCell = start_row.createCell(1);
						dateCell.setCellValue(baseDate.format(formatter));
						dateCell.setCellStyle(cellStyle);
						start_row.setRowNum(start_row.getRowNum() + 2);						
					} 
					
					
					baseDate = baseDate.plusDays(1);
				}
				index++;

			}

			// This code is to make sure that the columns are resized to fit any text that
			// went past the previously
			// set column boundaries before the new data was entered
			for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
//				workbook.getSheetAt(i).setColumnWidth(1, 15);
//				workbook.getSheetAt(i).setColumnWidth(3, 15);
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

	private Row getStartingRow(LocalDate baseDate, LocalDate[] endingPayPeriods, Sheet sheet) {
		Row dayOfWeekRow = null;

		// determine start row
		switch (baseDate.getDayOfWeek()) {
		case MONDAY:
			start_row = 5;
			break;
		case TUESDAY:
			start_row = 7;
			break;
		case WEDNESDAY:
			start_row = 9;
			break;
		case THURSDAY:
			start_row = 11;
			break;
		case FRIDAY:
			start_row = 13;
			break;
		default:
			break;
		}
		
		dayOfWeekRow = sheet.getRow(start_row);
		return dayOfWeekRow;
	}

	private LocalDate[] getSemiMonthlyDates() {
		return new LocalDate[] { LocalDate.of(year, 1, 8), LocalDate.of(year, 1, 24), LocalDate.of(year, 2, 8),
				LocalDate.of(year, 2, 21), LocalDate.of(year, 3, 8), LocalDate.of(year, 3, 24),
				LocalDate.of(year, 4, 8), LocalDate.of(year, 4, 23), LocalDate.of(year, 5, 8),
				LocalDate.of(year, 5, 24), LocalDate.of(year, 6, 8), LocalDate.of(year, 6, 23),
				LocalDate.of(year, 7, 8), LocalDate.of(year, 7, 24), LocalDate.of(year, 8, 8),
				LocalDate.of(year, 8, 24), LocalDate.of(year, 9, 8), LocalDate.of(year, 9, 23),
				LocalDate.of(year, 10, 8), LocalDate.of(year, 10, 24), LocalDate.of(year, 11, 8),
				LocalDate.of(year, 11, 23), LocalDate.of(year, 12, 8), LocalDate.of(year, 12, 24)

		};
	}
}
