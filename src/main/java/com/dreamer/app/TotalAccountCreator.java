package com.dreamer.app;

import java.io.FileOutputStream;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.Arrays;
import java.util.Date;
import java.util.List;
import java.util.stream.Collectors;
import java.util.stream.IntStream;

import org.apache.commons.cli.CommandLine;
import org.apache.commons.cli.CommandLineParser;
import org.apache.commons.cli.DefaultParser;
import org.apache.commons.cli.HelpFormatter;
import org.apache.commons.cli.Option;
import org.apache.commons.cli.Options;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

public class TotalAccountCreator {

	private static final Logger LOG = LoggerFactory.getLogger(TotalAccountCreator.class);

	private static final String DEFAULT_DATE_FORMAT = "yyyy-MM-dd";
	private static final String[] DEFAULT_ACCOUNTS = { "현금" };
	private static final int INDEX_A = 65;

	private Path inputFilePath;
	private Path outputFilePath;
	private int sheetIndex;
	private int headerRow;
	private int debitColumn;
	private int creditColumn;
	private String[] targetAccounts;
	private String dateFormat;

	private CellStyle numberStyle;
	private CellStyle dateStyle;

	public TotalAccountCreator(Path inputFilePath, Path outputFilePath, int sheetIndex, int headerRow, int debitColumn,
			int creditColumn, String[] targetAccount, String dateFormat) {
		this.inputFilePath = inputFilePath;
		this.outputFilePath = outputFilePath;
		this.sheetIndex = sheetIndex;
		this.headerRow = headerRow;
		this.debitColumn = debitColumn;
		this.creditColumn = creditColumn;
		this.targetAccounts = targetAccount;
		this.dateFormat = dateFormat;
	}

	public static void main(String[] args) throws Exception {
		Options options = new Options();
		options
				.addOption(Option.builder("i").longOpt("input").hasArg().required().desc("Input file for processing.").build());
		options.addOption(Option.builder("o").longOpt("output").hasArg().required().desc("Output file.").build());
		options.addOption(Option.builder("s").longOpt("sheet").hasArg().required(false).desc("Sheet index.").build());
		options.addOption(
				Option.builder("b").longOpt("beginRow").hasArg().required(false).desc("Start row of content.").build());
		options.addOption(Option.builder("h").longOpt("headerRow").hasArg().required(false).desc("row of header.").build());
		options.addOption(
				Option.builder("d").longOpt("debitColumn").hasArg().required(false).desc("Debit column index.").build());
		options.addOption(
				Option.builder("c").longOpt("creditColumn").hasArg().required(false).desc("Credit column index.").build());
		options.addOption(Option.builder("a").longOpt("targetAccounts").hasArgs().required(false)
				.desc("Processing account separated by space. \nex) 현금 카드").build());
		options.addOption(
				Option.builder("f").longOpt("dateFormat").hasArg().required(false).desc("Date format for use.").build());

		CommandLineParser parser = new DefaultParser();
		CommandLine cmd = null;
		try {
			cmd = parser.parse(options, args, true);
		} catch (org.apache.commons.cli.ParseException e) {
			System.err.println("error : " + e.getMessage());
			HelpFormatter formatter = new HelpFormatter();
			formatter.printHelp("TotalAccountCreator", options, true);
			System.exit(-1);
		}

		Path inputPath = Paths.get(cmd.getOptionValue("input"));
		Path outputPath = Paths.get(cmd.getOptionValue("output"));
		int sheetIndex = Integer.parseInt(cmd.getOptionValue("sheet", "1"));
		int headerRow = Integer.parseInt(cmd.getOptionValue("headerRow", "2"));
		int debitColumn = toIndex(cmd.getOptionValue("debitColumn", "D").toUpperCase());
		int creditColumn = toIndex(cmd.getOptionValue("creditColumn", "F").toUpperCase());
		String[] targetAccounts = cmd.getOptionValues("targetAccounts");
		String dateFormat = cmd.getOptionValue("dateFormat", DEFAULT_DATE_FORMAT);

		if (LOG.isInfoEnabled()) {
			LOG.info("Input: {}", inputPath);
			LOG.info("Output: {}", outputPath);
			LOG.info("Sheet index: {}", sheetIndex);
			LOG.info("Header row: {}", headerRow);
			LOG.info("Debit column: {}", toExcelIndex(debitColumn));
			LOG.info("Credit column: {}", toExcelIndex(creditColumn));
			LOG.info("Target accounts: {}",
					Arrays.stream(targetAccounts == null ? DEFAULT_ACCOUNTS : targetAccounts).collect(Collectors.joining(",")));
			LOG.info("Date format: {}", dateFormat);
		}

		TotalAccountCreator totalAccountCreator = new TotalAccountCreator(inputPath, outputPath, sheetIndex, headerRow,
				debitColumn, creditColumn, targetAccounts == null ? DEFAULT_ACCOUNTS : targetAccounts, dateFormat);

		totalAccountCreator.create();
		totalAccountCreator.openFile();
	}

	private void openFile() throws Exception {
		String[] commands = null;
		if (System.getProperty("os.name").contains("Mac")) {
			commands = new String[] { "open", outputFilePath.toString() };
		} else {
			// maybe windows
			commands = new String[] { "start", "", outputFilePath.toString() }; // "" is window title
		}

		ProcessBuilder processBuilder = new ProcessBuilder(commands);
		processBuilder.start();
	}

	private void create() throws Exception {
		List<String> targetAccountList = Arrays.stream(targetAccounts).collect(Collectors.toList());

		try (Workbook outputWorkbook = createOutputWorkbook()) {
			try (Workbook workbook = new XSSFWorkbook(inputFilePath.toFile())) {
				Sheet outputSheet = outputWorkbook.getSheetAt(0);
				Sheet sheet = workbook.getSheetAt(sheetIndex - 1);
				for (Row row : sheet) {
					if (row.getRowNum() < headerRow) {
						continue;
					}

					if (LOG.isTraceEnabled()) {
						LOG.trace("row: {}", row.getRowNum());
					}

					Cell cell = row.getCell(debitColumn);
					if (targetAccountList.contains(getCellValue(cell))) {
						writeDebit(outputSheet, row);
					}

					cell = row.getCell(creditColumn);
					if (targetAccountList.contains(getCellValue(cell))) {
						writeCredit(outputSheet, row);
					}
				}
			} catch (Exception e) {
				LOG.error(e.getMessage(), e);
			} finally {
				outputWorkbook.write(new FileOutputStream(outputFilePath.toFile()));
				if (LOG.isInfoEnabled()) {
					LOG.info("Result file created: {}", outputFilePath.toString());
				}
			}
		}
	}

	private Workbook createOutputWorkbook() throws Exception {
		Workbook workbook = new XSSFWorkbook();
		CellStyle centerAlign = workbook.createCellStyle();
		centerAlign.setAlignment(HorizontalAlignment.CENTER);

		Sheet sheet = workbook.createSheet("Total Account");
		Row row = sheet.createRow(0);
		CellRangeAddress cellRangeAddress = new CellRangeAddress(0, 0, 0, 5);
		createTitleCell(sheet, row, 0, "현금계정", cellRangeAddress, centerAlign);
		sheet.addMergedRegion(applyBorder(sheet, cellRangeAddress));

		row = sheet.createRow(1);
		cellRangeAddress = new CellRangeAddress(1, 1, 0, 2);
		createTitleCell(sheet, row, 0, "차변", cellRangeAddress, centerAlign);
		sheet.addMergedRegion(applyBorder(sheet, cellRangeAddress));

		cellRangeAddress = new CellRangeAddress(1, 1, 3, 5);
		createTitleCell(sheet, row, 3, "대변", cellRangeAddress, centerAlign);
		sheet.addMergedRegion(applyBorder(sheet, cellRangeAddress));

		row = sheet.createRow(2);
		createTitleCell(sheet, row, 0, "순번", centerAlign);
		createTitleCell(sheet, row, 1, "날짜", 3000, centerAlign);
		createTitleCell(sheet, row, 2, "금액", centerAlign);
		createTitleCell(sheet, row, 3, "금액", centerAlign);
		createTitleCell(sheet, row, 4, "날짜", 3000, centerAlign);
		createTitleCell(sheet, row, 5, "순번", centerAlign);

		numberStyle = workbook.createCellStyle();
		numberStyle.setDataFormat(workbook.createDataFormat().getFormat("0,000"));

		dateStyle = workbook.createCellStyle();
		dateStyle.setDataFormat(workbook.getCreationHelper().createDataFormat().getFormat(dateFormat));

		return workbook;
	}

	private CellRangeAddress applyBorder(Sheet sheet, CellRangeAddress cellRangeAddress) {
		RegionUtil.setBorderTop(BorderStyle.THIN, cellRangeAddress, sheet);
		RegionUtil.setBorderBottom(BorderStyle.THIN, cellRangeAddress, sheet);
		RegionUtil.setBorderLeft(BorderStyle.THIN, cellRangeAddress, sheet);
		RegionUtil.setBorderRight(BorderStyle.THIN, cellRangeAddress, sheet);

		return cellRangeAddress;
	}

	private Cell applyBorder(Sheet sheet, Cell cell) {
		int rowIndex = cell.getRowIndex();
		int columnIndex = cell.getColumnIndex();

		applyBorder(sheet, new CellRangeAddress(rowIndex, rowIndex, columnIndex, columnIndex));
		return cell;
	}

	private Cell createTitleCell(Sheet sheet, Row row, int column, String title, CellStyle cellStyle) {
		Cell cell = row.createCell(column);
		cell.setCellValue(title);
		cell.setCellStyle(cellStyle);

		return applyBorder(sheet, cell);
	}

	private Cell createTitleCell(Sheet sheet, Row row, int column, String title, int width, CellStyle cellStyle) {
		Cell cell = row.createCell(column);
		cell.setCellValue(title);
		cell.setCellStyle(cellStyle);
		sheet.setColumnWidth(column, width);

		return applyBorder(sheet, cell);
	}

	private Cell createTitleCell(Sheet sheet, Row row, int column, String title, CellRangeAddress cellRangeAddress,
			CellStyle cellStyle) {
		Cell cell = row.createCell(column);
		cell.setCellValue(title);
		cell.setCellStyle(cellStyle);

		applyBorder(sheet, cellRangeAddress);

		return cell;
	}

	private void writeDebit(Sheet sheet, Row row) throws Exception {
		if (LOG.isDebugEnabled()) {
			LOG.debug("debit: {}", IntStream.range(0, row.getLastCellNum()).mapToObj(i -> {
				return getCellValue(row.getCell(i)).toString();
			}).collect(Collectors.joining(" | ")));
		}

		Row outputRow = sheet.createRow(sheet.getLastRowNum() + 1);

		applyBorder(sheet, outputRow.createCell(0)).setCellValue((Double) getCellValue(row.getCell(0)));
		applyBorder(sheet, setDateValue(outputRow.createCell(1), getCellValue(row.getCell(1))));
		applyBorder(sheet, setNumberValue(outputRow.createCell(2), getCellValue(row.getCell(4))));
	}

	private void writeCredit(Sheet sheet, Row row) throws Exception {
		if (LOG.isDebugEnabled()) {
			LOG.debug("credit: {}", IntStream.range(0, row.getLastCellNum()).mapToObj(i -> {
				return getCellValue(row.getCell(i)).toString();
			}).collect(Collectors.joining(" | ")));
		}

		Row outputRow = sheet.getRow(sheet.getLastRowNum());
		applyBorder(sheet, setNumberValue(outputRow.createCell(3), getCellValue(row.getCell(6))));
		applyBorder(sheet, setDateValue(outputRow.createCell(4), getCellValue(row.getCell(1))));
		applyBorder(sheet, outputRow.createCell(5)).setCellValue((Double) getCellValue(row.getCell(0)));
	}

	private Cell setNumberValue(Cell cell, Number number) {
		cell.setCellValue(number.doubleValue());
		cell.setCellStyle(numberStyle);

		return cell;
	}

	private Cell setDateValue(Cell cell, Date date) {
		cell.setCellValue(date);
		cell.setCellStyle(dateStyle);

		return cell;
	}

	@SuppressWarnings("unchecked")
	private <T> T getCellValue(Cell cell) {
		T cellValue = null;

		if (cell == null) {
			return (T) "";
		}
		switch (cell.getCellType()) {
			case STRING:
				cellValue = (T) cell.getRichStringCellValue().getString();
				break;
			case NUMERIC:
				if (DateUtil.isCellDateFormatted(cell)) {
					cellValue = (T) cell.getDateCellValue();
				} else {
					cellValue = (T) Double.valueOf(cell.getNumericCellValue());
				}
				break;
			case BOOLEAN:
				cellValue = (T) Boolean.valueOf(cell.getBooleanCellValue());
				break;
			case FORMULA:
				cellValue = (T) cell.getCellFormula();
				break;
			default:
				cellValue = (T) "";
		}

		return cellValue;
	}

	private static char toExcelIndex(int index) {
		return (char) (index + INDEX_A);
	}

	private static int toIndex(String excelIndex) {
		return IntStream.range(0, excelIndex.length()).map(i -> {
			return i * 65 + (excelIndex.charAt(i) - INDEX_A);
		}).sum();
	}
}
