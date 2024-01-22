package com.example.ICE_Trading;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

import java.io.BufferedWriter;
import java.io.FileInputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.SQLException;

@SpringBootApplication
public class IceTradingApplication {

	public static void main(String[] args) {

		SpringApplication.run(IceTradingApplication.class, args);
		String excelFilePath = "src/main/resources/ICE.xlsx";
		String jdbcUrl = "jdbc:h2:mem:test;MODE=MySQL";
		String username = "sa";
		String password = "password";

		try (Connection connection = DriverManager.getConnection(jdbcUrl, username, password);
			 FileInputStream excelFile = new FileInputStream(excelFilePath);
			 Workbook workbook = new XSSFWorkbook(excelFile)) {
			Sheet sheet = workbook.getSheetAt(0);
			createTable(connection);
			insertData(connection, sheet);
			System.out.println("ICE Trading Data imported successfully.");

		} catch ( Exception e) {
			e.printStackTrace();
		}


	}


	private static void createTable(Connection connection) throws SQLException {
		String createTableSQL = "CREATE TABLE IF NOT EXISTS ICE_WAREHOUSE ("
				+ "TRADE_DATE VARCHAR(255),"
				+ "HUB VARCHAR(255),"
				+ "PRODUCT VARCHAR(255),"
				+ "STRIP VARCHAR(255),"
				+ "CONTRACT VARCHAR(255),"
				+ "CONTRACT_TYPE VARCHAR(255),"
				+ "STRIKE VARCHAR(255),"
				+ "SETTLEMENT_PRICE FLOAT,"
				+ "NET_CHANGE FLOAT,"
				+ "EXPIRATION_DATE VARCHAR(255),"
				+ "PRODUCT_ID VARCHAR(255)"
				+ ")";
		try (PreparedStatement preparedStatement = connection.prepareStatement(createTableSQL)) {
			preparedStatement.executeUpdate();
		}
	}

	private static void insertData(Connection connection, Sheet sheet) throws SQLException, IOException {

		String insertSQL = "INSERT INTO ICE_WAREHOUSE (" +
				"TRADE_DATE," +
				" HUB," +
				" PRODUCT," +
				" STRIP," +
				" CONTRACT," +
				" CONTRACT_TYPE," +
				" STRIKE," +
				" SETTLEMENT_PRICE," +
				" NET_CHANGE," +
				" EXPIRATION_DATE," +
				" PRODUCT_ID) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)";

		try (PreparedStatement preparedStatement = connection.prepareStatement(insertSQL)) {

			for (Row row : sheet) {
				if (row.getRowNum() == 0) {
					continue;
				}

				preparedStatement.setString(1, getStringValue(row.getCell(0)));
				preparedStatement.setString(2, getStringValue(row.getCell(1)));
				preparedStatement.setString(3, getStringValue(row.getCell(2)));
				preparedStatement.setString(4, getStringValue(row.getCell(3)));
				preparedStatement.setString(5, getStringValue(row.getCell(4)));
				preparedStatement.setString(6, getStringValue(row.getCell(5)));
				preparedStatement.setString(7, getStringValue(row.getCell(6)));
				preparedStatement.setString(8, getStringValue(row.getCell(7)));
				preparedStatement.setString(9, getStringValue(row.getCell(8)));
				preparedStatement.setString(10, getStringValue(row.getCell(9)));
				preparedStatement.setString(11, getStringValue(row.getCell(10)));
				preparedStatement.executeUpdate();

			}
		}

	}

	private static String getStringValue(Cell cell) {
		if (cell == null) {
			return null;
		}

		switch (cell.getCellType()) {
			case STRING:
				return cell.getStringCellValue();
			case NUMERIC:
				return String.valueOf(cell.getNumericCellValue());
			case BOOLEAN:
				return String.valueOf(cell.getBooleanCellValue());
			default:
				return null;
		}

}
}



