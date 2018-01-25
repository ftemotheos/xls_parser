package com.ftemotheos;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.IOException;
import java.sql.*;
import java.text.ParseException;
import java.util.*;

public class XlsParser {

    public static final String FILE_SEPARATOR = System.getProperty("file.separator");

    public static final String FILE_PATH = "src" + FILE_SEPARATOR + "main" + FILE_SEPARATOR + "resources" + FILE_SEPARATOR;

    public static final String[] TABLE_HEADER = new String[] {"Тип шины", "Сезон", "Название", "Остаток", "Цена", "Страна", "Год"};

    public static final String DDL_STATEMENT =
            "DROP TABLE IF EXISTS tires;" +
            "DROP TYPE IF EXISTS TIRES_TYPE;" +
            "CREATE TYPE TIRES_TYPE AS ENUM ('CAR','BIKE','FREIGHT','AGRICULTURAL','INDUSTRIAL');" +
            "CREATE TABLE IF NOT EXISTS tires (" +
                    "id serial PRIMARY KEY," +
                    "tires_type TIRES_TYPE," +
                    "season VARCHAR(10)," +
                    "width VARCHAR(3)," +
                    "height VARCHAR(2)," +
                    "diameter VARCHAR(2)," +
                    "brand_name VARCHAR(25)," +
                    "model_name VARCHAR(50)," +
                    "weight_index VARCHAR(2)," +
                    "speed_index VARCHAR(1)," +
                    "strengthened VARCHAR(2)," +
                    "is_studded BOOLEAN," +
                    "additional VARCHAR(2)," +
                    "remainder VARCHAR(15)," +
                    "price VARCHAR(20)," +
                    "country VARCHAR(25)," +
                    "production_year DATE" +
            ");";

    public static final String DML_STATEMENT =
            "INSERT INTO tires(" +
                    "tires_type, season, width, height, diameter, brand_name, model_name, " +
                    "weight_index, speed_index, strengthened, is_studded, " +
                    "additional, remainder, price, country, production_year" +
            ") VALUES (CAST(? AS TIRES_TYPE),?,?,?,?,?,?,?,?,?,?,?,?,?,?,?);";

    private String url;

    private String user;

    private String password;

    public XlsParser(String url, String user, String password) {
        this.url = url;
        this.user = user;
        this.password = password;
    }

    public void parseXls(String fileName, String sheetName) throws IOException, InvalidFormatException, SQLException, ParseException {

        Connection con = DriverManager.getConnection(url, user, password);
        PreparedStatement pst = con.prepareStatement(DDL_STATEMENT);
        pst.execute();

        Workbook wb = WorkbookFactory.create(new File(FILE_PATH + fileName));

        Sheet sheet = wb.getSheet(sheetName);

        DataFormatter formatter = new DataFormatter();

        int rowStart = sheet.getFirstRowNum();
        int rowEnd = sheet.getLastRowNum();
        int columnsAmount = TABLE_HEADER.length;

        boolean headerPassed = false;

        for (int rowNum = rowStart; rowNum <= rowEnd; rowNum++) {

            Row row = sheet.getRow(rowNum);

            if (row == null) {
                continue;
            }

            String[] entry = new String[16];
            Arrays.fill(entry, "");

            TiresType tiresType = TiresType.CAR;

            int shift = 0;

            boolean headerFound = true;

            for (int colNum = 0; colNum < columnsAmount; colNum++) {

                Cell cell = row.getCell(colNum, Row.CREATE_NULL_AS_BLANK);

                String textCellValue = formatter.formatCellValue(cell);

                if (!headerPassed) {
                    headerFound = headerFound && textCellValue.equalsIgnoreCase(TABLE_HEADER[colNum]);
                }
                else {
                    if (colNum == 0) {
                        if (textCellValue.matches("[Мм][Оо][Тт][Оо]")) {
                            tiresType = TiresType.BIKE;
                        }
                        else if (textCellValue.matches("[Гг][Рр][Уу][Зз][Оо][Вв].*")) {
                            tiresType = TiresType.FREIGHT;
                        }
                        else if (textCellValue.matches("[Сс]/[Хх]")) {
                            tiresType = TiresType.AGRICULTURAL;
                        }
                        else if (textCellValue.matches("[Ии][Нн][Дд][Уу][Сс][Тт][Рр].*}")) {
                            tiresType = TiresType.INDUSTRIAL;
                        }
                    }
                    if (colNum == 6) {
                        if (textCellValue.matches(".*[0-9][0-9]")) {
                            int length = textCellValue.length();
                            textCellValue = "20" + textCellValue.substring(length - 2, length) + "-1-1";
                        }
                        else {
                            textCellValue = "1970-1-1";
                        }
                    }
                    if (colNum == 2) {
                        parseAbbreviation(entry, textCellValue.trim());
                        shift = 9;
                    }
                    else {
                        entry[colNum + shift] = textCellValue;
                    }
                }
            }
            if (!headerPassed) {
                headerPassed = headerFound;
            }
            else {
                System.out.println(Arrays.asList(entry));
                pst = con.prepareStatement(DML_STATEMENT);
                pst.setString(1, tiresType.toString());
                pst.setString(2, entry[1]);
                pst.setString(3, entry[2]);
                pst.setString(4, entry[3]);
                pst.setString(5, entry[4]);
                pst.setString(6, entry[5]);
                pst.setString(7, entry[6]);
                pst.setString(8, entry[7]);
                pst.setString(9, entry[8]);
                pst.setString(10, entry[9]);
                pst.setObject(11, Boolean.valueOf(entry[10]));
                pst.setString(12, entry[11]);
                pst.setString(13, entry[12]);
                pst.setString(14, entry[13]);
                pst.setString(15, entry[14]);
                pst.setDate(16, java.sql.Date.valueOf(entry[15]));
                pst.executeUpdate();
            }
        }
    }

    private void parseAbbreviation(String[] entry, String text) {

        int startIndex = 0;
        int currentIndex = text.indexOf(" ");

        String token;

        if (currentIndex != -1) {
            token = text.substring(startIndex, currentIndex);
            if (token.contains("/")) {
                int index = token.indexOf("/");
                entry[2] = token.substring(0, index);
                entry[3] = token.substring(index + 1);
            }
        }
        startIndex = currentIndex + 1;
        currentIndex = text.indexOf(" ", startIndex);
        if (currentIndex != -1) {
            token = text.substring(startIndex, currentIndex);
            if (token.startsWith("R")) {
                if (token.endsWith("С")) {
                    entry[9] = "XL";
                    token = token.substring(0, currentIndex - startIndex - 1);
                }
                entry[4] = token.substring(1);
            }
        }
        startIndex = currentIndex + 1;
        currentIndex = text.indexOf(" ", startIndex);
        if (currentIndex != -1) {
            token = text.substring(startIndex, currentIndex);
            entry[5] = token;
        }
        startIndex = currentIndex + 1;
        token = text.substring(startIndex);

        int endIndex = token.length();

        if (token.endsWith(" FR")) {
            currentIndex = endIndex - 3;
            entry[11] = token.substring(currentIndex + 1, endIndex);
            endIndex = currentIndex;
            token = token.substring(0, endIndex);
        }
        if (token.endsWith(" шип")) {
            entry[10] = "true";
            endIndex = endIndex - 4;
            token = token.substring(0, endIndex);
        }
        else {
            entry[10] = "false";
        }
        if (token.endsWith(" XL")) {
            currentIndex = endIndex - 3;
            entry[9] = token.substring(currentIndex + 1, endIndex);
            endIndex = currentIndex;
            token = token.substring(0, endIndex);
            endIndex = token.length();
        }
        if (token.matches(".* [0-9][0-9][A-Z]")) {
            currentIndex = endIndex - 1;
            entry[8] = token.substring(currentIndex, endIndex);
            endIndex = currentIndex;
            token = token.substring(0, endIndex);
            endIndex = token.length();
            currentIndex = endIndex - 3;
            entry[7] = token.substring(currentIndex + 1, endIndex);
            endIndex = currentIndex;
            token = token.substring(0, endIndex);
        }
        if (token.length() != 0) {
            entry[6] = token;
        }
    }

}
