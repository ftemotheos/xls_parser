package com.ftemotheos;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import java.io.IOException;
import java.sql.SQLException;
import java.text.ParseException;

public class Main {
    public static void main(String[] args) {
        XlsParser parser = new XlsParser("jdbc:postgresql://localhost:5432/xls_storage",
                "user_timofey", "12345");
        try {
            parser.parseXls("Прайс шины.xls", "ACBA-2017-02-22-tyres");
        } catch (IOException e) {
            e.printStackTrace();
        } catch (InvalidFormatException e) {
            e.printStackTrace();
        } catch (SQLException e) {
            e.printStackTrace();
        } catch (ParseException e) {
            e.printStackTrace();
        }
    }
}
