package com.support.service;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;

public class ExcelService {
    private String inputFilePath = "/home/decimal/Downloads/DVES_Ticket_Report14.xlsx";
    private String outputFilePath = "/home/decimal/Documents/14-04-2025.xlsx";
    private String grid = "/home/decimal/Downloads/Daily Report updated Grid.xlsx";

    public void monthlyReportCreation(String issueType){
        try (FileInputStream inputStream = new FileInputStream(inputFilePath);
             Workbook workbook = new XSSFWorkbook(inputStream);
             Workbook newWorkbook = new XSSFWorkbook()) {

            Sheet originalSheet = workbook.getSheetAt(0);
            Sheet newSheet = newWorkbook.createSheet("KPI Report");

            String[] headers = {
                    "Issue Key", "Project Name", "Client", "English Name", "Reporters EmailId","Feasible Status ","Active PM/BUG", "Status", "Assignee",
                    "Created Date", "Updated Date", "Components", "Labels", "Tickets Aging",
                    "CSR", "Status CSR", "Assignee CSR",
                    "Created Date CSR", "Updated Date CSR", "DevOps", "Status DevOps", "Assignee DevOps", "Created Date DevOps", "Updated Date DevOps",
                    "Infra", "Status Infra", "Assignee Infra", "Created Date Infra", "Updated Date Infra", "L2", "Status L2", "Assignee L2",
                    "Created Date L2", "Updated Date L2", "Developer", "Status Developer", "Assignee Developer", "Created Date Developer", "Updated Date Developer",
                    "BUG", "Status Bug", "Assignee Bug", "Created Date Bug", "Updated Date Bug", "PM", "Status PM", "Assignee PM",
                    "Created Date PM", "Updated Date PM"
            };

            Row headerRow = newSheet.createRow(0);

            for (int i = 0; i < headers.length; i++) {
                Cell headerCell = headerRow.createCell(i);
                headerCell.setCellValue(headers[i]);
            }
            for (Row row:originalSheet){
                Cell key = row.getCell(0);
                Cell issuetype= row.getCell(1);
                Cell projectName= row.getCell(2);
            }

        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }
}
