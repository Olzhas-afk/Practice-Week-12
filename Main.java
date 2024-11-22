package org.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.List;

public class Main {
    public static void main(String[] args) {
        String filePath = "C:\\Users\\Asus\\IdeaProjects\\untitled\\src\\students.xlsx";
        List<Student> students = new ArrayList<>();

        try (FileInputStream fis = new FileInputStream(new File(filePath));
             Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = workbook.getSheet("Sheet1");
            for (Row row : sheet) {
                if (row.getRowNum() == 0) continue;

                try {
                    Cell nameCell = row.getCell(2);
                    Cell currentScholarshipCell = row.getCell(4);
                    Cell newScholarshipCell = row.getCell(7);

                    if (nameCell == null || currentScholarshipCell == null || newScholarshipCell == null) {
                        System.out.println("Ошибка: Пустая ячейка в строке " + (row.getRowNum() + 1));
                        continue;
                    }

                    String name = nameCell.getStringCellValue();
                    double currentScholarship = currentScholarshipCell.getNumericCellValue();
                    double newScholarship = newScholarshipCell.getNumericCellValue();

                    students.add(new Student(name, currentScholarship, newScholarship));
                } catch (Exception e) {
                    System.out.println("Ошибка в строке: " + (row.getRowNum() + 1));
                    e.printStackTrace();
                }
            }
        } catch (Exception e) {
            System.out.println("Ошибка при чтении файла: " + e.getMessage());
            e.printStackTrace();
        }

        for (Student student : students) {
            System.out.printf("Name: %s, Scholarship Increase: %.2f%n",
                    student.getName(),
                    student.getScholarshipIncrease());
        }
    }
}
