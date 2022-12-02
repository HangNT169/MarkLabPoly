package com.poly.hangnt169.service.impl;

import com.poly.hangnt169.constant.Constants;
import com.poly.hangnt169.model.MarkStudent;
import com.poly.hangnt169.model.Student;
import com.poly.hangnt169.service.MainService;
import jakarta.servlet.http.HttpServletResponse;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.io.OutputStream;
import java.net.URLEncoder;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Objects;

import static com.poly.hangnt169.util.VNCharacterUtils.removeAccent;

@Service
public class MainServiceImpl implements MainService {

    @Override
    public List<Student> readExcelLab(MultipartFile fileName) throws IOException {
        if (fileName.isEmpty()) {
            return null;
        }
        List<Student> students = new ArrayList<>();
        // doc excel
        XSSFWorkbook workbook = new XSSFWorkbook(fileName.getInputStream());
        XSSFSheet worksheet = workbook.getSheetAt(0);

        for (int i = 1; i < worksheet.getPhysicalNumberOfRows(); i++) {
            students.add(getStudentFromExcelLab(worksheet, i));
        }
        return students;
    }

    @Override
    public List<Student> readExcelQuiz(MultipartFile[] fileName) throws IOException {
        if (fileName.length == 0 || fileName[0].getOriginalFilename().isEmpty()) {
            return null;
        }
        List<Student> students = new ArrayList<>();
        for (int i = 0; i < fileName.length; i++) {
            readAFileQuiz(students, fileName[i]);
        }
        return students;
    }

    @Override
    public List<Student> readMark(MultipartFile fileName) throws IOException {
        if (fileName.isEmpty()) {
            return null;
        }
        List<Student> students = new ArrayList<>();
        // doc excel
        XSSFWorkbook workbook = new XSSFWorkbook(fileName.getInputStream());
        XSSFSheet worksheet = workbook.getSheetAt(0);

        for (int i = 1; i < worksheet.getPhysicalNumberOfRows(); i++) {
            students.add(getStudentFromMark(worksheet, i));
        }
        return students;
    }

    @Override
    public void exportExcel(HttpServletResponse response, List<Student> listMark, List<Student> listLab, List<Student> listQuiz) {
        //Phan 1merge diem lab vao list tong
        //B1: Tao map
        Map<String, List<MarkStudent>> mapLabs = new HashMap<>();
        if (Objects.nonNull(listLab)) {
            for (Student s : listLab) {
                List<MarkStudent> markStudents = new ArrayList<>();
                if (Objects.nonNull(s.getLab1())) {
                    markStudents.add(new MarkStudent("lab1", s.getLab1()));
                }
                if (Objects.nonNull(s.getLab2())) {
                    markStudents.add(new MarkStudent("lab2", s.getLab2()));
                }
                if (Objects.nonNull(s.getLab3())) {
                    markStudents.add(new MarkStudent("lab3", s.getLab3()));
                }
                if (Objects.nonNull(s.getLab4())) {
                    markStudents.add(new MarkStudent("lab4", s.getLab4()));
                }
                if (Objects.nonNull(s.getLab5())) {
                    markStudents.add(new MarkStudent("lab5", s.getLab5()));
                }
                if (Objects.nonNull(s.getLab6())) {
                    markStudents.add(new MarkStudent("lab6", s.getLab6()));
                }
                if (Objects.nonNull(s.getLab7())) {
                    markStudents.add(new MarkStudent("lab7", s.getLab7()));
                }
                if (Objects.nonNull(s.getLab8())) {
                    markStudents.add(new MarkStudent("lab8", s.getLab8()));
                }
                if (Objects.nonNull(s.getQuiz1())) {
                    markStudents.add(new MarkStudent("quiz1", s.getQuiz1()));
                }
                if (Objects.nonNull(s.getQuiz2())) {
                    markStudents.add(new MarkStudent("quiz2", s.getQuiz2()));
                }
                if (Objects.nonNull(s.getQuiz3())) {
                    markStudents.add(new MarkStudent("quiz3", s.getQuiz3()));
                }
                if (Objects.nonNull(s.getQuiz4())) {
                    markStudents.add(new MarkStudent("quiz4", s.getQuiz4()));
                }
                if (Objects.nonNull(s.getQuiz5())) {
                    markStudents.add(new MarkStudent("quiz5", s.getQuiz5()));
                }
                if (Objects.nonNull(s.getQuiz6())) {
                    markStudents.add(new MarkStudent("quiz6", s.getQuiz6()));
                }
                if (Objects.nonNull(s.getQuiz7())) {
                    markStudents.add(new MarkStudent("quiz7", s.getQuiz7()));
                }
                if (Objects.nonNull(s.getQuiz8())) {
                    markStudents.add(new MarkStudent("quiz8", s.getQuiz8()));
                }
                if (Objects.nonNull(s.getAssGD1())) {
                    markStudents.add(new MarkStudent("ass1", s.getAssGD1()));
                }
                if (Objects.nonNull(s.getAssGD2())) {
                    markStudents.add(new MarkStudent("ass2", s.getAssGD2()));
                }
                mapLabs.put(s.getEmail().toUpperCase(), markStudents);
            }
            // B2: Check map.
            for (int i = 0; i < listMark.size(); i++) {
                Student student = listMark.get(i);
                List<MarkStudent> listsMarkStudent = mapLabs.get(student.getEmail().toUpperCase());
                // ton tai sinh vien trong file diem lab
                if (Objects.nonNull(listsMarkStudent)) {
                    // update student trong listMark
                    for (MarkStudent markStudent : listsMarkStudent) {
                        if (markStudent.getTenDiem().equalsIgnoreCase("lab1")) {
                            student.setLab1(markStudent.getDiem());
                        }
                        if (markStudent.getTenDiem().equalsIgnoreCase("lab2")) {
                            student.setLab2(markStudent.getDiem());
                        }
                        if (markStudent.getTenDiem().equalsIgnoreCase("lab3")) {
                            student.setLab3(markStudent.getDiem());
                        }
                        if (markStudent.getTenDiem().equalsIgnoreCase("lab4")) {
                            student.setLab4(markStudent.getDiem());
                        }
                        if (markStudent.getTenDiem().equalsIgnoreCase("lab5")) {
                            student.setLab5(markStudent.getDiem());
                        }
                        if (markStudent.getTenDiem().equalsIgnoreCase("lab6")) {
                            student.setLab6(markStudent.getDiem());
                        }
                        if (markStudent.getTenDiem().equalsIgnoreCase("lab7")) {
                            student.setLab7(markStudent.getDiem());
                        }
                        if (markStudent.getTenDiem().equalsIgnoreCase("lab8")) {
                            student.setLab8(markStudent.getDiem());
                        }
                        if (markStudent.getTenDiem().equalsIgnoreCase("quiz1")) {
                            student.setQuiz1(markStudent.getDiem());
                        }
                        if (markStudent.getTenDiem().equalsIgnoreCase("quiz2")) {
                            student.setQuiz2(markStudent.getDiem());
                        }
                        if (markStudent.getTenDiem().equalsIgnoreCase("quiz3")) {
                            student.setQuiz3(markStudent.getDiem());
                        }
                        if (markStudent.getTenDiem().equalsIgnoreCase("quiz4")) {
                            student.setQuiz4(markStudent.getDiem());
                        }
                        if (markStudent.getTenDiem().equalsIgnoreCase("quiz5")) {
                            student.setQuiz5(markStudent.getDiem());
                        }
                        if (markStudent.getTenDiem().equalsIgnoreCase("quiz6")) {
                            student.setQuiz6(markStudent.getDiem());
                        }
                        if (markStudent.getTenDiem().equalsIgnoreCase("quiz7")) {
                            student.setQuiz7(markStudent.getDiem());
                        }
                        if (markStudent.getTenDiem().equalsIgnoreCase("quiz8")) {
                            student.setQuiz8(markStudent.getDiem());
                        }
                        if (markStudent.getTenDiem().equalsIgnoreCase("ass1")) {
                            student.setAssGD1(markStudent.getDiem());
                        }
                        if (markStudent.getTenDiem().equalsIgnoreCase("ass2")) {
                            student.setAssGD2(markStudent.getDiem());
                        }
                    }
                    listMark.set(i, student);
                }
            }
        }

        // Phan 2 merge diem quiz vao list tong
        Map<String, List<MarkStudent>> mapQuiz = new HashMap<>();
        if (Objects.nonNull(listQuiz)) {
            for (Student s : listQuiz) {
                List<MarkStudent> markStudents = new ArrayList<>();
                if (Objects.nonNull(s.getQuiz1())) {
                    markStudents.add(new MarkStudent("quiz1", s.getQuiz1()));
                }
                if (Objects.nonNull(s.getQuiz2())) {
                    markStudents.add(new MarkStudent("quiz2", s.getQuiz2()));
                }
                if (Objects.nonNull(s.getQuiz3())) {
                    markStudents.add(new MarkStudent("quiz3", s.getQuiz3()));
                }
                if (Objects.nonNull(s.getQuiz4())) {
                    markStudents.add(new MarkStudent("quiz4", s.getQuiz4()));
                }
                if (Objects.nonNull(s.getQuiz5())) {
                    markStudents.add(new MarkStudent("quiz5", s.getQuiz5()));
                }
                if (Objects.nonNull(s.getQuiz6())) {
                    markStudents.add(new MarkStudent("quiz6", s.getQuiz6()));
                }
                if (Objects.nonNull(s.getQuiz7())) {
                    markStudents.add(new MarkStudent("quiz7", s.getQuiz7()));
                }
                if (Objects.nonNull(s.getQuiz8())) {
                    markStudents.add(new MarkStudent("quiz8", s.getQuiz8()));
                }
                mapQuiz.put(s.getEmail().toUpperCase(), markStudents);
            }
            for (int i = 0; i < listMark.size(); i++) {
                Student student = listMark.get(i);
                List<MarkStudent> listsMarkStudent = mapQuiz.get(student.getEmail().toUpperCase());
                // ton tai sinh vien trong file diem quiz
                if (Objects.nonNull(listsMarkStudent)) {
                    // update student trong listMark
                    for (MarkStudent markStudent : listsMarkStudent) {
                        if (markStudent.getTenDiem().equalsIgnoreCase("quiz1")) {
                            student.setQuiz1(markStudent.getDiem());
                        }
                        if (markStudent.getTenDiem().equalsIgnoreCase("quiz2")) {
                            student.setQuiz2(markStudent.getDiem());
                        }
                        if (markStudent.getTenDiem().equalsIgnoreCase("quiz3")) {
                            student.setQuiz3(markStudent.getDiem());
                        }
                        if (markStudent.getTenDiem().equalsIgnoreCase("quiz4")) {
                            student.setQuiz4(markStudent.getDiem());
                        }
                        if (markStudent.getTenDiem().equalsIgnoreCase("quiz5")) {
                            student.setQuiz5(markStudent.getDiem());
                        }
                        if (markStudent.getTenDiem().equalsIgnoreCase("quiz6")) {
                            student.setQuiz6(markStudent.getDiem());
                        }
                        if (markStudent.getTenDiem().equalsIgnoreCase("quiz7")) {
                            student.setQuiz7(markStudent.getDiem());
                        }
                        if (markStudent.getTenDiem().equalsIgnoreCase("quiz8")) {
                            student.setQuiz8(markStudent.getDiem());
                        }
                    }
                    listMark.set(i, student);
                }
            }
        }

        // tao file excel luc in ra
        downloadBillingApartmentByMonth(response, listMark);
    }

    public void downloadBillingApartmentByMonth(HttpServletResponse response, List<Student> students) {
        try {
            String fileName = URLEncoder.encode(String.format("MarkStudent.xlsx"), "UTF-8");
            response.setContentType("application/ms-excel; charset=UTF-8");
            String headerValue = String.format("attachment; filename=\"%s\"", fileName);
            response.setHeader(Constants.HEADER_KEY, headerValue);
            writeBillingApartmentByMonth(response, students);
        } catch (Exception ex) {
            ex.printStackTrace(System.out);
        }
    }

    private void writeBillingApartmentByMonth(HttpServletResponse response, List<Student> students) {
        try (XSSFWorkbook workbook = new XSSFWorkbook(); OutputStream os = response.getOutputStream()) {
            createExcelMark(workbook, students);
            workbook.write(os);
        } catch (Exception ex) {
            ex.printStackTrace(System.out);
        }
    }

    private void createExcelMark(XSSFWorkbook workbook, List<Student> students) {
        Sheet sheet = workbook.createSheet("Mark Student");
        sheet.setColumnWidth(1, 25 * 256);
        sheet.setColumnWidth(2, 25 * 256);
        sheet.setColumnWidth(3, 25 * 256);
        sheet.setColumnWidth(4, 25 * 256);

        // Title
        CellStyle style = workbook.createCellStyle();
        style.setWrapText(true);
        XSSFFont font = workbook.createFont();
        font.setFontName("Arial");
        font.setBold(true);
        style.setFont(font);

        // Table
        CellStyle headerStyle = workbook.createCellStyle();
        headerStyle.setWrapText(true);

        font = workbook.createFont();
        font.setFontName("Times New Roman");
        headerStyle.setFont(font);
        headerStyle.setBorderBottom(BorderStyle.THIN);
        headerStyle.setBorderTop(BorderStyle.THIN);
        headerStyle.setBorderRight(BorderStyle.THIN);
        headerStyle.setBorderLeft(BorderStyle.THIN);

        // Set headers
        Row row = sheet.createRow(0);
        Cell headerCell = row.createCell(0);
        headerCell.setCellValue("#");
        headerCell.setCellStyle(headerStyle);

        headerCell = row.createCell(1);
        headerCell.setCellValue("Mã sinh viên");
        headerCell.setCellStyle(headerStyle);

        headerCell = row.createCell(2);
        headerCell.setCellValue("Họ và tên");
        headerCell.setCellStyle(headerStyle);

        headerCell = row.createCell(3);
        headerCell.setCellValue("Đánh giá Assignment GĐ 1 (10%)");
        headerCell.setCellStyle(headerStyle);

        headerCell = row.createCell(4);
        headerCell.setCellValue("Đánh giá Assignment GĐ 2 (10%)");
        headerCell.setCellStyle(headerStyle);

        headerCell = row.createCell(5);
        headerCell.setCellValue("Lab 1 (3.5%)");
        headerCell.setCellStyle(headerStyle);

        headerCell = row.createCell(6);
        headerCell.setCellValue("Lab 2 (3.5%)");
        headerCell.setCellStyle(headerStyle);

        headerCell = row.createCell(7);
        headerCell.setCellValue("Lab 3 (3.5%)");
        headerCell.setCellStyle(headerStyle);

        headerCell = row.createCell(8);
        headerCell.setCellValue("Lab 4 (3.5%)");
        headerCell.setCellStyle(headerStyle);

        headerCell = row.createCell(9);
        headerCell.setCellValue("Lab 5 (3.5%)");
        headerCell.setCellStyle(headerStyle);

        headerCell = row.createCell(10);
        headerCell.setCellValue("Lab 6 (3.5%)");
        headerCell.setCellStyle(headerStyle);

        headerCell = row.createCell(11);
        headerCell.setCellValue("Lab 7 (3.5%)");
        headerCell.setCellStyle(headerStyle);

        headerCell = row.createCell(12);
        headerCell.setCellValue("Lab 8 (3.5%)");
        headerCell.setCellStyle(headerStyle);

        headerCell = row.createCell(13);
        headerCell.setCellValue("Quiz 1 (1.5%)");
        headerCell.setCellStyle(headerStyle);

        headerCell = row.createCell(14);
        headerCell.setCellValue("Quiz 2 (1.5%)");
        headerCell.setCellStyle(headerStyle);

        headerCell = row.createCell(15);
        headerCell.setCellValue("Quiz 3 (1.5%)");
        headerCell.setCellStyle(headerStyle);

        headerCell = row.createCell(16);
        headerCell.setCellValue("Quiz 4 (1.5%)");
        headerCell.setCellStyle(headerStyle);

        headerCell = row.createCell(17);
        headerCell.setCellValue("Quiz 5 (1.5%)");
        headerCell.setCellStyle(headerStyle);

        headerCell = row.createCell(18);
        headerCell.setCellValue("Quiz 6 (1.5%)");
        headerCell.setCellStyle(headerStyle);

        headerCell = row.createCell(19);
        headerCell.setCellValue("Quiz 7 (1.5%)");
        headerCell.setCellStyle(headerStyle);

        headerCell = row.createCell(20);
        headerCell.setCellValue("Quiz 8 (1.5%)");
        headerCell.setCellStyle(headerStyle);

        // Set value
        style = workbook.createCellStyle();
        style.setWrapText(true);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);

        //Doc noi dung
        int i = 1;

        for (Student student : students) {
            int index = i;
            row = sheet.createRow(i++);
            Cell cell = row.createCell(0);

            cell.setCellStyle(style);
            cell.setCellValue(index++);

            cell = row.createCell(1);
            cell.setCellStyle(style);
            cell.setCellValue(student.getMssv());

            cell = row.createCell(2);
            cell.setCellStyle(style);
            cell.setCellValue(student.getTen());

            cell = row.createCell(3);
            cell.setCellStyle(style);
            cell.setCellValue(student.getAssGD1());

            cell = row.createCell(4);
            cell.setCellStyle(style);
            cell.setCellValue(student.getAssGD2());

            cell = row.createCell(5);
            cell.setCellStyle(style);
            cell.setCellValue(student.getLab1());

            cell = row.createCell(6);
            cell.setCellStyle(style);
            cell.setCellValue(student.getLab2());

            cell = row.createCell(7);
            cell.setCellStyle(style);
            cell.setCellValue(student.getLab3());

            cell = row.createCell(8);
            cell.setCellStyle(style);
            cell.setCellValue(student.getLab4());

            cell = row.createCell(9);
            cell.setCellStyle(style);
            cell.setCellValue(student.getLab5());

            cell = row.createCell(10);
            cell.setCellStyle(style);
            cell.setCellValue(student.getLab6());

            cell = row.createCell(11);
            cell.setCellStyle(style);
            cell.setCellValue(student.getLab7());

            cell = row.createCell(12);
            cell.setCellStyle(style);
            cell.setCellValue(student.getLab8());

            cell = row.createCell(13);
            cell.setCellStyle(style);
            cell.setCellValue(student.getQuiz1());

            cell = row.createCell(14);
            cell.setCellStyle(style);
            cell.setCellValue(student.getQuiz2());

            cell = row.createCell(15);
            cell.setCellStyle(style);
            cell.setCellValue(student.getQuiz3());

            cell = row.createCell(16);
            cell.setCellStyle(style);
            cell.setCellValue(student.getQuiz4());

            cell = row.createCell(17);
            cell.setCellStyle(style);
            cell.setCellValue(student.getQuiz5());

            cell = row.createCell(18);
            cell.setCellStyle(style);
            cell.setCellValue(student.getQuiz6());

            cell = row.createCell(19);
            cell.setCellStyle(style);
            cell.setCellValue(student.getQuiz7());

            cell = row.createCell(20);
            cell.setCellStyle(style);
            cell.setCellValue(student.getQuiz8());
        }
    }

    private void readAFileQuiz(List<Student> students, MultipartFile fileName) throws IOException {
        XSSFWorkbook workbook = new XSSFWorkbook(fileName.getInputStream());
        XSSFSheet worksheet = workbook.getSheetAt(0);
        String nameExcel = fileName.getOriginalFilename();
        for (int i = 1; i < worksheet.getPhysicalNumberOfRows(); i++) {
            XSSFRow row = worksheet.getRow(i);
            String email = row.getCell(1).getStringCellValue() + "@fpt.edu.vn";
            Student duplicateStudent = checkDuplicate(students, email);
            // trung doc data roi => chi update
            if (Objects.nonNull(duplicateStudent)) {
                // update data
                readQuiz(duplicateStudent, row, nameExcel);
                // update vao list
                int indexFind = indexStudent(students, duplicateStudent);
                students.set(indexFind, duplicateStudent);
            } else {
                students.add(getStudentFromExcelQuiz(worksheet, i, nameExcel));
            }

        }
    }

    private Student checkDuplicate(List<Student> students, String email) {
        for (Student s : students) {
            if (s.getEmail().equalsIgnoreCase(email)) {
                return s;
            }
        }
        return null;
    }

    private int indexStudent(List<Student> students, Student student) {
        for (int i = 0; i < students.size(); i++) {
            if (students.get(i).getEmail().equalsIgnoreCase(student.getEmail())) {
                return i;
            }
        }
        return -1;
    }

    private Student getStudentFromMark(XSSFSheet worksheet, int index) {
        Student student = new Student();
        XSSFRow row = worksheet.getRow(index);
        String name = removeAccent(row.getCell(2).toString());
        String mssv = row.getCell(1).toString();
        String email = createEmail(name, mssv);
        student.setMssv(mssv);
        student.setTen(name);
        student.setEmail(email);
        student.setAssGD1(row.getCell(3).toString());
        student.setAssGD2(row.getCell(4).toString());
        student.setLab1(row.getCell(5).toString());
        student.setLab2(row.getCell(6).toString());
        student.setLab3(row.getCell(7).toString());
        student.setLab4(row.getCell(8).toString());
        student.setLab5(row.getCell(9).toString());
        student.setLab6(row.getCell(10).toString());
        student.setLab7(row.getCell(11).toString());
        student.setLab8(row.getCell(12).toString());
        student.setQuiz1(row.getCell(13).toString());
        student.setQuiz2(row.getCell(14).toString());
        student.setQuiz3(row.getCell(15).toString());
        student.setQuiz4(row.getCell(16).toString());
        student.setQuiz5(row.getCell(17).toString());
        student.setQuiz6(row.getCell(18).toString());
        student.setQuiz7(row.getCell(19).toString());
        student.setQuiz8(row.getCell(20).toString());
        return student;
    }

    private Student getStudentFromExcelQuiz(XSSFSheet worksheet, int index, String nameExcel) {
        Student student = new Student();
        XSSFRow row = worksheet.getRow(index);
        String name = removeAccent(row.getCell(0).toString()).substring(5).trim();
        String email = row.getCell(1).toString() + "@fpt.edu.vn";
        readQuiz(student, row, nameExcel);
        student.setTen(name);
        student.setEmail(email);
        return student;
    }

    private void readQuiz(Student student, XSSFRow row, String nameExcel) {
        if ("Quiz_1_results.xlsx".equalsIgnoreCase(nameExcel)) {
            student.setQuiz1(row.getCell(2).toString());
        }
        if ("Quiz_2_results.xlsx".equalsIgnoreCase(nameExcel)) {
            student.setQuiz2(row.getCell(2).toString());
        }
        if ("Quiz_3_results.xlsx".equalsIgnoreCase(nameExcel)) {
            student.setQuiz3(row.getCell(2).toString());
        }
        if ("Quiz_4_results.xlsx".equalsIgnoreCase(nameExcel)) {
            student.setQuiz4(row.getCell(2).toString());
        }
        if ("Quiz_5_results.xlsx".equalsIgnoreCase(nameExcel)) {
            student.setQuiz5(row.getCell(2).toString());
        }
        if ("Quiz_6_results.xlsx".equalsIgnoreCase(nameExcel)) {
            student.setQuiz6(row.getCell(2).toString());
        }
        if ("Quiz_7_results.xlsx".equalsIgnoreCase(nameExcel)) {
            student.setQuiz7(row.getCell(2).toString());
        }
        if ("Quiz_8_results.xlsx".equalsIgnoreCase(nameExcel)) {
            student.setQuiz8(row.getCell(2).toString());
        }
    }

    private String createEmail(String name, String rollNumber) {
        String result;
        String temp = "";
        String[] strs = name.split(" ");
        result = strs[strs.length - 1];
        for (int i = 0; i < strs.length - 1; i++) {
            temp += strs[i].charAt(0);
        }
        result = result + temp + rollNumber.toUpperCase() + "@fpt.edu.vn";
        return result;
    }

    private Student getStudentFromExcelLab(XSSFSheet worksheet, int index) {
        Student student = new Student();
        XSSFRow row = worksheet.getRow(index);
        Row headerRow = worksheet.getRow(0);
        Map<String, Integer> listHeader = new HashMap<>();
        for (int i = 0; i < headerRow.getLastCellNum(); i++) {
            listHeader.put(String.valueOf(headerRow.getCell(i)), i);
        }
        List<String> headers = new ArrayList<>(listHeader.keySet());
        for (String header : headers) {
            if (header.equalsIgnoreCase("Email Address")) {
                student.setEmail(row.getCell(listHeader.get(header)).toString());
            }
            if (header.equalsIgnoreCase("Lab 1")) {
                student.setLab1(row.getCell(listHeader.get(header)).toString());
            }
            if (header.equalsIgnoreCase("Lab 2")) {
                student.setLab2(row.getCell(listHeader.get(header)).toString());
            }
            if (header.equalsIgnoreCase("Lab 3")) {
                student.setLab3(row.getCell(listHeader.get(header)).toString());
            }
            if (header.equalsIgnoreCase("Lab 4")) {
                student.setLab4(row.getCell(listHeader.get(header)).toString());
            }
            if (header.equalsIgnoreCase("Lab 5")) {
                student.setLab5(row.getCell(listHeader.get(header)).toString());
            }
            if (header.equalsIgnoreCase("Lab 6")) {
                student.setLab6(row.getCell(listHeader.get(header)).toString());
            }
            if (header.equalsIgnoreCase("Lab 7")) {
                student.setLab7(row.getCell(listHeader.get(header)).toString());
            }
            if (header.equalsIgnoreCase("Lab 8")) {
                student.setLab8(row.getCell(listHeader.get(header)).toString());
            }
            if (header.equalsIgnoreCase("Quiz 1")) {
                student.setQuiz1(row.getCell(listHeader.get(header)).toString());
            }
            if (header.equalsIgnoreCase("Quiz 2")) {
                student.setQuiz2(row.getCell(listHeader.get(header)).toString());
            }
            if (header.equalsIgnoreCase("Quiz 3")) {
                student.setQuiz3(row.getCell(listHeader.get(header)).toString());
            }
            if (header.equalsIgnoreCase("Quiz 4")) {
                student.setQuiz4(row.getCell(listHeader.get(header)).toString());
            }
            if (header.equalsIgnoreCase("Quiz 5")) {
                student.setQuiz5(row.getCell(listHeader.get(header)).toString());
            }
            if (header.equalsIgnoreCase("Quiz 6")) {
                student.setQuiz6(row.getCell(listHeader.get(header)).toString());
            }
            if (header.equalsIgnoreCase("Quiz 7")) {
                student.setQuiz7(row.getCell(listHeader.get(header)).toString());
            }
            if (header.equalsIgnoreCase("Quiz 8")) {
                student.setQuiz8(row.getCell(listHeader.get(header)).toString());
            }
            if (header.equalsIgnoreCase("Assignment 1 ")) {
                student.setAssGD1(row.getCell(listHeader.get(header)).toString());
            }
            if (header.equalsIgnoreCase("Assignment 2 ")) {
                student.setAssGD2(row.getCell(listHeader.get(header)).toString());
            }
        }
        return student;
    }

}