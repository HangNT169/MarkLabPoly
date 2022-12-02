package com.poly.hangnt169.service;

import com.poly.hangnt169.model.Student;
import jakarta.servlet.http.HttpServletResponse;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.util.List;

public interface MainService {

    List<Student> readExcelLab(MultipartFile fileName) throws IOException;

    List<Student> readExcelQuiz(MultipartFile[] fileName) throws IOException;

    List<Student> readMark(MultipartFile fileName) throws IOException;

    void exportExcel(HttpServletResponse response,List<Student> listMark, List<Student> listLab, List<Student> listQuiz);

}
