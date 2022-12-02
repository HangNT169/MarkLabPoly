package com.poly.hangnt169.controller;

import com.poly.hangnt169.model.Student;
import com.poly.hangnt169.service.impl.MainServiceImpl;
import jakarta.servlet.http.HttpServletResponse;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.util.List;

@Controller
public class MainController {

    @Autowired
    private MainServiceImpl mainServiceImpl;

    @GetMapping("/index")
    public String index() {
        return "index.html";
    }

    @PostMapping("/import-mark")
    public String exportMark(@RequestParam("fileLayDiemLab") MultipartFile fileLayDiemLab,
                              @RequestParam("fileLayDiemQuiz") MultipartFile[] fileLayDiemQuiz,
                              @RequestParam("fileLayDiem") MultipartFile fileLayDiem, HttpServletResponse response)
            throws IOException {
        List<Student> listLab = mainServiceImpl.readExcelLab(fileLayDiemLab);
        System.out.println("ListLab");
        System.out.println(listLab.size());
        listLab.forEach(s -> System.out.println(s.toString()));
        List<Student> listQuiz = mainServiceImpl.readExcelQuiz(fileLayDiemQuiz);
//        System.out.println("listQuiz");
//        System.out.println(listQuiz.size());
//        listQuiz.forEach(s -> System.out.println(s.toString()));
        List<Student> listMark = mainServiceImpl.readMark(fileLayDiem);
//        System.out.println("listMark");
//        System.out.println(listMark.size());
//        listMark.forEach(s -> System.out.println(s.toString()));
        mainServiceImpl.exportExcel(response, listMark, listLab, listQuiz);
//        System.out.println("listMark after");
//        System.out.println(listMark.size());
//        listMark.forEach(s -> System.out.println(s.toString()));
        return "index.html";
    }
}