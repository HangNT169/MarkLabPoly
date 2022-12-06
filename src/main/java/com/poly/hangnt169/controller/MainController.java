package com.poly.hangnt169.controller;

import com.poly.hangnt169.model.Student;
import com.poly.hangnt169.service.MainService;
import jakarta.servlet.http.HttpServletResponse;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.util.List;

@Controller
public class MainController {

    @Autowired
    private MainService mainServiceImpl;

    @GetMapping("/index")
    public String index() {
        return "index.html";
    }

    @GetMapping("/download-template-mark")
    @ResponseBody
    public void downloadTemplateMark(HttpServletResponse response) {
        mainServiceImpl.downloadTemplateMark(response);
    }

    @GetMapping("/download-template-mark-com108")
    @ResponseBody
    public void downloadTemplateMarkCom(HttpServletResponse response) {
        mainServiceImpl.downloadTemplateMarkCOM108(response);
    }

    @PostMapping("/import-mark")
    public String exportMark(@RequestParam("fileLayDiemLab") MultipartFile fileLayDiemLab,
                             @RequestParam("fileLayDiemQuiz") MultipartFile[] fileLayDiemQuiz,
                             @RequestParam("fileLayDiem") MultipartFile fileLayDiem, HttpServletResponse response)
            throws IOException {
        List<Student> listLab = mainServiceImpl.readExcelLab(fileLayDiemLab);
        List<Student> listQuiz = mainServiceImpl.readExcelQuiz(fileLayDiemQuiz);
        List<Student> listMark = mainServiceImpl.readMark(fileLayDiem);
        mainServiceImpl.exportExcel(response, listMark, listLab, listQuiz);
        return "index.html";
    }

    @PostMapping("/import-mark-COM108")
    public String exportMarkCOM(@RequestParam("fileLayDiemLabCOM108") MultipartFile fileLayDiemLab,
                                @RequestParam("fileLayDiemBaiHocOnlineCOM") MultipartFile fileLayDiemBaiHocOnline,
                                @RequestParam("fileLayDiemCOM108") MultipartFile fileLayDiem, HttpServletResponse response)
            throws IOException {
        List<Student> listLab = mainServiceImpl.readExcelLabCOM108(fileLayDiemLab);
        List<Student> listQuiz = mainServiceImpl.readExcelBaiHocOnlineCOM108(fileLayDiemBaiHocOnline);
        List<Student> listMark = mainServiceImpl.readMarkCOM108(fileLayDiem);
        mainServiceImpl.exportExcelCOM108(response, listMark, listLab, listQuiz);
        return "index.html";
    }
}
