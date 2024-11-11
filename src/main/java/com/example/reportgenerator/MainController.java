package com.example.reportgenerator;

import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;

@Controller
public class MainController {

    // Главная страница приложения
    @GetMapping("/")
    public String showHomePage() {
        // Возвращает главную страницу с шаблоном отчета
        return "index";  
    }
}
