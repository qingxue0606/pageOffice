package com.example.demo.controller;


import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.servlet.ModelAndView;

import javax.servlet.http.HttpSession;
import java.util.Map;

@Controller
public class MusicController {
    @RequestMapping(value = "/index3", method = RequestMethod.GET)
    public ModelAndView showIndex3(HttpSession session, Map<String, Object> map) {
        Object user=session.getAttribute("user");
        System.out.println("urser:"+ user);

        int min=1;
        int max=10;
        int num=min+(int)(Math.random()*(max-min+1));

        //num=1;
        String note="";

        switch(num)
        {
            case 1 :
                note="C";
                break;
            case 2 :
                note="D";
                break;
            case 3 :
                note="E";
                break;
            case 4 :
                note="F";
                break;
            case 5 :
                note="G";
                break;
            case 6 :
                note="A";
                break;
            case 7 :
                note="B";
                break;
            case 8 :
                note="C2";
                break;
                //重点练的音
            case 9 :
                note="F";
                break;
            case 10 :
                note="E";
                break;
            default :
                note="H";
        }


        map.put("note",note);

        ModelAndView mv = new ModelAndView("Index3");









        return mv;
    }



    @RequestMapping(value = "/index4", method = RequestMethod.GET)
    public ModelAndView showIndex4(HttpSession session, Map<String, Object> map) {

        ModelAndView mv = new ModelAndView("Index4");




        return mv;
    }

}
