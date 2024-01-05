package org.example.report;

import java.io.File;

public class Test {
    public static void main(String[] args) {
        File f = new File("C:\\Users\\SerPivas\\Desktop\\Ошибки");
        try{
            if(f.mkdir()) {
                System.out.println("Directory Created");
            } else {
                System.out.println("Directory is not created");
            }
        } catch(Exception e){
            e.printStackTrace();
        }
    }

}
