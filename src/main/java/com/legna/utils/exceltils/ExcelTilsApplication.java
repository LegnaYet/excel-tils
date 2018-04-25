package com.legna.utils.exceltils;

import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

import java.io.File;
import java.util.ArrayList;
import java.util.Objects;

//@SpringBootApplication
public class ExcelTilsApplication {

	public static void main(String[] args) {
//		SpringApplication.run(ExcelTilsApplication.class, args);
		File file = new File("C:\\Users\\Administrator\\Desktop\\TestExcel.xls");
		ArrayList<ArrayList<Object>> arrayLists = ReadExcel.readExcel(file);
		ArrayList<Object> objects = arrayLists.get(2);
		for (Object object:objects) {
			System.out.println(object.toString());
		}
	}
}
