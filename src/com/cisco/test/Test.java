package com.cisco.test;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.InputStream;
import java.util.List;

import com.cisco.util.excel.SimpleParser;

public class Test {

	public static void main(String[] args) throws FileNotFoundException {
		List<Student> result;
		InputStream is = new FileInputStream("student.xlsx");
		SimpleParser<Student> sp = new SimpleParser<>(Student.class);
		sp.setIs(is);
		result = sp.parse();
		System.out.println(result);
	}

}
