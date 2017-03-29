package com.cisco.test;

import com.cisco.util.excel.Alias;
import com.cisco.util.excel.Validable;

public class Student implements Validable {
	
	@Alias("姓名")
	private String name;
	
	@Alias("年龄")
	private Integer age;
	
	@Alias({"昵称","描述"})
	private String description;
	
	@Alias("成绩")
	private Double score;
	
	public String getName() {
		return name;
	}
	public void setName(String name) {
		this.name = name;
	}
	public Integer getAge() {
		return age;
	}
	public void setAge(Integer age) {
		this.age = age;
	}
	public String getDescription() {
		return description;
	}
	public void setDescription(String description) {
		this.description = description;
	}
	public Double getScore() {
		return score;
	}
	public void setScore(Double score) {
		this.score = score;
	}
	@Override
	public boolean isValid() {
		boolean valid = true;
		if(name == null || name.isEmpty()) {
			valid = false;
		}
		if(age == null || age <7) {
			valid = false;
		}
		if(score > 100 || score < 0) {
			valid = false;
		}
		return valid;
	}
	
	public String toString() {
		return String.format("{name:%s, age:%d, description:%s, score:%f}", name, age, description, score);
	}

	
}
