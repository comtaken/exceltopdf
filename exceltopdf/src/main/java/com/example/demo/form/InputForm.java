package com.example.demo.form;

import lombok.Data;

@Data
public class InputForm {
	//日程
	private String SCHEDULE;
	//出発地点
	private String STARTPOINT;
	//到着地点
	private String ENDPOINT;
	//経由１
	private String VIA1;
	//経由２
	private String VIA2;
	//経由３
	private String VIA3;
	//経由４
	private String VIA4;
	//経由５
	private String VIA5;
	//経由６
	private String VIA6;
	//食事：朝
	private String BREAKFAST;
	//食事：昼
	private String LUNCH;
	//食事：夜
	private String DINNER;
	//宿泊場：施設名
	private String HOTEL;

}
