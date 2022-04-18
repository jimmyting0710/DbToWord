package com.example.demo.ThisIsService;

import java.util.List;
import java.util.Map;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.context.annotation.ComponentScan;
import org.springframework.stereotype.Service;

import com.example.demo.ToWord;
import com.example.demo.ThisIsRepository.Findrepository;

@Service
public class Thisisservice {
	@Autowired
	Findrepository findrepository;

	public void getWord() {

		// 取得資料庫資料
//		List<Map<String, String>> data = findrepository.findalldata();
//		List<Map<String, String>> input =findrepository.findinput();
//		List<Map<String, String>> output =findrepository.findoutput();
//		List<Map<String, String>> catalog =findrepository.findcatalog();
//		ToWord toWord = new ToWord();
		// 轉成word
//		toWord.outPutToWork(input,output,data,catalog);
	}

	public void getWord1() {

//		// 取得資料庫資料
//		List<Map<String, String>> data = findrepository.findalldata();
//		List<Map<String, String>> data2 = findrepository.findgetinclude();
//		ToWord toWord = new ToWord();
//		// 轉成word
//		toWord.outPutToWork2(data, data2);
	}

	public void getWord2() {
		List<Map<String, String>> data = findrepository.findalldata();
		List<Map<String, String>> db2data = findrepository.finddb2data();
		List<Map<String, String>> imsdbdata = findrepository.findimsdbdata();
		ToWord toWord = new ToWord();
//		// 轉成word
		toWord.outPutToWork3(data,db2data,imsdbdata);
	}
}
