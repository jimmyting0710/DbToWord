package com.example.demo;

import java.io.IOException;

import javax.annotation.PostConstruct;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

import com.example.demo.ThisIsService.Thisisservice;

@SpringBootApplication
public class DbToWordApplication {
	@Autowired
	Thisisservice thisisservice;

	public static void main(String[] args) throws IOException {
		SpringApplication.run(DbToWordApplication.class, args);
	}
	
//	@PostConstruct
//	public void hehe() {
//		thisisservice.getWord();
//	}

	@PostConstruct
	public void AA() {
		thisisservice.getWord2();
	}
}
