package com.dag.service;

import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

//@SpringBootApplication
@SpringBootApplication(scanBasePackages = {"com.dag"})
//Same as @Configuration, @EnableAutoConfiguration, and @ComponentScan.
public class DagServiceApplication {

	public static void main(String[] args) {
		SpringApplication.run(DagServiceApplication.class, args);
	}
}
