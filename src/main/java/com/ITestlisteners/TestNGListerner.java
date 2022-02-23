package com.ITestlisteners;

import org.testng.ISuiteListener;
import org.testng.ITestContext;

import org.testng.ITestListener;
import org.testng.ITestResult;
import org.testng.annotations.Listeners;

//@Listeners(com.ITestlisteners.MyfirsttTestcase.class)

public class TestNGListerner implements ITestListener,ISuiteListener {
	 
	@Override
	 public void onFinish(ITestContext resut) {
		System.out.println("TestCases finish and details are " +resut.getName());
		
	}
	@Override
	public void onStart(ITestContext resut) {
		System.out.println("TestCases started and details are " +resut.getName());
		
	}
   
	//@Override
	public void onTestFaieldButWithinSuccessPercentage(ITestResult resut) {
		System.out.println("TestCases failed and details are " +resut.getName());
		
		
	}
	 
	@Override
	public void onTestFailure(ITestResult resut) {
		System.out.println("TestCases failed and details are " +resut.getName());
		
	}
	
	@Override
	public void onTestSkipped(ITestResult arg0) {
		
	}
	
	@Override
	public void onTestStart(ITestResult arg0) {
		
	}
	
	@Override
	public void onTestSuccess(ITestResult arg0) {
		
		
	}
}
