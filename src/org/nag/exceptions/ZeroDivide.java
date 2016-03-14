package org.nag.exceptions;

public class ZeroDivide {
	
	public static void main(String arg[]){
		int firstNumber;
		int secondNumber; 
		int result;
		
		
		try{
			firstNumber=10;
			secondNumber=5;
			result=firstNumber/secondNumber;
			System.out.println(" Result "+result);
			
			
			secondNumber=0;
			result=firstNumber/secondNumber;
			System.out.println(" Result "+result);
			
		}catch(ArithmeticException ae){
			System.out.println(ae.getMessage());
		}
	}

}
