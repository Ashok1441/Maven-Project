package com.BaseTest;

import com.codoid.products.exception.FilloException;
import com.codoid.products.fillo.Connection;
import com.codoid.products.fillo.Fillo;
import com.codoid.products.fillo.Recordset;

public class Datafetching {
public static void main(String[] args) throws FilloException {
		
		Fillo fillo = new Fillo();
		Connection connection =fillo.getConnection("D:\\Selenium\\ApachePOI\\DataFiles\\Sample2.xlsx");
		
		String strQuari ="Select * from Sheet2";
		Recordset rs =connection.executeQuery(strQuari);
		
		//print total Excell dsta
		while(rs.next()) {
			System.out.println(rs.getField("FirstName")+"----"+rs.getField("LastName")+"----"+rs.getField("MailId")+"----"+rs.getField("Phone NO"));
			
		}
		//Total Rows in Excell Sheet
		System.out.println("Total Rows in excel" + rs.getCount());
		
		rs.moveFirst();
		System.out.println(rs.getField("FirstName")+"----"+rs.getField("LastName")+"----"+rs.getField("MailId")+"----"+rs.getField("Phone NO"));
		
		rs.moveNext();
		System.out.println(rs.getField("FirstName")+"----"+rs.getField("LastName")+"----"+rs.getField("MailId")+"----"+rs.getField("Phone NO"));
         
		rs.movePrevious();
		System.out.println(rs.getField("FirstName")+"----"+rs.getField("LastName")+"----"+rs.getField("MailId")+"----"+rs.getField("Phone NO"));
		

		

	
		
		rs.close();
		connection.close();
	}


}
