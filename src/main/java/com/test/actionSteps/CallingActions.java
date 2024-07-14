package com.test.actionSteps;

public class CallingActions {

	public static String calling(String col2,String col3,String col4) {
		String status=null;
		switch(col2) {
		case "openBrowser":
			status=Actions.openBrowser();
			break;
		case "navigate":
			status=Actions.navigate();
			break;
		case "verifyText":
			status=Actions.verifyText();
			break;
		case "buttonClick":
			status=Actions.buttonClick();
			break;
		}
		return status;
	}
}
