package util.excel.app;

import util.excel.logic.BusinessLogic;

/**
 * Sheet Data Copy
 *
 */
public class App 
{
    public static void main( String[] args ){
    	try {
    		new BusinessLogic().copySheetData(args[0], args[1]);
		} catch (Exception exception) {
			System.out.println("Error while copying sheet data "+exception.getMessage());
			return;
		}
    	System.out.println("Sheet Data Copy Completed Successfully!! :)");
    }
}
