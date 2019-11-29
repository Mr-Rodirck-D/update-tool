package projectCode;

import java.io.File;
import java.io.FileWriter;
import java.io.IOException;
import java.io.PrintWriter;
import java.io.StringWriter;
import java.io.UnsupportedEncodingException;
import java.net.URISyntaxException;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;
import javax.swing.JOptionPane;

public class ErrorLog {
	public static String Path = null;
	String ErrorType = null;
	
	public ErrorLog(String ErrorLogType) {
		this.ErrorType = ErrorLogType;
	}
	
	public void createErrorLog(String LogContent) throws URISyntaxException, UnsupportedEncodingException {
		Path = ErrorLog.class.getProtectionDomain().getCodeSource().getLocation().toURI().getPath();
		//decode to deal with Chinese characters and blanks
		Path = java.net.URLDecoder.decode(Path, "utf-8");
		Path = Path.replace("%20", " ");
		//acquire parent path
		Path = Path.substring(1, Path.lastIndexOf("/"));
		Path = Path + "/Errorlog";
		File LogDir = new File(Path);
		if(!LogDir.exists()) {
			LogDir.mkdir();
		}
		FileWriter writer;
		SimpleDateFormat sdf = new SimpleDateFormat("MM-dd-yyyy EEE HH.mm.ss");	//set date format
		Calendar calendar = Calendar.getInstance();	
		Date date = calendar.getTime();	//get system date
		String dateStringPase = sdf.format(date);	//change date to string according to set format
		String errorLogName = dateStringPase + " " + this.ErrorType;
		String suffix = ".txt";

		File errorLog = new File(Path + "/" + errorLogName + suffix);
		if(!errorLog.exists()) {
			try {
				errorLog.createNewFile();
				writer = new FileWriter(errorLog);
				writer.write("");
				writer.write(this.ErrorType + "\n" + LogContent);
				writer.flush();
				writer.close();
			} catch (IOException e) {
				JOptionPane.showMessageDialog(null, "IO operation failed when creating error log!\nError log create failed!", "Warning!", JOptionPane.WARNING_MESSAGE);
				e.printStackTrace();
			}
		}
		
	}
	
	public static String getErrorInfo(Exception e){
		StringWriter sw = new StringWriter();
		PrintWriter pw = new PrintWriter(sw);
		e.printStackTrace(pw);
		return sw.toString();
	}
}