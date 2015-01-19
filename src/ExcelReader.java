/*
 * Main Excel Reader
 */

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;




//xls
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

//xlsx
import org.apache.poi.xssf.usermodel.XSSFCell;  
import org.apache.poi.xssf.usermodel.XSSFRow;  
import org.apache.poi.xssf.usermodel.XSSFSheet;  
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class ExcelReader {

	//public static String type = "xls";
	//public static String name = "16.006-012";
	public List<String> resultContent = new ArrayList<String>();

	public static void main(String[] args) throws IOException
	{
		ExcelReader excelReader = new ExcelReader();
		RandomSelect randomSelect = new RandomSelect();
		
		List<String> filenames = excelReader.readAllFileName("../lottery_data");
		
		for(int i=0; i<filenames.size(); i++)
		{	
			String nameOfFile = filenames.get(i);
			System.out.println(" ");
			System.out.println("Analysing " + nameOfFile + ".....................");
			excelReader.analyseData(excelReader, nameOfFile);
		}
		
		System.out.println(" ");
		System.out.println("Size of the saving list: " + excelReader.resultContent.size());
		
		int selectedSize = 3000;
		List<Integer> generatedRandomNum = randomSelect.GenRandomNum(selectedSize);
		List<String> finalSelectedNum = new ArrayList<String>();
		System.out.println("SSSSSSSSSSSSSSSSSSSSSSSSS");
		for(int i=0; i<selectedSize; i++)
		{
			int index = generatedRandomNum.get(i);
			//System.out.print(index + " ");
			
			finalSelectedNum.add(excelReader.resultContent.get(index));
		}
		
		//System.out.println("SSSSSSSSSSSSSSSSSSSSSSSSS" + finalSelectedNum.size());
		
		try
		{
			FileWriter fileWriter = new FileWriter("./data.log");
			/*
			String s = new String("This is a test!  \n" + "aaaa");
			fileWriter.write(s);
			String b = new String("test !!!!!!");
			fileWriter.write(b);
			*/
			for (int i=0; i<selectedSize; i++)
			{
				String s = finalSelectedNum.get(i);
				fileWriter.write(s);
				fileWriter.write("\n");
			}
			fileWriter.close(); // 关闭数据流
		}
		catch (IOException e) 
		{
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
		
	}
	
	private List<LotteryDto> readXls(String filepath , String type, int flag) throws IOException
	{
		InputStream is = new FileInputStream(filepath);
		
		LotteryDto lotteryDto = null;
		List<LotteryDto> list = new ArrayList<LotteryDto>();
		
		/*
		 * different between xls & xlsx
		 */
		
		if (type.equals("xls"))
		{
			HSSFWorkbook hssfworkbook = new HSSFWorkbook(is);
			
			/*
			 * sheet
			 */
			for (int numSheet = 0; numSheet < hssfworkbook.getNumberOfSheets(); numSheet++)
			{
				HSSFSheet hssfSheet = hssfworkbook.getSheetAt(numSheet);
				
				if (hssfSheet == null)
				{
					continue;
				}
				
				/*
				 * row
				 */
				//System.out.println("***"+hssfSheet.getLastRowNum());
				
				for (int rowNum =1; rowNum <= hssfSheet.getLastRowNum(); rowNum++)
				{
					HSSFRow hssfRow = hssfSheet.getRow(rowNum);
					if (hssfRow == null)
					{
						continue;
					}
					
					lotteryDto = new LotteryDto();
					
					/*
					 * cell
					 */
					
					/*Modified for new requirement.*/
					//Differ by the first letter in filename.
					HSSFCell cell_result;
					if (flag == 1)
					{
						cell_result = hssfRow.getCell(7);
					}
					else if (flag == 2)
					{
						cell_result = hssfRow.getCell(10);
					}
					else
					{
						cell_result = hssfRow.getCell(9);
					}
					
					
					if (cell_result == null)
					{
						continue;
					}
					lotteryDto.setResult(getValue(cell_result));
					
					
					list.add(lotteryDto);
					
				}
				
			}
		}
		
		else if (type.equals("xlsx"))
		{
			XSSFWorkbook xssfWorkbook = new XSSFWorkbook(filepath);
			
			/*
			 * sheet
			 */
			for (int numSheet = 0; numSheet < xssfWorkbook.getNumberOfSheets(); numSheet++)
			{
				XSSFSheet xssfSheet = xssfWorkbook.getSheetAt(numSheet);
				
				if (xssfSheet == null)
				{
					continue;
				}
				
				/*
				 * row
				 */
				//System.out.println("***"+xssfSheet.getLastRowNum());
				
				for (int rowNum =1; rowNum <= xssfSheet.getLastRowNum(); rowNum++)
				{
					XSSFRow xssfRow = xssfSheet.getRow(rowNum);
					if (xssfRow == null)
					{
						continue;
					}
					
					lotteryDto = new LotteryDto();
					
					/*
					 * cell
					 */
					
					//XSSFCell cell_result = xssfRow.getCell(7);
					//Differ by the first letter in filename.
					XSSFCell cell_result;
					if (flag == 1)
					{
						cell_result = xssfRow.getCell(7);
					}
					else
					{
						cell_result = xssfRow.getCell(9);
					}
					
					if (cell_result == null)
					{
						continue;
					}
					lotteryDto.setResult(getValue_xlsx(cell_result));
					
					
					list.add(lotteryDto);
					
				}
				
			}
			
		}
		
		return list;
	}
	
	//return cell value
	private String getValue(HSSFCell hssfCell)
	{
		if (hssfCell.getCellType() == hssfCell.CELL_TYPE_BOOLEAN)
		{
			//Return bool
			return String.valueOf(hssfCell.getBooleanCellValue());
		}
		else if (hssfCell.getCellType() == hssfCell.CELL_TYPE_NUMERIC)
		{
			//Return numberic
			return String.valueOf(hssfCell.getNumericCellValue());
		}
		else
		{
			//Return string
			return String.valueOf(hssfCell.getStringCellValue());
		}
	}
	
	private String getValue_xlsx(XSSFCell xssfCell)
	{
		if (xssfCell.getCellType() == xssfCell.CELL_TYPE_BOOLEAN)
		{
			//Return bool
			return String.valueOf(xssfCell.getBooleanCellValue());
		}
		else if (xssfCell.getCellType() == xssfCell.CELL_TYPE_NUMERIC)
		{
			//Return numberic
			return String.valueOf(xssfCell.getNumericCellValue());
		}
		else
		{
			//Return string
			return String.valueOf(xssfCell.getStringCellValue());
		}
	}
	
	//split by [
	private String splitResult(String result)
	{
		String splited_content = null;
		String[] temp = result.split("\\[");
		splited_content = temp[0];
		return splited_content;
	}
	
	//split by .
	private String[] splitResultDot(String result)
	{
		String[] temp = result.split(".");
		return temp;
	}
	
	//Get the name of the file, separated by .
	private String getName(String filename)
	{
		String tempName = "";
		String[] temp = filename.split("\\.");
		for (int i=0; i<temp.length-1; i++)
		{
			tempName = tempName + temp[i] + ".";
		}
		
		//System.out.println("@@@@@@@@@@@@@@@@" + tempName);
		
		return tempName;
	}
	
	//Get the type of the file, separated by .
	private String getType(String filename)
	{
		String tempType;
		String[] temp = filename.split("\\.");
		tempType = temp[temp.length - 1];
		
		return tempType;
	}
	
	//Read all file names
	private List<String> readAllFileName(String path)
	{
		File file = new File(path);
		File[] array = file.listFiles();
		
		List<String> fileNames = new ArrayList<String>();
		for(int i=0; i<array.length; i++)
		{
			if(array[i].isFile())
			{
				fileNames.add(array[i].getName());
				System.out.println(array[i].getName());
			}
		}
		
		return fileNames;
		
		
	}
	
	private void analyseData(ExcelReader excelReader, String filename) throws IOException
	{
		
		//String[] temSplit = excelReader.splitResultDot(filename);
		
		String name = excelReader.getName(filename);
		String type = excelReader.getType(filename);
		
		int flag = 0;
		if (name.substring(0, 1).equals("B"))
		{
			flag = 1;
		}
		else if (name.substring(0,1).equals("C"))
		{
			flag = 2;
		}

		
		LotteryDto lxls = null;
		
		//Read excel file 
		
		String filePath = "../lottery_data/" + name + type;
		
		List<LotteryDto> list = excelReader.readXls(filePath, type, flag);

		//excelReader.readAllFileName("../lottery_data");

		
		//System.out.println(list.size());
		
		//obtain list contents
		int total = list.size();
		int zhuang = 0;
		int xian = 0;
		
		for (int j=0; j<list.size(); j++)
		{
			lxls = (LotteryDto) list.get(j);
			String finalResult = excelReader.splitResult(lxls.getResult());
			//Push result into the saving list.
			excelReader.resultContent.add(finalResult);

			if (finalResult.equals("庄") || finalResult.equals("庄 "))
			{
				zhuang++;
			}
			else if (finalResult.equals("闲") || finalResult.equals("闲 "))
			{
				xian++;
			}
			
			//System.out.println("************" + finalResult);
		}
		
		System.out.println("Total number: " + total);
		System.out.println("***zhuang: " + zhuang);
		System.out.println("*****xian: " + xian);
		
		float ratio_zhuang = (float)zhuang/total;
		float ratio_xian = (float)xian/total;
		
		System.out.println("****ratio_zhuang: " + ratio_zhuang);
		
		System.out.println("****ratio_xian: " + ratio_xian);
	}
	
	
}
