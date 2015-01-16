/*
 * Main Excel Reader
 */

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
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

	public static String type = "xls";
	public static String name = "16.006-012";
	
	public static void main(String[] args) throws IOException
	{
		ExcelReader excelReader = new ExcelReader();
		
		System.out.println("!!!!!");
		
		LotteryDto lxls = null;
		
		//Read excel file 
		
		List<LotteryDto> list = excelReader.readXls("./" + name + "." + type);
		

		
		//System.out.println(list.size());
		
		//obtain list contents
		int total = list.size();
		int zhuang = 0;
		int xian = 0;
		
		for (int j=0; j<list.size(); j++)
		{
			lxls = (LotteryDto) list.get(j);
			String finalResult = excelReader.splitResult(lxls.getResult());
			
			if (finalResult.equals("в╞"))
			{
				zhuang++;
			}
			else if (finalResult.equals("оп"))
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
	
	private List<LotteryDto> readXls(String filepath) throws IOException
	{
		InputStream is = new FileInputStream(filepath);
		
		LotteryDto lotteryDto = null;
		List<LotteryDto> list = new ArrayList<LotteryDto>();
		
		/*
		 * different between xls & xlsx
		 */
		
		if (type == "xls")
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
				System.out.println("***"+hssfSheet.getLastRowNum());
				
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
					/*
					HSSFCell ch = hssfRow.getCell(10);
					if (ch == null)
					{
						continue;
					}
					lotteryDto.setChoice(getValue(ch));
					
					//System.out.println("#####"+ getValue(ch));
					
					HSSFCell bm = hssfRow.getCell(11);
					if (bm == null)
					{
						continue;
					}
					if (bm.getCellType() == bm.CELL_TYPE_NUMERIC)
					{
						lotteryDto.setBuyed_money((int)bm.getNumericCellValue());
					}
					else
					{
						lotteryDto.setBuyed_money(Integer.parseInt(getValue(bm)));
					}
					
					HSSFCell rm = hssfRow.getCell(12);
					if (rm == null)
					{
						continue;
					}
					if (rm.getCellType() == rm.CELL_TYPE_NUMERIC)
					{
						lotteryDto.setResult_money((int)rm.getNumericCellValue());
					}
					else
					{
						lotteryDto.setResult_money(Integer.parseInt(getValue(rm)));
					}
					
					//System.out.println("&&&&&&"+ getValue(rm));
					*/
					
					/*Modified for new requirement.*/
					
					//HSSFCell cell_result = hssfRow.getCell(7);
					HSSFCell cell_result = hssfRow.getCell(9);
					if (cell_result == null)
					{
						continue;
					}
					lotteryDto.setResult(getValue(cell_result));
					
					
					list.add(lotteryDto);
					
				}
				
			}
		}
		
		else if (type == "xlsx")
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
				System.out.println("***"+xssfSheet.getLastRowNum());
				
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
					
					XSSFCell cell_result = xssfRow.getCell(7);
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
	
}
