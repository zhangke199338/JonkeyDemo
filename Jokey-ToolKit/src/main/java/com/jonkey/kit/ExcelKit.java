package com.jonkey.kit;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;

import org.apache.commons.codec.binary.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.seaway.common.string.StringUtil;

import net.sourceforge.pinyin4j.PinyinHelper;

public class ExcelKit {
	//默认单元格内容为数字时格式  
    private static DecimalFormat df = new DecimalFormat("0");  
    // 默认单元格格式化日期字符串   
    private static SimpleDateFormat sdf = new SimpleDateFormat(  "yyyy-MM-dd HH:mm:ss");   
    // 格式化数字  
    private static DecimalFormat nf = new DecimalFormat("0.00");
	/*  
     * @return 将返回结果存储在ArrayList内，存储结构与二位数组类似  
     * lists.get(0).get(0)表示过去Excel中0行0列单元格  
     */  
    public static ArrayList<ArrayList<Object>> readExcel2003(File file){  
        try{  
            ArrayList<ArrayList<Object>> rowList = new ArrayList<ArrayList<Object>>();  
            ArrayList<Object> colList;  
            HSSFWorkbook wb = new HSSFWorkbook(new FileInputStream(file));  
            HSSFSheet sheet = wb.getSheetAt(0);  
            HSSFRow row;  
            HSSFCell cell;  
            Object value;  
            for(int i = sheet.getFirstRowNum() , rowCount = 0; rowCount < sheet.getPhysicalNumberOfRows() ; i++ ){  
                row = sheet.getRow(i);  
                colList = new ArrayList<Object>();  
                if(row == null){  
                    //当读取行为空时  
                    if(i != sheet.getPhysicalNumberOfRows()){//判断是否是最后一行  
                        rowList.add(colList);  
                    }  
                    continue;  
                }else{  
                    rowCount++;  
                }  
                for( int j = row.getFirstCellNum() ; j <= row.getLastCellNum() ;j++){  
                    cell = row.getCell(j);  
                    if(cell == null || cell.getCellType() == HSSFCell.CELL_TYPE_BLANK){  
                        //当该单元格为空  
                        if(j != row.getLastCellNum()){//判断是否是该行中最后一个单元格  
                            colList.add("");  
                        }  
                        continue;  
                    }  
                    switch(cell.getCellType()){  
                     case XSSFCell.CELL_TYPE_STRING:    
//                            System.out.println(i + "行" + j + " 列 is String type");    
                            value = cell.getStringCellValue();    
                            break;    
                        case XSSFCell.CELL_TYPE_NUMERIC:    
                            if ("@".equals(cell.getCellStyle().getDataFormatString())) {    
                                value = df.format(cell.getNumericCellValue());    
                            } else if ("General".equals(cell.getCellStyle()    
                                    .getDataFormatString())) {    
                                value = nf.format(cell.getNumericCellValue());    
                            } else {    
                                value = sdf.format(HSSFDateUtil.getJavaDate(cell    
                                        .getNumericCellValue()));    
                            }    
//                            System.out.println(i + "行" + j + " 列 is Number type ; DateFormt:"+ value.toString());   
                            break;    
                        case XSSFCell.CELL_TYPE_BOOLEAN:    
//                            System.out.println(i + "行" + j + " 列 is Boolean type");    
                            value = Boolean.valueOf(cell.getBooleanCellValue());  
                            break;    
                        case XSSFCell.CELL_TYPE_BLANK:    
//                            System.out.println(i + "行" + j + " 列 is Blank type");    
                            value = "";    
                            break;    
                        default:    
//                            System.out.println(i + "行" + j + " 列 is default type");    
                            value = cell.toString();    
                    }// end switch  
                    colList.add(value);  
                }//end for j  
                rowList.add(colList);  
            }//end for i  
              
            return rowList;  
        }catch(Exception e){  
            return null;  
        }  
    }  
    
    public static ArrayList<ArrayList<Object>> readExcel2007(File file){  
        try{  
            ArrayList<ArrayList<Object>> rowList = new ArrayList<ArrayList<Object>>();  
            ArrayList<Object> colList;  
            XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream(file));  
            XSSFSheet sheet = wb.getSheetAt(0);  
            XSSFRow row;  
            XSSFCell cell;  
            Object value;  
            for(int i = sheet.getFirstRowNum() , rowCount = 0; rowCount < sheet.getPhysicalNumberOfRows() ; i++ ){  
                row = sheet.getRow(i);  
                colList = new ArrayList<Object>();  
                if(row == null){  
                    //当读取行为空时  
                    if(i != sheet.getPhysicalNumberOfRows()){//判断是否是最后一行  
                        rowList.add(colList);  
                    }  
                    continue;  
                }else{  
                    rowCount++;  
                }  
                for( int j = row.getFirstCellNum() ; j <= row.getLastCellNum() ;j++){  
                    cell = row.getCell(j);  
                    if(cell == null || cell.getCellType() == HSSFCell.CELL_TYPE_BLANK){  
                        //当该单元格为空  
                        if(j != row.getLastCellNum()){//判断是否是该行中最后一个单元格  
                            colList.add("");  
                        }  
                        continue;  
                    }  
                    switch(cell.getCellType()){  
                     case XSSFCell.CELL_TYPE_STRING:    
//                            System.out.println(i + "行" + j + " 列 is String type");    
                            value = cell.getStringCellValue();    
                            break;    
                        case XSSFCell.CELL_TYPE_NUMERIC:    
                            if ("@".equals(cell.getCellStyle().getDataFormatString())) {    
                                value = df.format(cell.getNumericCellValue());    
                            } else if ("General".equals(cell.getCellStyle()    
                                    .getDataFormatString())) {    
                                value = nf.format(cell.getNumericCellValue());    
                            } else {    
                                value = sdf.format(HSSFDateUtil.getJavaDate(cell    
                                        .getNumericCellValue()));    
                            }    
//                            System.out.println(i + "行" + j + " 列 is Number type ; DateFormt:"+ value.toString());   
                            break;    
                        case XSSFCell.CELL_TYPE_BOOLEAN:    
//                            System.out.println(i + "行" + j + " 列 is Boolean type");    
                            value = Boolean.valueOf(cell.getBooleanCellValue());  
                            break;    
                        case XSSFCell.CELL_TYPE_BLANK:    
//                            System.out.println(i + "行" + j + " 列 is Blank type");    
                            value = "";    
                            break;    
                        default:    
//                            System.out.println(i + "行" + j + " 列 is default type");    
                            value = cell.toString();    
                    }// end switch  
                    colList.add(value);  
                }//end for j  
                rowList.add(colList);  
            }//end for i  
              
            return rowList;  
        }catch(Exception e){  
            return null;  
        }  
    }  
    
    public static String getPinYinHeadChar(String str)  
    {  
        String convert = "";  
        for (int j = 0; j < str.length(); j++)  
        {  
            char word = str.charAt(j);  
            String[] pinyinArray = PinyinHelper.toHanyuPinyinStringArray(word);  
            if (pinyinArray != null)  
            {  
                convert += pinyinArray[0].charAt(0);  
            } else  
            {  
                convert += word;  
            }  
        }  
        return convert;  
    }  
    
    public static void writeExcel(ArrayList<ArrayList<Object>> result,String path){  
        if(result == null){  
            return;  
        }  
        HSSFWorkbook wb = new HSSFWorkbook();  
        HSSFSheet sheet = wb.createSheet("sheet1");  
        for(int i = 0 ;i < result.size() ; i++){  
             HSSFRow row = sheet.createRow(i);  
            if(result.get(i) != null){  
                for(int j = 0; j < result.get(i).size() ; j ++){  
                    HSSFCell cell = row.createCell(j);  
                    cell.setCellValue(result.get(i).get(j).toString());  
                }  
            }  
        }  
        ByteArrayOutputStream os = new ByteArrayOutputStream();  
        try  
        {  
            wb.write(os);  
        } catch (IOException e){  
            e.printStackTrace();  
        }  
        byte[] content = os.toByteArray();  
        File file = new File(path);//Excel文件生成后存储的位置。  
        OutputStream fos  = null;  
        try  
        {  
            fos = new FileOutputStream(file);  
            fos.write(content);  
            os.close();  
            fos.close();  
        }catch (Exception e){  
            e.printStackTrace();  
        }             
    }  
    
	public static void main(String[] args) {
//		File file = new File("/Users/docse/Downloads/代发测试1(2).xls");  
//        ArrayList<ArrayList<Object>> result = ExcelKit.readExcel2003(file);  
////        ArrayList<ArrayList<Object>> other = new ArrayList<ArrayList<Object>>();
//        for(int i = 0 ;i < result.size() ;i++){  
////            for(int j = 0;j<result.get(i).size(); j++){  
////                System.out.println(i+"行 "+2+"列  "+ result.get(i).get(2).toString());  
////        	System.out.println("{\"bankName\": \""+result.get(i).get(2).toString().trim()+"\",\"firstLetter\": \""+getPinYinHeadChar(result.get(i).get(2).toString()).trim().substring(0, 1).toUpperCase()+"\"},");
////        	System.out.println("{\"bankName\": \""+result.get(i).get(0).toString().trim()+"\",\"firstLetter\": \""+result.get(i).get(1).toString().trim()+"\"},");
////        	ArrayList<Object> l = new ArrayList<Object>();
////        	l.add(result.get(i).get(2).toString().trim());
////        	l.add(getPinYinHeadChar(result.get(i).get(2).toString()).trim().substring(0, 1).toUpperCase());
////        	other.add(l);
////            } 
//        	System.out.println("INSERT INTO tb_student_agent_info VALUES ("+(i+100)+",'"+result.get(i).get(0).toString().trim()+"','"
//        																	+result.get(i).get(1).toString().trim()+"','"
//        																	+result.get(i).get(2).toString().trim()+"','"
//        																	+result.get(i).get(3).toString().trim()+"','"
//        																	+result.get(i).get(4).toString().trim()+"','"
//        																	+result.get(i).get(5).toString().trim()+"','"
//        																	+result.get(i).get(6).toString().trim()+"','"
//        																	+result.get(i).get(7).toString().trim()+"','20171226162900','20171226162900');");
//        }  
////        ExportExcel.writeExcel(other,"/Users/docse/Downloads/bb.xls");
        
		File file = new File("/Users/docse/Downloads/ChinaCityName.xlsx");
		ArrayList<ArrayList<Object>> result = ExcelKit.readExcel2007(file); 
		
		String pid = "";
		String pname = "";
		String peng = "";
		
		for(int i = 1 ;i < result.size() ;i++){
//			for(int j = 0;j<result.get(i).size(); j++){ 
				if(!"".equals(result.get(i).get(0).toString().trim())){
					pid = result.get(i).get(0).toString().trim();
					pname = result.get(i).get(1).toString().trim();
					peng = result.get(i).get(2).toString().trim();
				}
				
				
				System.out.println("INSERT INTO tb_city_name VALUES ("+(i+100)+",'"+pid+"','"
						+pname+"','"
						+peng+"','"
						+result.get(i).get(3).toString().trim()+"','"
						+result.get(i).get(4).toString().trim()+"','"
						+result.get(i).get(5).toString().trim()+"','"
						+result.get(i).get(6).toString().trim()+"','"
						+result.get(i).get(7).toString().trim()+"','"
						+result.get(i).get(8).toString().trim()+"','20180626162900','20180626162900');");
				
//			}
		}
        
        
        
	}

}
