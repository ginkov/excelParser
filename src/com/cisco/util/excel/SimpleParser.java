package com.cisco.util.excel;

import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jboss.logging.Logger;


/**
 * 根据输入的泛型类, 解析 Excel 表
 * @author xinyin
 *
 */
public class SimpleParser<T extends Validable> {
	private static final Logger log = Logger.getLogger(SimpleParser.class);

	private final Class<T> clazz;
	private InputStream is;
	
	private List<T> validRows;
	
	/*
	 *  把字段的名称，或者别名，与字段对应起来.
	 *  比如，字段可能叫 stundentName, 在 Excel 里可能是 stundentName, 也可能是 name, 
	 *  	或者是 学生姓名, 姓名.
	 *  这个 Map 就是把 stundentName, name, 学生姓名, 姓名 都映射成 studentName 这个 Field.
	 */
	private Map<String, Field> name2Fld;
	private Map<Field, Integer> fldPosition;

	public SimpleParser(Class<T> clazz) {
		this.clazz = clazz;
		this.name2Fld = new HashMap<>();
		this.validRows = new ArrayList<>();
	}
	
	/**
	 * 解析 Excel 的每一行，返回结果.
	 * @return
	 */
	public List<T> parse() {
		
		/*
		 * 目标类可以通过 @Alias 注解，说明本字段的别名.
		 */
		for(Field f: clazz.getDeclaredFields()) {
			name2Fld.put(f.getName(), f);
			Alias a = f.getDeclaredAnnotation(Alias.class);
			if( a != null && ! isEmpty(a.value())) {
				for(String s: a.value()) {
					String sTrim = s.trim();
					if(!sTrim.isEmpty()) {
						if(name2Fld.containsKey(sTrim)) {
							log.errorv("字段{0}的别名{1}与其它字段冲突.", f.getName(), sTrim);
						}
						else {
							name2Fld.put(sTrim, f);
						}
					}
				}
			}
		}
	    boolean headerFound = false;
		try {
			XSSFWorkbook wb = new XSSFWorkbook(is);
			XSSFSheet sheet = wb.getSheetAt(0);
			wb.close();
			is.close();
			Iterator<Row> rowIterator = sheet.iterator();
				
		    // 循环每一行
	        while(rowIterator.hasNext()) {
	            Row row = rowIterator.next();
	            // 扫描一行是不是一个合格的标题行，如果不是，就继续找头
	            if(!headerFound) {
	                // 如是找到头，返回 True，同时设定了每个字段的位置.
	                headerFound = parseHeader(row);
	            }
	            else {
	                T t = parseRow(row);
	                if(t != null) {
	                	validRows.add(t);
	                }
	            }
	        }
		} catch (IOException e) {
			log.errorv("读取excel文件错误: {0}", e.getMessage());
		}
		
		return validRows;
	}		
		
	// 找头部
	// 循环每一行.  解析这一列. 读取内容, 看看是不是 clazz 里的 Field.
	// 如果是 Field 填充 fieldPosition
	// 如果每个 field 都有 Position 则这个就定下来了，头找到了。
	// 把这个 field Position 记下来
	// 解析一行
	// 根据记录下来的 fieldPosition, 读取每个 Position 上的 str
	// 根据 Field 的类型, 设置 T 上的值，Field 只支持 Integer, String, Boolean, Double 四种类型
	
	private T parseRow(Row row) {
		T t = null;
		try {
			t = clazz.newInstance();
		} catch (InstantiationException | IllegalAccessException e ) {
			log.errorv("通过反射创建类{0}的实例时出错:{1}",clazz.getName(),e.getMessage());
			return null;
		}
		for(Field f: fldPosition.keySet()) {
			f.setAccessible(true);
			Cell cell = row.getCell(fldPosition.get(f));
			try {
				if(f.getType().equals(String.class)) {
					f.set(t, cell.getStringCellValue().trim());
				}
				else if(f.getType().equals(Integer.class)) {
					f.set(t, (int)cell.getNumericCellValue());
				}
				else if(f.getType().equals(Double.class)) {
					f.set(t, cell.getNumericCellValue());
				}
				else if(f.getType().equals(Boolean.class)) {
					f.set(t, cell.getBooleanCellValue());
				}
			} catch (IllegalArgumentException | IllegalAccessException e) {
				log.errorv("通过反射设置类{0}的属性{1}时出错:{2}", clazz.getName(),f.getName(),e.getMessage());
			}
		}
		if(t.isValid()) {
			return t;
		}
		else {
        	log.warnv("行内容不正确: {0}", t.toString());
			return null;
		}
	}
	/**
	 * 解析头部.
	 * @param row
	 * @return
	 * 循环每个 Cell， 如果和 Field 名相匹配，就加到 fieldPosition 里面
	 * 如果这个一行 fieldPosition 长度大于0，则说明头找到了。
	 */
	private boolean parseHeader(Row row) {
		int pos = 0;
		Map<Field, Integer> localFp = new HashMap<>();
		Iterator<Cell> cellIterator = row.cellIterator();
	    while(cellIterator.hasNext()) {
	        Cell cell = cellIterator.next();
	        switch(cell.getCellType()) {
            case Cell.CELL_TYPE_STRING:
               	String cellStr = cell.getStringCellValue().trim();
               	if(name2Fld.keySet().contains(cellStr)) {
               		localFp.put(name2Fld.get(cellStr), pos);
               	}
            default:
            	break;
	        }
	        ++ pos;
	    }
	    if(localFp.size() > 0) {
	    	this.fldPosition = localFp;
	    	return true;
	    }
	    else {
	    	return false;
	    }
	}
	
	public void setIs(InputStream is) {
		this.is = is;
	}
	
	/**
	 * 判断一个 String[] alias 是否为空.
	 * @param a
	 * @return
	 */
	private boolean isEmpty(String [] alias) {
		boolean result = true;
		for(String s: alias) {
			if(s!=null && !s.isEmpty()) {
				result = false;
			}
		}
		return result;
	}

}
