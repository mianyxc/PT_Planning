import java.io.*;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Iterator;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Planning {
	
	static int hq = 0;
	
	public static void main(String[] args) throws InvalidFormatException, IOException {
		/*
		String[] sj = {"T308Q", "D16Q", "T405NB", "T818Q", "D16H", "D06Q", "T308NB", 
				"T308HB", "T405HB", "D06H", "T405Q"};
		String[] color = {"亮白", "亮银", "活力橙", "火炬红", "金黄", "格兰纳达黑", "元黑", "辣椒红",
				"天空蓝", "扎西库拉黄"};
				*/
		ArrayList<YQJ> yqj = new ArrayList<YQJ>();
		ArrayList<SJ> sj = new ArrayList<SJ>();
		ArrayList<PaintingEvent> plan = new ArrayList<PaintingEvent>();
		
		//导入油漆件需求
		Workbook wb = WorkbookFactory.create(new File("demand.forecast_18.xlsx"));
		Sheet sheet = wb.getSheetAt(0);
		for (Row row : sheet) {
			if(row.getCell(2).toString().equals("油漆件")) {
				int id = Integer.valueOf(row.getCell(0).toString());
				String type = parseType(row.getCell(1).toString());
				String color = parseColor(row.getCell(1).toString());
				int[] demand = new int[24];
				for(int i = 12; i < 36; i++) {
					Cell cell = row.getCell(i);
					if(cell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
						demand[i-12] = (int) cell.getNumericCellValue();
					} else {
						demand[i-12] = 0;
					}
				}
				YQJ temp = new YQJ(id, type, color, demand, 0);
				yqj.add(temp);
			}
		}
		
		//导入油漆件库存
		Workbook wb2 = WorkbookFactory.create(new File("inventoryBalanceRecord.xlsx"));
		Sheet sheet2 = wb2.getSheetAt(0);
		for (int i = 1; i < sheet2.getPhysicalNumberOfRows()-1; i++) {
			Row row = sheet2.getRow(i);
			int id = Integer.valueOf(row.getCell(0).toString());
			for(YQJ y : yqj) {
				if(y.id == id) {
					y.inventory = (int) row.getCell(8).getNumericCellValue();
				}
			}
		}
		
		//导入塑件生产计划
		Workbook wb3 = WorkbookFactory.create(new File("zsscjh.xlsx"));
		Sheet sheet3 = wb3.getSheetAt(0);
		//System.out.println(sheet3.getRow(49).getCell(2).getNumericCellValue() * 3600 * 24);
		for(int i = 25; i < 48; i += 2) {
			Row row1 = sheet3.getRow(i);
			Row row2 = sheet3.getRow(i+1);
			String type = row1.getCell(1).toString();
			if(type.equals("T308Q_EN3200")) {
				type = "T308Q";
			}
			if(type.equals("T308Q_KM2700")) {
				continue;
			}
			int amount = (int) row1.getCell(2).getNumericCellValue();
			int startTime = 0;
			if(amount != 0) {
				int sequence = (int) row2.getCell(2).getNumericCellValue();
				if(sequence == 1) {
					startTime = -5400;
				} else if(sequence == 2) {
					if(i < 28) {
						startTime = (int) (sheet3.getRow(49).getCell(2).getNumericCellValue() * 24 + 0.75 - 9) * 3600;
					} else if(i < 38) {
						startTime = (int) (sheet3.getRow(50).getCell(2).getNumericCellValue() * 24 + 0.75 - 9) *3600;
					} else {
						startTime = (int) (sheet3.getRow(52).getCell(2).getNumericCellValue() * 24 + 0.75 - 9) *3600;
					}
				} else if(sequence == 3) {
					if(i < 38) {
						startTime = (int) (sheet3.getRow(51).getCell(2).getNumericCellValue() * 24 + 0.75 - 9) * 3600;
					} else {
						startTime = (int) (sheet3.getRow(53).getCell(2).getNumericCellValue() * 24 + 0.75 -9) * 3600;
					}
				}
			}
			ArrayList<YQJ> yqj_temp = new ArrayList<YQJ>();
			for(YQJ y : yqj) {
				if(type.equals(y.type)) {
					yqj_temp.add(y);
				}
			}
			SJ temp = new SJ(type, 0, 0, 30, yqj_temp, startTime, amount);
			//System.out.println(temp.type);
			temp.getYQJProduction();
			for(YQJ y : temp.yqj) {
				y.sj = temp;
			}
			sj.add(temp);
		}
		
		//导出目标产量
		Workbook wb_pt = new XSSFWorkbook();
		Sheet sheet_pt = wb_pt.createSheet("目标产量");
		Row row0_pt = sheet_pt.createRow(0);
		row0_pt.createCell(0).setCellValue("塑件");
		row0_pt.createCell(1).setCellValue("颜色");
		row0_pt.createCell(2).setCellValue("数量");
		int rowNo_pt = 1;
		for(SJ s : sj) {
			for(YQJ y : s.yqj) {
				Row row_pt = sheet_pt.createRow(rowNo_pt);
				row_pt.createCell(0).setCellValue(y.type);
				row_pt.createCell(1).setCellValue(y.color);
				row_pt.createCell(2).setCellValue(y.production);
				rowNo_pt++;
			}
		}
		FileOutputStream fileOut_pt = new FileOutputStream("目标产量.xlsx");
	    wb_pt.write(fileOut_pt);
	    fileOut_pt.close();
		
		
		//导入生产节拍
		Workbook wb4 = WorkbookFactory.create(new File("数据输入.xlsx"));
		Sheet sheet4 = wb4.getSheetAt(0);
		for(int i = 5; i < 16; i++) {
			Row row = sheet4.getRow(i);
			String type = row.getCell(1).toString();
			for(SJ s : sj) {
				if(s.type.equals(type)) {
					System.out.println(row.getCell(11).toString());
					s.rate = (int) row.getCell(20).getNumericCellValue();
					//计算生产结束时间
					s.getEndtime();
				}
			}
		}
		
		//算法循环
		int currentTime = -5400;
		String currentColor = "";
		PaintingEvent event;
		while(currentTime < 3600*24) {
			if(currentTime == -5400) {
				updateSJ(sj, currentTime, 5400);
				currentTime += 5400;
				//outputTheMoment(sj, yqj, currentTime);
			}
			else {
				event = getPaintingEvent(currentTime, sj, yqj, currentColor);
				if(event == null) {
					//System.out.println(currentTime);
					currentTime += updateYQJ(sj, yqj, event, currentTime, currentColor);
					currentColor = "";
				} else {
					//event.time = currentTime;
					currentTime += updateYQJ(sj, yqj, event, currentTime, currentColor);
					plan.add(event);
					currentColor = event.yqj.color;
				}
			}
		}
		
		//输出结果
		Workbook wbOut = new XSSFWorkbook();
		Sheet sheetOut = wbOut.createSheet("喷漆计划");
		Row row0 = sheetOut.createRow(0);
		row0.createCell(0).setCellValue("时间");
		row0.createCell(1).setCellValue("塑件");
		row0.createCell(2).setCellValue("颜色");
		row0.createCell(3).setCellValue("数量");
		
		String color = "";
		int time = 0;
		ArrayList<Lot> lot = new ArrayList<Lot>();
		Row row = null;
		int rowNo = 1;
		for(PaintingEvent e : plan) {
			if(e.yqj.color.equals(color)) {
				boolean flag = true;
				for(Lot l : lot) {
					if(l.type.equals(e.yqj.type)) {
						l.amount += e.amount;
						flag = false;
						break;
					}
				}
				if(flag) {
					lot.add(new Lot(time, e.yqj.type, e.yqj.color, e.amount));
				}
			} else {
				for(Lot l : lot) {
					row = sheetOut.createRow(rowNo);
					row.createCell(0).setCellValue(l.time);
					row.createCell(1).setCellValue(l.type);
					row.createCell(2).setCellValue(l.color);
					row.createCell(3).setCellValue(l.amount);
					rowNo++;
					//System.out.println(l.time + "\t" + l.type + "\t" + l.color + "\t" + l.amount);
				}
				lot = new ArrayList<Lot>();
				lot.add(new Lot(e.time, e.yqj.type, e.yqj.color, e.amount));
				color = e.yqj.color;
				time = e.time;
			}
			//System.out.println(e.time + "\t" + e.yqj.type + "\t" + e.yqj.color + "\t" + e.amount);
		}
		for(Lot l : lot) {
			row = sheetOut.createRow(rowNo);
			row.createCell(0).setCellValue(l.time);
			row.createCell(1).setCellValue(l.type);
			row.createCell(2).setCellValue(l.color);
			row.createCell(3).setCellValue(l.amount);
			rowNo++;
			//System.out.println(l.time + "\t" + l.type + "\t" + l.color + "\t" + l.amount);
		}
		
	    FileOutputStream fileOut = new FileOutputStream("喷漆计划.xlsx");
	    wbOut.write(fileOut);
	    fileOut.close();
		/*
	    for(SJ s : sj) {
	    	int a = 0;
	    	for(YQJ y : s.yqj) {
	    		a += y.production;
	    	}
	    	System.out.println(s.type + "\t" + s.amount + "\t" + a);
	    }
	    */
	}

	private static void outputTheMoment(ArrayList<SJ> sj, ArrayList<YQJ> yqj, int currentTime) {
		System.out.println("当前时间: " + getTime(currentTime));
		System.out.println("塑件"+"\t"+"注塑余量");
		for(SJ s : sj) {
			System.out.println(s.type + "\t" + s.buffer);
		}
		System.out.println("油漆件"+"\t"+"颜色"+"\t"+"库存"+"\t"+"未来三小时需求"+"\t"+"在喷涂量" + "\t" + "剩余目标产量");
		for(YQJ y : yqj) {
			System.out.println(y.type + "\t" + y.color + "\t" + y.inventory + "\t" + y.getDemandInNext3Hours(currentTime)
					+ "\t" + y.online.size() + "\t" + y.production);
		}
	}

	private static String getTime(int currentTime) {
		int h = currentTime / 3600;
		int m = (currentTime - h * 3600) / 60;
		int s = currentTime - h * 3600 - m * 60;
		h += 9;
		if(h >= 24) {
			h -= 24;
		}
		return h + ":" + m + ":" + s;
	}

	private static int updateYQJ(ArrayList<SJ> sj, ArrayList<YQJ> yqj, PaintingEvent event, int currentTime, String currentColor) {		
		int idleTime = 0;
		int passTime;
		//int amount;
		
		if(event == null) {
			passTime = 1800;
		} else {
			if(!event.yqj.color.equals(currentColor)) {
				idleTime = event.yqj.rate * 2;
				hq = 0;
				passTime = event.yqj.rate * ((event.amount)/3) + idleTime;
			} else {
				passTime = event.yqj.rate * ((int)Math.floor((event.amount-3+hq)/3.0)+1);
				//System.out.println(hq);
			}
			event.time = currentTime + idleTime;
			int amount = event.amount;
			//把新一轮油漆件放上去
			for(int t = currentTime+idleTime; t <= currentTime + passTime; t += event.yqj.rate) {
				if(hq != 0) {
					for(int j = hq; j < 3; j++) {
						event.yqj.online.add(new Integer(t+event.yqj.rate*170));
						amount--;
						hq++;
						if(amount == 0) {
							break;
						}
					}
					if(hq == 3) hq = 0;
					continue;
				}
				if(amount >=3 ) {
					for(int i = 0; i < 3; i++) {
						event.yqj.online.add(new Integer(t+event.yqj.rate*170));
					}
					amount -= 3;
				} else {
					for(int i = 0; i < amount; i++) {
						event.yqj.online.add(new Integer(t+event.yqj.rate*170));
						hq++;
					}
					amount = 0;
				}
			}
		}
		
		
		//更新塑件缓存
		updateSJ(sj, currentTime, passTime);
		if(event != null) {
			event.yqj.sj.buffer -= event.amount;
			event.yqj.production -= event.amount;
		}
		//更新油漆件库存
		for(YQJ y : yqj) {
			int h1 = (int) Math.ceil(currentTime/3600);
			int h2 = (int) Math.ceil((currentTime+passTime)/3600);
			for(int h = h1; h < h2; h++) {
				y.inventory -= y.demand[h];
			}
			Iterator<Integer> iter = y.online.iterator();
			while(iter.hasNext()){
			    Integer offTime = iter.next();
			    if(offTime <= currentTime + passTime){
			        iter.remove();
			    }
			}
		}
		return passTime;
	}

	private static PaintingEvent getPaintingEvent(int currentTime,
			ArrayList<SJ> sj, ArrayList<YQJ> yqj, String currentColor) {
		ArrayList<YQJ> yqj_short = new ArrayList<YQJ>();
		//需求限制
		for(YQJ y : yqj) {
			if(y.getShortage(currentTime) > 0) {
				System.out.println("需求限制触发："+"\t"+y.type+"\t"+y.color+"\t"+y.getShortage(currentTime));
				yqj_short.add(y);
			}
		}
		if(yqj_short.isEmpty()) { //如果不触发需求限制
			ArrayList<YQJ> yqj_todo = new ArrayList<YQJ>();
			for(YQJ y : yqj) {
				if(y.production > 0 && y.sj.buffer > 0) {
					yqj_todo.add(y);
				}
			}
			if(yqj_todo.isEmpty()) {
				return null; // 结束
			}
			//如果有相同颜色的需求，就生产相同颜色的油漆件
			for(YQJ y : yqj_todo) {
				if(y.color.equals(currentColor)) {
					//outputTheMoment(sj, yqj, currentTime);
					if(y.production <= y.sj.buffer) {
						PaintingEvent event = new PaintingEvent(y, y.production);
						return event;
					} else {
						PaintingEvent event = new PaintingEvent(y, y.sj.buffer);
						return event;
					}
				}
			}
			//如果没有相同颜色的需求，就找缓存最多的塑件来喷漆
			YQJ temp = yqj_todo.get(0);
			for(YQJ y : yqj_todo) {
				if(temp.sj.buffer < y.sj.buffer) {
					temp = y;
				}
				if(temp.sj.buffer == y.sj.buffer) {
					if(temp.production < y.production) {
						temp = y;
					}
				}
			}
			if(temp.production <= temp.sj.buffer) {
				PaintingEvent event = new PaintingEvent(temp, temp.production);
				//outputTheMoment(sj, yqj, currentTime);
				return event;
			} else {
				PaintingEvent event = new PaintingEvent(temp, temp.sj.buffer);
				return event;
			}
		} else { //如果触发需求限制
			YQJ temp = yqj_short.get(0);
			for(YQJ y : yqj_short) {
				if(y.color.equals(currentColor)) {
					PaintingEvent event = new PaintingEvent(y, y.production);
					return event;
				}
				if(temp.getShortage(currentTime) > y.getShortage(currentTime)) { // 缺口小的先生产
					temp = y;
				}
			}
			PaintingEvent event = new PaintingEvent(temp, temp.production);
			return event;
		}
	}

	private static void updateSJ(ArrayList<SJ> sj, int currentTime, int passTime) {
		for(SJ s : sj) {
			if(currentTime >= s.startTime && currentTime < s.endTime) {
				if(currentTime + passTime < s.endTime) {
					s.buffer += (currentTime+passTime-s.startTime) / s.rate - (currentTime - s.startTime) / s.rate;
				} else {
					s.buffer += (s.endTime - currentTime) / s.rate;
				}
			}
		}
		
	}

	private static String parseColor(String string) {
		return string.substring(string.indexOf('\\', 8)+1);
	}

	private static String parseType(String string) {
		//System.out.println(string.substring(3,string.indexOf('\\')+2));
		ArrayList<String> list = new ArrayList<String>(Arrays.asList("308\\前", "308\\N", "308\\H", 
				"405\\前", "405\\N", "405\\H", "818\\前", "D16\\前", "D16\\后", "D06\\前", "D06\\后"));
		ArrayList<String> sj = new ArrayList<String>(Arrays.asList("T308Q", "T308NB", "T308HB", "T405Q", 
				"T405NB", "T405HB", "T818Q", "D16Q", "D16H", "D06Q", "D06H"));
		return sj.get(list.indexOf(string.substring(3,string.indexOf('\\')+2)));
	}

}