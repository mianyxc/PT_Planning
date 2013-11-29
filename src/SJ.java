import java.util.ArrayList;

//塑件

public class SJ {
	//int id;
	String type;
	int rate;
	int buffer;
	int inventory;
	ArrayList<YQJ> yqj;
	int startTime;
	int endTime;
	int amount;
	
	public SJ(String type, int rate, int buffer, int inventory, ArrayList<YQJ> yqj, int startTime, int amount) {
		//this.id = id;
		this.type = type;
		this.rate = rate;
		this.buffer = buffer;
		this.inventory = inventory;
		this.yqj = yqj;
		this.startTime = startTime;
		this.amount = amount;
	}
	
	public int getEndtime() {
		endTime = startTime + amount * rate;
		return endTime;
	}
	
	//分配各种颜色的产量
	public void getYQJProduction() {
		int totalDemand = 0;
		int totalInventory = 0;
		for(YQJ y : yqj) {
			totalDemand += y.totalDemand;
			totalInventory += y.inventory;
		}
		YQJ temp = yqj.get(0);
		int totalProduction = 0;
		if(totalDemand != 0) {
			for(YQJ y : yqj) {
				y.production = amount * y.totalDemand / totalDemand;
				totalProduction += y.production;
				if(temp.totalDemand < y.totalDemand) {
					temp = y;
				}
			}
			temp.production += amount - totalProduction; // 把零头加到需求量最大的油漆件上。
		} else {
			for(YQJ y : yqj) {
				y.production = amount * y.inventory / totalInventory;
				totalProduction += y.production;
				if(temp.inventory < y.inventory) {
					temp = y;
				}
			}
			temp.production += amount - totalProduction; // 把零头加到库存量最大的油漆件上。
		}
	}
}
