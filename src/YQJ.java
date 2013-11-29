import java.util.ArrayList;

//ÓÍÆá¼þ

public class YQJ {
	int id;
	String type;
	String color;
	int[] demand;
	int inventory;
	ArrayList<Integer> online;
	int production;
	int totalDemand;
	SJ sj;
	int rate;
	
	public YQJ(int id, String type, String color, int[] demand, int inventory) {
		this.id = id;
		this.type = type;
		this.color = color;
		this.demand = demand;
		this.inventory = inventory;
		this.online = new ArrayList<Integer>();
		getTotalDemand();
		rate = 56;
	}
	
	public int getShortage(int currentTime) {
		return getDemandInNext3Hours(currentTime) - inventory - online.size();
	}
	
	public int getDemandInNext3Hours(int currentTime) {
		int t = (int) (currentTime/3600);
		int demandInNext3Hours = 0;
		for(int i = 0; i < 3; i++) {
			if(t+i >= 24) break;
			demandInNext3Hours += demand[t+i];
		}
		return demandInNext3Hours;
	}
	
	public int getTotalDemand() {
		totalDemand = 0;
		for(int i = 0; i < 24; i++) {
			totalDemand += demand[i];
		}
		return totalDemand;
	}
}
