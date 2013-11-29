
public class Lot {
	String time;
	String type;
	String color;
	int amount;
	
	public Lot(int time, String type, String color, int amount) {
		int h = time / 3600;
		int m = (time - h * 3600) / 60;
		int s = time - h * 3600 - m * 60;
		h += 9;
		if(h >= 24) {
			h -= 24;
		}
		this.time = h + ":" + m + ":" + s;
		this.type = type;
		this.color = color;
		this.amount = amount;
	}
}
