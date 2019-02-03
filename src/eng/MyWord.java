package eng;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.sql.Array;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.LinkedList;
import java.util.List;
import java.util.Random;

import jxl.Workbook;
import jxl.format.PageOrientation;
import jxl.write.Label;
import jxl.write.WritableCellFormat;
import jxl.write.WritableFont;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;

/*
 * 给一个单词
 * 
 * 1.正序输出，扣掉其中1个字母(必须配图)
 * 
 * 2.乱序输出，按正序填写
 * 
 * 
 * */

public class MyWord {

	public static ArrayList<Integer> getRandom(int n) {

		ArrayList<Integer> list = new ArrayList<Integer>();

		Random rand = new Random();

		boolean[] bool = new boolean[n];
		int num = 0;
		for (int i = 0; i < n; i++) {

			do {
				num = rand.nextInt(n);

			} while (bool[num]);

			bool[num] = true;
			list.add(num);

		}
		// System.out.println(list);

		return list;
	}

	public static String get1word(String[] inwords) {
		String outword = null;

		return outword;
	}

	public static void createExcel(OutputStream os, int sheetnum, String[] words) throws WriteException, IOException {
		// 创建工作薄
		WritableWorkbook workbook = Workbook.createWorkbook(os);

		for (int k = 0; k < sheetnum; k++) {

			String sheetname = "sheet" + k;
			// 创建新的一页
			WritableSheet sheet = workbook.createSheet(sheetname, 0);
			sheet.setPageSetup(PageOrientation.LANDSCAPE);
			List<String> getNumber = new ArrayList<String>();

			// 构造表头
			// sheet.mergeCells(0, 0, 29, 0);//
			// 添加合并单元格，第一个参数是起始列，第二个参数是起始行，第三个参数是终止列，第四个参数是终止行
			// WritableFont bold = new WritableFont(WritableFont.ARIAL, 16,
			// WritableFont.BOLD);// 设置字体种类和黑体显示,字体为Arial,字号大小为10,采用黑体显示
			// WritableCellFormat titleFormate = new WritableCellFormat(bold);//
			// 生成一个单元格样式控制对象
			// titleFormate.setAlignment(jxl.format.Alignment.CENTRE);// 单元格中的内容水平方向居中
			// titleFormate.setVerticalAlignment(jxl.format.VerticalAlignment.CENTRE);//
			// 单元格的内容垂直方向居中
			// Label title = new Label(0, 0, "ADD & MINUS in 10", titleFormate);
			// sheet.setRowView(0, 1200, false);// 设置第一行的高度
			// sheet.addCell(title);

			WritableFont cellbold = new WritableFont(WritableFont.ARIAL, 16, WritableFont.NO_BOLD);// 设置字体种类和黑体显示,字体为Arial,字号大小为10,采用黑体显示
			WritableCellFormat cellFormate = new WritableCellFormat(cellbold);// 生成一个单元格样式控制对象
			cellFormate.setAlignment(jxl.format.Alignment.CENTRE);// 单元格中的内容水平方向居中
			cellFormate.setVerticalAlignment(jxl.format.VerticalAlignment.CENTRE);// 单元格的内容垂直方向居中

			// 设置列宽

			sheet.setColumnView(0, 15);
			sheet.setColumnView(1, 15);
			sheet.setColumnView(2, 15);
			sheet.setColumnView(3, 15);
			sheet.setColumnView(4, 15);
			sheet.setColumnView(5, 15);
			sheet.setColumnView(6, 15);
			sheet.setColumnView(7, 15);

			// 行高
			sheet.setRowView(0, 1100, false);// 设置第一行的高度
			sheet.setRowView(1, 1100, false);// 设置第一行的高度
			sheet.setRowView(2, 1100, false);// 设置第一行的高度
			sheet.setRowView(3, 1100, false);// 设置第一行的高度
			sheet.setRowView(4, 1100, false);// 设置第一行的高度
			sheet.setRowView(5, 1100, false);// 设置第一行的高度
			sheet.setRowView(6, 1100, false);// 设置第一行的高度
			sheet.setRowView(7, 1100, false);// 设置第一行的高度


//			System.out.println(sheet.getColumns());
//			System.out.println(sheet.getRows());

			// Collections.shuffle(list); // 打乱字符串
			// System.out.println(list);

			String word = null;

			// 共有9行 第1,3,5,7行填字母 2,4,6,8自己写
			for (int i = 0; i < sheet.getRows();) {

				for (int a = 1; a < 9; ) {

					int randNumber = new Random().nextInt(words.length - 1) + 1;

					System.out.println("get words rannumber:" + randNumber);
					word = words[randNumber];

					int len = word.length();

					char[] cc = word.toCharArray();

					// 得到一组0到字符串长度的随机数
					ArrayList<?> l = getRandom(len);

					List<String> list = new LinkedList<String>();
					for (int j = 0; j < l.size(); j++) {

						list.add(String.valueOf(cc[(int) l.get(j)]));
					}

					Collections.shuffle(list);
					Label label = new Label(a, i, list.toString(), cellFormate);
					sheet.addCell(label);

					a=a+2;
				}

				i = i + 2;
			}

		}

		// 把创建的内容写入到输出流中，并关闭输出流
		workbook.write();
		workbook.close();
		os.close();

	}

	public static void main(String[] args) throws WriteException, IOException {

		String[] words = { "add", "arm", "apple", "bag", "ball", "banana", "bath", "basket", "bed", "big", "bike",
				"bird", "black", "blue", "boat", "body", "book", "box", "boy", "bus", "burger", "can", "cat", "cake",
				"car", "chair", "circle", "crocdile", "cream", "dad", "day", "dance", "dirty", "dog", "doll", "dot",
				"duck", "ear", "eat", "eraser", "eye", "face", "father", "fish", "fly", "food", "foot", "football",
				"frog", "game", "giraffe", "girl", "good", "grape", "green", "grey", "guitar", "hair", "hall", "hand",
				"head", "hippo", "home", "horse", "house", "hot", "ice", "igloo", "iguana", "jacket", "jellyfish",
				"kitchen", "knee", "leg", "lemon", "like", "listen", "long", "look", "lorry", "love", "make", "man",
				"maths", "monkey", "mom", "mother", "monster", "mouse", "mouth", "music", "name", "no", "nose",
				"number", "old", "open", "orange", "pear", "pen", "pencil", "pet", "piano", "picture", "pink", "plane",
				"purple", "rainbow", "red", "ride", "river", "room", "sad", "say", "school", "see", "shoe", "short",
				"sing", "sit", "skirt", "small", "snake", "smile", "sock", "sofa", "stand", "star", "swim", "stop",
				"street", "table", "tail", "tennis", "tick", "tiger", "toe", "toy", "train", "tooth", "tree", "water",
				"watermelon", "white", "woman", "yellow", "zebra", "zoo" };

		OutputStream out = null;
		try {
			out = new FileOutputStream("C:/eng.xls");

		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

		// String word = "apple";

		createExcel(out, 100, words);

	}

}

/*
 * 
 * alphabet 单词表
 * 
 */

// String[] wordsA = { "add", "arm", "apple" };
// String[] wordsB = { "bag", "ball", "banana", "bath", "basket", "bed", "big",
// "bike", "bird", "black", "blue",
// "boat", "body", "book", "box", "boy", "bus", "burger" };
//// "child", "correct", "computer", "clean", "close", "colour",
// String[] wordsC = { "can", "cat", "cake", "car", "chair", "circle",
// "crocdile", "cream" };
//// "draw", "drive",
// String[] wordsD = { "dad", "day", "dance", "dirty", "dog", "doll", "dot",
// "duck" };
// String[] wordsE = { "ear", "eat", "eraser", "eye" };
//// "family", "finger","fruit"
// String[] wordsF = { "face", "father", "fish", "fly", "food", "foot",
// "football", "frog" };
//// "grandfather", "grandmother", "grandma", "grandpa",
// String[] wordsG = { "game", "giraffe", "girl", "good", "grape", "green",
// "grey", "guitar" };
//// "helicopter",
// String[] wordsH = { "hair", "hall", "hand", "head", "hippo", "home", "horse",
// "house", "hot" };
// String[] wordsI = { "ice", "igloo", "iguana" };
// String[] wordsJ = { "jacket", "jellyfish" };
// String[] wordsK = { "kitchen", "knee" };
// String[] wordsL = { "leg", "lemon", "like", "listen", "long", "look",
// "lorry", "love" };
// String[] wordsM = { "make", "man", "maths", "monkey", "mom", "mother",
// "monster", "mouse", "mouth", "music" };
// String[] wordsN = { "name", "no", "nose", "number" };
// String[] wordsO = { "old", "open", "orange" };
//// "painting",
// String[] wordsP = { "pear", "pen", "pencil", "pet", "piano", "picture",
// "pink", "plane", "purple" };
//// String[] wordsQ = {};
// String[] wordsR = { "rainbow", "red", "ride", "river", "room" };
// String[] wordsS = { "sad", "say", "school", "see", "shoe", "short", "sing",
// "sit", "skirt", "small", "snake",
// "smile", "sock", "sofa", "stand", "star", "swim", "stop", "street" };
//// "television",
// String[] wordsT = { "table", "tail", "tennis", "tick", "tiger", "toe", "toy",
// "train", "tooth", "tree" };
//// String[] wordsU = {"umbrella"};
//// String[] wordsV = {};
// String[] wordsW = { "water", "watermelon", "white", "woman" };
//// String[] wordsX = {};
// String[] wordsY = { "yellow" };
// String[] wordsZ = { "zebra" ,"zoo"};