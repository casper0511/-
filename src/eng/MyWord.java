package eng;

import java.io.File;
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
import jxl.write.WritableImage;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;

/*
 * ��һ������
 * 
 * 1.����������۵�����1����ĸ(������ͼ)
 * 
 * 2.�����������������д
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
		// ����������
		WritableWorkbook workbook = Workbook.createWorkbook(os);

		for (int k = 0; k < sheetnum; k++) {

			String sheetname = "sheet" + k;
			// �����µ�һҳ
			WritableSheet sheet = workbook.createSheet(sheetname, 0);
			sheet.setPageSetup(PageOrientation.LANDSCAPE);
			// List<String> getNumber = new ArrayList<String>();

			// �����ͷ
			// sheet.mergeCells(0, 0, 29, 0);//
			// ��Ӻϲ���Ԫ�񣬵�һ����������ʼ�У��ڶ�����������ʼ�У���������������ֹ�У����ĸ���������ֹ��
			// WritableFont bold = new WritableFont(WritableFont.ARIAL, 16,
			// WritableFont.BOLD);// ������������ͺ�����ʾ,����ΪArial,�ֺŴ�СΪ10,���ú�����ʾ
			// WritableCellFormat titleFormate = new WritableCellFormat(bold);//
			// ����һ����Ԫ����ʽ���ƶ���
			// titleFormate.setAlignment(jxl.format.Alignment.CENTRE);// ��Ԫ���е�����ˮƽ�������
			// titleFormate.setVerticalAlignment(jxl.format.VerticalAlignment.CENTRE);//
			// ��Ԫ������ݴ�ֱ�������
			// Label title = new Label(0, 0, "ADD & MINUS in 10", titleFormate);
			// sheet.setRowView(0, 1200, false);// ���õ�һ�еĸ߶�
			// sheet.addCell(title);

			WritableFont cellbold = new WritableFont(WritableFont.ARIAL, 10, WritableFont.NO_BOLD);// ������������ͺ�����ʾ,����ΪArial,�ֺŴ�СΪ10,���ú�����ʾ
			WritableCellFormat cellFormate = new WritableCellFormat(cellbold);// ����һ����Ԫ����ʽ���ƶ���
			cellFormate.setAlignment(jxl.format.Alignment.CENTRE);// ��Ԫ���е�����ˮƽ�������
			cellFormate.setVerticalAlignment(jxl.format.VerticalAlignment.CENTRE);// ��Ԫ������ݴ�ֱ�������

			// �����п�
			// 4��ͼ 4�е��� �Ų��£�����3��ͼ 3�е��� 
//			sheet.setColumnView(0, 15);
//			sheet.setColumnView(1, 15);
//			sheet.setColumnView(2, 15);
//			sheet.setColumnView(3, 15);
//			sheet.setColumnView(4, 15);
//			sheet.setColumnView(5, 15);
//			sheet.setColumnView(6, 15);
//			sheet.setColumnView(7, 15);

			sheet.setColumnView(0, 20);
			sheet.setColumnView(1, 20);
			sheet.setColumnView(2, 20);
			sheet.setColumnView(3, 20);
			sheet.setColumnView(4, 20);
			sheet.setColumnView(5, 20);

			
			// �и�
			sheet.setRowView(0, 1100, false);// ���õ�һ�еĸ߶�
			sheet.setRowView(1, 1100, false);// ���õ�һ�еĸ߶�
			sheet.setRowView(2, 1100, false);// ���õ�һ�еĸ߶�
			sheet.setRowView(3, 1100, false);// ���õ�һ�еĸ߶�
			sheet.setRowView(4, 1100, false);// ���õ�һ�еĸ߶�
			sheet.setRowView(5, 1100, false);// ���õ�һ�еĸ߶�
			sheet.setRowView(6, 1100, false);// ���õ�һ�еĸ߶�
			sheet.setRowView(7, 1100, false);// ���õ�һ�еĸ߶�

			// System.out.println(sheet.getColumns());
			// System.out.println(sheet.getRows());

			// Collections.shuffle(list); // �����ַ���
			// System.out.println(list);

			String word = null;

			// ����9�� ��1,3,5,7������ĸ 2,4,6,8�Լ�д
			for (int i = 0; i < sheet.getRows();) {

				//a���еı��� 4��a<9 3��a<7
				for (int a = 1; a < 7;) {

					// �ӵ��ʱ������ȡһ������
					int randNumber = new Random().nextInt(words.length - 1) + 1;
					// System.out.println("get words rannumber:" + randNumber);
					word = words[randNumber];

					System.out.println("get file path:" + System.getProperty("user.dir")+"/pic/" + word + ".png");
					
//					String a1=System.getProperty("user.dir");
//					System.out.println("user.dir:"+a1);
					
					File imgFile = new File(System.getProperty("user.dir")+"/pic/" + word + ".png");

					int len = word.length();

					char[] cc = word.toCharArray();

					// �õ�һ��0���ַ������ȵ������
					ArrayList<?> l = getRandom(len);

					List<String> list = new LinkedList<String>();
					for (int j = 0; j < l.size(); j++) {

						list.add(String.valueOf(cc[(int) l.get(j)]));
					}

					// ������ĸ˳��
					Collections.shuffle(list);
					// a���к� i���к�
					// ����ͼƬ
					WritableImage image = new WritableImage(a - 1, i, 1, 2, imgFile);
					sheet.addImage(image);

					// ���뵥��
					Label label = new Label(a, i, list.toString(), cellFormate);
					sheet.addCell(label);

					a = a + 2;
				}

				i = i + 2;
			}

		}

		// �Ѵ���������д�뵽������У����ر������
		workbook.write();
		workbook.close();
		os.close();

	}

	public static void InitWordlist() {
		// ��ʼ�����ʱ�

		ArrayList<ArrayList<Wordlist>> biglist = new ArrayList<ArrayList<Wordlist>>();
		ArrayList<Wordlist> list = new ArrayList<Wordlist>();

		Wordlist word = new Wordlist("add", new File(""));

	}

	public static void main(String[] args) throws WriteException, IOException {

		InitWordlist();

		String[] words = { "arm", "apple", "bag", "ball", "banana", "bath", "basket", "bed", "big", "bike", "bird",
				"black", "blue", "boat", "body", "book", "box", "boy", "bus", "burger", "can", "cat", "cake", "car",
				"chair", "circle", "crocdile", "dad", "dance", "dog", "doll", "dot", "duck", "ear", "eat", "eraser",
				"eye", "face", "father", "fish", "fly", "food", "foot", "football", "frog", "giraffe", "girl", "good",
				"grape", "green", "grey", "guitar", "hair", "hand", "head", "hippo", "horse", "house", "hot", "ice",
				"igloo", "iguana", "jacket", "jellyfish", "kitchen", "knee", "leg", "lemon", "listen", "look", "lorry",
				"love", "man", "maths", "monkey", "mom", "monster", "mouse", "mouth", "music", "nose", "number", "old",
				"orange", "pear", "pen", "pencil", "piano", "pink", "plane", "purple", "rainbow", "red", "river",
				"room", "sad", "say", "school", "see", "shoe", "short", "sing", "sit", "skirt", "small", "snake",
				"smile", "sock", "sofa", "stand", "star", "swim", "stop", "street", "table", "tall", "tennis", "tick",
				"tiger", "toe", "toy", "train", "tooth", "tree", "water", "watermelon", "white", "woman", "yellow",
				"zebra", "zoo" };

		OutputStream out = null;
		try {
			out = new FileOutputStream("C:/eng.xls");

		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

		// String word = "apple";

		createExcel(out, 10, words);

	}

}

/*
 * 
 * alphabet ���ʱ�
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